/* ═══════════════════════════════════════════════════════════
   Azure File Attacher – Outlook Add-in Task Pane Logic
   ═══════════════════════════════════════════════════════════

   Two modes:
   1. DROP ZONE  – User drags files from Salesforce / desktop into this pane,
                   we read them as base64 and attach via Office.js.
   2. BROWSE     – Connects to Salesforce Apex (AzureFileStorageController)
                   to list and fetch Azure files, then attaches via Office.js.
   ═══════════════════════════════════════════════════════════ */

// ─── Globals ──────────────────────────────────────────────
let officeReady = false;
let azureCurrentPath = '';
let azureAllFiles = [];      // raw listing from Apex for current dir
let azureFilteredFiles = []; // after client-side search filter

// Salesforce connection config (persisted in localStorage)
let config = {
    sfUrl: '',
    sfToken: ''
};

// ─── Office.js Initialization ─────────────────────────────
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        officeReady = true;
        setStatus('Connected to Outlook – ready.');
        initDropZone();
        loadConfig();
        registerOfficeDragDrop();
    } else {
        setStatus('Error: This add-in only works in Outlook.');
    }
});


/* ════════════════════════════════════════════════════════════
   TAB SWITCHING
   ════════════════════════════════════════════════════════════ */
function switchTab(tab) {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.pane').forEach(p => p.classList.remove('active'));

    document.getElementById(tab === 'drop' ? 'tabDrop' : 'tabBrowse').classList.add('active');
    document.getElementById(tab === 'drop' ? 'paneDrop' : 'paneBrowse').classList.add('active');

    if (tab === 'browse' && config.sfUrl && config.sfToken && azureAllFiles.length === 0) {
        azureNavigate('');
    }
}


/* ════════════════════════════════════════════════════════════
   DROP ZONE – HTML5 Drag & Drop
   ════════════════════════════════════════════════════════════ */
function initDropZone() {
    const zone = document.getElementById('dropZone');

    zone.addEventListener('dragenter', (e) => {
        e.preventDefault();
        e.stopPropagation();
        zone.classList.add('drag-over');
    });

    zone.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.stopPropagation();
        e.dataTransfer.dropEffect = 'copy';
        zone.classList.add('drag-over');
    });

    zone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        e.stopPropagation();
        // Only remove if we're actually leaving the zone
        if (!zone.contains(e.relatedTarget)) {
            zone.classList.remove('drag-over');
        }
    });

    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        e.stopPropagation();
        zone.classList.remove('drag-over');

        const files = e.dataTransfer.files;
        if (files && files.length > 0) {
            handleDroppedFiles(files);
        } else {
            setStatus('No files detected in drop.');
        }
    });

    // Click to open file picker
    zone.addEventListener('click', () => {
        document.getElementById('filePicker').click();
    });
}

/**
 * Register the Office.js DragAndDropEvent handler for OWA / new Outlook.
 * This is a separate API from the HTML5 DragEvent used in classic Outlook.
 */
function registerOfficeDragDrop() {
    try {
        if (Office.context.mailbox && Office.context.mailbox.addHandlerAsync) {
            Office.context.mailbox.addHandlerAsync(
                Office.EventType.DragAndDropEvent,
                (event) => {
                    const eventData = event.dragAndDropEventData;
                    if (eventData && eventData.type === 'drop') {
                        const files = eventData.dataTransfer.files;
                        if (files && files.length > 0) {
                            handleDroppedOfficeFiles(files);
                        }
                    }
                },
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        console.warn('Office DragDrop handler not supported:', result.error.message);
                    } else {
                        console.log('Office DragDrop handler registered.');
                    }
                }
            );
        }
    } catch (err) {
        console.warn('Office DragDrop API not available:', err);
    }
}

/**
 * Handle files dropped via the HTML5 drop event (classic Outlook / desktop files).
 */
function handleDroppedFiles(fileList) {
    const count = fileList.length;
    setStatus(`Processing ${count} file(s)…`);

    for (let i = 0; i < fileList.length; i++) {
        const file = fileList[i];
        readAndAttach(file, file.name, file.type || 'application/octet-stream');
    }
}

/**
 * Handle files dropped via Office.js DragAndDropEvent (OWA / new Outlook).
 * Each file has .name, .type, and .fileContent (a Blob).
 */
async function handleDroppedOfficeFiles(files) {
    setStatus(`Processing ${files.length} file(s) from Outlook…`);

    for (const file of files) {
        try {
            // fileContent is a Blob in the Office.js DragDrop API
            const blob = file.fileContent;
            const name = file.name;
            const type = file.type || 'application/octet-stream';

            const base64 = await blobToBase64(blob);
            attachToEmail(base64, name, type);
        } catch (err) {
            console.error('Error processing Office drop file:', err);
            addAttachmentRow(file.name, file.type, 'error', err.message);
        }
    }
}

/**
 * Handle file picker selection.
 */
function handleFilePick(event) {
    const files = event.target.files;
    if (files && files.length > 0) {
        handleDroppedFiles(files);
    }
    // Reset so the same file can be picked again
    event.target.value = '';
}

/**
 * Read a File object as base64 and attach it to the current email.
 */
function readAndAttach(file, fileName, mimeType) {
    const rowId = addAttachmentRow(fileName, mimeType, 'queued');

    const reader = new FileReader();

    reader.onload = function () {
        // result is "data:<mime>;base64,<data>" — strip the prefix
        const base64 = reader.result.split(',')[1];
        updateAttachmentStatus(rowId, 'attaching');
        attachToEmail(base64, fileName, mimeType, rowId);
    };

    reader.onerror = function () {
        console.error('FileReader error for', fileName);
        updateAttachmentStatus(rowId, 'error', 'Failed to read file');
        setStatus(`Error reading ${fileName}`);
    };

    reader.readAsDataURL(file);
}


/* ════════════════════════════════════════════════════════════
   ATTACH TO EMAIL — Office.js
   ════════════════════════════════════════════════════════════ */

/**
 * Attach a base64-encoded file to the current compose item.
 *
 * @param {string} base64     – raw base64 string (no data-url prefix)
 * @param {string} fileName   – display name including extension
 * @param {string} mimeType   – MIME type
 * @param {string} [rowId]    – optional UI row ID to update status
 */
function attachToEmail(base64, fileName, mimeType, rowId) {
    if (!officeReady) {
        const msg = 'Office.js not ready – cannot attach.';
        console.error(msg);
        if (rowId) updateAttachmentStatus(rowId, 'error', msg);
        setStatus(msg);
        return;
    }

    if (!rowId) {
        rowId = addAttachmentRow(fileName, mimeType, 'attaching');
    }

    setStatus(`Attaching ${fileName}…`);

    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
        base64,
        fileName,
        { isInline: false },
        (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log('Attached:', fileName, '| ID:', result.value);
                updateAttachmentStatus(rowId, 'attached');
                setStatus(`✓ ${fileName} attached`);
            } else {
                console.error('Attach failed:', result.error.message);
                updateAttachmentStatus(rowId, 'error', result.error.message);
                setStatus(`✗ Failed to attach ${fileName}`);
            }
        }
    );
}


/* ════════════════════════════════════════════════════════════
   BROWSE AZURE – Calls Salesforce Apex via REST
   ════════════════════════════════════════════════════════════ */

/**
 * Navigate to a directory in Azure storage via the Salesforce Apex endpoint.
 */
async function azureNavigate(path) {
    azureCurrentPath = path || '';
    updateBreadcrumb(azureCurrentPath);

    const listEl = document.getElementById('azureFileList');
    const loadEl = document.getElementById('azureLoading');

    if (!config.sfUrl || !config.sfToken) {
        listEl.innerHTML = `
            <div class="empty-state">
                <i class="ms-Icon ms-Icon--PlugDisconnected" aria-hidden="true"></i>
                <p>Please configure your Salesforce connection in Settings above.</p>
            </div>`;
        return;
    }

    loadEl.style.display = 'flex';
    setStatus('Loading directory…');

    try {
        const result = await callApex('browseDirectory', { directoryPath: azureCurrentPath });
        azureAllFiles = result || [];
        azureFilteredFiles = [...azureAllFiles];
        renderAzureFiles();
        setStatus(`Loaded ${azureAllFiles.length} items`);
    } catch (err) {
        console.error('Azure browse error:', err);
        listEl.innerHTML = `
            <div class="empty-state">
                <i class="ms-Icon ms-Icon--ErrorBadge" aria-hidden="true"></i>
                <p>Error: ${escapeHtml(err.message || String(err))}</p>
            </div>`;
        setStatus('Error loading files');
    } finally {
        loadEl.style.display = 'none';
    }
}

/**
 * Filter the current directory listing client-side.
 */
function filterAzureFiles() {
    const term = (document.getElementById('azureSearch').value || '').toLowerCase();
    if (!term) {
        azureFilteredFiles = [...azureAllFiles];
    } else {
        azureFilteredFiles = azureAllFiles.filter(f =>
            f.name && f.name.toLowerCase().includes(term)
        );
    }
    renderAzureFiles();
}

/**
 * Render the file/folder list in the Browse tab.
 */
function renderAzureFiles() {
    const listEl = document.getElementById('azureFileList');

    if (azureFilteredFiles.length === 0) {
        listEl.innerHTML = `
            <div class="empty-state">
                <i class="ms-Icon ms-Icon--FabricFolderSearch" aria-hidden="true"></i>
                <p>No files found</p>
            </div>`;
        return;
    }

    // Sort: folders first, then files alphabetically
    const sorted = [...azureFilteredFiles].sort((a, b) => {
        if (a.isDirectory && !b.isDirectory) return -1;
        if (!a.isDirectory && b.isDirectory) return 1;
        return (a.name || '').localeCompare(b.name || '');
    });

    let html = '';
    for (const file of sorted) {
        if (file.isDirectory) {
            html += `
                <div class="azure-file-item" onclick="azureNavigate('${escapeAttr(file.path)}')">
                    <i class="ms-Icon ms-Icon--FabricFolder icon-folder" aria-hidden="true"></i>
                    <div class="azure-file-info">
                        <span class="azure-file-name">${escapeHtml(file.name)}</span>
                        <span class="azure-file-meta">Folder</span>
                    </div>
                </div>`;
        } else {
            const size = formatFileSize(file.size);
            html += `
                <div class="azure-file-item">
                    <i class="ms-Icon ms-Icon--Page icon-file" aria-hidden="true"></i>
                    <div class="azure-file-info">
                        <span class="azure-file-name">${escapeHtml(file.name)}</span>
                        <span class="azure-file-meta">${size}</span>
                    </div>
                    <button class="azure-attach-btn"
                            onclick="azureAttachFile('${escapeAttr(file.path)}', '${escapeAttr(file.name)}'); event.stopPropagation();">
                        <i class="ms-Icon ms-Icon--Attach" aria-hidden="true"></i> Attach
                    </button>
                </div>`;
        }
    }

    listEl.innerHTML = html;
}

/**
 * Fetch a single file from Azure (via Apex) and attach it to the current email.
 */
async function azureAttachFile(filePath, fileName) {
    const rowId = addAttachmentRow(fileName, '', 'attaching');
    setStatus(`Fetching ${fileName} from Azure…`);

    try {
        const result = await callApex('getFileContent', { filePath: filePath });
        const base64 = result.base64Content;
        const mimeType = result.contentType || 'application/octet-stream';

        attachToEmail(base64, fileName, mimeType, rowId);
    } catch (err) {
        console.error('Azure fetch error:', err);
        updateAttachmentStatus(rowId, 'error', err.message || 'Fetch failed');
        setStatus(`Error fetching ${fileName}`);
    }
}


/* ════════════════════════════════════════════════════════════
   SALESFORCE APEX CALLOUT (REST API)
   ════════════════════════════════════════════════════════════ */

/**
 * Call a Salesforce Apex @AuraEnabled method via the standard Aura/REST endpoint.
 *
 * This uses the Salesforce REST API pattern:
 *   POST /services/apexrest/AzureFileStorage/<methodName>
 *
 * OR you can use the Aura endpoint if the Apex class is @AuraEnabled:
 *   POST /aura?r=1  (with appropriate message payload)
 *
 * For simplicity, we'll assume you expose a lightweight REST endpoint.
 * Adjust the URL pattern to match your org's setup.
 *
 * @param {string} method – Apex method name (e.g., 'browseDirectory')
 * @param {object} params – method parameters
 * @returns {Promise<any>}
 */
async function callApex(method, params) {
    // ──────────────────────────────────────────────────────────
    // OPTION A: Custom Apex REST endpoint
    // You'd create an @RestResource class in Salesforce like:
    //
    //   @RestResource(urlMapping='/AzureFileStorage/*')
    //   global class AzureFileStorageRest { ... }
    //
    // Then call it here:
    // ──────────────────────────────────────────────────────────
    const url = `${config.sfUrl}/services/apexrest/AzureFileStorage/${method}`;

    const response = await fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${config.sfToken}`
        },
        body: JSON.stringify(params)
    });

    if (!response.ok) {
        const errText = await response.text();
        throw new Error(`Salesforce API error (${response.status}): ${errText}`);
    }

    return await response.json();
}


/* ════════════════════════════════════════════════════════════
   CONFIG PERSISTENCE
   ════════════════════════════════════════════════════════════ */

function saveConfig() {
    config.sfUrl   = (document.getElementById('cfgSfUrl').value || '').replace(/\/+$/, '');
    config.sfToken = document.getElementById('cfgSfToken').value || '';

    try {
        localStorage.setItem('azureFileAttacher_config', JSON.stringify(config));
    } catch (e) {
        // localStorage might be blocked in some Outlook contexts
        console.warn('Could not persist config:', e);
    }

    setStatus('Configuration saved.');

    // Reload if we have both values
    if (config.sfUrl && config.sfToken) {
        azureNavigate('');
    }
}

function loadConfig() {
    try {
        const saved = localStorage.getItem('azureFileAttacher_config');
        if (saved) {
            config = JSON.parse(saved);
            document.getElementById('cfgSfUrl').value = config.sfUrl || '';
            document.getElementById('cfgSfToken').value = config.sfToken || '';
        }
    } catch (e) {
        console.warn('Could not load config:', e);
    }
}


/* ════════════════════════════════════════════════════════════
   UI HELPERS
   ════════════════════════════════════════════════════════════ */

let attachmentCounter = 0;

/**
 * Add a row to the attachment list UI and return its ID.
 */
function addAttachmentRow(fileName, mimeType, status, detail) {
    const id = 'att-' + (++attachmentCounter);
    const listEl = document.getElementById('attachmentList');
    const iconClass = getFileIconClass(fileName);

    const row = document.createElement('div');
    row.className = 'attachment-item';
    row.id = id;
    row.innerHTML = `
        <i class="ms-Icon ${iconClass}" aria-hidden="true"></i>
        <div class="attachment-info">
            <span class="attachment-name" title="${escapeAttr(fileName)}">${escapeHtml(fileName)}</span>
            <span class="attachment-meta">${escapeHtml(mimeType || '')}</span>
        </div>
        <span class="attachment-status status-${status}">${statusLabel(status, detail)}</span>
    `;

    listEl.prepend(row);
    return id;
}

/**
 * Update an existing attachment row's status.
 */
function updateAttachmentStatus(rowId, status, detail) {
    const row = document.getElementById(rowId);
    if (!row) return;

    const badge = row.querySelector('.attachment-status');
    if (badge) {
        badge.className = `attachment-status status-${status}`;
        badge.textContent = statusLabel(status, detail);
    }
}

function statusLabel(status, detail) {
    switch (status) {
        case 'queued':    return 'Queued';
        case 'attaching': return 'Attaching…';
        case 'attached':  return '✓ Attached';
        case 'error':     return '✗ ' + (detail || 'Error');
        default:          return status;
    }
}

/**
 * Update breadcrumb UI for Azure browsing.
 */
function updateBreadcrumb(path) {
    const el = document.getElementById('azureBreadcrumb');
    let html = '<a href="#" onclick="azureNavigate(\'\'); return false;">Root</a>';

    if (path) {
        const parts = path.split('/');
        let cumulative = '';
        for (let i = 0; i < parts.length; i++) {
            if (!parts[i]) continue;
            cumulative = cumulative ? cumulative + '/' + parts[i] : parts[i];
            html += '<span class="sep">›</span>';
            if (i === parts.length - 1) {
                html += `<span class="current">${escapeHtml(parts[i])}</span>`;
            } else {
                html += `<a href="#" onclick="azureNavigate('${escapeAttr(cumulative)}'); return false;">${escapeHtml(parts[i])}</a>`;
            }
        }
    }

    el.innerHTML = html;
}

function setStatus(msg) {
    const el = document.getElementById('statusText');
    if (el) el.textContent = msg;
}

function getFileIconClass(fileName) {
    if (!fileName) return 'ms-Icon--Page';
    const ext = fileName.split('.').pop().toLowerCase();
    const map = {
        pdf:  'ms-Icon--PDF',
        doc:  'ms-Icon--WordDocument',  docx: 'ms-Icon--WordDocument',
        xls:  'ms-Icon--ExcelDocument', xlsx: 'ms-Icon--ExcelDocument',
        ppt:  'ms-Icon--PowerPointDocument', pptx: 'ms-Icon--PowerPointDocument',
        jpg:  'ms-Icon--FileImage', jpeg: 'ms-Icon--FileImage',
        png:  'ms-Icon--FileImage', gif:  'ms-Icon--FileImage',
        svg:  'ms-Icon--FileImage', webp: 'ms-Icon--FileImage',
        zip:  'ms-Icon--ZipFolder', rar:  'ms-Icon--ZipFolder',
        txt:  'ms-Icon--TextDocument',
        csv:  'ms-Icon--ExcelDocument',
        html: 'ms-Icon--FileHTML',
        xml:  'ms-Icon--FileCode',
        mp4:  'ms-Icon--Video', mov: 'ms-Icon--Video',
        mp3:  'ms-Icon--MusicInCollection', wav: 'ms-Icon--MusicInCollection'
    };
    return map[ext] || 'ms-Icon--Page';
}

function formatFileSize(bytes) {
    if (!bytes || bytes === 0) return '0 B';
    const units = ['B', 'KB', 'MB', 'GB'];
    let size = bytes;
    let idx = 0;
    while (size >= 1024 && idx < units.length - 1) { size /= 1024; idx++; }
    return size.toFixed(1) + ' ' + units[idx];
}


/* ════════════════════════════════════════════════════════════
   UTILITY
   ════════════════════════════════════════════════════════════ */

function blobToBase64(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result.split(',')[1]);
        reader.onerror = () => reject(new Error('Blob read failed'));
        reader.readAsDataURL(blob);
    });
}

function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str || '';
    return div.innerHTML;
}

function escapeAttr(str) {
    return (str || '').replace(/'/g, "\\'").replace(/"/g, '&quot;');
}
