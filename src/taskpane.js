/* ═══════════════════════════════════════════════════════════
   Azure File Attacher – Outlook Add-in Task Pane Logic
   ═══════════════════════════════════════════════════════════

   Drop zone – User drags files from desktop into this pane,
               we read them as base64 and attach via Office.js.
   ═══════════════════════════════════════════════════════════ */

// ─── Globals ──────────────────────────────────────────────
let officeReady = false;

// ─── Office.js Initialization ─────────────────────────────
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        officeReady = true;
        setStatus('Connected to Outlook – ready.');
        initDropZone();
    } else {
        setStatus('Error: This add-in only works in Outlook.');
    }
});


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

    // Paste support (Ctrl+V) — works in OWA where drag/drop is blocked by iframe
    document.addEventListener('paste', (e) => {
        const files = e.clipboardData && e.clipboardData.files;
        if (files && files.length > 0) {
            e.preventDefault();
            handleDroppedFiles(files);
        }
    });
}

/**
 * Handle files dropped via the HTML5 drop event.
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


/* ════════════════════════════════════════════════════════════
   UTILITY
   ════════════════════════════════════════════════════════════ */

function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str || '';
    return div.innerHTML;
}

function escapeAttr(str) {
    return (str || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
