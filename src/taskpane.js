/* ===========================================================
   Azure File Attacher - Outlook Add-in Task Pane Logic
   ===========================================================

   ES5-SAFE BUILD. This file must parse and run in the IE11 /
   Trident webview used by classic Outlook 2016/2019 (perpetual,
   requirement set 1.6). That means:

     - NO arrow functions
     - NO template literals
     - NO async / await
     - NO fetch()   -> use XMLHttpRequest
     - NO new URL() -> parse manually
     - NO Element.prepend() -> use insertBefore

   ATTACH STRATEGY
   ---------------
   Primary path is URL-based attach:
       Office.context.mailbox.item.addFileAttachmentAsync(url, name, ...)
   This requires only requirement set 1.1, so it works on EVERY
   Outlook version including 2019. Outlook/Exchange downloads the
   file from the URL itself. For Exchange Online the fetch happens
   server-side, so the URL MUST be reachable without the user's
   session -> i.e. an Azure SAS URL (?sv=...&sig=...).

   The base64 path (addFileAttachmentFromBase64Async) requires
   requirement set 1.8 and DOES NOT EXIST on Outlook 2019. We
   feature-detect it: on clients that lack it we surface a clear
   message instead of failing silently.
   =========================================================== */

/* --- Globals -------------------------------------------------- */
var officeReady = false;
var canAttachBase64 = false; // set true only if req set 1.8 API exists

// Build version — bump on every published change to the hosted files.
// Shown in the task pane footer so you can confirm at a glance which build
// actually loaded (Office caches add-in web content). If the footer shows this
// number, the new taskpane.js is live; an older number or nothing = cached/old.
// (This is the JS build version; the manifest <Version> only changes when the
// manifest itself does and the admin re-uploads it.)
var ADDIN_VERSION = '1.0.4';

/* --- Office.js Initialization -------------------------------- */
Office.onReady(function (info) {
    // Stamp the build version into the footer (and console) regardless of host,
    // so you can verify which build loaded.
    showBuildVersion();

    if (info.host === Office.HostType.Outlook) {
        officeReady = true;

        // Feature-detect the base64 attach API (requirement set 1.8).
        // Absent on Outlook 2019 (req set 1.6).
        try {
            canAttachBase64 = !!(Office.context &&
                Office.context.mailbox &&
                Office.context.mailbox.item &&
                typeof Office.context.mailbox.item.addFileAttachmentFromBase64Async === 'function');
        } catch (e) {
            canAttachBase64 = false;
        }

        setStatus('Connected to Outlook - ready.');
        initDropZone();
    } else {
        setStatus('Error: This add-in only works in Outlook.');
    }
});


/* ============================================================
   DROP ZONE - HTML5 Drag & Drop
   ============================================================ */
function initDropZone() {
    var zone = document.getElementById('dropZone');

    // Listen on document level so drops anywhere in the pane are caught
    document.addEventListener('dragenter', function (e) {
        e.preventDefault();
        zone.className = addClass(zone.className, 'drag-over');
    });

    document.addEventListener('dragover', function (e) {
        e.preventDefault();
        if (e.dataTransfer) {
            e.dataTransfer.dropEffect = 'copy';
        }
    });

    document.addEventListener('dragleave', function (e) {
        e.preventDefault();
        if (!document.documentElement.contains(e.relatedTarget)) {
            zone.className = removeClass(zone.className, 'drag-over');
        }
    });

    document.addEventListener('drop', function (e) {
        e.preventDefault();
        e.stopPropagation();
        zone.className = removeClass(zone.className, 'drag-over');
        handleDrop(e);
    });

    zone.addEventListener('dragenter', function (e) {
        e.preventDefault();
        e.stopPropagation();
        zone.className = addClass(zone.className, 'drag-over');
    });

    zone.addEventListener('dragover', function (e) {
        e.preventDefault();
        e.stopPropagation();
        if (e.dataTransfer) {
            e.dataTransfer.dropEffect = 'copy';
        }
        zone.className = addClass(zone.className, 'drag-over');
    });

    zone.addEventListener('dragleave', function (e) {
        e.preventDefault();
        e.stopPropagation();
        if (!zone.contains(e.relatedTarget)) {
            zone.className = removeClass(zone.className, 'drag-over');
        }
    });

    // Click to open file picker
    zone.addEventListener('click', function () {
        var picker = document.getElementById('filePicker');
        if (picker) {
            picker.click();
        }
    });

    // Paste support (Ctrl+V)
    document.addEventListener('paste', function (e) {
        var files = e.clipboardData && e.clipboardData.files;
        if (files && files.length > 0) {
            e.preventDefault();
            handleDroppedFiles(files);
        }
    });
}

/**
 * Resolve a drop event into an attach action.
 *
 * Order is URL-FIRST because URL attach is the only path that works
 * on every Outlook version. On Outlook 2019 the custom-MIME and
 * file-object branches will be empty/unsupported anyway, so a SAS
 * URL dragged as plain text is what lands here.
 */
function handleDrop(e) {
    var dt = e.dataTransfer;
    if (!dt) {
        setStatus('No data in drop.');
        return;
    }

    // IE11 / Trident (classic Outlook 2019) does NOT recognize MIME format names in
    // getData — it only knows the legacy 'Text' and 'URL'. So read both: the MIME
    // name (modern Outlook / Chromium) and the legacy name (Outlook 2019).
    var plainText = safeGet(dt, 'text/plain') || safeGet(dt, 'Text');
    var uriList = safeGet(dt, 'text/uri-list') || safeGet(dt, 'URL');

    // 1a. Obfuscated gateway blob (the normal Salesforce path). It is NOT an http
    //     URL on the clipboard — so casually dragging it into a browser does nothing.
    //     Decode it back to "<url>|<filename>" and attach. (Key is public; this is
    //     anti-accident obfuscation, not security.)
    if (plainText && plainText.indexOf('http') !== 0 && !firstUrl(uriList)) {
        var decoded = unscrambleLink(trim(plainText));
        if (decoded && decoded.indexOf('http') === 0) {
            var sep = decoded.indexOf('|');
            var dUrl = sep > -1 ? decoded.substring(0, sep) : decoded;
            var dName = sep > -1 ? decoded.substring(sep + 1) : '';
            handleDroppedUrl(dUrl, dName);
            return;
        }
    }

    // 1b. Plain URL dragged as text / uri-list (legacy or direct drag).
    var url = firstUrl(uriList) || firstUrl(plainText);
    if (url) {
        handleDroppedUrl(url, nameHintFromText(plainText));
        return;
    }

    // 2. URL embedded in dragged HTML (anchor/img).
    var htmlData = safeGet(dt, 'text/html');
    if (htmlData) {
        var m = htmlData.match(/(?:href|src)=["']([^"']+)["']/i);
        if (m && /^https?:\/\//.test(m[1])) {
            handleDroppedUrl(m[1]);
            return;
        }
    }

    // 3. Custom base64 payload from Salesforce LWC.
    //    NOTE: IE11 (Outlook 2019) cannot read custom MIME types, so
    //    this branch only ever fires on modern Outlook. And it needs
    //    the 1.8 base64 API.
    var fileData = safeGet(dt, 'application/x-file-data');
    if (fileData) {
        try {
            var parsed = JSON.parse(fileData);
            if (parsed && parsed.base64Content && parsed.fileName) {
                if (!canAttachBase64) {
                    setStatus('This Outlook version cannot attach files by data. Drag the file as a link (SAS URL) instead.');
                    return;
                }
                setStatus('Attaching ' + parsed.fileName + '...');
                attachToEmail(parsed.base64Content, parsed.fileName, parsed.mimeType || 'application/octet-stream');
                return;
            }
        } catch (err) {
            if (window.console) { console.error('Failed to parse x-file-data:', err); }
        }
    }

    // 4. Real file objects (drag from desktop / file explorer).
    //    Reading file bytes -> base64 needs the 1.8 API, so this is
    //    modern-Outlook only.
    var files = dt.files;
    if (files && files.length > 0) {
        if (!canAttachBase64) {
            setStatus('This Outlook version cannot attach local files here. Use a file link (SAS URL) instead.');
            return;
        }
        handleDroppedFiles(files);
        return;
    }

    // Diagnostic (no console on Outlook 2019): show what the drop actually carried.
    var dbg = '';
    try {
        var t = dt.types ? Array.prototype.join.call(dt.types, ',') : 'none';
        var hadText = (safeGet(dt, 'text/plain') || safeGet(dt, 'Text')) ? 'y' : 'n';
        dbg = ' [types=' + t + '; text=' + hadText + ']';
    } catch (e) { /* ignore */ }
    setStatus('No files or links detected in drop.' + dbg);
}

/**
 * Attach a URL by letting Outlook/Exchange download it server-side.
 * Works on ALL Outlook versions (requirement set 1.1).
 */
function handleDroppedUrl(url, nameHint) {
    // Prefer the explicit name passed via the drag payload (keeps the filename
    // out of the URL); fall back to parsing the URL's last path segment.
    var fileName = nameHint || fileNameFromUrl(url);
    var rowId = addAttachmentRow(fileName, '', 'attaching');
    setStatus('Attaching ' + fileName + ' from URL...');

    Office.context.mailbox.item.addFileAttachmentAsync(
        url,
        fileName,
        { isInline: false },
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                if (window.console) { console.log('Attached from URL:', fileName, '| ID:', result.value); }
                updateAttachmentStatus(rowId, 'attached');
                setStatus('✓ ' + fileName + ' attached');
            } else {
                var msg = result.error ? result.error.message : 'unknown error';
                if (window.console) { console.warn('addFileAttachmentAsync failed:', msg); }

                // Only attempt a client-side download fallback if this
                // client actually supports base64 attach (modern Outlook).
                // On Outlook 2019 there is no fallback - the URL must be
                // server-fetchable (a valid SAS URL).
                if (canAttachBase64) {
                    fetchAndAttach(url, fileName, rowId);
                } else {
                    updateAttachmentStatus(rowId, 'error', 'URL not reachable by Outlook');
                    setStatus('✗ Could not attach ' + fileName + '. The link must be a public/SAS URL that Outlook can download.');
                }
            }
        }
    );
}

/**
 * Fallback: download the file client-side via XHR, convert to base64,
 * and attach. MODERN OUTLOOK ONLY (XHR-to-blob + base64 attach API).
 * Guarded by canAttachBase64 at the call site.
 */
function fetchAndAttach(url, fileName, rowId) {
    setStatus('Downloading ' + fileName + '...');
    try {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', url, true);
        xhr.responseType = 'blob';
        xhr.onload = function () {
            if (xhr.status < 200 || xhr.status >= 300) {
                updateAttachmentStatus(rowId, 'error', 'HTTP ' + xhr.status);
                setStatus('Error downloading ' + fileName + ' (HTTP ' + xhr.status + ')');
                return;
            }
            var blob = xhr.response;
            var mimeType = (blob && blob.type) || 'application/octet-stream';
            var reader = new FileReader();
            reader.onload = function () {
                var base64 = String(reader.result).split(',')[1];
                updateAttachmentStatus(rowId, 'attaching');
                attachToEmail(base64, fileName, mimeType, rowId);
            };
            reader.onerror = function () {
                updateAttachmentStatus(rowId, 'error', 'Failed to read file');
                setStatus('Error reading ' + fileName);
            };
            reader.readAsDataURL(blob);
        };
        xhr.onerror = function () {
            updateAttachmentStatus(rowId, 'error', 'Download failed (CORS/auth?)');
            setStatus('Error: Could not download ' + fileName + '. The file may require authentication or block CORS.');
        };
        xhr.send();
    } catch (err) {
        if (window.console) { console.error('fetchAndAttach error:', err); }
        updateAttachmentStatus(rowId, 'error', err && err.message);
        setStatus('Error: Could not download ' + fileName + '.');
    }
}

/**
 * Handle a list of dropped/picked File objects (base64 path).
 */
function handleDroppedFiles(fileList) {
    var count = fileList.length;
    setStatus('Processing ' + count + ' file(s)...');

    for (var i = 0; i < fileList.length; i++) {
        var file = fileList[i];
        readAndAttach(file, file.name, file.type || 'application/octet-stream');
    }
}

/**
 * Handle file picker selection.
 */
function handleFilePick(event) {
    if (!canAttachBase64) {
        setStatus('This Outlook version cannot attach local files. Use a file link (SAS URL) instead.');
        event.target.value = '';
        return;
    }
    var files = event.target.files;
    if (files && files.length > 0) {
        handleDroppedFiles(files);
    }
    event.target.value = '';
}

/**
 * Read a File object as base64 and attach it to the current email.
 */
function readAndAttach(file, fileName, mimeType) {
    var rowId = addAttachmentRow(fileName, mimeType, 'queued');
    var reader = new FileReader();

    reader.onload = function () {
        var base64 = String(reader.result).split(',')[1];
        updateAttachmentStatus(rowId, 'attaching');
        attachToEmail(base64, fileName, mimeType, rowId);
    };

    reader.onerror = function () {
        if (window.console) { console.error('FileReader error for', fileName); }
        updateAttachmentStatus(rowId, 'error', 'Failed to read file');
        setStatus('Error reading ' + fileName);
    };

    reader.readAsDataURL(file);
}


/* ============================================================
   ATTACH TO EMAIL (base64) - Office.js
   ============================================================ */

/**
 * Attach a base64-encoded file to the current compose item.
 * Requires requirement set 1.8 (NOT available on Outlook 2019).
 */
function attachToEmail(base64, fileName, mimeType, rowId) {
    if (!officeReady) {
        var msg = 'Office.js not ready - cannot attach.';
        if (window.console) { console.error(msg); }
        if (rowId) { updateAttachmentStatus(rowId, 'error', msg); }
        setStatus(msg);
        return;
    }

    if (!canAttachBase64) {
        var msg2 = 'This Outlook version cannot attach by data. Use a SAS URL link.';
        if (rowId) { updateAttachmentStatus(rowId, 'error', msg2); }
        setStatus(msg2);
        return;
    }

    if (!rowId) {
        rowId = addAttachmentRow(fileName, mimeType, 'attaching');
    }

    setStatus('Attaching ' + fileName + '...');

    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
        base64,
        fileName,
        { isInline: false },
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                if (window.console) { console.log('Attached:', fileName, '| ID:', result.value); }
                updateAttachmentStatus(rowId, 'attached');
                setStatus('✓ ' + fileName + ' attached');
            } else {
                var emsg = result.error ? result.error.message : 'unknown error';
                if (window.console) { console.error('Attach failed:', emsg); }
                updateAttachmentStatus(rowId, 'error', emsg);
                setStatus('✗ Failed to attach ' + fileName);
            }
        }
    );
}


/* ============================================================
   UI HELPERS
   ============================================================ */

var attachmentCounter = 0;

/**
 * Add a row to the attachment list UI and return its ID.
 */
function addAttachmentRow(fileName, mimeType, status, detail) {
    var id = 'att-' + (++attachmentCounter);
    var listEl = document.getElementById('attachmentList');
    var iconClass = getFileIconClass(fileName);

    var row = document.createElement('div');
    row.className = 'attachment-item';
    row.id = id;
    row.innerHTML =
        '<i class="ms-Icon ' + iconClass + '" aria-hidden="true"></i>' +
        '<div class="attachment-info">' +
            '<span class="attachment-name" title="' + escapeAttr(fileName) + '">' + escapeHtml(fileName) + '</span>' +
            '<span class="attachment-meta">' + escapeHtml(mimeType || '') + '</span>' +
        '</div>' +
        '<span class="attachment-status status-' + status + '">' + statusLabel(status, detail) + '</span>';

    // IE11 has no Element.prepend()
    if (listEl) {
        if (listEl.firstChild) {
            listEl.insertBefore(row, listEl.firstChild);
        } else {
            listEl.appendChild(row);
        }
    }
    return id;
}

/**
 * Update an existing attachment row's status.
 */
function updateAttachmentStatus(rowId, status, detail) {
    var row = document.getElementById(rowId);
    if (!row) { return; }

    var badge = row.querySelector('.attachment-status');
    if (badge) {
        badge.className = 'attachment-status status-' + status;
        badge.textContent = statusLabel(status, detail);
    }
}

function statusLabel(status, detail) {
    switch (status) {
        case 'queued':    return 'Queued';
        case 'attaching': return 'Attaching...';
        case 'attached':  return '✓ Attached';
        case 'error':     return '✗ ' + (detail || 'Error');
        default:          return status;
    }
}

function setStatus(msg) {
    var el = document.getElementById('statusText');
    if (el) { el.textContent = msg; }
}

/** Show the build version in the footer so the loaded build is verifiable. */
function showBuildVersion() {
    var el = document.getElementById('buildVersion');
    if (el) { el.textContent = 'v' + ADDIN_VERSION; }
    if (window.console) { console.log('Azure File Attacher build v' + ADDIN_VERSION); }
}

function getFileIconClass(fileName) {
    if (!fileName) { return 'ms-Icon--Page'; }
    var ext = fileName.split('.').pop().toLowerCase();
    var map = {
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


/* ============================================================
   UTILITY (ES5 / IE11 safe)
   ============================================================ */

/** Safely read a dataTransfer format (some throw in IE11). */
function safeGet(dt, type) {
    try {
        return dt.getData(type) || '';
    } catch (e) {
        return '';
    }
}

/** Return the first http(s) URL found in a (possibly multi-line) string. */
function firstUrl(text) {
    if (!text) { return ''; }
    var lines = text.split(/[\r\n]+/);
    for (var i = 0; i < lines.length; i++) {
        var line = trim(lines[i]);
        // uri-list comment lines start with '#'
        if (line && line.charAt(0) !== '#' && /^https?:\/\//i.test(line)) {
            return line;
        }
    }
    return '';
}

/**
 * Reverse of the LWC's scrambleLink: keyed XOR + base64url. The key is PUBLIC by
 * design — it only keeps a usable URL off the clipboard so a casual drag/paste into
 * a browser bar does nothing. Must stay identical to the LWC's key/algorithm.
 * Returns "" on any failure (so we fall through to the plain-URL path).
 */
function unscrambleLink(blob) {
    try {
        var key = 'RefuahAttachGw';
        var b64 = String(blob).replace(/-/g, '+').replace(/_/g, '/');
        while (b64.length % 4) { b64 += '='; }
        var x = atob(b64);
        var s = '';
        for (var i = 0; i < x.length; i++) {
            s += String.fromCharCode(x.charCodeAt(i) ^ key.charCodeAt(i % key.length));
        }
        return decodeURIComponent(s);
    } catch (e) {
        return '';
    }
}

/**
 * Find the filename hint in dragged text (legacy plain-URL path). Returns the first
 * line that is NOT a URL and not a uri-list comment.
 */
function nameHintFromText(text) {
    if (!text) { return ''; }
    var lines = text.split(/[\r\n]+/);
    for (var i = 0; i < lines.length; i++) {
        var line = trim(lines[i]);
        if (line && line.charAt(0) !== '#' && !/^https?:\/\//i.test(line)) {
            return line;
        }
    }
    return '';
}

/**
 * Derive a filename from a URL without using new URL().
 * Fallback only — used when no explicit name hint was dragged (e.g. a plain SAS
 * URL whose path ends in the real filename, or a dragged HTML anchor).
 */
function fileNameFromUrl(url) {
    var s = String(url);
    // strip query and fragment
    s = s.split('#')[0].split('?')[0];
    // take last path segment
    var parts = s.split('/');
    var last = parts[parts.length - 1] || 'attachment';
    try {
        last = decodeURIComponent(last);
    } catch (e) { /* keep raw */ }
    return last || 'attachment';
}

function trim(s) {
    return String(s).replace(/^\s+|\s+$/g, '');
}

/** className helpers (IE11 lacks classList on some elements). */
function addClass(className, cls) {
    var c = ' ' + className + ' ';
    if (c.indexOf(' ' + cls + ' ') === -1) {
        return trim(className + ' ' + cls);
    }
    return className;
}

function removeClass(className, cls) {
    return trim((' ' + className + ' ').replace(' ' + cls + ' ', ' '));
}

function escapeHtml(str) {
    var div = document.createElement('div');
    div.textContent = str || '';
    return div.innerHTML;
}

function escapeAttr(str) {
    return (str || '')
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}
