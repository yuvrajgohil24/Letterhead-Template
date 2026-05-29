const GAS_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbya_xwH1Xzdm8ErfCtnca_I6TcTuwcNq-BFdnU1_iKmXHuwWq7XMpvT9INMW_Ch3rkC/exec";

// ===================== IMAGE UTILS =====================
async function convertImagesToBase64(element) {
    const images = element.getElementsByTagName('img');
    for (let img of images) {
        if (img.src.startsWith('data:')) continue;
        try {
            const response = await fetch(img.src);
            const blob = await response.blob();
            const reader = new FileReader();
            await new Promise((resolve, reject) => {
                reader.onloadend = () => {
                    img.src = reader.result;
                    resolve();
                };
                reader.onerror = reject;
                reader.readAsDataURL(blob);
            });
        } catch (err) {
            console.warn('Failed to convert image to Base64:', img.src, err);
        }
    }
}

// ===================== ANNEXURE STATUS =====================
function updateAnnexureStatus() {
    const input = document.getElementById('annexureInput');
    const status = document.getElementById('annexureStatus');
    if (input.files.length > 0) {
        status.innerText = '📎 ' + input.files[0].name;
        status.style.color = 'var(--brand-orange)';
    } else {
        status.innerText = 'No Annexure';
        status.style.color = '#666';
    }
}

// ===================== DOWNLOAD PDF - FIXED =====================
async function downloadPDF() {
    const pageSelect = document.getElementById('pageSelect').value;
    const pdfRoot = document.getElementById('pdf-root');
    const annexureInput = document.getElementById('annexureInput');

    // Determine which pages to capture
    const allPages = Array.from(document.querySelectorAll('#pdf-root > .page'));
    let pagesToCapture = [];

    if (pageSelect === 'page1') {
        pagesToCapture = [allPages[0]];
    } else {
        pagesToCapture = allPages;
    }

    // Add capture mode
    pdfRoot.classList.add('pdf-capture-mode');
    await convertImagesToBase64(pdfRoot);

    // Scroll to top and wait
    window.scrollTo(0, 0);
    await new Promise(r => setTimeout(r, 150));

    try {
        const { PDFDocument } = PDFLib;
        const finalDoc = await PDFDocument.create();

        // A4 in points: 595 x 842 pt  (1mm = 2.8346pt)
        const A4_W_PT = 595.28;
        const A4_H_PT = 841.89;

        for (let i = 0; i < pagesToCapture.length; i++) {
            const pageEl = pagesToCapture[i];
            if (!pageEl) continue;

            // Scroll page into view so html2canvas can see it properly
            pageEl.scrollIntoView({ behavior: 'instant', block: 'start' });
            await new Promise(r => setTimeout(r, 100));

            const canvas = await html2canvas(pageEl, {
                scale: 2,
                useCORS: true,
                allowTaint: true,
                backgroundColor: '#ffffff',
                width: pageEl.offsetWidth,
                height: pageEl.offsetHeight,
                scrollX: 0,
                scrollY: -window.scrollY,
                logging: false,
                onclone: (clonedDoc) => {
                    // Hide page numbers and buttons in the cloned version
                    clonedDoc.querySelectorAll('.page-number, button').forEach(el => {
                        el.style.display = 'none';
                    });
                }
            });

            // Convert canvas → PNG bytes → embed in PDF
            const imgDataUrl = canvas.toDataURL('image/png');
            const imgBase64 = imgDataUrl.split(',')[1];
            const imgBytes = Uint8Array.from(atob(imgBase64), c => c.charCodeAt(0));

            const pngImage = await finalDoc.embedPng(imgBytes);
            const pdfPage = finalDoc.addPage([A4_W_PT, A4_H_PT]);
            pdfPage.drawImage(pngImage, {
                x: 0,
                y: 0,
                width: A4_W_PT,
                height: A4_H_PT
            });
        }

        let finalPdfBytes = await finalDoc.save();

        // Merge annexure if present
        if (annexureInput.files.length > 0) {
            try {
                const annexureFile = annexureInput.files[0];
                const annexureBuffer = await annexureFile.arrayBuffer();

                const mainPdfDoc = await PDFDocument.load(finalPdfBytes);
                const annexurePdfDoc = await PDFDocument.load(annexureBuffer);
                const mergedPdfDoc = await PDFDocument.create();

                const mainPages = await mergedPdfDoc.copyPages(mainPdfDoc, mainPdfDoc.getPageIndices());
                mainPages.forEach(p => mergedPdfDoc.addPage(p));

                const annexurePages = await mergedPdfDoc.copyPages(annexurePdfDoc, annexurePdfDoc.getPageIndices());
                annexurePages.forEach(p => mergedPdfDoc.addPage(p));

                finalPdfBytes = await mergedPdfDoc.save();
            } catch (mergeErr) {
                console.error('Merge error:', mergeErr);
                alert('Letterhead generated, but Annexure merge failed.');
            }
        }

        // Download the PDF
        const blob = new Blob([finalPdfBytes], { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'HOSPkart_Letterhead.pdf';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        // Email via GAS (optional)
        if (GAS_WEBAPP_URL && GAS_WEBAPP_URL !== "YOUR_GOOGLE_APPS_SCRIPT_URL_HERE") {
            let binary = '';
            const bytesArray = new Uint8Array(finalPdfBytes);
            for (let i = 0; i < bytesArray.byteLength; i++) {
                binary += String.fromCharCode(bytesArray[i]);
            }
            const base64Str = btoa(binary);
            const params = new URLSearchParams();
            params.append('pdf', base64Str);
            params.append('to', 'try.rajrathore@gmail.com');
            params.append('pageSelection', pageSelect);
            fetch(GAS_WEBAPP_URL, {
                method: 'POST',
                mode: 'no-cors',
                body: params
            }).catch(error => console.error('Email error:', error));
        }

    } catch (err) {
        console.error('PDF Error:', err);
        alert('Error generating PDF: ' + err.message);
    } finally {
        pdfRoot.classList.remove('pdf-capture-mode');
        window.scrollTo(0, 0);
    }
}


// ===================== DRAFTS & EXPORT LOGIC =====================
const EDITABLE_IDS = [
    'body-p1', 'body-p3'
];

function toggleSidebar() {
    const sidebar = document.getElementById('sidebar');
    sidebar.classList.toggle('closed');
}

function toggleEditorSidebar() {
    const editorSidebar = document.getElementById('editorSidebar');
    editorSidebar.classList.toggle('closed');
}

// --- Enhanced getFormData: also saves dynamic blank pages ---
function getFormData() {
    const data = {};
    EDITABLE_IDS.forEach(id => {
        const el = document.getElementById(id);
        if (el) data[id] = el.innerHTML;
    });
    const dynamicPages = document.querySelectorAll('[id^="page-blank-"]');
    data.__blankPages = [];
    dynamicPages.forEach(page => {
        const contentEl = page.querySelector('[contenteditable="true"]');
        data.__blankPages.push({
            id: page.id,
            contentId: contentEl ? contentEl.id : '',
            html: contentEl ? contentEl.innerHTML : ''
        });
    });
    return data;
}

// --- Enhanced setFormData: re-creates dynamic blank pages ---
function setFormData(data) {
    document.querySelectorAll('[id^="page-blank-"]').forEach(p => p.remove());

    if (data.__blankPages && data.__blankPages.length > 0) {
        const pdfRoot = document.getElementById('pdf-root');
        const page3 = document.getElementById('page3');
        data.__blankPages.forEach(bp => {
            const newPage = document.createElement('div');
            newPage.className = 'letterhead page page-break blank-page';
            newPage.id = bp.id;
            newPage.innerHTML = `
                <div class="content" style="flex: 1;">
                    <div class="blank-page-content" contenteditable="true" id="${bp.contentId}">
                        ${bp.html}
                    </div>
                </div>
                <button onclick="this.parentElement.remove(); updatePageNumbers();" style="position: absolute; top: 15px; right: 20px; background: #ffefef; color: #d32f2f; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 12px; z-index: 10;">🗑️ Remove Page</button>
            `;
            if (page3) {
                pdfRoot.insertBefore(newPage, page3);
            } else {
                pdfRoot.appendChild(newPage);
            }
            const editable = newPage.querySelector('[contenteditable="true"]');
            if (editable) {
                editable.addEventListener('click', () => {
                    const sidebar = document.getElementById('editorSidebar');
                    if (sidebar.classList.contains('closed')) {
                        toggleEditorSidebar();
                    }
                });
            }
            if (!EDITABLE_IDS.includes(bp.contentId)) {
                EDITABLE_IDS.push(bp.contentId);
            }
        });
        blankPageCount = Math.max(blankPageCount, data.__blankPages.length + 2);
    }

    EDITABLE_IDS.forEach(id => {
        const el = document.getElementById(id);
        if (el && data[id] !== undefined) {
            el.innerHTML = data[id];
        }
    });

    updatePageNumbers();
}

// --- LocalStorage Drafts ---
function saveDraft() {
    const name = prompt("Enter a name for this draft (e.g., 'Candidate John Doe'):");
    if (!name) return;

    const data = getFormData();
    const draft = {
        id: Date.now(),
        name: name,
        data: data,
        date: new Date().toLocaleString()
    };

    let drafts = JSON.parse(localStorage.getItem('letterhead_drafts') || '[]');
    const existingIndex = drafts.findIndex(d => d.name === name);
    if (existingIndex >= 0) {
        if (!confirm(`Draft '${name}' already exists. Overwrite?`)) return;
        drafts[existingIndex] = draft;
    } else {
        drafts.push(draft);
    }

    localStorage.setItem('letterhead_drafts', JSON.stringify(drafts));
    renderDraftList();
    alert('Draft saved successfully!');
}

function loadDraft(id) {
    if (!confirm("Loading a draft will overwrite current changes. Continue?")) return;
    const drafts = JSON.parse(localStorage.getItem('letterhead_drafts') || '[]');
    const draft = drafts.find(d => d.id === id);
    if (draft) {
        setFormData(draft.data);
    }
}

function deleteDraft(e, id) {
    e.stopPropagation();
    if (!confirm("Are you sure you want to delete this draft?")) return;
    let drafts = JSON.parse(localStorage.getItem('letterhead_drafts') || '[]');
    drafts = drafts.filter(d => d.id !== id);
    localStorage.setItem('letterhead_drafts', JSON.stringify(drafts));
    renderDraftList();
}

function renderDraftList() {
    const list = document.getElementById('draftList');
    const drafts = JSON.parse(localStorage.getItem('letterhead_drafts') || '[]');
    if (drafts.length === 0) {
        list.innerHTML = '<div style="padding:10px; color:#999; text-align:center; font-size:12px;">No saved drafts</div>';
        return;
    }
    list.innerHTML = drafts.map(d => `
        <div class="draft-item" onclick="loadDraft(${d.id})">
            <div class="draft-name" title="${d.name}">${d.name}</div>
            <div class="draft-actions">
                <div class="action-icon" onclick="deleteDraft(event, ${d.id})" title="Delete">🗑️</div>
            </div>
        </div>
    `).join('');
}

function clearCurrent() {
    if (!confirm("Clear all current fields?")) return;
    location.reload();
}

// --- Export / Import ---
function exportData() {
    const data = getFormData();
    const jsonStr = JSON.stringify(data, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'letterhead_data_' + new Date().toISOString().slice(0, 10) + '.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function importData(input) {
    const file = input.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = JSON.parse(e.target.result);
            setFormData(data);
            alert('Data imported successfully!');
        } catch (err) {
            alert('Error importing file: Invalid JSON');
            console.error(err);
        }
        input.value = '';
    };
    reader.readAsText(file);
}

// ===================== SELECTION & FORMATTING =====================
let lastSelection = null;

let focusedEditable = null;
let lastFocusedPage = null;

document.addEventListener('focusin', (e) => {
    if (e.target && e.target.hasAttribute && e.target.hasAttribute('contenteditable')) {
        focusedEditable = e.target;
        lastFocusedPage = e.target.closest('.page');
    }
});

document.addEventListener('focusout', (e) => {
    // Keep reference for a brief moment so paste events still have context
    setTimeout(() => {
        if (document.activeElement && document.activeElement.hasAttribute &&
            document.activeElement.hasAttribute('contenteditable')) {
            focusedEditable = document.activeElement;
        } else {
            focusedEditable = null;
        }
    }, 200);
});

function saveSelection() {
    const sel = window.getSelection();
    if (sel.rangeCount > 0) {
        lastSelection = sel.getRangeAt(0);
    }
}

function restoreSelection() {
    if (lastSelection) {
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(lastSelection);
    }
}

document.addEventListener('selectionchange', () => {
    const sel = window.getSelection();
    if (sel.rangeCount > 0) {
        let node = sel.anchorNode;
        while (node && node !== document.body) {
            if (node.nodeType === 1 && node.hasAttribute('contenteditable')) {
                saveSelection();
                break;
            }
            node = node.parentNode;
        }
    }
});

function formatDoc(cmd, value = null) {
    restoreSelection();
    if (cmd === 'fontSize' && value && !isNaN(value)) {
        document.execCommand('fontSize', false, '7');
        const fontElems = document.querySelectorAll('font[size="7"]');
        fontElems.forEach(el => {
            el.removeAttribute('size');
            el.style.fontSize = value + 'px';
        });
    } else {
        if (value) {
            document.execCommand(cmd, false, value);
        } else {
            document.execCommand(cmd, false, null);
        }
    }
    // After any formatting change, trigger a debounced reflow
    debouncedReflow();
}

// ===================== INSERT IMAGE =====================
function insertImageAtCursor() {
    saveSelection();
    document.getElementById('inlineImageInput').click();
}

function handleInlineImage(input) {
    const file = input.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (e) {
        restoreSelection();
        const imgHtml = `<img src="${e.target.result}" style="max-width:100%; height:auto; border-radius:4px; margin:8px 0;" />`;
        document.execCommand('insertHTML', false, imgHtml);
    };
    reader.readAsDataURL(file);
    input.value = '';
}

// ===================== INSERT TABLE =====================
function insertTableAtCursor() {
    saveSelection();
    const rows = prompt('Number of rows:', '3');
    const cols = prompt('Number of columns:', '3');
    if (!rows || !cols) return;
    const r = parseInt(rows), c = parseInt(cols);
    if (isNaN(r) || isNaN(c) || r < 1 || c < 1) return;

    let tableHtml = '<table>';
    tableHtml += '<tr>';
    for (let j = 0; j < c; j++) {
        tableHtml += `<th contenteditable="true">Header ${j + 1}</th>`;
    }
    tableHtml += '</tr>';
    for (let i = 0; i < r - 1; i++) {
        tableHtml += '<tr>';
        for (let j = 0; j < c; j++) {
            tableHtml += `<td contenteditable="true">Cell</td>`;
        }
        tableHtml += '</tr>';
    }
    tableHtml += '</table><br>';

    restoreSelection();
    document.execCommand('insertHTML', false, tableHtml);
}

// ===================== CHANGE LOGO =====================
function changeLogo(input) {
    const file = input.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (e) {
        const logos = document.querySelectorAll('img[alt="HOSPkart Logo"]');
        logos.forEach(logo => {
            logo.src = e.target.result;
        });
    };
    reader.readAsDataURL(file);
    input.value = '';
}

// ===================== PAGE NUMBERING =====================
function updatePageNumbers() {
    const pages = document.querySelectorAll('#pdf-root > .letterhead.page');
    const total = pages.length;
    pages.forEach((page, i) => {
        let pnEl = page.querySelector('.page-number');
        if (!pnEl) {
            pnEl = document.createElement('div');
            pnEl.className = 'page-number';
            page.appendChild(pnEl);
        }
        pnEl.textContent = `Page ${i + 1} of ${total}`;
    });
}

// ===================== AUTO-SAVE =====================
function autoSave() {
    const data = getFormData();
    localStorage.setItem('letterhead_autosave', JSON.stringify(data));
    const toast = document.getElementById('autosaveToast');
    if (toast) {
        toast.classList.add('show');
        setTimeout(() => toast.classList.remove('show'), 1500);
    }
}

function checkAutoSaveRestore() {
    const saved = localStorage.getItem('letterhead_autosave');
    if (!saved) return;
    const banner = document.createElement('div');
    banner.className = 'restore-banner';
    banner.innerHTML = `
        <span>📄 Unsaved work found from your last session.</span>
        <button class="btn-restore" onclick="restoreAutoSave(this.parentElement)">Restore</button>
        <button class="btn-dismiss" onclick="this.parentElement.remove()">Dismiss</button>
    `;
    document.body.prepend(banner);
}

function restoreAutoSave(bannerEl) {
    const saved = localStorage.getItem('letterhead_autosave');
    if (saved) {
        try {
            const data = JSON.parse(saved);
            setFormData(data);
        } catch (e) {
            console.error('Failed to restore auto-save:', e);
        }
    }
    if (bannerEl) bannerEl.remove();
}

// ===================== EDITOR SIDEBAR AUTO-OPEN =====================
function initAutoSidebar() {
    const editables = document.querySelectorAll('[contenteditable="true"]');
    editables.forEach(el => {
        el.addEventListener('click', () => {
            const sidebar = document.getElementById('editorSidebar');
            if (sidebar.classList.contains('closed')) {
                toggleEditorSidebar();
            }
        });
    });
}

// ===================== ADD BLANK PAGE =====================
let blankPageCount = 2;
function addBlankPage() {
    blankPageCount++;
    const pdfRoot = document.getElementById('pdf-root');
    const page3 = document.getElementById('page3');

    const newPage = document.createElement('div');
    newPage.className = 'letterhead page page-break blank-page';
    newPage.id = 'page-blank-' + blankPageCount;

    newPage.innerHTML = `
        <div class="content" style="flex: 1;">
            <div class="blank-page-content" contenteditable="true" id="body-blank-${blankPageCount}">
                <p>Click here to edit new blank page...</p>
            </div>
        </div>
        <button onclick="this.parentElement.remove(); updatePageNumbers();" style="position: absolute; top: 15px; right: 20px; background: #ffefef; color: #d32f2f; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 12px; z-index: 10;">🗑️ Remove Page</button>
    `;

    if (lastFocusedPage) {
        lastFocusedPage.insertAdjacentElement('afterend', newPage);
    } else if (page3) {
        pdfRoot.insertBefore(newPage, page3);
    } else {
        pdfRoot.appendChild(newPage);
    }

    const editable = newPage.querySelector('[contenteditable="true"]');
    editable.addEventListener('click', () => {
        const sidebar = document.getElementById('editorSidebar');
        if (sidebar.classList.contains('closed')) {
            toggleEditorSidebar();
        }
    });
    // Trigger reflow quickly on paste so overflow moves to next page immediately
    editable.addEventListener('paste', () => {
        clearTimeout(reflowTimer);
        reflowTimer = setTimeout(() => reflow(), 100);
    });

    if (!EDITABLE_IDS.includes('body-blank-' + blankPageCount)) {
        EDITABLE_IDS.push('body-blank-' + blankPageCount);
    }

    updatePageNumbers();
}

// ===================== INITIALIZATION =====================
document.addEventListener('DOMContentLoaded', () => {
    renderDraftList();
    initAutoSidebar();
    updatePageNumbers();
    checkAutoSaveRestore();

    setInterval(autoSave, 15000);

    // Use debounced reflow for all content changes
    document.addEventListener('input', (e) => {
        if (e.target.hasAttribute('contenteditable')) {
            debouncedReflow();
        }
    });

    // ===================== PASTE SANITIZER =====================
    // Intercept paste on ALL contenteditable elements.
    // Strips all inline styles, background colors, and unwanted tags
    // so that content pasted from dark-theme apps (Claude, VS Code, etc.)
    // does not break the letter layout.
    document.addEventListener('paste', (e) => {
        const target = e.target;
        if (!target || !target.isContentEditable) return;
        e.preventDefault();

        // Prefer HTML from clipboard so we preserve bold/italic/lists,
        // but we sanitize it heavily.
        let html = (e.clipboardData || window.clipboardData).getData('text/html');
        let text = (e.clipboardData || window.clipboardData).getData('text/plain');

        if (html) {
            html = sanitizePastedHTML(html);
        } else {
            // Fall back to plain text — convert newlines to <br> / <p>
            html = text
                .split(/\n/)
                .map(line => line.trim() ? `<p>${escapeHTML(line)}</p>` : '<br>')
                .join('');
        }

        // Insert the sanitized HTML at the cursor
        const sel = window.getSelection();
        if (sel && sel.rangeCount) {
            const range = sel.getRangeAt(0);
            range.deleteContents();
            const frag = range.createContextualFragment(html);
            range.insertNode(frag);
            range.collapse(false);
            sel.removeAllRanges();
            sel.addRange(range);
        }
        debouncedReflow();
        // Second safety pass after a longer delay to catch any layout measurements
        // that weren't ready during the first pass (e.g. images, complex flex layout).
        setTimeout(() => reflow(), 600);
    }, true);
});

// Escape special HTML chars for plain-text fallback
function escapeHTML(str) {
    return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

// Deep-clean pasted HTML:
// - Remove dangerous/structural tags entirely (script, style, head, meta, iframe...)
// - Keep text-level semantics: b, strong, i, em, u, s, p, br, ul, ol, li, h1-h6, a
// - Strip ALL inline styles, class, id, background, color attributes from every element
function sanitizePastedHTML(html) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');

    // Remove tags we never want
    const REMOVE_TAGS = ['script','style','head','meta','link','iframe','object','embed','form','input','button','select','textarea','svg','canvas'];
    REMOVE_TAGS.forEach(tag => {
        doc.querySelectorAll(tag).forEach(el => el.remove());
    });

    // Allowed tags — everything else gets replaced with its children
    const ALLOW_TAGS = new Set(['p','br','b','strong','i','em','u','s','strike','ul','ol','li','h1','h2','h3','h4','h5','h6','a','span','div','table','thead','tbody','tr','th','td','blockquote','code','pre']);

    function cleanNode(node) {
        if (node.nodeType === Node.TEXT_NODE) return;
        if (node.nodeType !== Node.ELEMENT_NODE) {
            node.remove();
            return;
        }
        const tag = node.tagName.toLowerCase();
        // Recursively clean children first
        Array.from(node.childNodes).forEach(cleanNode);

        if (!ALLOW_TAGS.has(tag)) {
            // Unwrap: replace element with its children
            const parent = node.parentNode;
            if (parent) {
                while (node.firstChild) parent.insertBefore(node.firstChild, node);
                parent.removeChild(node);
            }
            return;
        }

        // Strip all attributes except href on <a> and a minimal set on <td>/<th>
        const attrsToKeep = tag === 'a' ? ['href'] : (tag === 'td' || tag === 'th') ? ['colspan','rowspan'] : [];
        Array.from(node.attributes).forEach(attr => {
            if (!attrsToKeep.includes(attr.name)) {
                node.removeAttribute(attr.name);
            }
        });
        // Ensure no residual style/color
        node.style && (node.style.cssText = '');
    }

    Array.from(doc.body.childNodes).forEach(cleanNode);
    return doc.body.innerHTML;
}

// ===================== AUTO PAGINATION =====================
let isReflowing = false;
let reflowTimer = null;

function debouncedReflow() {
    clearTimeout(reflowTimer);
    reflowTimer = setTimeout(() => reflow(), 400);
}

// Get all editable content areas in page order
function getOrderedEditables() {
    const pages = Array.from(document.querySelectorAll('#pdf-root > .letterhead.page'));
    const editables = [];
    pages.forEach(page => {
        const el = page.querySelector('.letter-body, .letter-body-page2, .blank-page-content');
        if (el) editables.push(el);
    });
    return editables;
}

// Find the next editable after the given one, or create a blank page
function getOrCreateNextEditable(currentEl) {
    const editables = getOrderedEditables();
    const idx = editables.indexOf(currentEl);
    if (idx >= 0 && idx < editables.length - 1) {
        return editables[idx + 1];
    }
    // Need a new page
    addBlankPage();
    const newPageId = 'body-blank-' + blankPageCount;
    return document.getElementById(newPageId);
}

// ---- Shared overflow measurement helper ----
// Returns true if el's content exceeds its visible (client) height.
//
// KEY INSIGHT: Inside a flex column with overflow:hidden, browsers frequently
// report scrollHeight === clientHeight even when content is visually clipped.
// This makes the scrollHeight approach unreliable.
//
// SOLUTION: Use getBoundingClientRect(). With overflow:hidden on the parent,
// child elements are still laid out at their natural positions — only the
// *painting* is clipped. So a child whose getBoundingClientRect().bottom
// exceeds the parent's getBoundingClientRect().bottom is definitively overflowing.
function isOverflowing(el) {
    // 1. Fast path: check if scrollHeight exceeds clientHeight (now reliable due to CSS min-height constraint)
    if (el.scrollHeight > el.clientHeight + 5) return true;

    // Get element bounding rectangle
    const elRect = el.getBoundingClientRect();
    if (elRect.height < 2) return false; // Element not rendered yet — skip

    // 2. Range-based content bottom check:
    // Creates a range spanning all text and elements inside el, and gets its bounding box.
    // This ignores any scroll container limitations and yields the true layout bottom of all contents.
    if (el.firstChild) {
        try {
            const range = document.createRange();
            range.selectNodeContents(el);
            const rangeRect = range.getBoundingClientRect();
            if (rangeRect.height > 0 && rangeRect.bottom > elRect.bottom + 5) {
                return true;
            }
        } catch (e) {
            // Ignore Range errors
        }
    }

    // 3. Fallback path: walk backwards to check if the last rendered child extends below the bottom boundary
    return lastChildExceedsBottom(el, elRect.bottom);
}

// Check if the last meaningful node inside `el` extends below `bottomLimit` (px, viewport coords).
function lastChildExceedsBottom(el, bottomLimit) {
    // Walk backwards to find the last rendered child
    let node = el.lastChild;
    while (node) {
        if (node.nodeType === Node.ELEMENT_NODE) {
            const r = node.getBoundingClientRect();
            if (r.height > 0) return r.bottom > bottomLimit + 5;
        } else if (node.nodeType === Node.TEXT_NODE && node.textContent.trim()) {
            try {
                const range = document.createRange();
                range.selectNodeContents(node);
                const r = range.getBoundingClientRect();
                if (r.height > 0) return r.bottom > bottomLimit + 5;
            } catch (e) { /* ignore */ }
        }
        node = node.previousSibling;
    }
    return false;
}

// Forward overflow: move excess content from el to the next page
function checkOverflow(el) {
    if (!el || !isOverflowing(el)) return false;

    let contentToMoveNodes = [];

    // Remove nodes from the bottom until the element no longer overflows.
    while (isOverflowing(el) && el.lastChild) {
        let lastNode = el.lastChild;

        // Skip trailing empty text nodes
        while (lastNode && lastNode.nodeType === Node.TEXT_NODE && lastNode.textContent.trim() === '') {
            let prev = lastNode.previousSibling;
            el.removeChild(lastNode);
            lastNode = prev;
        }
        if (!lastNode) break;

        if (lastNode.nodeType === Node.TEXT_NODE) {
            const words = lastNode.textContent.split(' ');
            let movedTextChunks = [];
            while (isOverflowing(el) && words.length > 0) {
                movedTextChunks.unshift(words.pop());
                lastNode.textContent = words.join(' ');
            }
            if (movedTextChunks.length > 0) {
                contentToMoveNodes.unshift(movedTextChunks.join(' ') + ' ');
            }
        } else if (lastNode.tagName === 'BR') {
            contentToMoveNodes.unshift('<br>');
            el.removeChild(lastNode);
        } else if (lastNode.nodeType === Node.ELEMENT_NODE) {
            if (lastNode.tagName === 'TABLE') {
                const rows = Array.from(lastNode.rows);
                if (rows.length > 1) {
                    let movedRowsHtml = '';
                    while (isOverflowing(el) && rows.length > 1) {
                        const lastRow = rows.pop();
                        if (lastRow.parentNode.tagName === 'THEAD' && rows.length === 0) break;
                        movedRowsHtml = lastRow.outerHTML + movedRowsHtml;
                        lastRow.remove();
                    }
                    if (movedRowsHtml) {
                        const tableClone = lastNode.cloneNode(false);
                        const thead = lastNode.querySelector('thead');
                        if (thead) tableClone.appendChild(thead.cloneNode(true));
                        const newTbody = document.createElement('tbody');
                        newTbody.innerHTML = movedRowsHtml;
                        tableClone.appendChild(newTbody);
                        contentToMoveNodes.unshift(tableClone.outerHTML);
                    } else {
                        if (el.childNodes.length === 1) break;
                        contentToMoveNodes.unshift(lastNode.outerHTML);
                        el.removeChild(lastNode);
                    }
                } else {
                    if (el.childNodes.length === 1) break;
                    contentToMoveNodes.unshift(lastNode.outerHTML);
                    el.removeChild(lastNode);
                }
            } else {
                if (el.childNodes.length === 1 && lastNode.childNodes.length > 0) {
                    // Unwrap so children can be paginated individually
                    while (lastNode.firstChild) {
                        el.insertBefore(lastNode.firstChild, lastNode);
                    }
                    el.removeChild(lastNode);
                    continue;
                } else if (el.childNodes.length === 1) {
                    break;
                }
                contentToMoveNodes.unshift(lastNode.outerHTML);
                el.removeChild(lastNode);
            }
        }
    }

    if (contentToMoveNodes.length > 0) {
        const nextEditable = getOrCreateNextEditable(el);
        if (nextEditable) {
            // Clear the placeholder text that blank pages start with
            const placeholders = [
                '<p>Click here to edit new blank page...</p>',
                '<p>(Continue your letter content here...)</p>',
                '<p>Click to edit and add more content for this page.</p>'
            ];
            let currentHTML = nextEditable.innerHTML.trim();
            if (placeholders.some(p => currentHTML === p || currentHTML.startsWith(p))) {
                // Strip the leading placeholder paragraph before prepending
                placeholders.forEach(p => { currentHTML = currentHTML.replace(p, '').trim(); });
                nextEditable.innerHTML = contentToMoveNodes.join('') + (currentHTML ? currentHTML : '');
            } else {
                nextEditable.innerHTML = contentToMoveNodes.join('') + currentHTML;
            }
        }
        return true;
    }
    return false;
}

// Backward underflow: pull content from the next page back to this one
function checkUnderflow(el) {
    const editables = getOrderedEditables();
    const idx = editables.indexOf(el);
    if (idx < 0 || idx >= editables.length - 1) return false;

    const nextEditable = editables[idx + 1];
    const PLACEHOLDERS = [
        '<p>Click here to edit new blank page...</p>',
        '<p>(Continue your letter content here...)</p>',
        '<p>Click to edit and add more content for this page.</p>'
    ];
    const nextHTML = nextEditable ? nextEditable.innerHTML.trim() : '';
    if (!nextEditable || nextHTML === '' || PLACEHOLDERS.includes(nextHTML)) {
        return false;
    }

    // IMPORTANT FIX: If the user is currently editing a page BELOW this one,
    // do NOT pull content up from that page (or any page below it).
    // This prevents manually pasted content from being sucked up to a previous page.
    if (focusedEditable) {
        const focusedIdx = editables.indexOf(focusedEditable);
        // If the next editable is AT or AFTER the focused editable, skip underflow for it
        if (focusedIdx >= 0 && (idx + 1) >= focusedIdx) {
            return false;
        }
    }

    let pulled = false;

    // Try pulling nodes from the START of the next editable back to the END of this one
    while (nextEditable.firstChild && !isOverflowing(el)) {
        let firstNode = nextEditable.firstChild;

        // Skip leading empty text nodes
        while (firstNode && firstNode.nodeType === Node.TEXT_NODE && firstNode.textContent.trim() === '') {
            let next = firstNode.nextSibling;
            nextEditable.removeChild(firstNode);
            firstNode = next;
        }
        if (!firstNode) break;

        // Clone and tentatively append to current page
        const cloned = firstNode.cloneNode(true);
        el.appendChild(cloned);

        // Check if it fits
        if (isOverflowing(el)) {
            // Doesn't fit — put it back
            el.removeChild(cloned);

            // Try word-by-word for text nodes
            if (firstNode.nodeType === Node.TEXT_NODE) {
                const words = firstNode.textContent.split(' ');
                let fittedWords = [];
                const testNode = document.createTextNode('');
                el.appendChild(testNode);

                while (words.length > 0) {
                    fittedWords.push(words.shift());
                    testNode.textContent = fittedWords.join(' ');
                    if (isOverflowing(el)) {
                        words.unshift(fittedWords.pop());
                        break;
                    }
                }

                if (fittedWords.length > 0) {
                    testNode.textContent = fittedWords.join(' ') + ' ';
                    firstNode.textContent = words.join(' ');
                    pulled = true;
                } else {
                    el.removeChild(testNode);
                    break;
                }
            } else if (['DIV', 'P', 'UL', 'OL', 'LI', 'SPAN', 'B', 'I', 'U', 'STRONG', 'EM', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6'].includes(firstNode.tagName) && firstNode.childNodes.length > 0) {
                const wrapper = firstNode.cloneNode(false);
                el.appendChild(wrapper);
                
                let pulledChild = false;
                while (firstNode.firstChild) {
                    const childCloned = firstNode.firstChild.cloneNode(true);
                    wrapper.appendChild(childCloned);
                    if (isOverflowing(el)) {
                        wrapper.removeChild(childCloned);
                        if (firstNode.firstChild.nodeType === Node.TEXT_NODE) {
                            const words = firstNode.firstChild.textContent.split(' ');
                            let fittedWords = [];
                            const testNode = document.createTextNode('');
                            wrapper.appendChild(testNode);

                            while (words.length > 0) {
                                fittedWords.push(words.shift());
                                testNode.textContent = fittedWords.join(' ');
                                if (isOverflowing(el)) {
                                    words.unshift(fittedWords.pop());
                                    break;
                                }
                            }
                            if (fittedWords.length > 0) {
                                testNode.textContent = fittedWords.join(' ') + ' ';
                                firstNode.firstChild.textContent = words.join(' ');
                                pulledChild = true;
                            } else {
                                wrapper.removeChild(testNode);
                            }
                        }
                        break;
                    } else {
                        firstNode.removeChild(firstNode.firstChild);
                        pulledChild = true;
                    }
                }
                if (!pulledChild) {
                    el.removeChild(wrapper);
                    break;
                }
                pulled = true;
                if (firstNode.childNodes.length === 0) {
                    nextEditable.removeChild(firstNode);
                } else {
                    break;
                }
            } else {
                break;
            }
        } else {
            // It fits — remove from next page
            nextEditable.removeChild(firstNode);
            pulled = true;
        }
    }

    return pulled;
}

// Clean up empty dynamically-created blank pages
function cleanupEmptyPages() {
    const blankPages = document.querySelectorAll('[id^="page-blank-"]');
    blankPages.forEach(page => {
        const contentEl = page.querySelector('[contenteditable="true"]');
        if (contentEl) {
            const text = contentEl.innerText.trim();
            const html = contentEl.innerHTML.trim();
            if (text === '' || html === '' || html === '<br>' || html === '<p><br></p>') {
                page.remove();
            }
        }
    });
}

// Main reflow: run forward overflow, then backward underflow, then cleanup
function reflow() {
    if (isReflowing) return;
    isReflowing = true;

    // Freeze the viewport scroll so reflow DOM mutations don't cause the page to jump
    const savedScrollX = window.scrollX;
    const savedScrollY = window.scrollY;

    try {
        // Save cursor position
        const sel = window.getSelection();
        let savedAnchorNode = sel.anchorNode;
        let savedAnchorOffset = sel.anchorOffset;

        // Pass 1: Forward overflow — process each editable in order
        let maxPasses = 20;
        let changed = true;
        while (changed && maxPasses-- > 0) {
            changed = false;
            const editables = getOrderedEditables();
            for (const el of editables) {
                if (checkOverflow(el)) {
                    changed = true;
                }
            }
        }

        // Pass 2: Backward underflow — process each editable in order
        maxPasses = 20;
        changed = true;
        while (changed && maxPasses-- > 0) {
            changed = false;
            const editables = getOrderedEditables();
            for (const el of editables) {
                if (checkUnderflow(el)) {
                    changed = true;
                }
            }
        }

        // Pass 3: Cleanup empty blank pages
        cleanupEmptyPages();

        // Update page numbers
        updatePageNumbers();

        // Restore cursor if possible
        try {
            if (savedAnchorNode && savedAnchorNode.parentNode) {
                const range = document.createRange();
                range.setStart(savedAnchorNode, Math.min(savedAnchorOffset, savedAnchorNode.length || 0));
                range.collapse(true);
                sel.removeAllRanges();
                sel.addRange(range);
            }
        } catch (e) {
            // Cursor restoration failed, that's okay
        }
    } finally {
        isReflowing = false;
        // Restore the viewport scroll position so the user's view doesn't jump
        window.scrollTo(savedScrollX, savedScrollY);
    }
}