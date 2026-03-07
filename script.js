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

// ===================== DOWNLOAD PDF =====================
async function downloadPDF() {
    const pageSelect = document.getElementById('pageSelect').value;
    const pdfRoot = document.getElementById('pdf-root');
    const page3 = document.getElementById('page3');
    const annexureInput = document.getElementById('annexureInput');

    let removedPage = null;
    if (pageSelect === 'page1-2') {
        if (page3) {
            removedPage = page3;
            page3.remove();
        }
    }

    let element;
    if (pageSelect === 'page1') {
        element = document.getElementById('page1');
    } else {
        element = pdfRoot;
    }

    if (!element) return;

    await convertImagesToBase64(element);

    const allPages = document.querySelectorAll('.page, .letterhead');
    allPages.forEach(p => {
        p.style.boxShadow = 'none';
        p.style.border = 'none';
        p.style.margin = '0';
    });

    // Hide page numbers during PDF generation
    document.querySelectorAll('.page-number').forEach(pn => pn.style.display = 'none');

    const opt = {
        margin: 0,
        filename: 'HOSPkart_Letterhead.pdf',
        image: { type: 'jpeg', quality: 1.0 },
        html2canvas: {
            scale: 2,
            useCORS: true,
            backgroundColor: '#ffffff',
            scrollX: 0,
            scrollY: 0,
            x: 0,
            y: 0,
            width: element.offsetWidth,
            windowWidth: element.offsetWidth,
            logging: false
        },
        jsPDF: {
            unit: 'mm',
            format: 'a4',
            orientation: 'portrait'
        },
        pagebreak: { mode: ['css', 'legacy'] }
    };

    html2pdf().set(opt).from(element).output('arraybuffer').then(async function (pdfBuffer) {
        let finalPdfBytes = new Uint8Array(pdfBuffer);

        // Merge annexure if present
        if (annexureInput.files.length > 0) {
            try {
                const annexureFile = annexureInput.files[0];
                const annexureBuffer = await annexureFile.arrayBuffer();

                const { PDFDocument } = PDFLib;
                const mainPdfDoc = await PDFDocument.load(finalPdfBytes);
                const annexurePdfDoc = await PDFDocument.load(annexureBuffer);

                const mergedPdfDoc = await PDFDocument.create();

                const mainPages = await mergedPdfDoc.copyPages(mainPdfDoc, mainPdfDoc.getPageIndices());
                mainPages.forEach(page => mergedPdfDoc.addPage(page));

                const annexurePages = await mergedPdfDoc.copyPages(annexurePdfDoc, annexurePdfDoc.getPageIndices());
                annexurePages.forEach(page => mergedPdfDoc.addPage(page));

                finalPdfBytes = await mergedPdfDoc.save();
                console.log('Annexure merged successfully.');
            } catch (mergeErr) {
                console.error('Merge error:', mergeErr);
                alert('Error merging Annexure. Downloading letterhead only.');
            }
        }

        // Trigger download
        const blob = new Blob([finalPdfBytes], { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = opt.filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        // Restore UI
        allPages.forEach(p => {
            p.style.boxShadow = '0 0 15px rgba(0, 0, 0, 0.2)';
            p.style.border = '1px solid #ddd';
            p.style.margin = '0 0 40px 0';
        });
        document.querySelectorAll('.page-number').forEach(pn => pn.style.display = '');
        if (removedPage) {
            pdfRoot.appendChild(removedPage);
        }

        // Email via GAS
        let binary = '';
        const len = finalPdfBytes.byteLength;
        for (let i = 0; i < len; i++) {
            binary += String.fromCharCode(finalPdfBytes[i]);
        }
        const base64Str = btoa(binary);
        const params = new URLSearchParams();
        params.append('pdf', base64Str);
        params.append('to', 'try.rajrathore@gmail.com');
        params.append('pageSelection', pageSelect);

        if (GAS_WEBAPP_URL && GAS_WEBAPP_URL !== "YOUR_GOOGLE_APPS_SCRIPT_URL_HERE") {
            fetch(GAS_WEBAPP_URL, {
                method: 'POST',
                mode: 'no-cors',
                body: params
            }).catch(error => console.error('Email error:', error));
        }

    }).catch(error => {
        allPages.forEach(p => {
            p.style.boxShadow = '0 0 15px rgba(0, 0, 0, 0.2)';
            p.style.border = '1px solid #ddd';
            p.style.margin = '0 0 40px 0';
        });
        document.querySelectorAll('.page-number').forEach(pn => pn.style.display = '');
        if (removedPage) pdfRoot.appendChild(removedPage);
        console.error('PDF Error:', error);
        alert('Error: ' + error.message);
    });
}

// ===================== DRAFTS & EXPORT LOGIC =====================
const EDITABLE_IDS = [
    'ref-no', 'letter-date', 'recipient-block', 'body-p1',
    'sign-1', 'sign-2', 'body-p2', 'sign-3', 'sign-4',
    'body-p3', 'sign-5', 'sign-6'
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
    // Save all known editable IDs
    EDITABLE_IDS.forEach(id => {
        const el = document.getElementById(id);
        if (el) data[id] = el.innerHTML;
    });
    // Save dynamically added blank pages
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
    // First, remove any existing dynamic blank pages
    document.querySelectorAll('[id^="page-blank-"]').forEach(p => p.remove());

    // Re-create blank pages if any were saved
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
            // Wire up editor sidebar auto-open
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
        // Update blankPageCount so new pages get unique IDs
        blankPageCount = Math.max(blankPageCount, data.__blankPages.length + 2);
    }

    // Now restore content for all editable fields
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
    // Header row
    tableHtml += '<tr>';
    for (let j = 0; j < c; j++) {
        tableHtml += `<th contenteditable="true">Header ${j + 1}</th>`;
    }
    tableHtml += '</tr>';
    // Data rows
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
    // Show toast
    const toast = document.getElementById('autosaveToast');
    if (toast) {
        toast.classList.add('show');
        setTimeout(() => toast.classList.remove('show'), 1500);
    }
}

function checkAutoSaveRestore() {
    const saved = localStorage.getItem('letterhead_autosave');
    if (!saved) return;
    // Show restore banner
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

    if (page3) {
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

    // Auto-save every 15 seconds
    setInterval(autoSave, 15000);
});
