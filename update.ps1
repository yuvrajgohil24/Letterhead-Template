$file = 'c:\Users\acer\Desktop\Letterhead-Template\index.html'
$html = Get-Content $file -Raw

# 1. Add toggle button
$html = $html -replace '(<select id="pageSelect" class="page-select">)', '<button class="btn-primary" onclick="addBlankPage()" style="margin-right: 15px; padding: 12px 20px; font-weight: 600; font-size: 14px; border-radius: 6px; cursor: pointer; border: none; box-shadow: 0 4px 15px rgba(30, 42, 79, 0.3);">➕ Add Blank Page</button>`n            $1'

# 2. Fix Download PDF dropdown
$html = $html -replace '<option value="page1-2">Download 2 Pages \(1 & 2\)</option>', ''

# 3. Replace Page 2 completely
$startStr = '<!-- ================= PAGE 2 ================= -->'
$endStr = '<!-- ================= PAGE 3 ================= -->'
$startIdx = $html.IndexOf($startStr)
$endIdx = $html.IndexOf($endStr)

if ($startIdx -ge 0 -and $endIdx -gt $startIdx) {
    $before = $html.Substring(0, $startIdx)
    $after = $html.Substring($endIdx)
    $newPage2 = @"
        <!-- ================= PAGE 2 (BLANK) ================= -->
        <div class="letterhead page page-break blank-page" id="page2">
            <div class="blank-page-content" contenteditable="true" id="body-p2">
                <p>Click here to edit this blank page...</p>
                <br>
                <p>Use "Add Blank Page" button at the top to add more pages between the first and last cover pages.</p>
            </div>
        </div>

"@
    $html = $before + $newPage2 + $after
}

# 4. Modify download PDF pageSelect logic
$oldLogic = @"
            // Strategy: Remove unwanted pages from DOM temporarily
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
"@
$newLogic = @"
            let element;
            let removedPage = null;
            if (pageSelect === 'page1') {
                element = document.getElementById('page1');
            } else {
                element = pdfRoot;
            }
"@
$html = $html.Replace($oldLogic, $newLogic)

# 5. Add addBlankPage function at the end
$scriptEnd = @"
        // Initialize everything
        document.addEventListener('DOMContentLoaded', () => {
            renderDraftList();
            initAutoSidebar();
        });
"@
$scriptNew = @"
        // Initialize everything
        document.addEventListener('DOMContentLoaded', () => {
            renderDraftList();
            initAutoSidebar();
        });

        // Add blank page logic
        let blankPageCount = 2;
        function addBlankPage() {
            blankPageCount++;
            const pdfRoot = document.getElementById('pdf-root');
            const page3 = document.getElementById('page3'); 
            
            const newPage = document.createElement('div');
            newPage.className = 'letterhead page page-break blank-page';
            newPage.id = 'page-blank-' + blankPageCount;
            
            newPage.innerHTML = `
                <div class="blank-page-content" contenteditable="true" id="body-blank-` + '${blankPageCount}' + `">
                    <p>Click here to edit new blank page...</p>
                </div>
                <button onclick="this.parentElement.remove()" style="position: absolute; top: 15px; right: 20px; background: #ffefef; color: #d32f2f; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 12px; z-index: 10;">🗑️ Remove Page</button>
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
            
            if(!EDITABLE_IDS.includes('body-blank-' + blankPageCount)) {
                EDITABLE_IDS.push('body-blank-' + blankPageCount);
            }
        }
"@

$html = $html.Replace($scriptEnd, $scriptNew)

$html | Set-Content $file -Encoding UTF8
