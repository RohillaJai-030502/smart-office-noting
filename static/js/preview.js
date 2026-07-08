// ============================================================
// preview.js — Universal Offline Live Noting Preview Engine
// ============================================================

document.addEventListener("DOMContentLoaded", () => {
    const formContainer = document.querySelector(".form-container");
    const hasPreviewLayout = document.querySelector(".preview-layout-container");
    const isPasswordPage = window.location.pathname.includes("password");
    
    // Only execute on drafting form pages that don't already have a manual preview layout
    if (formContainer && !hasPreviewLayout && !isPasswordPage) {
        initializeUniversalPreview(formContainer);
    }
});

function initializeUniversalPreview(formContainer) {
    // 1. Determine template context (Noting vs Fax)
    const pageTitle = document.querySelector("h2")?.textContent || "";
    const isFax = pageTitle.toLowerCase().includes("fax") || pageTitle.toLowerCase().includes("email") || pageTitle.toLowerCase().includes("मेल") || pageTitle.toLowerCase().includes("फैक्स");
    
    // 2. Restructure DOM into split-screen layout
    const wrapper = document.createElement("div");
    wrapper.className = "preview-layout-container";
    formContainer.parentNode.insertBefore(wrapper, formContainer);
    
    const formSide = document.createElement("div");
    formSide.className = "form-side";
    const previewSide = document.createElement("div");
    previewSide.className = "preview-side";
    
    wrapper.appendChild(formSide);
    wrapper.appendChild(previewSide);
    
    // Move the original form container into the left side
    formSide.appendChild(formContainer);
    
    // 3. Inject A4 live preview panel on the right
    previewSide.innerHTML = `
        <div style="margin-bottom: 0.75rem; display: flex; justify-content: space-between; align-items: center;">
          <span style="font-weight: 700; font-size: 0.9rem; color: var(--primary-navy); text-transform: uppercase; letter-spacing: 0.5px;">Live Document Preview</span>
          <span style="font-size: 0.75rem; color: #2e7d32; font-weight: bold; background: #e8f5e9; padding: 2px 8px; border-radius: 10px;">Offline Mode</span>
        </div>
        <div class="noting-sheet-preview" id="global-preview-sheet" style="${isFax ? 'border-left: none; padding: 0.75in;' : ''}">
          
          <!-- Document Header -->
          <div class="noting-header">
            TECHNICAL BALLISTICS RESEARCH LABORATORY<br>
            <span style="font-size: 11pt; font-weight: normal; text-decoration: none;" id="prev-sub-module">HRD DIVISION</span><br>
            <span style="font-size: 12pt;" id="prev-doc-type">${isFax ? 'FAX MESSAGE' : 'INTER-OFFICE NOTE'}</span>
          </div>
          
          <!-- Meta Block -->
          <div class="noting-meta">
            <div>Ref No: <span id="prev-ref-no" style="font-weight: normal;">—</span></div>
            <div>Date: <span id="prev-date" style="font-weight: normal;">—</span></div>
          </div>
          
          <!-- Subject -->
          <div class="noting-subject">
            <span class="label">Subject:</span>
            <span id="prev-subject" style="font-weight: bold;">—</span>
          </div>
          
          <!-- Noting Body -->
          <div class="noting-body" id="prev-body-area">
            <!-- Dynamic variables summary or template text -->
          </div>
          
          <!-- Tables Container -->
          <div class="noting-table-container" id="prev-tables-area">
            <!-- Dynamic tables cloned here -->
          </div>
          
          <!-- Signatures Footer -->
          <div class="noting-signatures" id="prev-sig-area">
            <div class="sig-block" id="prev-sig-1">
              (<span id="prev-sig-name-1">—</span>)<br>
              <span class="desig" id="prev-sig-desig-1">—</span>
            </div>
            <div class="sig-block" id="prev-sig-2" style="display: none; margin-top: 1rem;">
              (<span id="prev-sig-name-2">—</span>)<br>
              <span class="desig" id="prev-sig-desig-2">—</span>
            </div>
          </div>
          
        </div>
    `;
    
    // Hide sub-module header for fax foms
    if (isFax) {
        const subModuleEl = document.getElementById("prev-sub-module");
        if (subModuleEl) subModuleEl.style.display = "none";
    }
    
    // 4. Setup change and input observers on the form
    const form = formSide.querySelector("form");
    if (form) {
        form.addEventListener("input", updateUniversalPreview);
        form.addEventListener("change", updateUniversalPreview);
        
        // Also observe dynamic table insertions/deletions
        const observer = new MutationObserver(() => {
            updateUniversalPreview();
            // Bind inputs in newly added table rows too
            form.querySelectorAll("table input, table select").forEach(input => {
                input.removeEventListener("input", updateUniversalPreview);
                input.addEventListener("input", updateUniversalPreview);
            });
        });
        
        form.querySelectorAll("table").forEach(table => {
            observer.observe(table, { childList: true, subtree: true });
        });
    }
    
    // 5. Initial rendering
    updateUniversalPreview();
    
    // 6. Bind automatic Hindi transliteration to matching Eng/Hindi pairs globally
    autoBindTransliteration();
}

function updateUniversalPreview() {
    // A. Reference and Date Mappings
    let refNo = document.getElementById("form_no")?.value || 
                document.getElementById("REF_NO")?.value || 
                document.getElementById("ref_no")?.value || "—";
    
    let rawDate = document.getElementById("ref_date")?.value || 
                  document.getElementById("DATE")?.value || 
                  document.getElementById("ref_date_1")?.value || "";
    
    let formattedDate = rawDate || "—";
    // Check if it is standard HTML date picker string
    if (/^\d{4}-\d{2}-\d{2}$/.test(rawDate)) {
        const d = new Date(rawDate);
        const day = String(d.getDate()).padStart(2, '0');
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const year = d.getFullYear();
        formattedDate = `${day}/${month}/${year}`;
    }
    
    document.getElementById("prev-ref-no").textContent = refNo;
    document.getElementById("prev-date").textContent = formattedDate;
    
    // B. Subject Mapping
    let subHindi = document.getElementById("subject_hindi")?.value || 
                   document.getElementById("SUBJECT_HINDI")?.value || "";
    let subEng = document.getElementById("subject_english")?.value || 
                 document.getElementById("SUBJECT_ENGLISH")?.value || "";
    let subGeneral = document.getElementById("subject")?.value || 
                     document.getElementById("SUBJECT")?.value || "";
    
    let subjectText = subGeneral;
    if (subHindi || subEng) {
        subjectText = subHindi;
        if (subEng) {
            subjectText += (subjectText ? " / " : "") + subEng;
        }
    }
    document.getElementById("prev-subject").textContent = subjectText || "—";
    
    // C. Signatories Mappings
    let sigName1 = document.getElementById("signatory_1_name")?.value || 
                    document.getElementById("SIGNATORY_NAME")?.value || 
                    document.getElementById("sig_name")?.value || "—";
    let sigDesig1 = document.getElementById("signatory_1_desig")?.value || 
                     document.getElementById("SIGNATORY_DESIG")?.value || 
                     document.getElementById("sig_desig")?.value || "—";
                     
    let sigName2 = document.getElementById("signatory_2_name")?.value || "";
    let sigDesig2 = document.getElementById("signatory_2_desig")?.value || "";
    
    document.getElementById("prev-sig-name-1").textContent = sigName1;
    document.getElementById("prev-sig-desig-1").textContent = sigDesig1;
    
    const sigBlock2 = document.getElementById("prev-sig-2");
    if (sigName2) {
        sigBlock2.style.display = "block";
        document.getElementById("prev-sig-name-2").textContent = sigName2;
        document.getElementById("prev-sig-desig-2").textContent = sigDesig2;
    } else {
        sigBlock2.style.display = "none";
    }
    
    // D. Replicate Dynamic Tables in Preview
    const tablesArea = document.getElementById("prev-tables-area");
    tablesArea.innerHTML = "";
    
    const formTables = document.querySelectorAll(".form-side table");
    formTables.forEach((table) => {
        const previewTable = document.createElement("table");
        previewTable.className = "noting-table";
        
        // Headers
        const thead = document.createElement("thead");
        const headerRow = document.createElement("tr");
        const formHeaders = table.querySelectorAll("th");
        formHeaders.forEach(th => {
            const previewTh = document.createElement("th");
            // Check if cell header itself has input (e.g. custom noting table headers)
            const input = th.querySelector("input");
            previewTh.textContent = input ? input.value : th.textContent.trim();
            headerRow.appendChild(previewTh);
        });
        thead.appendChild(headerRow);
        previewTable.appendChild(thead);
        
        // Body rows
        const tbody = document.createElement("tbody");
        const formRows = table.querySelectorAll("tbody tr");
        formRows.forEach(tr => {
            // Skip helper / template rows if they are hidden
            if (tr.style.display === "none" || tr.classList.contains("hidden")) return;
            
            const previewTr = document.createElement("tr");
            const cells = tr.querySelectorAll("td");
            cells.forEach(td => {
                const previewTd = document.createElement("td");
                
                // Read input, textarea, select, or plaintext value
                const input = td.querySelector("input[type='text'], input[type='number']");
                const textarea = td.querySelector("textarea");
                const select = td.querySelector("select");
                
                if (input) {
                    previewTd.textContent = input.value;
                    if (input.style.textAlign === "center") {
                        previewTd.className = "center";
                    }
                } else if (textarea) {
                    previewTd.textContent = textarea.value;
                } else if (select) {
                    previewTd.textContent = select.value;
                } else {
                    // Check if there is plain text inside
                    previewTd.textContent = td.textContent.trim();
                }
                previewTr.appendChild(previewTd);
            });
            tbody.appendChild(previewTr);
        });
        previewTable.appendChild(tbody);
        tablesArea.appendChild(previewTable);
    });
    
    // E. Variables Summary List (for forms without tables)
    const bodyArea = document.getElementById("prev-body-area");
    bodyArea.innerHTML = "";
    
    // Only generate variables summary list if there are no dynamic tables
    if (formTables.length === 0) {
        const formInputs = document.querySelectorAll(".form-side .form-group input, .form-side .form-group textarea, .form-side .form-group select");
        let listHTML = `<div style="font-weight: bold; border-bottom: 1px solid #ccc; padding-bottom: 4px; margin-bottom: 10px; font-size: 10.5pt; text-transform: uppercase;">Variables Summary / विवरण सारांश</div>`;
        listHTML += `<table style="width:100%; border-collapse:collapse; font-size:10pt;">`;
        
        const bilingMap = {
            "REF_NO": "Reference No / संदर्भ संख्या",
            "REF_DATE": "Date / दिनांक",
            "DATE": "Date / दिनांक",
            "START_DATE": "Start Date / आरंभ तिथि",
            "END_DATE": "End Date / समाप्ति तिथि",
            "ORG_INSTITUTE": "Institute Name / संस्थान का नाम",
            "COLLEGE_NAME": "Institute Name / संस्थान का नाम",
            "INSTITUTE_NAME": "Institute Name / संस्थान का नाम",
            "COURSE_TITLE": "Title / शीर्षक",
            "LECTURE_TITLE": "Title / शीर्षक",
            "PROJECT_TITLE": "Title / शीर्षक",
            "COURSE_TYPE": "Course Type / पाठ्यक्रम का प्रकार",
            "GROUP_NAME": "Group Name / समूह का नाम",
            "FEE_AMOUNT": "Fee Amount / शुल्क राशि",
            "BENEFICIARY_NAME": "Beneficiary / लाभार्थी का नाम",
            "BUDGET_HEAD": "Budget Head / बजट शीर्ष",
            "ORIGINAL_REF_NO": "Original Ref / मूल संदर्भ सं",
            "REASON_CANCELLATION": "Reason / निरस्तीकरण का कारण",
            "SPEAKER_NAME": "Speaker Name / वक्ता का नाम",
            "VENUE": "Venue / स्थान",
            "STUDENT_NAME_ENG": "Student Name / छात्र का नाम",
            "STUDENT_NAME": "Student Name / छात्र का नाम",
            "FATHER_NAME_ENG": "Father's Name / पिता का नाम",
            "FATHER_NAME": "Father's Name / पिता का नाम",
            "MOBILE_NO": "Mobile No / मोबाइल नंबर",
            "DATE_OF_JOINING": "Date of Joining / शामिल होने की तिथि",
            "DATE_OF_COMPLETION": "Date of Completion / पूर्ण होने की तिथि",
            "ATTENDANCE_GRADE": "Attendance Grade / उपस्थिति ग्रेड",
            "PERFORMANCE_RATING": "Performance Rating / प्रदर्शन रेटिंग",
            "REPORT_SUBMITTED": "Report Submitted / रिपोर्ट प्रस्तुत",
            "INTERNSHIP_DURATION": "Internship Duration / इंटर्नशिप अवधि",
            "LAST_DATE": "Last Date / अंतिम तिथि",
            "STIPEND_RATE": "Stipend Rate / स्टाइपेंड दर",
            "MENTOR_NAME_DESIG": "Mentor Details / मेंटर विवरण",
            "VISIT_DATE": "Visit Date / यात्रा की तिथि",
            "EMAIL_DATE": "Email Date / ईमेल की तिथि",
            "MENTOR_EMAIL": "Mentor Email / मेंटर ईमेल",
            "TRAINING_PERIOD": "Training Period / प्रशिक्षण अवधि",
            "TO_TEXT": "To (Recipient) / सेवा में",
            "FAX_NO": "Fax No / फैक्स संख्या",
            "REF_TEXT": "Reference Text / संदर्भ पाठ"
        };
        
        let hasVars = false;
        formInputs.forEach(input => {
            const id = input.id || input.name;
            const val = input.value || "—";
            
            // Skip metadata and button inputs
            if (["REF_NO", "form_no", "DATE", "ref_date", "template_id", "subject_hindi", "subject_english", "SUBJECT_HINDI", "SUBJECT_ENGLISH", "subject", "SUBJECT", "SIGNATORY_NAME", "SIG_NAME", "SIGNATORY_DESIG", "SIG_DESIG", "signatory_1_name", "signatory_1_desig", "signatory_2_name", "signatory_2_desig"].includes(id)) {
                return;
            }
            
            hasVars = true;
            const cleanLabel = bilingMap[id.toUpperCase()] || id.replace(/_/g, ' ');
            listHTML += `
              <tr style="border-bottom:1px solid #eee;">
                <td style="width:42%; padding:4px 0; font-weight:bold; color:#555; font-size:9.5pt;">${cleanLabel}:</td>
                <td style="padding:4px 0; font-size:9.5pt;">${val.replace(/\n/g, '<br>')}</td>
              </tr>
            `;
        });
        
        listHTML += `</table>`;
        if (hasVars) {
            bodyArea.innerHTML = listHTML;
        }
    }
}

// Automatically bind transliterating pairs across the form
function autoBindTransliteration() {
    // English/Hindi signatory names
    setupTransliteration("signatory_1_name", "signatory_1_name_hindi", "auto-translit-check");
    setupTransliteration("signatory_2_name", "signatory_2_name_hindi", "auto-translit-check");
    
    // Subject lines
    setupTransliteration("subject_english", "subject_hindi", "auto-translit-check");
    setupTransliteration("SUBJECT_ENGLISH", "SUBJECT_HINDI", "auto-translit-check");
    
    // Any English field mapping to a Hindi field
    const englishInputs = document.querySelectorAll("input[id*='eng'], input[id*='ENG']");
    englishInputs.forEach(engInput => {
        const baseId = engInput.id.replace(/_eng/i, "");
        const hindiEl = document.getElementById(baseId + "_hindi") || document.getElementById(baseId + "_HINDI");
        if (hindiEl) {
            setupTransliteration(engInput.id, hindiEl.id, "auto-translit-check");
        }
    });
}
