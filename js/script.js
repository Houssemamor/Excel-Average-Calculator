// Utility to sanitize and parse numeric values
function sanitizeNumber(value) {
    if (value === undefined || value === null || value === '') {
        return NaN;
    }
    const num = parseFloat(String(value).replace(/[¤,\s]/g, '').trim());
    return isNaN(num) ? NaN : num;
}

// Display grade with 2 decimal places or dash
function displayGrade(grade) {
    return isNaN(grade) ? '-' : grade.toFixed(2);
}

// Detect if grades are on 0-2000 scale by checking if any value exceeds 20
function detectGradeScale(gradesArray) {
    for (const grade of gradesArray) {
        if (!isNaN(grade) && grade > 20) {
            return 2000; // 0-2000 scale
        }
    }
    return 20; // 0-20 scale
}

// Convert grade from 0-2000 scale to 0-20 scale
function normalizeGrade(grade, scale) {
    if (isNaN(grade)) return NaN;
    return scale === 2000 ? grade / 100 : grade;
}

// Determine CSS class for grade color coding (using 0-20 scale)
function getGradeClass(grade) {
    if (grade >= 15) return "grade-good";     // 75% and above
    if (grade >= 10) return "grade-average";  // 50-74%
    return "grade-poor";                       // Below 50%
}

// Compute subject average based on exam and coursework components
// Weighting: Exam 60%, Project and Devoir combined 40%
// If Exam is missing, calculate from available components with equal weight
function compute_avg(exam, devoir, projet, tp) {
    const hasExam = !isNaN(exam);
    const hasDevoir = !isNaN(devoir);
    const hasProjet = !isNaN(projet);
    const hasTp = !isNaN(tp);
    
    // If exam exists, use standard weighting (60% exam, 40% coursework)
    if (hasExam) {
        if (hasProjet && hasDevoir) {
            return exam * 0.6 + projet * 0.2 + devoir * 0.2;
        } else if (hasProjet) {
            return exam * 0.6 + projet * 0.4;
        } else if (hasDevoir) {
            return exam * 0.6 + devoir * 0.4;
        } else if (hasTp) {
            return exam * 0.6 + tp * 0.4;
        }
        return exam;
    }
    
    // If exam is missing, calculate from available components with equal weighting
    const availableComponents = [];
    if (hasDevoir) availableComponents.push(devoir);
    if (hasProjet) availableComponents.push(projet);
    if (hasTp) availableComponents.push(tp);
    
    if (availableComponents.length === 0) return NaN;
    
    // Return average of available components
    return availableComponents.reduce((a, b) => a + b, 0) / availableComponents.length;
}

// Extract semester number from unit string
function extractSemester(unit) {
    // Handle both formats: "Name/Semestre 1" and "Name - Semestre 1"
    const match = unit.match(/[\/\-]?\s*[Ss]emestre\s*(\d+)/i);
    return match ? `Semester ${match[1]}` : "Other";
}

// Clean unit name by removing semester suffix
function cleanUnitName(unit) {
    // Remove everything after and including "/" or "-" followed by "Semestre"
    return unit.replace(/\s*[\/\-]?\s*[Ss]emestre\s*\d+.*$/i, '').trim();
}

// Auto-detect column indices from headers
function detectColumnIndices(headerRow) {
    const indices = {
        unit: -1,
        unitCoef: -1,
        subject: -1,
        coef: -1,
        devoir: -1,
        projet: -1,
        exam: -1,
        tp: -1
    };

    // Search headers for column names (case-insensitive)
    for (let i = 0; i < headerRow.length; i++) {
        const header = String(headerRow[i] || '').toLowerCase().trim();

        // More specific matching for unit - must start with "unité" and not include "coef" or "/"
        if (header === 'unité' || (header.startsWith('unité') && !header.includes('coef') && !header.includes('/'))) {
            indices.unit = i;
        }
        if (header.includes('coef') && header.includes('unité') && !header.includes('matière')) indices.unitCoef = i;
        if (header.includes('unité/matière') || header.includes('subject') || header.includes('matière')) indices.subject = i;
        if (header.includes('coef') && !header.includes('unité')) indices.coef = i;
        if (header.includes('devoir')) indices.devoir = i;
        if (header.includes('projet')) indices.projet = i;
        if (header.includes('examen')) indices.exam = i;
        if (header.includes('travaux') || header.includes('pratiques') || header === 'tp') indices.tp = i;
    }

    // Fallback to default indices if detection fails
    if (indices.unit === -1) indices.unit = 0;
    if (indices.subject === -1) indices.subject = 2;
    if (indices.coef === -1) indices.coef = 3;
    if (indices.devoir === -1) indices.devoir = 4;
    if (indices.projet === -1) indices.projet = 5;
    if (indices.exam === -1) indices.exam = 6;
    if (indices.tp === -1) indices.tp = 7;

    console.log('Detected column indices:', indices);
    return indices;
}

// Process and handle Excel file
function processExcelFile(file) {
    if (!file) return;

    const reader = new FileReader();
    const outputDiv = document.getElementById('output');
    outputDiv.innerHTML = '<div style="text-align:center; padding:20px;"><i class="fas fa-spinner fa-spin fa-2x"></i><p>Processing file...</p></div>';

    reader.onload = function (event) {
        processFileData(event.target.result, outputDiv);
    };

    reader.readAsArrayBuffer(file);
}

// Extract and process file data
function processFileData(data, outputDiv) {
    try {
        const uint8Data = new Uint8Array(data);
        const workbook = XLSX.read(uint8Data, { type: 'array' });

        // Find the 'Grades' sheet (fallback to second sheet if not found)
        let sheetName = workbook.SheetNames.find(name =>
            name.toLowerCase().includes('grade')
        );
        if (!sheetName) sheetName = workbook.SheetNames[1];

        const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

        const result = [];
        const unitData = new Map();
        const semesterData = new Map();
        const allGrades = []; // Collect all grades to detect scale

        // Auto-detect column indices from header row
        const colIndices = detectColumnIndices(sheet[1]);

        // First pass: collect all grades to detect scale
        for (let i = 2; i < sheet.length; i++) {
            const row = sheet[i];
            if (!row || row.length < 4) continue;

            allGrades.push(
                sanitizeNumber(row[colIndices.devoir]), // DEVOIR
                sanitizeNumber(row[colIndices.projet]), // PROJET
                sanitizeNumber(row[colIndices.exam]),   // EXAMEN
                sanitizeNumber(row[colIndices.tp])      // TRAVAUX PRATIQUES
            );
        }

        // Detect the grade scale (0-20 or 0-2000)
        const gradeScale = detectGradeScale(allGrades);

        // Second pass: process rows with normalized grades
        for (let i = 2; i < sheet.length; i++) {
            const row = sheet[i];
            if (!row || row.length < 4) continue;

            // Column mappings using auto-detected indices
            let unit = row[colIndices.unit] || '-';
            const name = row[colIndices.subject] || '-';
            const coef = sanitizeNumber(row[colIndices.coef]);
            
            // Skip header rows or rows with header keywords
            if (typeof unit === 'string' && (
                unit.toLowerCase() === 'unité' || 
                unit.toLowerCase().includes('devoir') ||
                unit.toLowerCase().includes('examen')
            )) {
                console.warn(`Skipping header row at ${i + 1}`);
                continue;
            }
            
            const devoir = normalizeGrade(sanitizeNumber(row[colIndices.devoir]), gradeScale);
            const projet = normalizeGrade(sanitizeNumber(row[colIndices.projet]), gradeScale);
            const exam = normalizeGrade(sanitizeNumber(row[colIndices.exam]), gradeScale);
            const tp = normalizeGrade(sanitizeNumber(row[colIndices.tp]), gradeScale);

            // Skip rows with invalid coefficient or no grades at all
            if (isNaN(coef)) {
                console.warn(`Skipping row ${i + 1}: Unit="${cleanUnitName(unit)}", Name="${name}" - Invalid coefficient`);
                continue;
            }
            
            // Check if at least one grade component exists
            const hasAnyGrade = !isNaN(exam) || !isNaN(devoir) || !isNaN(projet) || !isNaN(tp);
            if (!hasAnyGrade) {
                console.warn(`Skipping row ${i + 1}: Unit="${cleanUnitName(unit)}", Name="${name}" - No grade components`);
                continue;
            }

            const avg = compute_avg(exam, devoir, projet, tp);
            const weighted = avg * coef;
            const semester = extractSemester(unit);
            const cleanUnit = cleanUnitName(unit); // Clean unit name for display
            const hasExamGrade = !isNaN(exam); // Track if exam grade exists

            result.push({ unit: cleanUnit, name, exam, devoir, projet, tp, coef, avg, weighted, semester, hasExamGrade });

            // Accumulate unit statistics
            if (!unitData.has(cleanUnit)) {
                unitData.set(cleanUnit, { totalWeighted: 0, totalCoef: 0, semester });
            }

            const unitInfo = unitData.get(cleanUnit);
            unitInfo.totalWeighted += weighted;
            unitInfo.totalCoef += coef;

            // Accumulate semester statistics
            if (!semesterData.has(semester)) {
                semesterData.set(semester, { totalWeighted: 0, totalCoef: 0 });
            }

            const semesterInfo = semesterData.get(semester);
            semesterInfo.totalWeighted += weighted;
            semesterInfo.totalCoef += coef;
        }

        // Generate output HTML
        let html = '<h2><i class="fas fa-book"></i> Subject Results</h2>';
        //html += `<p style="color: #666; font-size: 0.9rem; margin-bottom: 15px;">Grade Scale: 0-20 | File detected: ${gradeScale === 2000 ? '0-2000 (converted)' : '0-20'}</p>`;
        html += '<table>';
        html += '<tr><th>Subject</th><th>DEVOIR</th><th>TP</th><th>PROJET</th><th>EXAMEN</th><th>Coef.</th><th>Average</th><th>Weighted</th></tr>';

        let totalWeighted = 0;
        let totalCoef = 0;

        // Group results by unit for visual division
        const unitGroups = new Map();
        result.forEach(row => {
            if (!unitGroups.has(row.unit)) {
                unitGroups.set(row.unit, []);
            }
            unitGroups.get(row.unit).push(row);
        });

        // Output subject results table with unit dividers
        let unitIndex = 0;
        unitGroups.forEach((rows, unit) => {
            // Add unit subheader
            const unitBgClass = unitIndex % 2 === 0 ? 'unit-bg-light' : 'unit-bg-dark';
            html += `<tr class="unit-subheader ${unitBgClass}">
          <td colspan="8"><strong><i class="fas fa-folder-open"></i> ${unit}</strong></td>
        </tr>`;
            
            // Add subject rows for this unit
            rows.forEach(row => {
                const avgDisplay = row.hasExamGrade ? 
                    row.avg.toFixed(2) : 
                    `${row.avg.toFixed(2)} <span class="uncertain-badge"><i class="fas fa-exclamation-circle"></i> Uncertain</span>`;
                
                html += `<tr class="${unitBgClass}">
          <td>${row.name}</td>
          <td class="subject-grade ${getGradeClass(row.devoir)}">${displayGrade(row.devoir)}</td>
          <td class="subject-grade ${getGradeClass(row.tp)}">${displayGrade(row.tp)}</td>
          <td class="subject-grade ${getGradeClass(row.projet)}">${displayGrade(row.projet)}</td>
          <td class="subject-grade ${getGradeClass(row.exam)}">${displayGrade(row.exam)}</td>
          <td>${Math.round(row.coef/100)}</td>
          <td class="subject-grade ${getGradeClass(row.avg)}">${avgDisplay}</td>
          <td>${row.weighted.toFixed(2)/100}</td>
        </tr>`;

                totalWeighted += row.weighted;
                totalCoef += row.coef;
            });
            
            unitIndex++;
        });

        html += '</table>';

        // Unit averages section
        html += '<h2><i class="fas fa-layer-group"></i> Unit Averages</h2>';
        html += '<table><tr><th>Unit</th><th>Semester</th><th>Average</th><th>Total Coef</th></tr>';

        unitData.forEach((data, unit) => {
            const unitAvg = data.totalWeighted / data.totalCoef;
            html += `<tr>
          <td>${unit}</td>
          <td>${data.semester}</td>
          <td class="unit-grade ${getGradeClass(unitAvg)}">${unitAvg.toFixed(2)}</td>
          <td>${Math.round(data.totalCoef/100)}</td>
        </tr>`;
        });
        html += '</table>';

        // Semester averages section
        html += '<h2><i class="fas fa-calendar-alt"></i> Semester Averages</h2>';
        semesterData.forEach((data, semester) => {
            const semesterAvg = data.totalWeighted / data.totalCoef;
            html += `<div class="semester-card">
          <h3>${semester}</h3>
          <table>
            <tr><th>Total Weighted</th><th>Total Coefficient</th><th>Semester Average</th></tr>
            <tr>
              <td class="semester-total-weighted">${data.totalWeighted.toFixed(2)/100}</td>
              <td class="semester-total-coef">${Math.round(data.totalCoef/100)}</td>
              <td class="unit-grade ${getGradeClass(semesterAvg)}">${semesterAvg.toFixed(2)}</td>
            </tr>
          </table>
        </div>`;
        });

        // Overall average section
        const generalAvg = totalWeighted / totalCoef;
        html += `<div class="overall-average">
        <h3>Overall Average</h3>
        <div class="average-value ${getGradeClass(generalAvg)}">${generalAvg.toFixed(2)}</div>
        <p>Based on ${result.length} subjects | Total Coefficient: ${Math.round(totalCoef/100)}</p>
      </div>`;

        // Store scale info in the output for reference
        outputDiv.dataset.gradeScale = gradeScale;
        outputDiv.innerHTML = html;

        // Hide upload section and scroll to results
        document.querySelector('.upload-container').style.display = 'none';
        document.querySelector('.instructions').style.display = 'none';
        outputDiv.scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (error) {
        outputDiv.innerHTML = `<div class="error">
        <i class="fas fa-exclamation-triangle fa-3x"></i>
        <h2>Error Processing File</h2>
        <p>${error.message}</p>
        <p>Please ensure you're uploading a valid Excel file with the correct format.</p>
      </div>`;
    }
}

// Main file upload handler via input
document.getElementById('upload').addEventListener('change', function (e) {
    const file = e.target.files[0];
    processExcelFile(file);
});

// Drag and drop support
const uploadContainer = document.querySelector('.upload-container');
const uploadInput = document.getElementById('upload');

// Prevent default drag behaviors
document.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.stopPropagation();
});

document.addEventListener('drop', (e) => {
    e.preventDefault();
    e.stopPropagation();
});

// Highlight drop zone on drag over
uploadContainer.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.stopPropagation();
    uploadContainer.classList.add('dragover');
});

uploadContainer.addEventListener('dragleave', (e) => {
    e.preventDefault();
    e.stopPropagation();
    uploadContainer.classList.remove('dragover');
});

// Handle dropped files
uploadContainer.addEventListener('drop', (e) => {
    e.preventDefault();
    e.stopPropagation();
    uploadContainer.classList.remove('dragover');

    const files = e.dataTransfer.files;
    const file = files[0];

    // Validate file type
    if (file && (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
        file.type === 'application/vnd.ms-excel' ||
        file.name.endsWith('.xlsx') ||
        file.name.endsWith('.xls'))) {
        processExcelFile(file);
    } else {
        alert('Please drop a valid Excel file (.xlsx or .xls)');
    }
});
