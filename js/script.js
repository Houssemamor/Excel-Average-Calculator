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
    const match = unit.match(/semestre\s*(\d+)/i);
    return match ? `Semester ${match[1]}` : "Other";
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

        // First pass: collect all grades to detect scale
        for (let i = 2; i < sheet.length; i++) {
            const row = sheet[i];
            if (!row || row.length < 10) continue;

            allGrades.push(
                sanitizeNumber(row[4]), // DEVOIR
                sanitizeNumber(row[5]), // PROJET
                sanitizeNumber(row[6]), // EXAMEN
                sanitizeNumber(row[7])  // TRAVAUX PRATIQUES
            );
        }

        // Detect the grade scale (0-20 or 0-2000)
        const gradeScale = detectGradeScale(allGrades);

        // Second pass: process rows with normalized grades
        for (let i = 2; i < sheet.length; i++) {
            const row = sheet[i];
            if (!row || row.length < 10) continue;

            // Column mappings
            const unit = row[0] || '-';          // Column A: Unit
            const name = row[2] || '-';          // Column C: Subject name
            const coef = sanitizeNumber(row[3]); // Column D: Coefficient
            const devoir = normalizeGrade(sanitizeNumber(row[4]), gradeScale); // Column E: DEVOIR
            const projet = normalizeGrade(sanitizeNumber(row[5]), gradeScale); // Column F: PROJET
            const exam = normalizeGrade(sanitizeNumber(row[6]), gradeScale);   // Column G: EXAMEN
            const tp = normalizeGrade(sanitizeNumber(row[7]), gradeScale);     // Column H: TRAVAUX PRATIQUES

            // Skip rows with invalid coefficient or no grades at all
            if (isNaN(coef)) {
                console.warn(`Skipping row ${i + 1}: Unit="${unit}", Name="${name}" - Invalid coefficient`);
                continue;
            }
            
            // Check if at least one grade component exists
            const hasAnyGrade = !isNaN(exam) || !isNaN(devoir) || !isNaN(projet) || !isNaN(tp);
            if (!hasAnyGrade) {
                console.warn(`Skipping row ${i + 1}: Unit="${unit}", Name="${name}" - No grade components`);
                continue;
            }

            const avg = compute_avg(exam, devoir, projet, tp);
            const weighted = avg * coef;
            const semester = extractSemester(unit);
            const hasExamGrade = !isNaN(exam); // Track if exam grade exists

            result.push({ unit, name, exam, devoir, projet, tp, coef, avg, weighted, semester, hasExamGrade });

            // Accumulate unit statistics
            if (!unitData.has(unit)) {
                unitData.set(unit, { totalWeighted: 0, totalCoef: 0, semester });
            }

            const unitInfo = unitData.get(unit);
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
        html += `<p style="color: #666; font-size: 0.9rem; margin-bottom: 15px;">Grade Scale: 0-20 | File detected: ${gradeScale === 2000 ? '0-2000 (converted)' : '0-20'}</p>`;
        html += '<table>';
        html += '<tr><th>Subject</th><th>EXAMEN</th><th>DEVOIR</th><th>PROJET</th><th>TP</th><th>Coef.</th><th>Average</th><th>Weighted</th></tr>';

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
          <td class="subject-grade ${getGradeClass(row.exam)}">${displayGrade(row.exam)}</td>
          <td class="subject-grade ${getGradeClass(row.devoir)}">${displayGrade(row.devoir)}</td>
          <td class="subject-grade ${getGradeClass(row.projet)}">${displayGrade(row.projet)}</td>
          <td class="subject-grade ${getGradeClass(row.tp)}">${displayGrade(row.tp)}</td>
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
              <td>${data.totalWeighted.toFixed(2)}</td>
              <td>${Math.round(data.totalCoef/100)}</td>
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
