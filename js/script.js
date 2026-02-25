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
function compute_avg(exam, devoir, projet, tp) {
    if (!isNaN(projet) && !isNaN(devoir)) {
        return exam * 0.6 + projet * 0.2 + devoir * 0.2;
    } else if (!isNaN(projet)) {
        return exam * 0.6 + projet * 0.4;
    } else if (!isNaN(devoir)) {
        return exam * 0.6 + devoir * 0.4;
    } else if (!isNaN(tp)) {
        return exam * 0.6 + tp * 0.4;
    }
    return exam;
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
                sanitizeNumber(row[5]), // EXAMEN
                sanitizeNumber(row[6]), // PROJET
                sanitizeNumber(row[7])  // TP
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
            const exam = normalizeGrade(sanitizeNumber(row[5]), gradeScale);   // Column F: EXAMEN
            const projet = normalizeGrade(sanitizeNumber(row[6]), gradeScale); // Column G: PROJET
            const tp = normalizeGrade(sanitizeNumber(row[7]), gradeScale);     // Column H: TP

            // Skip rows with invalid exam grade or coefficient
            if (isNaN(exam) || isNaN(coef)) continue;

            const avg = compute_avg(exam, devoir, projet, tp);
            const weighted = avg * coef;
            const semester = extractSemester(unit);

            result.push({ unit, name, exam, devoir, projet, tp, coef, avg, weighted, semester });

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

        // Output subject results table
        result.forEach(row => {
            html += `<tr>
          <td>${row.name}</td>
          <td class="subject-grade ${getGradeClass(row.exam)}">${displayGrade(row.exam)}</td>
          <td class="subject-grade ${getGradeClass(row.devoir)}">${displayGrade(row.devoir)}</td>
          <td class="subject-grade ${getGradeClass(row.projet)}">${displayGrade(row.projet)}</td>
          <td class="subject-grade ${getGradeClass(row.tp)}">${displayGrade(row.tp)}</td>
          <td>${row.coef}</td>
          <td class="subject-grade ${getGradeClass(row.avg)}">${row.avg.toFixed(2)}</td>
          <td>${row.weighted.toFixed(2)}</td>
        </tr>`;

            totalWeighted += row.weighted;
            totalCoef += row.coef;
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
          <td>${data.totalCoef}</td>
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
              <td>${data.totalCoef}</td>
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
        <p>Based on ${result.length} subjects | Total Coefficient: ${totalCoef}</p>
      </div>`;

        // Store scale info in the output for reference
        outputDiv.dataset.gradeScale = gradeScale;
        outputDiv.innerHTML = html;
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
