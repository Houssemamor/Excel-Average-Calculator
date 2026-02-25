# Excel Average Calculator

A web-based application to calculate subject, unit, and semester averages from Excel files.

## Live Demo

Try it now: [https://houssemamor.github.io/Excel-Average-Calculator/](https://houssemamor.github.io/Excel-Average-Calculator/)

## Features

- Upload Excel files with grade data
- Automatic calculation of subject averages using weighted formula
- Group subjects by unit and semester
- Display comprehensive statistics with color-coded grades
- Responsive design for mobile and desktop
- No server required - works entirely in the browser

## Formula

Subject averages are calculated using:
```
Average = (Exam × 60%) + (Coursework × 40%)
```

Where coursework includes Devoir, Projet, and TP components.

## Project Structure

```
Excel-Average-Calculator/
├── index.html              # Main HTML file
├── css/
│   └── styles.css         # Stylesheet
├── js/
│   └── script.js          # JavaScript logic
├── .gitignore             # Git ignore rules
└── README.md              # This file
```

## How to Use

### Online (No Installation Required)
Simply visit the live site: [https://houssemamor.github.io/Excel-Average-Calculator/](https://houssemamor.github.io/Excel-Average-Calculator/)

### Local Setup
1. Clone or download this repository
2. Open `index.html` in a web browser
3. Upload your Excel file and view results

### Usage Steps
1. Click "Choose Excel File" or drag and drop your grades file
2. The application will process and display:
   - Subject results with individual grades
   - Unit averages grouped by unit
   - Semester averages
   - Overall average

## Excel File Format

The Excel file should contain the following columns:
- Column A: Unit (e.g., "Semestre 1")
- Column B: (Reserved)
- Column C: Subject Name
- Column D: Coefficient
- Column E: DEVOIR
- Column F: EXAMEN
- Column G: PROJET
- Column H: TRAVAUX PRATIQUES

Data should start from row 3 (rows 1-2 are headers).

## Deployment

This project is already deployed on GitHub Pages at: [https://houssemamor.github.io/Excel-Average-Calculator/](https://houssemamor.github.io/Excel-Average-Calculator/)

Changes pushed to the `main` branch are automatically deployed.

## Technologies Used

- HTML5
- CSS3 with responsive design
- Vanilla JavaScript
- XLSX library for Excel parsing
- Font Awesome icons

## Browser Support

- Chrome/Edge (latest)
- Firefox (latest)
- Safari (latest)
- Mobile browsers

## License

MIT License - Feel free to use this project for your needs.

## Notes

- All processing is done client-side - no data is sent to any server
- Excel files are not stored - all calculations happen in memory
- Supports both .xlsx and .xls file formats
