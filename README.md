# Excel Average Calculator

A web-based application to calculate subject, unit, and semester averages from Excel files.

## Live Demo

Try it now: [https://houssemamor.github.io/Excel-Average-Calculator/](https://houssemamor.github.io/Excel-Average-Calculator/)

## Features

- **Auto-detect columns** - Automatically identifies grade columns regardless of order
- Upload Excel files with grade data
- Automatic calculation of subject averages using weighted formula
- Group subjects by unit and semester with visual separation
- Display comprehensive statistics with color-coded grades
- **Uncertain grade badges** - Marks subjects without exam grades
- Responsive design optimized for desktop, tablet, and mobile
- No server required - works entirely in the browser
- Dynamic version tracking with last commit info
- Support for 0-20 and 0-2000 grade scales

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

### Supported Columns (Order doesn't matter - auto-detected)
- **Unité** / Unit: Unit/semester name (e.g., "Administration des systèmes/Semestre 1")
- **Unité/Matière** / Subject: Subject/course name
- **Coef** / Coefficient: Subject coefficient (weight)
- **DEVOIR**: Devoir/Homework grade
- **PROJET**: Project grade
- **EXAMEN** / EXAM: Exam grade
- **TRAVAUX PRATIQUES** / TP: Lab/practical work grade

### Important Notes
- Headers should be in row 2
- Data should start from row 3
- The application automatically detects column positions, so any column order works
- Supports both 0-20 and 0-2000 grade scales
- Automatically converts 0-2000 scale to 0-20 for display

### Example Excel Structure
```
Row 1: [Notes] (title, optional)
Row 2: [Unité] [Coef. unité] [Unité/Matière] [Coef.] [DEVOIR] [PROJET] [EXAMEN] [TP] ...
Row 3: [Administration des systèmes/Semestre 1] [4,00] [Administration système Windows] [1,00] [] [15,00] [8,00] [] ...
```

## Deployment

This project is already deployed on GitHub Pages at: [https://houssemamor.github.io/Excel-Average-Calculator/](https://houssemamor.github.io/Excel-Average-Calculator/)

Changes pushed to the `main` branch are automatically deployed.

## Key Features Explained

### Auto-Detect Column Detection
The application intelligently identifies which columns contain grades, coefficients, and subject information. This means your Excel files can have columns in any order.

### Uncertain Grades
When a subject lacks an exam grade, the average is calculated from available components (Devoir, Projet, TP). These are marked with an "Uncertain" badge to indicate the grade might not be complete.

### Grade Scale Detection
Automatically detects whether grades are on a 0-20 or 0-2000 scale and converts accordingly for consistent display.

### Unit and Semester Organization
Subjects are grouped by unit with semester information extracted automatically. Unit names are cleaned up for better readability.

### Mobile-Optimized
- Responsive layouts for all screen sizes
- Horizontal scrolling for data tables on small screens
- Touch-friendly interface
- Optimized performance on mobile devices

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

## Changelog

### Latest Updates
- Auto-detect column indices for flexible Excel formats
- Support for subjects without exam grades (with "Uncertain" badges)
- Improved mobile/tablet responsive design
- Dynamic version tracking with commit messages
- Cleaned unit naming (automatic semester extraction)
- Support for 0-2000 grade scale detection and conversion

## Footer Info

The footer displays:
- Last update date (linked to GitHub commit)
- Current version (based on commit count)
- Latest commit message
- GitHub repository link

## Notes

- All processing is done client-side - no data is sent to any server
- Excel files are not stored - all calculations happen in memory
- Column detection is case-insensitive and flexible
- Empty grade cells are handled gracefully
- Automatic rounding of coefficients for cleaner display
- Console logs (F12 > Console) show detected columns and skipped rows for debugging

## Changelog

### Latest Updates
- Auto-detect column indices for flexible Excel formats
- Support for subjects without exam grades (with "Uncertain" badges)
- Improved mobile/tablet responsive design
- Dynamic version tracking with commit messages
- Cleaned unit naming (automatic semester extraction)
- Support for 0-2000 grade scale detection and conversion
- Supports both .xlsx and .xls file formats
