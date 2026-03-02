# Excel Average Calculator

A client-side web application that parses Excel grade files and calculates weighted subject, unit, semester, and overall averages. No server required.

## Live Demo

[https://houssemamor.github.io/Excel-Average-Calculator/](https://houssemamor.github.io/Excel-Average-Calculator/)

## Features

- **Auto-detect columns** -- identifies grade columns by header name regardless of order
- **Grade scale detection** -- supports both 0-20 and 0-2000 scales, normalizes automatically
- **Weighted averages** -- Exam 60%, Coursework (Devoir/Projet/TP) 40%
- **Uncertain grade badges** -- marks subjects missing an exam grade
- **Unit and semester grouping** -- subjects organized by unit with semester extraction
- **Color-coded grades** -- green (>=15), yellow (>=10), red (<10)
- **Drag-and-drop upload** -- or click to browse
- **Responsive design** -- optimized for desktop, tablet, and mobile
- **Dynamic footer** -- shows last commit info via GitHub API
- **Supports .xlsx and .xls** file formats

## Formula

```
Subject Average = (Exam * 0.6) + (Coursework * 0.4)
```

If both Devoir and Projet exist, coursework is split equally (0.2 each). If the exam is missing, available components are averaged with equal weight and flagged as "Uncertain."

## Project Structure

```
Excel-Average-Calculator/
├── index.html          # Entry point, footer with GitHub API commit fetch
├── css/
│   └── styles.css      # All styling, responsive breakpoints at 768px and 480px
├── js/
│   └── script.js       # Excel parsing, grade calculation, HTML rendering
├── .gitignore
└── README.md
```

## Usage

### Online
Visit the [live demo](https://houssemamor.github.io/Excel-Average-Calculator/).

### Local
1. Clone the repository
2. Open `index.html` in a browser
3. Upload an Excel file

### Steps
1. Click "Choose Excel File" or drag and drop
2. The app displays:
   - Subject results with individual component grades
   - Unit averages grouped by unit
   - Semester averages
   - Overall weighted average

## Excel File Format

### Supported Columns (auto-detected, any order)

| Column             | Description                          |
|--------------------|--------------------------------------|
| **Unité**          | Unit/semester name                   |
| **Coef. unité**    | Unit coefficient (optional)          |
| **Unité/Matière**  | Subject name                         |
| **Coef.**          | Subject coefficient                  |
| **DEVOIR**         | Homework grade                       |
| **PROJET**         | Project grade                        |
| **EXAMEN**         | Exam grade                           |
| **TRAVAUX PRATIQUES** / **TP** | Lab/practical grade     |

### Layout Rules
- Row 1: title (optional, ignored)
- Row 2: headers (used for auto-detection)
- Row 3+: data rows

### Scale Handling
If any grade value exceeds 20, the app treats the file as 0-2000 scale and divides all grades by 100 for display on a 0-20 scale.

## Deployment

Hosted on GitHub Pages. Pushes to `main` deploy automatically.

## Technologies

- HTML5, CSS3, Vanilla JavaScript
- [SheetJS (xlsx)](https://github.com/SheetJS/sheetjs) for Excel parsing
- [Font Awesome 6](https://fontawesome.com/) for icons

## Browser Support

- Chrome / Edge (latest)
- Firefox (latest)
- Safari (latest)
- Mobile browsers

## Notes

- All processing is client-side. No data leaves the browser.
- Column detection is case-insensitive with fallback defaults.
- Empty grade cells are handled gracefully (shown as `-`).
- Coefficients are rounded for display; raw values used in calculations.
- Open the browser console (F12) to see detected columns and skipped rows.

## Changelog

### 2026-03-02
- `732e4fa` Add section notes for uncertain grades; enhance UI with inline uncertain badges

### 2026-02-27
- `a13edd3` Enhance README with detailed Excel format instructions; clarify auto-detection and mobile responsiveness
- `f3ecbcb` Fix column order in subject results table; move coefficient column to end
- `8edae5a` Refactor footer layout and enhance responsive design across devices
- `a285e00` Enhance semester data display with bold styling; improve unit name extraction
- `c5f2fe6` Add footer with GitHub link and last commit info; implement auto-detection of column indices

### 2026-02-26
- `27d680a` Update meta tags for improved SEO and application description

### 2026-02-25
- `31dbd7f` Add unit subheader styling and uncertain grade badge; improve average calculation logic
- `59f1210` Enhance usage instructions and clarify deployment process in README
- `13d4839` Add live demo link to README
- `40a541a` Round coefficients and total values for cleaner display; hide upload section after processing
- `9f785b6` Fix grade column order in file processing for accurate calculations
- `bc9b6f2` Implement grade scale detection and normalization (0-20 / 0-2000)
- `d118380` Initial project: HTML, CSS, JavaScript, README, .gitignore

## License

MIT
