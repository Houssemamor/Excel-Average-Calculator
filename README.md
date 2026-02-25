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

1. Clone or download this repository
2. Open `index.html` in a web browser
3. Click "Choose Excel File" to upload your grades file
4. The application will process and display:
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

## Deploying to GitHub Pages

1. Create a new repository on GitHub
2. Initialize git in your project folder:
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/your-username/Excel-Average-Calculator.git
   git push -u origin main
   ```

3. Go to your repository settings
4. Navigate to "Pages" section
5. Select "Deploy from a branch"
6. Choose "main" branch and root folder
7. Save - your site will be live at: `https://your-username.github.io/Excel-Average-Calculator`

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
