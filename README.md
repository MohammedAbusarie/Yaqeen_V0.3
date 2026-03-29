# 🎓 Yaqeen Attendance Tool

<div align="center">

**Automated attendance & grades input • Time-saving • Privacy-first**

*Created by Mohammed Abusarie*

[![Static Site](https://img.shields.io/badge/Static%20Site-Yes-brightgreen)](https://netlify.com)
[![No Backend](https://img.shields.io/badge/Backend-None-success)](https://netlify.com)
[![Privacy First](https://img.shields.io/badge/Privacy-First-blue)](https://netlify.com)
[![Browser Only](https://img.shields.io/badge/Processing-Browser%20Only-orange)](https://netlify.com)

</div>

---

## 🌟 Overview

**Yaqeen Attendance Tool** is a sophisticated, client-side web application that automates the tedious process of marking attendance and inputting grades in Excel spreadsheets. Built entirely with vanilla JavaScript (ES Modules), this tool runs **100% in your browser**—no server, no backend, no data uploads. Your files never leave your device, ensuring complete privacy and security.

### 🎯 What Makes This Special?

This isn't just another attendance tool. It's a showcase of **advanced automation** and **complex programming techniques**:

- 🤖 **Intelligent Automation**: Automatically detects columns, matches student IDs, and processes multi-sheet workbooks
- 🧠 **Smart Pattern Recognition**: Handles email-formatted IDs, detects lecture/section boundaries, and normalizes data
- 🔍 **Advanced Search & Matching**: Fuzzy matching with manual override capabilities
- 📊 **Multi-Format Support**: Works with Excel (.xlsx), OpenDocument (.ods), CSV, Google Sheets, and SharePoint
- 🎨 **Interactive Preview System**: Real-time preview with highlighting, filtering, and editing before final export
- 📄 **Multi-Export Formats**: Generate JSON, TXT, and professionally formatted PDF reports
- 🔐 **Privacy-First Architecture**: Zero data transmission—everything processes locally

---

## ✨ Key Features

### 🚀 Core Automation Features

#### 1. **Automated Workbook Processing**
- **Multi-sheet scanning**: Automatically processes all sheets in a workbook
- **Dynamic column detection**: Intelligently identifies ID columns, name columns (single or dual), and target columns
- **Header detection**: Scans rows 2-5 to find attendance/grade columns, automatically detecting lecture vs. section boundaries
- **Week detection**: Automatically discovers available weeks (W1, W2, etc.) across all sheets

#### 2. **Intelligent Student ID Matching**
- **Email format support**: Extracts student IDs from email addresses (e.g., `123456@university.edu` → `123456`)
- **ID normalization**: Handles various ID formats (numeric strings, numbers, trailing zeros)
- **Duplicate detection**: Identifies and tracks duplicate IDs across sections
- **Ambiguous match handling**: Flags cases where an ID appears multiple times for manual review

#### 3. **Automated Attendance Marking**
- **Bulk processing**: Mark attendance for hundreds of students in seconds
- **Section-aware**: Automatically distinguishes between lecture and section attendance columns
- **Week-specific**: Targets specific week columns (W1, W2, etc.) automatically
- **Ordered input preservation**: Maintains the exact order from your input file, including section delimiters

#### 4. **Automated Grade Input**
- **CSV-style parsing**: Supports `id,grade` format with flexible comma handling
- **Bulk grade entry**: Input grades for entire classes at once
- **Preview before commit**: Review all grade changes before applying them
- **Manual override**: Edit individual grades in the preview interface

### 🛠️ Advanced Programming Features

#### 1. **In-Browser XLSX Processing**
```javascript
// Complex workbook manipulation using SheetJS
- Full workbook parsing and manipulation
- Cell-level read/write operations
- Style preservation (colors, formatting)
- Multi-sheet coordination
- Memory-efficient processing
```

#### 2. **Dynamic Column Detection Algorithm**
- **Heuristic-based detection**: Analyzes cell content patterns to identify ID vs. name columns
- **Multi-row header scanning**: Searches rows 1-5 for column headers
- **Boundary detection**: Automatically finds lecture/section column boundaries
- **Context-aware matching**: Uses target student IDs to improve detection accuracy

#### 3. **Smart URL Conversion System**
- **Google Sheets integration**: Converts edit URLs to XLSX export URLs automatically
- **SharePoint support**: Handles SharePoint/OneDrive URLs with authentication-aware fallbacks
- **CORS handling**: Graceful degradation when browser security blocks downloads
- **Error recovery**: Clear error messages with manual download instructions

#### 4. **State Management Architecture**
- **Modular state container**: Centralized state management without frameworks
- **Wizard-based workflow**: Multi-step process with validation at each stage
- **Preview state tracking**: Maintains edit history and match status for each row
- **Undo/redo capability**: Track changes before final commit

#### 5. **Interactive Preview System**
- **Real-time rendering**: Dynamically generates preview tables from workbook data
- **Dual view modes**: Grouped-by-sheet or ordered-by-input views
- **Advanced filtering**: Filter by sheet, delimiter, or match status
- **Inline editing**: Fix matches, mark wrong entries, edit grades—all in preview
- **Search functionality**: Find students by ID or name across all sheets

#### 6. **Error Handling & Validation**
- **Custom error types**: `ValidationError`, `DownloadError`, `FileError`, `ProcessingError`
- **Comprehensive validation**: Input validation, file format checking, data integrity verification
- **User-friendly messages**: Clear, actionable error messages with recovery suggestions
- **Graceful degradation**: Handles edge cases without crashing

#### 7. **Report Generation Engine**
- **Multi-format export**: JSON (structured data), TXT (human-readable), PDF (professional formatting)
- **Metadata tracking**: Timestamps, week, type, sheet URLs, and statistics
- **Ordered output**: Preserves input file structure with section delimiters
- **PDF generation**: Uses jsPDF + AutoTable for professional table formatting

### 🧩 Additional tools, input methods & UX

These capabilities are exposed from the home **feature cards** and related views (for example `#viewOcr`, sheet merger, formula panel, QR tool).

#### **OCR Attendance Parser (experimental)**
- **Tesseract.js + OpenCV.js**: Parse attendance sheet **images** in the browser
- **Batch images**: Queue multiple images with progress tracking
- **Confidence tiers**: Separates confident vs uncertain student IDs with **inline correction**
- **Text export**: Generate a `.txt` attendance list from OCR output
- **Dedicated workflow**: Step-by-step UI for load → process → review → export

#### **Sheet Merger**
- **Cross-sheet column mapping**: Drag-and-drop mapping across multiple sheets
- **Per-sheet color coding**: Golden-angle HSL tinting so columns stay visually distinct
- **Accordion + search**: Group by sheet, filter columns quickly at scale
- **Dynamic matrix**: Add/remove mapped columns; optional **auto-fill** from header names and column-index proximity
- **Duplicate headers**: Optional cleanup when the same header appears more than once
- **Paginated preview**: Preview merged rows with **load more**; download a merged **XLSX** workbook

#### **Online Sheet Formula Panel**
- **Google Sheets**: Per-sheet **`ARRAYFORMULA`** snippets ready to copy
- **Excel**: Per-sheet formulas using **`LET` / `XLOOKUP` / `SEQUENCE` / `INDEX`** patterns
- **Row-aware ranges**: Computes sensible end rows per sheet (e.g. `getFormulaMaxRowForSheet`)
- **Copy UX**: One-click copy with **toast** feedback

#### **QR tool / YaqeenScan**
- **Deployable scanner**: Feature entry for **`YaqeenScan.exe`** (placed under `downloads/` for static hosting)
- **Dedicated view**: Marketing-style page describing the companion tool (`#viewQrTool`)

#### **Search & Pick (third input method)**
- Besides **file upload** and **textarea**, you can **search the loaded workbook** by student **ID or name**
- **Build a pick list** with duplicate detection and toast feedback
- **Grade tasks**: Optional **grade dialog** when picking students for grade entry

#### **Home & wizard polish**
- **Feature search**: Filter feature cards via keywords (`#featureSearch` / `data-search`)
- **Splash screen**: Branded load animation (e.g. Yaqeen V0.3)
- **Toast notifications**: `#toastContainer` for copy, add, duplicate, and similar actions
- **Column search in wizard**: Live filter over column options plus **row-1 header scan** across sheets
- **JSON round-trip**: Load a previously exported **JSON report** to restore preview without re-processing the workbook
- **TXT exports — modified vs original**: Download **modified** attendance/grade lines and **original input** lines as separate text files
- **Dark theme (current UI)**: Black / red / white identity; tuned for long sessions

---

## 🏗️ Technical Architecture

### **Technology Stack**
- **Frontend**: Vanilla JavaScript (ES Modules) - No frameworks, pure performance
- **XLSX Processing**: [SheetJS (xlsx-js-style)](https://github.com/SheetJS/sheetjs) - Full workbook manipulation
- **PDF Generation**: [jsPDF](https://github.com/parallax/jsPDF) + [AutoTable](https://github.com/simonbengtsson/jsPDF-AutoTable)
- **OCR (experimental)**: [Tesseract.js](https://github.com/naptha/tesseract.js) + [OpenCV.js](https://docs.opencv.org/4.x/d5/d10/tutorial_js_root.html) (browser-side image parsing)
- **Deployment**: Static site (Netlify-compatible, no build step required)

### **Project Structure**
```
web/
├── index.html          # Main UI with wizard workflow & feature-card views
├── app.js              # Application entry point & DOM wiring (~323 lines)
├── attendance.js       # Core processing logic (~1,410 lines)
├── report.js           # Report generation engine (~200 lines)
├── downloads/          # Static assets for QR / YaqeenScan (e.g. README + optional .exe)
├── src/
│   ├── handlers.js     # Event handlers & orchestration (~3,556 lines)
│   ├── ocr.js          # OCR attendance parser (Tesseract + OpenCV pipeline)
│   ├── sheetMerger.js  # Multi-sheet column merger UI + XLSX export
│   ├── state.js        # State management
│   ├── navigation.js   # View switching
│   ├── metadata.js     # Report metadata & filenames
│   ├── fileRead.js     # File I/O helpers
│   ├── dom.js          # DOM utilities
│   └── uiStatus.js     # Status & loading UI
└── styles.css          # Styling (dark theme, responsive layout)
```

### **Key Algorithms & Techniques**

#### **1. Student ID Matching Algorithm**
```javascript
// Handles multiple ID formats:
- Pure numeric IDs: "123456"
- Email format: "123456@university.edu" → extracts "123456"
- Normalized IDs: Removes trailing zeros, handles number-to-string conversion
- Pattern-based detection: Identifies ID-like values when target set unavailable
```

#### **2. Column Detection Heuristic**
```javascript
// Multi-pass detection:
1. Name column detection: Scans first 10 columns, scores by text content (>50% text = name column)
2. ID column detection: Matches against target ID set, handles email format
3. Header detection: Scans rows 2-5, detects lecture/section boundaries
4. Context-aware refinement: Uses detected columns to improve accuracy
```

#### **3. Preview Generation Pipeline**
```javascript
// Complex multi-step process:
1. Load workbook into memory
2. Build student index per sheet (ID → row mapping)
3. Match input IDs against index (handles duplicates, ambiguous matches)
4. Generate preview rows with old/new values
5. Apply user edits (fix matches, edit grades)
6. Render interactive table with filtering
7. Apply edits to workbook with optional highlighting
8. Export modified workbook
```

#### **4. URL Conversion Logic**
```javascript
// Google Sheets:
/edit → /export?format=xlsx

// SharePoint:
Extract tenant, user path, file ID
Convert to download URL format
Handle authentication requirements gracefully
```

---

## 🎮 Usage Guide

### **Quick Start**

1. **Load Your Spreadsheet**
   - Upload an `.xlsx`, `.ods`, or `.csv` file, OR
   - Paste a Google Sheets or SharePoint URL (auto-downloads if CORS allows)

2. **Configure Settings**
   - Choose **Single-sheet** or **Multi-sheet** mode
   - Select target column (automatically detected from headers)
   - Choose task type: **Attendance** or **Grade input**

3. **Input Your Data**
   - **Attendance**: Upload a `.txt` file with student IDs (one per line, supports section delimiters)
   - **Grades**: Upload a `.txt` file with `id,grade` format
   - **Search & Pick**: Type or paste in the textarea, **or** search students by ID/name from the loaded workbook and build a pick list (with grade prompt when applicable)

4. **Review & Edit**
   - Preview all changes before applying
   - Fix incorrect matches using search
   - Edit individual grades if needed
   - Filter by sheet, delimiter, or status

5. **Export**
   - Download modified workbook (with optional cell highlighting)
   - Export JSON, TXT, or PDF reports
   - Optionally download **modified** vs **original** line lists as separate `.txt` files; reload a saved **JSON** report to restore preview

### **Input File Formats**

#### **Attendance Input (.txt)**
```
Section 1
123456
234567
345678

Section 2
456789
567890
```
- Empty lines reset section context
- Section titles are preserved in ordered output
- Duplicate IDs are tracked per section

#### **Grade Input (.txt)**
```
123456,15
234567,18
345678,20
```
- Format: `id,grade` (comma-separated)
- Supports grades with commas (uses first comma as separator)
- Invalid lines are treated as delimiters

---

## 🔧 Advanced Features

### **1. Wizard-Based Workflow**
A 4-step wizard guides you through the process:
- **Step 1**: Load file (upload or URL)
- **Step 2**: Configure (mode, sheet, column)
- **Step 3**: Input data (task type, input file)
- **Step 4**: Generate preview

### **2. Interactive Preview System**
- **Grouped view**: Organize by sheet for easy review
- **Ordered view**: Preserve input file order with section delimiters
- **Filtering**: Filter by sheet, delimiter, or match status
- **Search**: Find students by ID or name across all sheets
- **Inline editing**: Fix matches, mark wrong, edit grades

### **3. Smart Matching & Fixing**
- **Automatic matching**: Matches students by ID across sheets
- **Ambiguous detection**: Flags duplicate IDs for review
- **Manual fix dialog**: Search by ID or name to correct matches
- **Status tracking**: Tracks matched, not found, ambiguous, and manually fixed entries

### **4. Cell Highlighting**
- **Optional highlighting**: Highlight modified cells in exported workbook
- **Custom colors**: Choose highlight color (default: yellow)
- **Style preservation**: Maintains existing cell formatting

### **5. Multi-Format Export**
- **Modified Workbook**: Download `.xlsx` with all changes applied
- **JSON Report**: Structured data for programmatic use (can be **re-loaded** to restore preview state)
- **TXT Report**: Human-readable action report
- **TXT — modified vs original**: Separate downloads for lines that reflect **applied edits** vs **raw input**
- **PDF Report**: Professionally formatted table with summary

### **6. Toasts, search helpers & home discovery**
- **Toast notifications**: Inline feedback for copy, duplicates, list changes, and similar actions
- **Wizard column search**: Quickly find columns by label across the workbook
- **Home feature search**: Filter the feature grid by keywords

---

## 🚀 Deployment

### **Netlify Deployment** (Recommended)

1. **Connect Repository**
   ```bash
   # Netlify will auto-detect settings
   ```

2. **Configure Build Settings**
   - **Publish directory**: `web`
   - **Build command**: (none - static site)
   - **Deploy**: Automatic on push to main branch

3. **Environment Variables**
   - None required! This is a fully static site.

### **Alternative Hosting**
Works on any static hosting service:
- GitHub Pages
- Vercel
- Cloudflare Pages
- AWS S3 + CloudFront
- Any web server

---

## 🎓 Complex Programming Highlights

### **1. In-Browser File Processing**
Processing large Excel files entirely in the browser requires:
- Efficient memory management
- Streaming file reads
- Optimized parsing algorithms
- Error handling for corrupted files

### **2. Dynamic Column Detection**
The column detection algorithm uses:
- **Heuristic scoring**: Analyzes cell content patterns
- **Multi-pass scanning**: Searches multiple rows and columns
- **Context awareness**: Uses target IDs to improve accuracy
- **Boundary detection**: Finds lecture/section column boundaries automatically

### **3. State Management Without Frameworks**
Built a custom state management system:
- Centralized state container
- Reactive UI updates
- Wizard workflow state machine
- Preview state tracking with edit history

### **4. Error Handling Architecture**
Custom error types with user-friendly messages:
```javascript
ValidationError  // Input validation failures
DownloadError    // Network/CORS issues
FileError        // File parsing errors
ProcessingError  // Workbook processing failures
```

### **5. URL Conversion & CORS Handling**
- Converts Google Sheets edit URLs to export URLs
- Handles SharePoint URLs with authentication awareness
- Graceful CORS failure handling with clear user guidance
- Validates downloaded files (checks Excel signature)

### **6. Preview System Complexity**
The preview system involves:
- Building student indexes per sheet
- Matching algorithms with ambiguity detection
- Real-time filtering and sorting
- Inline editing with state tracking
- Delimiter preservation in ordered view

---

## 📊 Performance & Limitations

### **Performance**
- Processes workbooks with 1000+ students in seconds
- Efficient memory usage (processes in chunks where possible)
- Optimized DOM updates (minimal re-renders)

### **Known Limitations**
- **CORS restrictions**: Google Sheets/SharePoint downloads may fail due to browser security
  - **Solution**: Manual download and upload (clearly guided in UI)
- **Large files**: Very large workbooks (>10MB) may cause browser slowdowns
  - **Solution**: Process in smaller batches or use desktop Excel for very large files
- **Browser compatibility**: Requires modern browser with ES Module support
  - **Supported**: Chrome, Firefox, Safari, Edge (latest versions)

---

## 🤝 Contributing

This is a personal project, but suggestions and feedback are welcome!

**Contact**: [mohammed.abusarie@ecu.edu.eg](mailto:mohammed.abusarie@ecu.edu.eg)

---

## 📝 License

This project is created for educational and administrative use at the Egyptian Chinese University.

---

## 🙏 Acknowledgments

Built with:
- [SheetJS](https://sheetjs.com/) - Excel file processing
- [jsPDF](https://github.com/parallax/jsPDF) - PDF generation
- [AutoTable](https://github.com/simonbengtsson/jsPDF-AutoTable) - PDF table formatting
- [Tesseract.js](https://github.com/naptha/tesseract.js) - In-browser OCR (experimental attendance parser)
- [OpenCV.js](https://docs.opencv.org/4.x/d5/d10/tutorial_js_root.html) - Image preprocessing for OCR pipeline

---

## 🎯 Future Enhancements

Potential features for future versions:
- Batch processing multiple weeks at once
- Template-based column mapping
- Import/export of configuration presets
- Mobile-responsive improvements

---

<div align="center">

**Made with ❤️ by Mohammed Abusarie**

*Automating the tedious, so you can focus on what matters.*

</div>
