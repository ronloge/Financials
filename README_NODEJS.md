# Project Financials Consultant Metrics Analyzer - Node.js Web Application

## Overview
Web-based version of the Project Financials Consultant Metrics Analyzer built with Node.js, Express.js, and modern web technologies.

## 🚀 Features

### Core Functionality
- **Excel File Upload** - Upload and analyze Project Financials Excel files
- **Real-time Configuration** - Adjust thresholds and settings via web interface
- **Interactive Dashboard** - Visual metrics and performance charts
- **Consultant Analysis** - Detailed performance tables with color coding
- **Date Filtering** - Configure date ranges for analysis

### Web Interface
- **Responsive Design** - Works on desktop and mobile devices
- **Bootstrap UI** - Modern, professional interface
- **Chart.js Integration** - Interactive performance charts
- **Real-time Updates** - Instant feedback and analysis results

## 📋 Requirements
- Node.js 14+ 
- npm or yarn package manager

## 🛠️ Installation

### 1. Install Dependencies
```bash
npm install
```

### 2. Start Development Server
```bash
npm run dev
```

### 3. Start Production Server
```bash
npm start
```

## 🌐 Usage

1. **Open Browser**: Navigate to `http://localhost:3000`
2. **Upload File**: Select your Project Financials Excel file
3. **Configure Settings**: Adjust thresholds and date filtering as needed
4. **Run Analysis**: Click "Run Analysis" to process the data
5. **View Results**: Explore the Dashboard and Consultants tabs

## ⚙️ Configuration

### Thresholds
- **Efficiency Threshold**: Percentage threshold for consultant efficiency scoring
- **Success Threshold**: Percentage threshold for project success determination

### Date Filtering
- **Enable/Disable**: Toggle date-based filtering
- **Days from Today**: Include projects from last N days
- **Specific Date**: Include projects from a specific date onwards

## 📊 Analysis Features

### Dashboard
- Total consultants, hours, and projects
- Average efficiency metrics
- Top 10 consultant performance chart

### Consultant Analysis
- Performance table with color coding:
  - 🟢 Green: High performance (80%+ efficiency)
  - 🟡 Yellow: Average performance (60-79% efficiency)  
  - 🔴 Red: Poor performance (<60% efficiency)
- Sortable columns
- Detailed metrics per consultant

## 🗂️ File Structure
```
├── server.js              # Main Express server
├── package.json           # Node.js dependencies
├── config.json            # Configuration file
├── public/                # Web assets
│   ├── index.html         # Main HTML page
│   ├── style.css          # Styling
│   └── script.js          # Client-side JavaScript
├── uploads/               # Uploaded files (auto-created)
└── backup/                # Backup of old files
```

## 🔧 API Endpoints

### POST /upload
Upload and analyze Excel file
- **Body**: FormData with 'excelFile'
- **Response**: Analysis results JSON

### GET /config
Get current configuration
- **Response**: Configuration JSON

### POST /config  
Save configuration
- **Body**: Configuration JSON
- **Response**: Success/error status

## 🚨 Error Handling
- File upload validation
- Excel parsing error handling
- Configuration validation
- User-friendly error messages

## 📈 Performance
- Efficient Excel parsing with XLSX library
- Minimal server-side processing
- Client-side chart rendering
- Responsive UI updates

## 🔒 Security Notes
- File upload size limits
- File type validation (.xlsx, .xls only)
- Input sanitization
- No sensitive data storage

## 🆕 Migration from Streamlit
This Node.js version provides:
- **Better Performance**: Faster file processing and UI updates
- **Modern Interface**: Responsive web design with Bootstrap
- **Easier Deployment**: Standard web application deployment
- **Cross-Platform**: Works on any device with a web browser

## 🚀 Deployment Options
- **Local**: Run on localhost for personal use
- **Network**: Deploy on company network for team access
- **Cloud**: Deploy to Heroku, AWS, or other cloud platforms

## 📝 Development
- **Hot Reload**: Use `npm run dev` for development with auto-restart
- **Debugging**: Console logs and error handling throughout
- **Extensible**: Easy to add new features and analysis types