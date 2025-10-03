![Version](https://img.shields.io/badge/version-1.0.2-blue)
![AI Powered](https://img.shields.io/badge/AI-Powered-green)
![Node.js](https://img.shields.io/badge/Node.js-18+-brightgreen)
![License](https://img.shields.io/badge/license-MIT-blue)

# Project Financials Consultant Metrics Analyzer - Node.js Edition

## Overview
Comprehensive Node.js web application that analyzes consultant and solution architect performance from Project Financials Excel data, featuring advanced analytics, predictive modeling, and detailed reporting with sophisticated scoring algorithms.

## 🚀 Key Features

### Core Analysis
- ✅ **Consultant Performance Analysis** - Dual-threshold efficiency and success scoring
- ✅ **Solution Architect Analysis** - Volume-weighted composite performance ranking
- ✅ **Customer Performance Analysis** - Practice vs Company performance with risk scoring
- ✅ **DAS+ Analysis** - Delivery Accuracy Score Plus with completion-aware performance measurement
- ✅ **Team Combination Analysis** - SA-Consultant and SA-Customer success rate analysis
- ✅ **Advanced Risk Scoring** - Multi-factor customer risk assessment

### Data Processing Intelligence
- ✅ **Smart Name Normalization** - Handles spacing inconsistencies and formatting variations
- ✅ **Comma-separated Field Processing** - Multiple consultants/SAs per project with hour splitting
- ✅ **Missing Data Handling** - Prefers valid values over N/A entries with intelligent merging
- ✅ **Exclusion Management** - Configurable project/consultant exclusions via file upload
- ✅ **Date-based Filtering** - Configurable project age filtering with status-aware logic

### Interactive Web Interface
- ✅ **Bootstrap UI** - Modern, responsive web interface with NetSync branding
- ✅ **Sortable Tables** - All columns sortable with ascending/descending indicators
- ✅ **Configurable Thresholds** - Real-time GUI controls for all performance thresholds
- ✅ **Color-coded Analysis** - Green/Yellow/Red highlighting based on configurable performance
- ✅ **Interactive Modals** - Detailed project views with Savant and SSRS links
- ✅ **Chart Visualizations** - Performance charts using Chart.js
- ✅ **Export Functionality** - CSV and Excel export for all analysis tabs
- ✅ **Random Project Review** - 2 random projects per consultant with filtering options
- ✅ **Exclusion Management** - Real-time project exclusions with auto-recalculation
- ✅ **Consultant of Quarter** - Gamified performance rankings with DQ functionality

## 📊 Analysis Modules

### 1. Consultant Performance Analysis
**Metrics Calculated:**
- **Efficiency Score**: Projects within efficiency threshold (default 15%)
- **Success Rate**: Projects within success threshold (default 30%)
- **Total Hours**: Billable hours contributed
- **Project Count**: Number of projects worked on

**Features:**
- **Clickable consultant names** open detailed project modals with:
  - Full project details with Savant and SSRS links
  - Real-time exclusion functionality
  - **Random Projects for Review** - 2 random projects with filtered/all data options
  - Color-coded project rows based on variance thresholds
- **Export functionality** - CSV and Excel export options
- **Consultant of Quarter Rankings** - Gamified performance highlights with DQ management
- Color-coded rows based on performance levels
- Sortable by all metrics

### 2. Solution Architect Analysis
**Metrics Calculated:**
- **Success Rate**: Projects designed that stayed within budget
- **Volume Score**: Total hours designed/sold
- **Project Count**: Number of projects designed
- **Budgeted vs Actual Hours**: Comprehensive hour tracking
- **Variance Analysis**: Budget performance measurement

**Features:**
- **Performance Highlights** - Top 3 SAs with medal rankings (🥇🥈🥉)
- **Clickable SA names** open detailed project portfolios
- **Export functionality** - CSV and Excel export options
- Volume-weighted performance ranking
- Success rate tracking by SA
- Color-coded project variance analysis

### 3. Customer Performance Analysis
**Metrics Calculated:**
- **Success Rate**: Percentage of projects within budget
- **Average Variance**: Mean budget variance across projects
- **Risk Score**: Multi-factor risk assessment (NEW)
- **Project Volume**: Total projects and hours
- **Configurable minimum project count filter** with slider control (1-20 projects)
- **Auto-refresh** when filter settings change

**Risk Scoring Algorithm:**
The risk scoring system evaluates customers across four key dimensions:

1. **Success Rate Risk (40% weight)**
   - <50% success rate = 4 points (High Risk)
   - 50-70% success rate = 2 points (Medium Risk)
   - 70-85% success rate = 1 point (Low Risk)
   - >85% success rate = 0 points (Minimal Risk)

2. **Project Volume Risk (20% weight)**
   - <5 projects = 2 points (Limited data, higher uncertainty)
   - 5-10 projects = 1 point (Moderate data)
   - >10 projects = 0 points (Sufficient data for reliable assessment)

3. **Variance Consistency Risk (25% weight)**
   - >50% average variance = 3 points (Highly unpredictable)
   - 30-50% average variance = 2 points (Moderately unpredictable)
   - 15-30% average variance = 1 point (Some variability)
   - <15% average variance = 0 points (Consistent performance)

4. **Budget Size Risk (15% weight)**
   - >5000 hours = 1 point (Very large projects carry execution risk)
   - 2000-5000 hours = 0 points (Optimal project size)
   - <2000 hours = 1 point (Very small projects may indicate scope issues)

**Risk Level Classification:**
- **Low Risk (0-2 points)**: Reliable customer with consistent, successful project outcomes
- **Medium Risk (3-5 points)**: Generally reliable but with some areas of concern
- **High Risk (6-8 points)**: Significant risk factors present, requires close monitoring
- **Critical Risk (9+ points)**: Multiple severe risk factors, immediate attention required

### 4. DAS+ (Delivery Accuracy Score Plus) Analysis
**Purpose**: Advanced performance metric that measures alignment between budget usage and actual delivery progress.

**Calculation Method:**
```
Budget Ratio = Actual Hours / Budget Hours
Completion Ratio = Project Complete % / 100
DAS+ Score = 1 - |Budget Ratio - Completion Ratio|
```

**Completion Normalization:**
- Closed projects = 100% complete (regardless of reported %)
- 95%+ complete projects = 100% complete
- All others use actual completion percentage

**Score Interpretation:**
- **1.0**: Perfect alignment of budget usage and delivery
- **0.85+**: Excellent performance
- **0.75-0.84**: Good performance
- **0.50-0.74**: Moderate performance
- **0.30-0.49**: Poor performance
- **<0.30**: Critical performance issues

**Metrics Displayed:**
- Average DAS+ Score
- Median DAS+ Score
- Count of low-performing projects (<0.75)
- Count of high-performing projects (≥0.85)
- Count of projects needing review (0.3-0.9 range)

### 5. Team Combination Analysis (NEW)
**SA-Consultant Combinations:**
- Analyzes success rates when specific Solution Architects work with specific Consultants
- Identifies optimal team pairings for future project assignments
- **Configurable minimum project count filter** with slider control (1-10 projects)
- **Clickable rows** to view detailed projects for each SA-Consultant combination
- **Auto-refresh** when filter settings change
- Sorted by success rate to highlight best-performing combinations

**SA-Customer Analysis:**
- Tracks Solution Architect performance with specific customers
- Identifies which SAs are most successful with which customers
- Helps with SA assignment decisions for customer projects
- Enables customer relationship optimization

**Business Value:**
- **Resource Optimization**: Assign proven successful SA-Consultant combinations
- **Customer Success**: Match SAs with customers where they have historical success
- **Risk Mitigation**: Avoid combinations with poor historical performance
- **Training Insights**: Identify why certain combinations work better

## 🎨 Configuration Options

### Threshold Controls (GUI Configurable)
- **Efficiency Threshold**: Strict performance standard (default 15%)
- **Success Threshold**: Project completion standard (default 30%)
- **Green Threshold**: Under-budget performance (default -10%)
- **Yellow Threshold**: Acceptable over-budget (default +10%)
- **Red Threshold**: Failed project threshold (default +30%)

### Exclusions Management
- **View Current Exclusions**: Modal showing all active consultant-project exclusions
- **Manage DQ List**: Disqualify/re-qualify consultants from Consultant of Quarter
- **Clear All Caches**: Force reload of filter files and exclusions
- **Real-time Exclusions**: Add exclusions directly from project modals with auto-recalculation

### Date Filtering Options
- **Enable/Disable**: Toggle date-based project filtering
- **Filter Type**: Days from today OR specific date
- **Days Range**: Configurable number of days (default 365)
- **Status-Aware**: Always include open projects regardless of age

### Data Source Options
- **📁 Upload Excel File**: Main project financials data (required)
  - Single file for current analysis
  - Multiple files for trending analysis (when enabled)
- **🌐 Connect to SSRS Report**: Direct connection to SQL Server Reporting Services
  - Server: NSSQL1601.netsync.com
  - Database: DWPROD
  - Windows or SQL Server Authentication

### Filter Files (Optional)
- **Engineers List**: Filter to specific consultants (.txt file or text input)
- **Exclusions**: Consultant-Project exclusions (.csv file or text input)
- **Solution Architects**: Filter to specific SAs (.txt file or text input)
- **Config File**: Upload JSON configuration (.json file)

## 🔧 Technical Architecture

### Backend (Node.js/Express)
- **Express.js**: Web server framework
- **Multer**: File upload handling
- **XLSX**: Excel file parsing
- **JSON Configuration**: Persistent settings storage

### Frontend (Vanilla JavaScript)
- **Bootstrap 5**: Responsive UI framework
- **Chart.js**: Performance visualizations
- **Vanilla JS**: No framework dependencies for maximum performance
- **Font Awesome**: Icon library

### Data Processing
- **Smart Parsing**: Automatic header detection and data normalization
- **Duplicate Handling**: Intelligent project merging and deduplication
- **Multi-pass Analysis**: Separate passes for practice vs company metrics
- **Real-time Configuration**: Dynamic threshold updates without restart

## 📋 Installation & Setup

### Prerequisites
- Node.js 14+ 
- npm or yarn package manager

### Installation Steps
```bash
# Clone or download the project
cd "Financials Folder"

# Install dependencies
npm install

# Start the server
npm start
# OR
node server.js
```

### Access the Application
- Open browser to `http://localhost:3000`
- Upload Excel file and configure settings
- Run analysis and explore results

## 📁 File Structure
```
Financials Folder/
├── server.js                 # Main Node.js server
├── package.json             # Dependencies and scripts
├── config.json              # Configuration storage
├── public/
│   ├── index.html          # Main web interface
│   ├── script.js           # Client-side JavaScript
│   └── style.css           # Custom styling
├── uploads/                # Temporary file storage
└── backup/                 # Original Python files
```

## 🎯 Usage Workflow

1. **Configure Settings**: Adjust thresholds using GUI sliders
2. **Upload Files**: 
   - Required: Excel project financials file
   - Optional: Engineer lists, exclusions, SA lists, config files
3. **Run Analysis**: Click "🚀 Run Analysis" button
4. **Explore Results**: Navigate through tabs:
   - **Dashboard**: Overview metrics and charts
   - **Consultants**: Individual consultant performance
   - **Solution Architects**: SA performance and project details
   - **Customers**: Customer analysis with risk scoring
   - **DAS+ Analysis**: Advanced delivery accuracy metrics
   - **Team Analysis**: SA-Consultant and SA-Customer combinations
5. **Drill Down**: Click names to view detailed project information with:
   - Full project details and external links (Savant/SSRS)
   - Real-time exclusion capabilities
   - Random project review functionality
6. **Export**: Multiple export options:
   - Individual tab exports (CSV/Excel)
   - Complete analysis export (all tabs in one file)
   - Export buttons available on each analysis tab
7. **Manage Exclusions**: 
   - View and remove current exclusions
   - Add exclusions directly from project modals
   - Manage Consultant of Quarter disqualifications

## 🎮 Gamification Features

### Consultant of the Quarter
- **Composite Scoring**: Multi-factor performance ranking system
- **Medal Rankings**: 🥇 Gold, 🥈 Silver, 🥉 Bronze performance highlights
- **Disqualification Management**: Remove/restore consultants from rankings
- **Performance Cards**: Visual highlights with key metrics
- **Clickable Rankings**: Direct access to consultant project details

### Random Project Review
- **2 Random Projects**: Automatically selected for each consultant
- **Filtered vs All Data**: Toggle between practice-filtered and complete datasets
- **Refresh Functionality**: Generate new random selections
- **Direct Links**: Savant and SSRS integration for immediate access
- **Exclusion Integration**: Add exclusions directly from random project reviews

## 📊 Business Intelligence Insights

### For Pre-Sales Teams
- **SA-Customer Success Rates**: Assign SAs to customers where they have proven success
- **Customer Risk Assessment**: Identify high-risk customers before engagement
- **Team Optimization**: Use SA-Consultant combination data for project staffing
- **Estimation Accuracy**: Track SA estimation performance over time

### For Post-Sales Teams
- **Project Health Monitoring**: Use DAS+ scores to identify projects needing intervention
- **Resource Allocation**: Optimize consultant assignments based on historical performance
- **Customer Relationship Management**: Proactively manage high-risk customer relationships
- **Performance Coaching**: Use detailed metrics for consultant development

### For Management
- **Strategic Planning**: Understand practice performance vs company-wide metrics
- **Risk Management**: Multi-factor customer risk scoring for portfolio management
- **Team Development**: Identify successful team combinations and replicate them
- **Performance Tracking**: Comprehensive metrics for consultant and SA evaluation

## 🔍 Advanced Features

### Sortable Tables
- All table columns are sortable with visual indicators (↕ ↑ ↓)
- Click once for ascending, click again for descending
- Smart sorting: numeric columns sort numerically, text columns alphabetically

### Interactive Modals
- **Consultant Project Details** with:
  - Variance analysis and color-coded performance
  - Savant and SSRS external links
  - Real-time exclusion functionality
  - Random Projects for Review (2 projects with filtered/all data options)
  - Sortable project tables
- **SA Project Portfolios** with success tracking and variance analysis
- **Customer Project Histories** with team assignments and resource details
- **SA-Consultant Combination Projects** showing shared project details
- **Exclusions Management** - View and remove current exclusions
- **Disqualified Consultants** - Manage Consultant of Quarter DQ list

### Color-Coded Analysis
- **Performance Rows**: Green (excellent), Yellow (good), Red (needs attention)
- **Risk Badges**: Color-coded risk levels for immediate visual assessment
- **Variance Indicators**: Project-level color coding based on budget performance

### Real-time Configuration
- GUI sliders update thresholds immediately
- Configuration persists between sessions
- No server restart required for threshold changes

## 🚨 Data Quality Features

### Intelligent Data Processing
- **Name Normalization**: Handles spacing and formatting inconsistencies
- **Duplicate Detection**: Merges duplicate project entries intelligently
- **Missing Data Handling**: Prefers valid values over N/A entries
- **Exclusion Processing**: Applies consultant-project exclusions automatically

### Validation & Error Handling
- **File Format Validation**: Ensures proper Excel file structure
- **Data Type Conversion**: Safely converts text to numbers with fallbacks
- **Missing Field Detection**: Graceful handling of missing columns
- **Error Reporting**: Clear error messages for troubleshooting

## 🔄 Migration from Python/Streamlit

This Node.js version provides all functionality from the original Python/Streamlit application with additional benefits:

### Enhanced Features
- **Better Performance**: Faster processing and rendering
- **Improved UI**: More responsive and professional interface
- **Enhanced Sorting**: All tables fully sortable with better UX
- **Real-time Config**: GUI-based configuration without file editing
- **Team Analysis**: New SA-Consultant and SA-Customer combination analysis
- **Risk Scoring**: Advanced customer risk assessment

### Maintained Features
- **All Original Calculations**: Identical algorithms and scoring methods
- **DAS+ Analysis**: Complete implementation with all original features
- **File Processing**: Same intelligent data handling and normalization
- **Export Capabilities**: Equivalent reporting and analysis depth

## 🛠️ Troubleshooting

### Common Issues
1. **File Upload Errors**: Ensure Excel file has proper structure with "Resources Engaged" column
2. **Missing Data**: Check that Budget Hrs and Total Hrs Posted columns exist
3. **Configuration Issues**: Use "Upload Config File" to restore settings
4. **Performance Issues**: Large files (>10MB) may take longer to process

### Support
- Check browser console for detailed error messages
- Verify Excel file format matches expected structure
- Ensure all required columns are present in the data
- Use smaller data sets for testing if performance issues occur

## 📈 Future Enhancements

### Planned Features
- **Database Integration**: Direct connection to project management systems
- **Automated Reporting**: Scheduled report generation and email delivery
- **Predictive Analytics**: ML models for project outcome prediction
- **API Integration**: Connect to CRM and project management tools
- **Multi-tenant Support**: Support for multiple organizations/practices

### Enhancement Opportunities
- **Real-time Data**: Live project status updates
- **Mobile Optimization**: Enhanced mobile device support
- **Advanced Visualizations**: More chart types and interactive dashboards
- **Collaboration Features**: Shared analysis and commenting
- **Integration APIs**: Connect with Salesforce, Jira, and other business tools

---

**Version**: 2.0 (Node.js Edition)  
**Last Updated**: January 2025  
**Compatibility**: Node.js 14+, Modern Browsers (Chrome, Firefox, Safari, Edge)# Financials


