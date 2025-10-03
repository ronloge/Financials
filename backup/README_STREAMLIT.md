# 📊 Project Financials Consultant Metrics Analyzer - Web Interface

## Overview
Interactive web-based dashboard for analyzing consultant and solution architect performance using Streamlit. Provides real-time data exploration, filtering, and visualization capabilities.

## 🚀 Quick Start

### Installation
```bash
pip install -r requirements_streamlit.txt
```

### Launch Application
```bash
streamlit run streamlit_app.py
```

The web interface will open at `http://localhost:8501`

## 📁 Data Sources

### Excel File Upload
- Upload PMO Project Financials Excel files
- Supports .xlsx and .xls formats
- Automatic column detection and processing

### SQL Server/OLAP Connection
- Connect to Analysis Services cubes
- Server: NSSQL1601.netsync.com
- Database: DWPROD
- Supports Windows and SQL Server authentication

## 🎛️ Configuration

### Interactive Settings
- **Thresholds**: Adjust efficiency (15%) and success (30%) thresholds
- **Analysis Options**: Enable/disable customer and DAS+ analysis
- **Minimum Projects**: Set thresholds for meaningful analysis

### Config File Upload
- Upload custom `config.json` files
- Download current configuration
- Reset to defaults or original settings

## 📈 Dashboard Tabs

### 1. Dashboard
- Key performance metrics overview
- Success score distribution charts
- Top performer rankings
- Total hours and project counts

### 2. Consultants
- **Performance Table**: Sortable consultant metrics with filtering
- **Project Drill-down**: Click any consultant to see detailed projects
- **Color-coded Projects**: Green (under budget), Yellow (acceptable), Red (over budget)
- **Random Project Review**: Automatic selection of 2 projects per consultant
- **Quarter Selection**: Choose current or previous quarter for reviews
- **Exclusion Management**: Add/remove project exclusions

### 3. Solution Architects
- SA performance rankings
- Project design success rates
- Volume-weighted scoring
- Individual SA project drill-down

### 4. Customers
- Customer performance analysis (when enabled)
- Success rates by client
- Consultant-customer combinations
- Project volume and variance metrics

### 5. DAS+ Analysis
- Delivery Accuracy Score Plus calculations
- Performance distribution analysis
- Project-level DAS+ scores
- High/low performance categorization

### 6. Consultant vs SA
- **Success Rate Pairings**: Best SA-Consultant team combinations
- **DAS+ Pairings**: Teams with highest DAS+ scores when working together
- **Project Drill-down**: View shared projects for any pairing
- Minimum project filters for meaningful comparisons

## 🎯 Interactive Features

### Project Selection & Review
- **Automatic Review Selection**: 2 random projects per consultant
- **Smart Selection Logic**: Prefers 1 good + 1 poor project
- **Quarter Filtering**: Current or previous quarter options
- **Performance Indicators**: 🟢 Good, 🔴 Poor, 🟡 Average projects

### Data Filtering
- **Minimum Projects**: Filter consultants/SAs by project count
- **Performance Thresholds**: Set minimum efficiency/success rates
- **Date Ranges**: Focus on specific time periods
- **Role Selection**: Filter by consultant or SA

### Exclusion Management
- **Session Exclusions**: Temporary exclusions for current session
- **File Exclusions**: Permanent exclusions saved to Exclude.csv
- **Project-level Control**: Exclude specific projects per consultant
- **Bulk Operations**: Select multiple projects for exclusion

## 📊 Visualizations

### Charts & Graphs
- **Bar Charts**: Success scores and efficiency rankings
- **Scatter Plots**: Performance vs volume analysis
- **Box Plots**: Distribution comparisons between roles
- **Histograms**: Variance distribution analysis

### Color Coding
- **🟢 Green**: Projects under budget (>10% under)
- **🟡 Yellow**: Acceptable performance (10-30% over)
- **🔴 Red**: Failed projects (>30% over budget)
- **Performance Indicators**: Visual status for quick assessment

## 💾 Data Export

### Download Options
- **CSV Exports**: All summary tables downloadable
- **Excel Reports**: Detailed project breakdowns
- **Configuration Files**: Save current settings
- **Project Details**: Individual consultant/SA project lists

### File Management
- **Automatic Cleanup**: Remove old report files
- **Timestamped Outputs**: Prevent file overwrites
- **Session Management**: Clear cached data

## 🔧 Advanced Features

### Session State Management
- **Persistent Selections**: Maintain consultant/SA selections across interactions
- **Cached Calculations**: Store DAS+ and complex analysis results
- **Configuration Memory**: Remember user settings during session

### Debug Mode
- **Debug Toggle**: Enable detailed processing logs
- **Error Tracking**: View application errors and warnings
- **Performance Monitoring**: Track calculation times

### Connection Management
- **Credential Storage**: Save database connections for Q Developer
- **Connection Testing**: Verify database accessibility
- **Endpoint Discovery**: Automatically find available services

## 🎨 User Interface

### Responsive Design
- **Wide Layout**: Optimized for large datasets
- **Collapsible Sidebar**: Configuration and upload controls
- **Tabbed Interface**: Organized feature access
- **Interactive Tables**: Sortable, filterable data grids

### User Experience
- **Progress Indicators**: Visual feedback during processing
- **Status Messages**: Clear success/error notifications
- **Help Text**: Contextual guidance throughout interface
- **Keyboard Shortcuts**: Efficient navigation options

## 🔍 Troubleshooting

### Common Issues
- **File Upload Errors**: Check Excel file format and column headers
- **Database Connection**: Verify credentials and network access
- **Performance Issues**: Use filters to reduce data volume
- **Missing Data**: Check for required columns in source data

### Debug Information
- Enable debug mode for detailed error logs
- Check browser console for JavaScript errors
- Verify Python package versions match requirements
- Review connection logs for database issues

## 📋 Requirements

### System Requirements
- Python 3.7+
- 4GB+ RAM recommended for large datasets
- Modern web browser (Chrome, Firefox, Safari, Edge)

### Dependencies
```
streamlit>=1.28.0
pandas>=1.5.0
numpy>=1.24.0
openpyxl>=3.1.0
plotly>=5.15.0
requests>=2.31.0
```

## 🚀 Deployment

### Local Development
```bash
streamlit run streamlit_app.py --server.port 8501
```

### Production Deployment
```bash
streamlit run streamlit_app.py --server.address 0.0.0.0 --server.port 80
```

### Docker Deployment
```dockerfile
FROM python:3.9-slim
COPY requirements_streamlit.txt .
RUN pip install -r requirements_streamlit.txt
COPY . .
EXPOSE 8501
CMD ["streamlit", "run", "streamlit_app.py"]
```

## 🔐 Security Considerations

### Data Protection
- Credentials stored locally only during session
- No persistent storage of sensitive data
- Database connections use encrypted protocols
- File uploads processed in memory only

### Access Control
- No built-in authentication (add reverse proxy if needed)
- Session isolation between users
- Temporary file cleanup on exit

## 📞 Support

### Getting Help
- Check debug logs for error details
- Verify data format matches expected structure
- Review configuration settings for accuracy
- Test with sample data to isolate issues

### Feature Requests
- Submit enhancement requests with business justification
- Include specific use cases and expected outcomes
- Provide sample data for testing new features