#!/usr/bin/env python3
"""
Streamlit Web GUI for Project Financials Consultant Metrics Analyzer

A web-based interface for analyzing consultant and solution architect performance
from Project Financials Excel data with interactive filtering and visualization.
"""

import streamlit as st
import pandas as pd
import numpy as np
import json
import io
import tempfile
import os
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from consultant_metrics_analyzer_fixed import ConsultantMetricsAnalyzer

# Page configuration
st.set_page_config(
    page_title="Consultant Metrics Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

def debug_log(message):
    """Add debug message to log"""
    if 'debug_log' not in st.session_state:
        st.session_state.debug_log = []
    
    from datetime import datetime
    timestamp = datetime.now().strftime('%H:%M:%S.%f')[:-3]
    st.session_state.debug_log.append(f"[{timestamp}] {message}")
    
    # Keep only last 50 entries
    if len(st.session_state.debug_log) > 50:
        st.session_state.debug_log = st.session_state.debug_log[-50:]

def main():
    debug_log("App started/rerun")
    st.title("📊 Project Financials Consultant Metrics Analyzer")
    st.markdown("---")
    
    # Sidebar for file uploads and configuration
    with st.sidebar:
        # Debug toggle
        debug_mode = st.checkbox("🔍 Debug Mode", key="debug_mode")
        if debug_mode:
            if st.button("📋 Clear Debug Log"):
                st.session_state.debug_log = []
        
        # Application stop button
        st.markdown("---")
        if st.button("⛔ Stop Application", type="secondary"):
            st.success("Stopping application...")
            import os
            os._exit(0)
        
        st.header("📁 Data Source")
        
        # Data source selection
        data_source = st.radio(
            "Select Data Source",
            ["📁 Upload Excel File", "🌐 Connect to SSRS Report"],
            help="Choose between uploading an Excel file or connecting to SSRS"
        )
        
        if data_source == "📁 Upload Excel File":
            # Excel file upload
            excel_file = st.file_uploader(
                "Upload Project Financials Excel File",
                type=['xlsx', 'xls'],
                help="Upload your PMO Project Financials Excel file"
            )
        else:
            # SQL Server connection
            excel_file = None
            st.subheader("💾 SQL Server Connection")
            
            col1, col2 = st.columns(2)
            with col1:
                sql_server = st.text_input("Server", value="NSSQL1601.netsync.com", help="SQL Server instance")
            with col2:
                sql_database = st.text_input("Database", value="DWPROD", help="Database name")
            
            auth_type = st.radio(
                "Authentication Type",
                ["Windows Authentication", "SQL Server Authentication"],
                help="Choose authentication method"
            )
            
            if auth_type == "SQL Server Authentication":
                col1, col2 = st.columns(2)
                with col1:
                    sql_username = st.text_input("Username", help="SQL Server username")
                with col2:
                    sql_password = st.text_input("Password", type="password", help="SQL Server password")
            else:
                col1, col2 = st.columns(2)
                with col1:
                    sql_username = st.text_input("Username", help="Format: DOMAIN\\username or leave blank for current user")
                with col2:
                    sql_password = st.text_input("Password", type="password", help="Windows domain password (optional)")
            
            if st.button("🔗 Connect to SQL Server"):
                if sql_username and sql_password:
                    with st.spinner("Connecting to SQL Server..."):
                        excel_file = fetch_sql_data(sql_server, sql_database, sql_username, sql_password, auth_type)
                        if excel_file:
                            st.success("✅ SQL Server data fetched successfully!")
                            st.session_state.sql_data = excel_file
                            # Store credentials for Q Developer
                            store_credentials(sql_server, sql_database, sql_username, sql_password, auth_type)
                        else:
                            st.error("❌ Failed to fetch SQL Server data")
                else:
                    st.warning("Please enter username and password")
            
            # Use cached SQL data if available
            if 'sql_data' in st.session_state:
                excel_file = st.session_state.sql_data
                st.info("📊 Using cached SQL Server data")
        
        # Config file upload (optional) - only show if Excel is selected
        if data_source == "📁 Upload Excel File":
            config_file = st.file_uploader(
                "Upload Config File (Optional)",
                type=['json'],
                help="Upload custom config.json or use default settings"
            )
        else:
            config_file = None
        
        # Config editor section
        st.subheader("⚙️ Configuration Editor")
        
        # Initialize config in session state if not exists
        if 'current_config' not in st.session_state:
            st.session_state.current_config = get_default_config()
            st.session_state.original_config = st.session_state.current_config.copy()
        
        # Handle config file upload
        if config_file:
            uploaded_config = json.load(config_file)
            st.session_state.current_config = uploaded_config
            st.session_state.original_config = uploaded_config.copy()
            st.success("✅ Config file uploaded and loaded!")
        
        # Config reset buttons
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔄 Reset to Original"):
                st.session_state.current_config = st.session_state.original_config.copy()
                st.success("Config reset to original settings")
                st.rerun()
        
        with col2:
            if st.button("🏭 Reset to Defaults"):
                st.session_state.current_config = get_default_config()
                st.success("Config reset to default settings")
                st.rerun()
        
        # Editable config sections
        config = st.session_state.current_config
        
        with st.expander("📊 Thresholds", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                config['thresholds']['efficiency_threshold'] = st.slider(
                    "Efficiency Threshold", 0.0, 1.0, 
                    config['thresholds'].get('efficiency_threshold', 0.15), 0.01,
                    help="Threshold for consultant efficiency scoring (15% = 0.15)"
                )
                config['thresholds']['success_threshold'] = st.slider(
                    "Success Threshold", 0.0, 1.0, 
                    config['thresholds'].get('success_threshold', 0.3), 0.01,
                    help="Threshold for project success determination (30% = 0.3)"
                )
            with col2:
                config['thresholds']['green_threshold'] = st.slider(
                    "Green Threshold", -0.5, 0.0, 
                    config['thresholds'].get('green_threshold', -0.1), 0.01,
                    help="Under budget threshold for green highlighting"
                )
                config['thresholds']['red_threshold'] = st.slider(
                    "Red Threshold", 0.0, 1.0, 
                    config['thresholds'].get('red_threshold', 0.3), 0.01,
                    help="Over budget threshold for red highlighting"
                )
        
        with st.expander("🎯 Analysis Settings"):
            col1, col2 = st.columns(2)
            with col1:
                config['client_analysis']['enable_client_analysis'] = st.checkbox(
                    "Enable Customer Analysis", 
                    config.get('client_analysis', {}).get('enable_client_analysis', False)
                )
                config['client_analysis']['min_projects_threshold'] = st.number_input(
                    "Min Projects (Customer)", 1, 50, 
                    config.get('client_analysis', {}).get('min_projects_threshold', 3)
                )
            with col2:
                if 'das_plus_analysis' not in config:
                    config['das_plus_analysis'] = {}
                config['das_plus_analysis']['enable_das_plus'] = st.checkbox(
                    "Enable DAS+ Analysis", 
                    config.get('das_plus_analysis', {}).get('enable_das_plus', False)
                )
                config['das_plus_analysis']['min_projects_for_review'] = st.number_input(
                    "Min Projects (DAS+)", 1, 20, 
                    config.get('das_plus_analysis', {}).get('min_projects_for_review', 3)
                )
        
        with st.expander("📅 Date Filtering"):
            if 'project_filtering' not in config:
                config['project_filtering'] = {}
            
            config['project_filtering']['enable_date_filter'] = st.checkbox(
                "Enable Date Filter", 
                config.get('project_filtering', {}).get('enable_date_filter', False)
            )
            
            if config['project_filtering']['enable_date_filter']:
                filter_type = st.radio(
                    "Filter Type",
                    ["Days from Today", "Specific Date"],
                    index=0 if config.get('project_filtering', {}).get('filter_type', 'days') == 'days' else 1
                )
                
                if filter_type == "Days from Today":
                    config['project_filtering']['filter_type'] = 'days'
                    config['project_filtering']['days_from_today'] = st.number_input(
                        "Days from Today", 1, 3650, 
                        config.get('project_filtering', {}).get('days_from_today', 365)
                    )
                else:
                    config['project_filtering']['filter_type'] = 'date'
                    from datetime import datetime as dt, date
                    default_date = dt.strptime(config.get('project_filtering', {}).get('specific_date', get_current_quarter_start()), '%Y-%m-%d').date()
                    selected_date = st.date_input(
                        "Projects After Date",
                        value=default_date
                    )
                    config['project_filtering']['specific_date'] = selected_date.strftime('%Y-%m-%d')
        
        # Download current config
        config_json = json.dumps(config, indent=2)
        st.download_button(
            label="📥 Download Current Config",
            data=config_json,
            file_name=f"config_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            mime="application/json"
        )
        
        # Optional filter files
        st.subheader("📋 Filter Files (Optional)")
        engineers_file = st.file_uploader("Engineers List", type=['txt'])
        exclude_file = st.file_uploader("Exclusions", type=['csv'])
        sa_file = st.file_uploader("Solution Architects", type=['txt'])
    
    # Show tabs if analyzer is available (prioritize analyzer over file state)
    if 'analyzer' in st.session_state:
        debug_log("Analyzer found in session state")
        analyzer = st.session_state.analyzer
        
        # Update analyzer config with current session config
        analyzer.config = st.session_state.current_config.copy()
        debug_log(f"Updated analyzer config - success_threshold: {analyzer.config['thresholds']['success_threshold']}")
        
        # Update analyzer with current session exclusions
        if 'session_exclusions' in st.session_state:
            debug_log(f"Applying {len(st.session_state.session_exclusions)} session exclusions")
            # Reset exclusions to original + session exclusions
            original_exclusions = [excl for excl in analyzer.exclusions if excl not in getattr(st.session_state, 'session_exclusions', [])]
            analyzer.exclusions = original_exclusions + st.session_state.session_exclusions
        
        # Main content area with tabs
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "📈 Dashboard", "👥 Consultants", "🏗️ Solution Architects", 
            "🏢 Customers", "🎯 DAS+ Analysis", "⚖️ Consultant vs SA"
        ])
        
        with tab1:
            show_dashboard(analyzer)
        
        with tab2:
            show_consultant_analysis(analyzer)
        
        with tab3:
            show_sa_analysis(analyzer)
        
        with tab4:
            show_customer_analysis(analyzer)
        
        with tab5:
            show_das_analysis(analyzer)
        
        with tab6:
            show_consultant_sa_comparison(analyzer)
    
    elif excel_file is not None:
        # Show run button
        st.markdown("---")
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            run_analysis = st.button("🚀 Run Analysis", type="primary", use_container_width=True)
        
        if run_analysis:
            debug_log("Run Analysis button clicked")
            # Process the uploaded files
            analyzer = process_files(excel_file, config_file, engineers_file, exclude_file, sa_file)
            
            if analyzer:
                # Store analyzer in session state
                st.session_state.analyzer = analyzer
                debug_log("Analyzer stored in session state")
                st.success("✅ Analysis completed! Explore the tabs below.")
                st.rerun()
        
        debug_log("Excel file uploaded, waiting for Run Analysis")
        st.info("👆 Click 'Run Analysis' button above to process your files")
        
        # Show debug log if enabled
        if st.session_state.get('debug_mode', False) and 'debug_log' in st.session_state:
            with st.expander("🔍 Debug Log", expanded=False):
                for log_entry in reversed(st.session_state.debug_log[-20:]):
                    st.text(log_entry)
    else:
        # Welcome screen
        st.info("👆 Please upload a Project Financials Excel file to begin analysis")
        
        # Session and file management buttons
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🗑️ Clear Session"):
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                st.rerun()
        
        with col2:
            if st.button("📋 Clean Old Files"):
                cleanup_result = cleanup_old_excel_files()
                if cleanup_result['deleted_count'] > 0:
                    st.success(f"✅ Cleaned up {cleanup_result['deleted_count']} old files")
                    with st.expander("Deleted Files"):
                        for file in cleanup_result['deleted_files']:
                            st.text(f"- {file}")
                else:
                    st.info("📋 No old files to clean up")
    
    # Show file cleanup option when files are uploaded
    if excel_file is not None:
        st.sidebar.markdown("---")
        st.sidebar.subheader("📋 File Management")
        
        if st.sidebar.button("🗑️ Clean Old Reports"):
            cleanup_result = cleanup_old_excel_files()
            if cleanup_result['deleted_count'] > 0:
                st.sidebar.success(f"Cleaned up {cleanup_result['deleted_count']} files")
            else:
                st.sidebar.info("No old files to clean up")
        
        # Only clear if no analyzer exists
        debug_log("No Excel file and no analyzer - showing welcome")
        if 'selected_consultant' in st.session_state:
            del st.session_state.selected_consultant
            debug_log("Selected consultant cleared")
        
        # Show sample data format
        st.subheader("📋 Expected Excel Format")
        sample_data = pd.DataFrame({
            'Job Number': ['J00001', 'J00002', 'J00003'],
            'Resources Engaged': ['John Doe', 'Jane Smith, Bob Wilson', 'Alice Johnson'],
            'Budget Hrs': [100, 200, 150],
            'Total Hrs Posted': [95, 220, 180],
            'Project Status': ['Closed', 'Open', 'Implementation'],
            'Project Complete %': [100, 85, 60],
            'Customer': ['Microsoft', 'Google', 'Amazon']
        })
        st.dataframe(sample_data)

def process_files(excel_file, config_file, engineers_file, exclude_file, sa_file):
    """Process uploaded files and return configured analyzer"""
    try:
        with st.spinner("Processing files..."):
            # Create temporary directory for files
            temp_dir = tempfile.mkdtemp()
            
            # Save Excel file
            excel_path = os.path.join(temp_dir, "data.xlsx")
            with open(excel_path, "wb") as f:
                f.write(excel_file.getbuffer())
            
            # Use current config from session state
            config = st.session_state.current_config.copy()
            config['files']['excel_file'] = excel_path
            config['files']['header_row'] = st.sidebar.number_input(
                "Header Row (0-based)", min_value=0, max_value=20, 
                value=config['files'].get('header_row', 13)
            )
            
            # Create config file
            config_path = os.path.join(temp_dir, "config.json")
            
            # Handle optional files
            if engineers_file:
                engineers_path = os.path.join(temp_dir, "engineers.txt")
                with open(engineers_path, "wb") as f:
                    f.write(engineers_file.getbuffer())
                config['files']['engineers_file'] = engineers_path
            
            if exclude_file:
                exclude_path = os.path.join(temp_dir, "exclude.csv")
                with open(exclude_path, "wb") as f:
                    f.write(exclude_file.getbuffer())
                config['files']['exclude_file'] = exclude_path
            
            if sa_file:
                sa_path = os.path.join(temp_dir, "sa.txt")
                with open(sa_path, "wb") as f:
                    f.write(sa_file.getbuffer())
                config['files']['solution_architects_file'] = sa_path
            
            # Save updated config
            with open(config_path, 'w') as f:
                json.dump(config, f, indent=2)
            
            # Initialize analyzer with Streamlit mode
            analyzer = ConsultantMetricsAnalyzer(config_path)
            analyzer.streamlit_mode = True  # Disable interactive prompts
            
            # Add session exclusions to analyzer
            if 'session_exclusions' in st.session_state:
                analyzer.exclusions.extend(st.session_state.session_exclusions)
            
            # Create progress tracking
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                # Step 1: Load data
                status_text.text("Loading Excel data...")
                progress_bar.progress(25)
                analyzer.load_data()
                
                # Step 2: Analyze consultants
                status_text.text("Analyzing consultant performance...")
                progress_bar.progress(50)
                analyzer.analyze_consultants()
                
                # Step 3: Complete processing
                status_text.text("Finalizing analysis...")
                progress_bar.progress(100)
                
            except Exception as e:
                import traceback
                st.error(f"Analysis Error: {str(e)}")
                with st.expander("Error Details", expanded=True):
                    st.text(traceback.format_exc())
                return None
            
            # Clear progress indicators
            progress_bar.empty()
            status_text.empty()
            
            st.success("✅ Files processed successfully!")
            return analyzer
            
    except Exception as e:
        st.error(f"❌ Error processing files: {str(e)}")
        return None

def get_current_quarter_start():
    """Get the start date of the current quarter"""
    from datetime import datetime
    current_date = datetime.now()
    current_quarter_start = datetime(current_date.year, ((current_date.month - 1) // 3) * 3 + 1, 1)
    return current_quarter_start.strftime('%Y-%m-%d')

def get_default_config():
    """Get default configuration dictionary"""
    return {
        "thresholds": {
            "efficiency_threshold": 0.15,
            "success_threshold": 0.3,
            "green_threshold": -0.1,
            "yellow_threshold": 0.1,
            "red_threshold": 0.3
        },
        "scoring": {
            "hours_per_bonus_point": 1000,
            "max_hours_multiplier": 2.0,
            "bonus_points_per_1000_hours": 25
        },
        "files": {
            "excel_file": "",
            "engineers_file": "",
            "exclude_file": "",
            "solution_architects_file": "",
            "header_row": 13
        },
        "display": {
            "top_performers_count": 5
        },
        "solution_architect_scoring": {
            "hours_per_multiplier": 1000,
            "max_volume_multiplier": 3.0
        },
        "client_analysis": {
            "enable_client_analysis": False,
            "min_projects_threshold": 3,
            "track_consultant_client_performance": True
        },
        "das_plus_analysis": {
            "enable_das_plus": False,
            "min_projects_for_review": 3,
            "review_das_min": 0.3,
            "review_das_max": 0.9,
            "sample_projects_per_consultant": 2,
            "current_year_offset": 0
        },
        "project_filtering": {
            "enable_date_filter": False,
            "filter_type": "date",
            "days_from_today": 365,
            "specific_date": get_current_quarter_start(),
            "exclude_closed_before_date": True
        }
    }

def create_default_config(temp_dir):
    """Create default configuration file"""
    default_config = get_default_config()
    
    config_path = os.path.join(temp_dir, "config.json")
    with open(config_path, 'w') as f:
        json.dump(default_config, f, indent=2)
    
    return config_path

def show_dashboard(analyzer):
    """Show main dashboard with key metrics"""
    st.header("📈 Performance Dashboard")
    
    # Show current date filter status
    if analyzer.config.get('project_filtering', {}).get('enable_date_filter', False):
        filter_config = analyzer.config['project_filtering']
        if filter_config.get('filter_type') == 'days':
            from datetime import datetime, timedelta
            cutoff_date = datetime.now() - timedelta(days=filter_config.get('days_from_today', 365))
            st.info(f"📅 **Date Filter Active**: Projects from {cutoff_date.strftime('%Y-%m-%d')} onwards ({filter_config.get('days_from_today', 365)} days)")
        else:
            st.info(f"📅 **Date Filter Active**: Projects from {filter_config.get('specific_date', get_current_quarter_start())} onwards")
    else:
        st.info("📅 **Date Filter**: All projects included (no date filtering)")
    
    # Key metrics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_consultants = len(analyzer.consultant_metrics)
        st.metric("Total Consultants", total_consultants)
    
    with col2:
        total_hours = sum(m['total_hours'] for m in analyzer.consultant_metrics.values())
        st.metric("Total Hours", f"{total_hours:,.0f}")
    
    with col3:
        avg_efficiency = np.mean([m['efficiency_score'] for m in analyzer.consultant_metrics.values()])
        st.metric("Average Efficiency", f"{avg_efficiency:.1f}%")
    
    with col4:
        total_projects = sum(m['unique_projects'] for m in analyzer.consultant_metrics.values())
        st.metric("Total Projects", total_projects)
    
    # Charts
    col1, col2 = st.columns(2)
    
    with col1:
        # Success score by consultant
        consultant_names = list(analyzer.consultant_metrics.keys())
        success_scores = [m['success_ratio'] for m in analyzer.consultant_metrics.values()]
        
        fig = px.bar(
            x=consultant_names,
            y=success_scores,
            title="Success Score by Consultant",
            labels={'x': 'Consultant', 'y': 'Success Score (%)'}
        )
        fig.update_layout(xaxis_tickangle=45)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Top performers
        sorted_consultants = sorted(
            analyzer.consultant_metrics.items(),
            key=lambda x: x[1]['efficiency_score'],
            reverse=True
        )[:10]
        
        names = [item[0] for item in sorted_consultants]
        scores = [item[1]['efficiency_score'] for item in sorted_consultants]
        
        fig = px.bar(
            x=scores,
            y=names,
            orientation='h',
            title="Top 10 Consultants by Efficiency",
            labels={'x': 'Efficiency Score (%)', 'y': 'Consultant'}
        )
        fig.update_layout(yaxis={'categoryorder': 'total ascending'})
        st.plotly_chart(fig, use_container_width=True)

def show_consultant_analysis(analyzer):
    """Show detailed consultant analysis with filtering"""
    st.header("👥 Consultant Performance Analysis")
    
    # Filters
    col1, col2, col3 = st.columns(3)
    
    with col1:
        min_projects = st.slider("Minimum Projects", 1, 50, 3)
    
    with col2:
        min_efficiency = st.slider("Minimum Efficiency %", 0, 100, 0)
    
    with col3:
        selected_consultants = st.multiselect(
            "Select Consultants",
            options=list(analyzer.consultant_metrics.keys()),
            default=[]
        )
    
    # Filter data
    filtered_data = []
    for consultant, metrics in analyzer.consultant_metrics.items():
        if (metrics['unique_projects'] >= min_projects and 
            metrics['efficiency_score'] >= min_efficiency and
            (not selected_consultants or consultant in selected_consultants)):
            
            filtered_data.append({
                'Consultant': consultant,
                'Projects': metrics['unique_projects'],
                'Hours': metrics['total_hours'],
                'Efficiency %': f"{metrics['efficiency_score']:.1f}%",
                'Goal Attainment %': f"{metrics['success_ratio']:.1f}%",
                'Within Budget': metrics['projects_within_budget'],
                'Over Budget': metrics['projects_over_budget'],
                'On Hold': metrics.get('projects_on_hold', 0)
            })
    
    if filtered_data:
        df = pd.DataFrame(filtered_data)
        
        # Display consultant table with selection
        event = st.dataframe(
            df, 
            use_container_width=True,
            on_select="rerun",
            selection_mode="single-row",
            key="consultant_table"
        )
        
        # Store selected consultant in session state
        if event.selection and len(event.selection.rows) > 0:
            selected_row = event.selection.rows[0]
            selected_consultant = df.iloc[selected_row]['Consultant']
            st.session_state.selected_consultant = selected_consultant
            debug_log(f"Consultant selected: {selected_consultant}")
        
        # Show project details when consultant is selected
        if 'selected_consultant' in st.session_state:
            selected_consultant = st.session_state.selected_consultant
            debug_log(f"Showing details for consultant: {selected_consultant}")
            
            st.subheader(f"📋 Projects for {selected_consultant}")
            
            # Get project details for selected consultant
            project_details = get_consultant_projects(analyzer, selected_consultant)
            
            if not project_details.empty:
                # Add variance calculation and color coding info
                project_details = add_project_variance(project_details)
                
                # Display projects table with clickable links
                display_projects_with_links(project_details, analyzer.config)
                
                # Add exclusion functionality
                st.subheader("🚫 Manage Project Exclusions")
                
                # Show existing exclusions for this consultant
                existing_exclusions = get_consultant_exclusions(selected_consultant)
                if existing_exclusions:
                    st.write("**Current Exclusions:**")
                    
                    for exclusion in existing_exclusions:
                        col1, col2, col3 = st.columns([3, 1, 1])
                        with col1:
                            exclusion_type = "📁 File" if exclusion['type'] == 'file' else "🔒 Session"
                            st.text(f"{exclusion_type}: {exclusion['project']}")
                        with col2:
                            if st.button("Remove", key=f"remove_{selected_consultant}_{exclusion['project']}_{exclusion['type']}"):
                                remove_exclusion(selected_consultant, exclusion['project'], exclusion['type'])
                                if 'analyzer' in st.session_state:
                                    st.session_state.analyzer.analyze_consultants()
                                # Clear selected consultant to refresh project table
                                if 'selected_consultant' in st.session_state:
                                    del st.session_state.selected_consultant
                                st.success(f"Removed exclusion for {exclusion['project']}")
                                st.rerun()
                
                st.write("**Add New Exclusions:**")
                
                # Project selection for exclusion
                if 'Job Number' in project_details.columns:
                    project_options = project_details['Job Number'].tolist()
                    selected_projects = st.multiselect(
                        f"Select projects to exclude for {selected_consultant}",
                        options=project_options,
                        key=f"exclude_{selected_consultant}"
                    )
                    
                    if selected_projects:
                        col1, col2 = st.columns(2)
                        with col1:
                            if st.button("🔒 Exclude for Session Only", key=f"session_{selected_consultant}"):
                                add_session_exclusions(selected_consultant, selected_projects)
                                # Recalculate metrics with new exclusions
                                if 'analyzer' in st.session_state:
                                    st.session_state.analyzer.analyze_consultants()
                                # Clear selected consultant to refresh project table
                                if 'selected_consultant' in st.session_state:
                                    del st.session_state.selected_consultant
                                st.success(f"Excluded {len(selected_projects)} projects for session")
                                st.rerun()
                        
                        with col2:
                            if st.button("💾 Exclude & Update File", key=f"file_{selected_consultant}"):
                                if update_exclusion_file(selected_consultant, selected_projects):
                                    # Recalculate metrics with new exclusions
                                    if 'analyzer' in st.session_state:
                                        st.session_state.analyzer.analyze_consultants()
                                    # Clear selected consultant to refresh project table
                                    if 'selected_consultant' in st.session_state:
                                        del st.session_state.selected_consultant
                                    st.success(f"Updated exclusion file with {len(selected_projects)} projects")
                                    st.rerun()
                                else:
                                    st.error("Failed to update exclusion file")
                
                # Projects table displayed by display_projects_with_links function above
                
                # Project summary metrics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Projects", len(project_details))
                with col2:
                    total_budget = project_details['Budget Hrs'].sum()
                    st.metric("Total Budget Hours", f"{total_budget:,.0f}")
                with col3:
                    total_actual = project_details['Total Hrs Posted'].sum()
                    st.metric("Total Actual Hours", f"{total_actual:,.0f}")
                with col4:
                    overall_variance = ((total_actual - total_budget) / total_budget * 100) if total_budget > 0 else 0
                    st.metric("Overall Variance", f"{overall_variance:+.1f}%")
                
                # Performance charts for consultant
                if 'Variance %' in project_details.columns:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # Variance distribution
                        fig_var = px.histogram(
                            project_details,
                            x='Variance %',
                            title=f"{selected_consultant} - Variance Distribution",
                            labels={'Variance %': 'Variance (%)', 'count': 'Number of Projects'}
                        )
                        st.plotly_chart(fig_var, use_container_width=True)
                    
                    with col2:
                        # Budget vs Actual scatter
                        if 'Budget Hrs' in project_details.columns and 'Total Hrs Posted' in project_details.columns:
                            fig_scatter = px.scatter(
                                project_details,
                                x='Budget Hrs',
                                y='Total Hrs Posted',
                                title=f"{selected_consultant} - Budget vs Actual",
                                labels={'Budget Hrs': 'Budget Hours', 'Total Hrs Posted': 'Actual Hours'}
                            )
                            # Add diagonal line for perfect accuracy
                            max_hours = max(project_details['Budget Hrs'].max(), project_details['Total Hrs Posted'].max())
                            fig_scatter.add_shape(
                                type="line",
                                x0=0, y0=0, x1=max_hours, y1=max_hours,
                                line=dict(color="red", dash="dash")
                            )
                            st.plotly_chart(fig_scatter, use_container_width=True)
                
                # Random Project Review Selection
                st.subheader(f"🎯 Random Projects for Review - {selected_consultant}")
                
                # Quarter selection for review
                review_quarter = st.radio(
                    "Select Quarter for Review",
                    ["Current Quarter", "Previous Quarter"],
                    key=f"quarter_{selected_consultant}",
                    help="Choose which quarter to select review projects from"
                )
                
                use_previous_quarter = (review_quarter == "Previous Quarter")
                review_projects = select_review_projects(project_details, use_previous_quarter)
                
                if not review_projects.empty:
                    st.write(f"**Selected for Management Review ({review_quarter}):**")
                    for idx, row in review_projects.iterrows():
                        variance = row.get('Variance %', 'N/A')
                        status = "🟢 Good" if isinstance(variance, str) and "-" in variance else "🔴 Poor" if isinstance(variance, str) and "+" in variance and float(variance.replace('%', '').replace('+', '')) > 30 else "🟡 Average"
                        st.write(f"• **{row.get('Job Number', 'N/A')}** - {row.get('Job Description', 'N/A')} ({status}, Variance: {variance})")
                    
                    display_projects_with_links(review_projects, analyzer.config)
                else:
                    st.info(f"No projects available for review selection in {review_quarter.lower()}")
                
                # Download project details
                csv_projects = project_details.to_csv(index=False)
                st.download_button(
                    label=f"📥 Download {selected_consultant} Projects",
                    data=csv_projects,
                    file_name=f"{selected_consultant.replace(' ', '_')}_projects_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
            else:
                st.info(f"No project details available for {selected_consultant}")
        
        # Download summary button
        csv = df.to_csv(index=False)
        st.download_button(
            label="📥 Download Summary CSV",
            data=csv,
            file_name=f"consultant_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )
    else:
        st.info("No consultants match the selected filters")

def show_sa_analysis(analyzer):
    """Show Solution Architect analysis"""
    st.header("🏗️ Solution Architect Performance")
    
    if not analyzer.sa_metrics:
        st.info("No Solution Architect data available")
        return
    
    # SA data
    sa_data = []
    for sa_name, metrics in analyzer.sa_metrics.items():
        sa_data.append({
            'Solution Architect': sa_name,
            'Projects': metrics['total_projects'],
            'Goal Attainment %': f"{metrics['success_rate']:.1f}%",
            'Budgeted Hours': metrics['total_budgeted_hours'],
            'Actual Hours': metrics['total_actual_hours'],
            'Variance %': f"{((metrics['total_actual_hours'] - metrics['total_budgeted_hours']) / metrics['total_budgeted_hours'] * 100):+.1f}%" if metrics['total_budgeted_hours'] > 0 else "0.0%"
        })
    
    df = pd.DataFrame(sa_data)
    
    # Display SA table with selection
    event = st.dataframe(
        df, 
        use_container_width=True,
        on_select="rerun",
        selection_mode="single-row"
    )
    
    # Show project details when SA is selected
    if event.selection and len(event.selection.rows) > 0:
        selected_row = event.selection.rows[0]
        selected_sa = df.iloc[selected_row]['Solution Architect']
        
        st.subheader(f"🏗️ Projects for {selected_sa}")
        
        # Get project details for selected SA
        sa_project_details = get_sa_projects(analyzer, selected_sa)
        
        if not sa_project_details.empty:
            # Add variance calculation
            sa_project_details = add_project_variance(sa_project_details)
            
            # Display projects table with color highlighting and exclusion options
            styled_sa_projects = apply_color_styling(sa_project_details, analyzer.config)
            
            # Add exclusion functionality for SA
            st.subheader("🚫 Manage Project Exclusions")
            
            # Show existing exclusions for this SA
            existing_sa_exclusions = get_consultant_exclusions(selected_sa)
            if existing_sa_exclusions:
                st.write("**Current Exclusions:**")
                
                for exclusion in existing_sa_exclusions:
                    col1, col2, col3 = st.columns([3, 1, 1])
                    with col1:
                        exclusion_type = "📁 File" if exclusion['type'] == 'file' else "🔒 Session"
                        st.text(f"{exclusion_type}: {exclusion['project']}")
                    with col2:
                        if st.button("Remove", key=f"remove_sa_{selected_sa}_{exclusion['project']}_{exclusion['type']}"):
                            remove_exclusion(selected_sa, exclusion['project'], exclusion['type'])
                            if 'analyzer' in st.session_state:
                                st.session_state.analyzer.analyze_consultants()
                            st.success(f"Removed exclusion for {exclusion['project']}")
                            st.rerun()
            
            st.write("**Add New Exclusions:**")
            
            # Project selection for exclusion
            if 'Job Number' in sa_project_details.columns:
                sa_project_options = sa_project_details['Job Number'].tolist()
                selected_sa_projects = st.multiselect(
                    f"Select projects to exclude for {selected_sa}",
                    options=sa_project_options,
                    key=f"exclude_sa_{selected_sa}"
                )
                
                if selected_sa_projects:
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button("🔒 Exclude for Session Only", key=f"session_sa_{selected_sa}"):
                            add_session_exclusions(selected_sa, selected_sa_projects)
                            # Recalculate metrics with new exclusions
                            if 'analyzer' in st.session_state:
                                st.session_state.analyzer.analyze_consultants()
                            st.success(f"Excluded {len(selected_sa_projects)} projects for session")
                            st.rerun()
                    
                    with col2:
                        if st.button("💾 Exclude & Update File", key=f"file_sa_{selected_sa}"):
                            if update_exclusion_file(selected_sa, selected_sa_projects):
                                # Recalculate metrics with new exclusions
                                if 'analyzer' in st.session_state:
                                    st.session_state.analyzer.analyze_consultants()
                                st.success(f"Updated exclusion file with {len(selected_sa_projects)} projects")
                                st.rerun()
                            else:
                                st.error("Failed to update exclusion file")
            
            st.dataframe(styled_sa_projects, use_container_width=True)
            
            # Project summary metrics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Projects", len(sa_project_details))
            with col2:
                total_budget = sa_project_details['Budget Hrs'].sum()
                st.metric("Total Budget Hours", f"{total_budget:,.0f}")
            with col3:
                total_actual = sa_project_details['Total Hrs Posted'].sum()
                st.metric("Total Actual Hours", f"{total_actual:,.0f}")
            with col4:
                overall_variance = ((total_actual - total_budget) / total_budget * 100) if total_budget > 0 else 0
                st.metric("Overall Variance", f"{overall_variance:+.1f}%")
            
            # Download project details
            csv_sa_projects = sa_project_details.to_csv(index=False)
            st.download_button(
                label=f"📥 Download {selected_sa} Projects",
                data=csv_sa_projects,
                file_name=f"{selected_sa.replace(' ', '_')}_SA_projects_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        else:
            st.info(f"No project details available for {selected_sa}")
    
    # SA performance chart
    fig = px.scatter(
        df,
        x='Budgeted Hours',
        y='Goal Attainment %',
        size='Projects',
        hover_name='Solution Architect',
        title="SA Performance: Goal Attainment vs Volume"
    )
    st.plotly_chart(fig, use_container_width=True)

def show_customer_analysis(analyzer):
    """Show customer performance analysis"""
    st.header("🏢 Customer Performance Analysis")
    
    # Enable customer analysis for web interface
    if not analyzer.config.get('client_analysis', {}).get('enable_client_analysis', False):
        if st.button("🔧 Enable Customer Analysis"):
            analyzer.config['client_analysis'] = {
                'enable_client_analysis': True,
                'min_projects_threshold': 3,
                'track_consultant_client_performance': True
            }
            st.rerun()
    
    if analyzer.config.get('client_analysis', {}).get('enable_client_analysis', False):
        try:
            # Run customer analysis and get data
            customer_data, consultant_customer_data = run_customer_analysis_data(analyzer)
            
            if customer_data:
                # Customer performance filters
                col1, col2 = st.columns(2)
                with col1:
                    min_projects_customer = st.slider("Minimum Projects (Customer)", 1, 20, 3)
                with col2:
                    min_success_rate = st.slider("Minimum Success Rate %", 0, 100, 0)
                
                # Filter customer data
                filtered_customers = [
                    customer for customer in customer_data 
                    if customer['Total_Projects'] >= min_projects_customer and 
                       customer['Goal_Attainment'] >= min_success_rate
                ]
                
                if filtered_customers:
                    # Customer performance table
                    customer_df = pd.DataFrame(filtered_customers)
                    
                    # Display customer table with selection
                    event = st.dataframe(
                        customer_df, 
                        use_container_width=True,
                        on_select="rerun",
                        selection_mode="single-row"
                    )
                    
                    # Show consultant-customer combinations when customer selected
                    if event.selection and len(event.selection.rows) > 0:
                        selected_row = event.selection.rows[0]
                        selected_customer = customer_df.iloc[selected_row]['Customer']
                        
                        st.subheader(f"👥 Consultant Performance with {selected_customer}")
                        
                        # Filter consultant-customer data for selected customer
                        customer_consultants = [
                            combo for combo in consultant_customer_data 
                            if combo['Customer'] == selected_customer
                        ]
                        
                        if customer_consultants:
                            combo_df = pd.DataFrame(customer_consultants)
                            st.dataframe(combo_df, use_container_width=True)
                            
                            # Customer metrics
                            col1, col2, col3, col4 = st.columns(4)
                            selected_metrics = next(c for c in filtered_customers if c['Customer'] == selected_customer)
                            
                            with col1:
                                st.metric("Total Projects", selected_metrics['Total_Projects'])
                            with col2:
                                st.metric("Goal Attainment", f"{selected_metrics['Goal_Attainment']:.1f}%")
                            with col3:
                                st.metric("Efficiency Rate", f"{selected_metrics['Efficiency_Rate']:.1f}%")
                            with col4:
                                st.metric("Avg Variance", f"{selected_metrics['Avg_Variance_Pct']:+.1f}%")
                        else:
                            st.info(f"No consultant combinations found for {selected_customer}")
                    
                    # Customer performance chart
                    fig = px.scatter(
                        customer_df,
                        x='Total_Projects',
                        y='Goal_Attainment',
                        size='Total_Budgeted_Hours',
                        hover_name='Customer',
                        title="Customer Performance: Goal Attainment vs Project Volume",
                        labels={'Goal_Attainment': 'Goal Attainment (%)', 'Total_Projects': 'Number of Projects'}
                    )
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Download buttons
                    col1, col2 = st.columns(2)
                    with col1:
                        csv_customers = customer_df.to_csv(index=False)
                        st.download_button(
                            label="📥 Download Customer Data",
                            data=csv_customers,
                            file_name=f"customer_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                            mime="text/csv"
                        )
                    
                    if consultant_customer_data:
                        with col2:
                            csv_combos = pd.DataFrame(consultant_customer_data).to_csv(index=False)
                            st.download_button(
                                label="📥 Download Consultant-Customer Data",
                                data=csv_combos,
                                file_name=f"consultant_customer_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                mime="text/csv"
                            )
                else:
                    st.info("No customers match the selected filters")
            else:
                st.info("No customer data available")
            
        except Exception as e:
            st.error(f"❌ Error running customer analysis: {str(e)}")
    else:
        st.info("💡 Customer analysis disabled - click button above to enable")

def show_das_analysis(analyzer):
    """Show DAS+ analysis with interactive filtering"""
    st.header("🎯 DAS+ Analysis")
    
    # Enable DAS+ analysis for web interface
    if not analyzer.config.get('das_plus_analysis', {}).get('enable_das_plus', False):
        if st.button("🔧 Enable DAS+ Analysis"):
            analyzer.config['das_plus_analysis'] = {
                'enable_das_plus': True,
                'min_projects_for_review': 3,
                'review_das_min': 0.3,
                'review_das_max': 0.9,
                'sample_projects_per_consultant': 2,
                'current_year_offset': 0
            }
            st.rerun()
    
    if analyzer.config.get('das_plus_analysis', {}).get('enable_das_plus', False):
        # DAS+ filters
        col1, col2, col3 = st.columns(3)
        
        with col1:
            min_das = st.slider("Minimum DAS+ Score", 0.0, 1.0, 0.0, 0.1)
        
        with col2:
            max_das = st.slider("Maximum DAS+ Score", 0.0, 1.0, 1.0, 0.1)
        
        with col3:
            min_projects_das = st.slider("Minimum Projects (DAS+)", 1, 20, 3)
        
        if st.button("🚀 Run DAS+ Analysis"):
            try:
                with st.spinner("Calculating DAS+ scores..."):
                    # Get DAS+ data and store in session state
                    das_data = run_das_plus_analysis_data(analyzer, min_projects_das)
                    st.session_state.das_data = das_data
                
                st.success("✅ DAS+ analysis completed!")
                
            except Exception as e:
                st.error(f"❌ Error running DAS+ analysis: {str(e)}")
        
        # Display DAS+ results if available
        if 'das_data' in st.session_state and st.session_state.das_data:
            das_data = st.session_state.das_data
            
            # Filter DAS+ data
            filtered_das = [
                consultant for consultant in das_data 
                if min_das <= consultant['Avg_DAS_Plus'] <= max_das
            ]
            
            if filtered_das:
                # DAS+ performance table
                das_df = pd.DataFrame(filtered_das)
                
                # Display DAS+ table with selection
                event = st.dataframe(
                    das_df, 
                    use_container_width=True,
                    on_select="rerun",
                    selection_mode="single-row"
                )
                
                # Show consultant details when selected
                if event.selection and len(event.selection.rows) > 0:
                    selected_row = event.selection.rows[0]
                    selected_consultant = das_df.iloc[selected_row]['Consultant']
                    
                    st.subheader(f"🎯 DAS+ Details for {selected_consultant}")
                    
                    # DAS+ metrics for selected consultant
                    col1, col2, col3, col4 = st.columns(4)
                    selected_metrics = das_df.iloc[selected_row]
                    
                    with col1:
                        st.metric("Projects Analyzed", selected_metrics['Project_Count'])
                    with col2:
                        st.metric("Average DAS+", f"{selected_metrics['Avg_DAS_Plus']:.3f}")
                    with col3:
                        st.metric("High Performance", f"{selected_metrics['High_DAS_Count']} projects")
                    with col4:
                        st.metric("Low Performance", f"{selected_metrics['Low_DAS_Count']} projects")
                    
                    # Show individual DAS+ projects for selected consultant
                    st.subheader(f"📈 DAS+ Projects for {selected_consultant}")
                    consultant_das_projects = get_consultant_das_projects(analyzer, selected_consultant)
                    
                    if not consultant_das_projects.empty:
                        # Color code DAS+ projects
                        def color_das_score(val):
                            try:
                                score = float(val)
                                if score >= 0.85:
                                    return 'background-color: #90EE90'  # Green
                                elif score < 0.75:
                                    return 'background-color: #FF6B6B'  # Red
                                elif score < 0.85:
                                    return 'background-color: #FFFF99'  # Yellow
                                else:
                                    return ''
                            except:
                                return ''
                        
                        styled_das = consultant_das_projects.style.applymap(color_das_score, subset=['DAS_Plus'])
                        st.dataframe(styled_das, use_container_width=True)
                    else:
                        st.info(f"No DAS+ project data available for {selected_consultant}")
                
                # DAS+ performance chart
                fig = px.scatter(
                    das_df,
                    x='Project_Count',
                    y='Avg_DAS_Plus',
                    size='High_DAS_Count',
                    color='Low_DAS_Count',
                    hover_name='Consultant',
                    title="DAS+ Performance: Average Score vs Project Volume",
                    labels={'Avg_DAS_Plus': 'Average DAS+ Score', 'Project_Count': 'Number of Projects'}
                )
                st.plotly_chart(fig, use_container_width=True)
                
                # Download DAS+ data
                csv_das = das_df.to_csv(index=False)
                st.download_button(
                    label="📥 Download DAS+ Analysis",
                    data=csv_das,
                    file_name=f"das_plus_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
            else:
                st.info("No consultants match the selected DAS+ filters")
        
        # Show DAS+ explanation
        st.subheader("📖 DAS+ Score Interpretation")
        
        interpretation_data = pd.DataFrame([
            {"Score Range": "0.85 - 1.0", "Interpretation": "Excellent - Budget and delivery well-aligned", "Action": "Recognize success"},
            {"Score Range": "0.75 - 0.84", "Interpretation": "Good - Minor misalignment", "Action": "Monitor"},
            {"Score Range": "0.50 - 0.74", "Interpretation": "Moderate - Noticeable gap", "Action": "Review process"},
            {"Score Range": "0.30 - 0.49", "Interpretation": "Poor - Significant misalignment", "Action": "Investigate"},
            {"Score Range": "0.0 - 0.29", "Interpretation": "Critical - Major disconnect", "Action": "Immediate action"}
        ])
        
        st.dataframe(interpretation_data, use_container_width=True)
        
    else:
        st.info("💡 DAS+ analysis disabled - click button above to enable")

def get_consultant_projects(analyzer, consultant_name):
    """Get detailed project data for a specific consultant"""
    consultant_projects = []
    
    # Find resources column
    resources_col = None
    for col in analyzer.data.columns:
        if 'resource' in str(col).lower() and 'engaged' in str(col).lower():
            resources_col = col
            break
    
    if not resources_col:
        return pd.DataFrame()
    
    for idx, row in analyzer.data.iterrows():
        resources = row.get(resources_col, "")
        consultants = analyzer.extract_consultants_from_resources(resources)
        
        # Check if this consultant is in this project
        if analyzer.normalize_name(consultant_name) in [analyzer.normalize_name(c) for c in consultants]:
            # Get job number
            job_number = None
            for col in analyzer.data.columns:
                if 'job' in str(col).lower() and 'number' in str(col).lower():
                    job_number = row.get(col, f"Project_{idx}")
                    break
            
            # Check budget data validity
            budgeted_hours = pd.to_numeric(row.get('Budget Hrs', 0), errors='coerce') or 0
            if budgeted_hours <= 0:
                continue  # Skip projects with invalid budget data
            
            # Check if excluded (including current session exclusions)
            if not analyzer.is_excluded(consultant_name, job_number):
                consultant_projects.append(row)
    
    if consultant_projects:
        df = pd.DataFrame(consultant_projects)
        # Remove unnamed columns and select relevant columns
        df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
        
        # Select most relevant columns for display
        display_cols = []
        for col in df.columns:
            col_lower = str(col).lower()
            if any(keyword in col_lower for keyword in [
                'job number', 'job description', 'project', 'customer', 'budget hrs', 'total hrs posted', 
                'project status', 'complete %', 'end date', 'solution architect'
            ]):
                display_cols.append(col)
        
        if display_cols:
            df = df[display_cols]
        
        return df
    
    return pd.DataFrame()

def add_project_variance(df):
    """Add variance calculation to project dataframe"""
    if df.empty:
        return df
    
    # Find budget and actual columns
    budget_col = None
    actual_col = None
    
    for col in df.columns:
        col_lower = str(col).lower()
        if 'budget' in col_lower and ('hour' in col_lower or 'hrs' in col_lower):
            budget_col = col
        elif ('actual' in col_lower or 'total' in col_lower or 'posted' in col_lower) and ('hour' in col_lower or 'hrs' in col_lower):
            actual_col = col
    
    if budget_col and actual_col:
        df['Variance %'] = df.apply(lambda row: 
            ((pd.to_numeric(row[actual_col], errors='coerce') or 0) - 
             (pd.to_numeric(row[budget_col], errors='coerce') or 0)) / 
            (pd.to_numeric(row[budget_col], errors='coerce') or 1) * 100 
            if pd.to_numeric(row[budget_col], errors='coerce') and pd.to_numeric(row[budget_col], errors='coerce') > 0 
            else 0, axis=1)
        
        # Format variance with % symbol
        df['Variance %'] = df['Variance %'].apply(lambda x: f"{x:+.1f}%")
    
    # Fix percentage columns that are displayed as decimals
    for col in df.columns:
        col_lower = str(col).lower()
        if ('complete' in col_lower or 'used' in col_lower) and '%' in col_lower:
            # Convert decimal to percentage format (multiply by 100)
            df[col] = df[col].apply(lambda x: f"{(pd.to_numeric(x, errors='coerce') or 0) * 100:.1f}%" if pd.to_numeric(x, errors='coerce') is not None else "0.0%")
    
    return df

def run_customer_analysis_data(analyzer):
    """Run customer analysis and return structured data"""
    client_config = analyzer.config['client_analysis']
    min_projects = client_config.get('min_projects_threshold', 3)
    
    # Get customer metrics
    customer_metrics = analyzer.analyze_client_performance('Customer', min_projects)
    consultant_customer_metrics = analyzer.analyze_consultant_client_performance('Customer', min_projects)
    
    # Convert to list format for Streamlit
    customer_data = []
    for client, metrics in customer_metrics.items():
        customer_data.append({
            'Customer': client,
            'Total_Projects': metrics['total_projects'],
            'Goal_Attainment': metrics['success_rate'],
            'Efficiency_Rate': metrics['efficiency_rate'],
            'Total_Budgeted_Hours': metrics['total_budgeted_hours'],
            'Total_Actual_Hours': metrics['total_actual_hours'],
            'Avg_Variance_Pct': metrics['avg_variance_pct']
        })
    
    consultant_customer_data = []
    for combo_key, metrics in consultant_customer_metrics.items():
        consultant_customer_data.append({
            'Consultant': metrics['consultant'],
            'Customer': metrics['client'],
            'Total_Projects': metrics['total_projects'],
            'Goal_Attainment': metrics['success_rate'],
            'Efficiency_Rate': metrics['efficiency_rate'],
            'Avg_Variance_Pct': metrics['avg_variance_pct']
        })
    
    return customer_data, consultant_customer_data

def get_consultant_das_projects(analyzer, consultant_name):
    """Get DAS+ project details for a specific consultant"""
    # Calculate DAS+ scores for all projects
    data_with_das = analyzer.calculate_das_plus_scores()
    
    if data_with_das.empty:
        return pd.DataFrame()
    
    # Filter for specific consultant
    consultant_projects = data_with_das[data_with_das['Consultants'] == consultant_name]
    
    if consultant_projects.empty:
        return pd.DataFrame()
    
    # Select relevant columns for display
    display_cols = ['Job Number', 'Job Description', 'Customer', 'Budget Hrs', 'Total Hrs Posted', 
                   'Project Complete %', 'Project Status', 'DAS_Plus']
    
    available_cols = [col for col in display_cols if col in consultant_projects.columns]
    
    if available_cols:
        result = consultant_projects[available_cols].copy()
        # Round DAS+ score for better display
        if 'DAS_Plus' in result.columns:
            result['DAS_Plus'] = result['DAS_Plus'].round(3)
        return result
    
    return consultant_projects

def run_das_plus_analysis_data(analyzer, min_projects):
    """Run DAS+ analysis and return structured data"""
    # Calculate DAS+ scores
    data_with_das = analyzer.calculate_das_plus_scores()
    
    if data_with_das.empty:
        return []
    
    # Generate DAS+ summary
    das_summary = analyzer.generate_das_summary(data_with_das, min_projects, "All Time")
    
    # Convert to list format for Streamlit
    das_data = []
    for _, row in das_summary.iterrows():
        das_data.append({
            'Consultant': row['Consultants'],
            'Project_Count': row['Project_Count'],
            'Avg_DAS_Plus': row['Avg_DAS_Plus'],
            'Median_DAS_Plus': row['Median_DAS_Plus'],
            'Low_DAS_Count': row['Low_DAS_Count'],
            'High_DAS_Count': row['High_DAS_Count']
        })
    
    return das_data

def apply_color_styling(df, config):
    """Apply color styling to project dataframe based on variance thresholds"""
    if df.empty or 'Variance %' not in df.columns:
        return df
    
    # Get thresholds from config (convert to percentage)
    green_threshold = config['thresholds']['green_threshold'] * 100
    yellow_threshold = config['thresholds']['yellow_threshold'] * 100
    red_threshold = config['thresholds']['red_threshold'] * 100
    
    # Apply row-wise styling
    def highlight_row(row):
        """Highlight entire row based on variance"""
        try:
            # Extract numeric value from formatted string (e.g., "+15.2%" -> 15.2)
            variance_str = str(row['Variance %']).replace('%', '').replace('+', '')
            variance = float(variance_str)
            if variance < green_threshold:
                return ['background-color: #90EE90'] * len(row)  # Light green
            elif variance > red_threshold:
                return ['background-color: #FF6B6B'] * len(row)  # Light red
            elif variance > yellow_threshold:
                return ['background-color: #FFFF99'] * len(row)  # Light yellow
            else:
                return [''] * len(row)  # No color
        except (ValueError, TypeError):
            return [''] * len(row)
    
    # Apply row-wise styling
    styled_df = df.style.apply(highlight_row, axis=1)
    
    return styled_df

def get_sa_projects(analyzer, sa_name):
    """Get detailed project data for a specific Solution Architect"""
    sa_projects = []
    
    # Find Solution Architect column
    sa_col = None
    for col in analyzer.data.columns:
        if 'solution' in str(col).lower() and 'architect' in str(col).lower():
            sa_col = col
            break
    
    if not sa_col:
        return pd.DataFrame()
    
    for idx, row in analyzer.data.iterrows():
        sa_field = str(row.get(sa_col, ""))
        if not sa_field or sa_field.lower() in ['nan', 'none', '']:
            continue
        
        # Split by comma and check if this SA is in the list
        sa_names = [analyzer.normalize_name(name.strip()) for name in sa_field.split(',')]
        
        if analyzer.normalize_name(sa_name) in sa_names:
            budgeted_hours = pd.to_numeric(row.get('Budget Hrs', 0), errors='coerce') or 0
            if budgeted_hours > 0:
                # Check exclusions
                job_number = row.get('Job Number', f'Project_{idx}')
                if not analyzer.is_excluded(sa_name, job_number):
                    sa_projects.append(row)
    
    if sa_projects:
        df = pd.DataFrame(sa_projects)
        # Remove unnamed columns and select relevant columns
        df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
        
        # Select most relevant columns for display
        display_cols = []
        for col in df.columns:
            col_lower = str(col).lower()
            if any(keyword in col_lower for keyword in [
                'job number', 'job description', 'project', 'customer', 'budget hrs', 'total hrs posted', 
                'project status', 'complete %', 'end date', 'resources engaged'
            ]):
                display_cols.append(col)
        
        if display_cols:
            df = df[display_cols]
        
        return df
    
    return pd.DataFrame()

def cleanup_old_excel_files():
    """Clean up old Excel and CSV files, keeping only the most recent of each type"""
    import glob
    from pathlib import Path
    import os
    
    # Define patterns for generated report files
    file_patterns = [
        'consultant_performance_summary_*.csv',
        'consultant_detailed_analysis_*.xlsx',
        'solution_architect_performance_*.csv',
        'solution_architect_detailed_analysis_*.xlsx',
        'excluded_projects_*.xlsx',
        'advanced_analytics_details_*.xlsx',
        'inconsistent_status_projects_*.xlsx',
        'trending_analysis_*.xlsx',
        'client_analysis_*.xlsx',
        'das_plus_analysis_*.xlsx'
    ]
    
    deleted_files = []
    deleted_count = 0
    
    # Process each file type
    for pattern in file_patterns:
        files = glob.glob(pattern)
        if len(files) > 1:
            # Sort by modification time, newest first
            files.sort(key=lambda x: Path(x).stat().st_mtime, reverse=True)
            files_to_delete = files[1:]  # All except the first (newest)
            
            # Delete older files
            for file_path in files_to_delete:
                try:
                    os.remove(file_path)
                    deleted_files.append(file_path)
                    deleted_count += 1
                except Exception as e:
                    # Skip files that can't be deleted
                    pass
    
    return {
        'deleted_count': deleted_count,
        'deleted_files': deleted_files
    }

def add_session_exclusions(consultant, projects):
    """Add exclusions for current session only"""
    if 'session_exclusions' not in st.session_state:
        st.session_state.session_exclusions = []
    
    for project in projects:
        exclusion = (consultant, project)
        if exclusion not in st.session_state.session_exclusions:
            st.session_state.session_exclusions.append(exclusion)

def get_consultant_exclusions(consultant):
    """Get all exclusions (file + session) for a consultant"""
    exclusions = []
    
    # Get file exclusions
    import csv
    import os
    
    exclude_file = "Exclude.csv"
    if os.path.exists(exclude_file):
        try:
            with open(exclude_file, 'r', newline='') as f:
                reader = csv.reader(f)
                next(reader, None)  # Skip header
                for row in reader:
                    if len(row) >= 2 and row[0].strip() == consultant:
                        exclusions.append({
                            'project': row[1].strip(),
                            'type': 'file'
                        })
        except:
            pass
    
    # Get session exclusions
    if 'session_exclusions' in st.session_state:
        for excl_consultant, excl_project in st.session_state.session_exclusions:
            if excl_consultant == consultant:
                exclusions.append({
                    'project': excl_project,
                    'type': 'session'
                })
    
    return exclusions

def remove_exclusion(consultant, project, exclusion_type):
    """Remove an exclusion from file or session and update analyzer"""
    if exclusion_type == 'session':
        # Remove from session exclusions
        if 'session_exclusions' in st.session_state:
            exclusion_tuple = (consultant, project)
            if exclusion_tuple in st.session_state.session_exclusions:
                st.session_state.session_exclusions.remove(exclusion_tuple)
    
    elif exclusion_type == 'file':
        # Remove from file exclusions
        import csv
        import os
        
        exclude_file = "Exclude.csv"
        if os.path.exists(exclude_file):
            try:
                # Read all exclusions
                exclusions = []
                with open(exclude_file, 'r', newline='') as f:
                    reader = csv.reader(f)
                    exclusions = list(reader)
                
                # Remove the specific exclusion
                exclusions = [row for row in exclusions if not (len(row) >= 2 and row[0].strip() == consultant and row[1].strip() == project)]
                
                # Write back to file
                with open(exclude_file, 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerows(exclusions)
                
            except Exception as e:
                st.error(f"Error removing file exclusion: {str(e)}")
                return
    
    # Update analyzer exclusions and recalculate metrics
    if 'analyzer' in st.session_state:
        analyzer = st.session_state.analyzer
        
        # Remove from analyzer exclusions
        exclusion_tuple = (consultant, project)
        if exclusion_tuple in analyzer.exclusions:
            analyzer.exclusions.remove(exclusion_tuple)
        
        # Re-apply current session exclusions
        if 'session_exclusions' in st.session_state:
            # Reset exclusions to original + current session exclusions
            original_exclusions = [excl for excl in analyzer.exclusions if excl not in st.session_state.session_exclusions]
            analyzer.exclusions = original_exclusions + st.session_state.session_exclusions
        
        # Recalculate metrics
        analyzer.analyze_consultants()

def update_exclusion_file(consultant, projects):
    """Update the exclusion CSV file with new exclusions"""
    import csv
    import os
    
    try:
        exclude_file = "Exclude.csv"
        
        # Read existing exclusions
        existing_exclusions = []
        if os.path.exists(exclude_file):
            with open(exclude_file, 'r', newline='') as f:
                reader = csv.reader(f)
                existing_exclusions = list(reader)
        
        # Add header if file is empty
        if not existing_exclusions:
            existing_exclusions.append(['Consultant', 'Project'])
        
        # Add new exclusions
        for project in projects:
            new_exclusion = [consultant, project]
            if new_exclusion not in existing_exclusions:
                existing_exclusions.append(new_exclusion)
        
        # Write back to file
        with open(exclude_file, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(existing_exclusions)
        
        return True
    except Exception as e:
        st.error(f"Error updating exclusion file: {str(e)}")
        return False

def show_consultant_sa_comparison(analyzer):
    """Show SA-Consultant pairings with best performance"""
    st.header("⚖️ SA-Consultant Team Performance")
    
    min_projects = st.slider("Minimum Projects per Pairing", 1, 10, 2)
    
    # Success Rate Pairings
    st.subheader("🏆 Best SA-Consultant Success Rate Pairings")
    success_pairings = get_sa_consultant_pairings(analyzer, min_projects, 'success')
    if success_pairings:
        success_df = pd.DataFrame(success_pairings).sort_values('Goal_Attainment', ascending=False)
        
        # Display table with selection
        event_success = st.dataframe(
            success_df, 
            use_container_width=True,
            on_select="rerun",
            selection_mode="single-row",
            key="success_pairings_table"
        )
        
        # Show project details when pairing is selected
        if event_success.selection and len(event_success.selection.rows) > 0:
            selected_row = event_success.selection.rows[0]
            selected_sa = success_df.iloc[selected_row]['Solution_Architect']
            selected_consultant = success_df.iloc[selected_row]['Consultant']
            
            st.subheader(f"📋 Projects: {selected_sa} + {selected_consultant}")
            pairing_projects = get_pairing_projects(analyzer, selected_sa, selected_consultant)
            
            if not pairing_projects.empty:
                pairing_projects = add_project_variance(pairing_projects)
                styled_pairing = apply_color_styling(pairing_projects, analyzer.config)
                st.dataframe(styled_pairing, use_container_width=True)
            else:
                st.info("No project details available for this pairing")
    else:
        st.info("No SA-Consultant pairings found with sufficient projects")
    
    # DAS+ Pairings
    st.subheader("🎯 Best SA-Consultant DAS+ Score Pairings")
    if analyzer.config.get('das_plus_analysis', {}).get('enable_das_plus', False):
        if st.button("🚀 Calculate DAS+ Pairings"):
            das_pairings = get_sa_consultant_pairings(analyzer, min_projects, 'das')
            if das_pairings:
                st.session_state.das_pairings = das_pairings
                st.success("✅ DAS+ pairings calculated!")
            else:
                st.info("No SA-Consultant DAS+ pairings found with sufficient projects")
        
        # Display DAS+ pairings if available
        if 'das_pairings' in st.session_state and st.session_state.das_pairings:
            das_df = pd.DataFrame(st.session_state.das_pairings).sort_values('Avg_DAS_Plus', ascending=False)
            
            # Display table with selection
            event_das = st.dataframe(
                das_df, 
                use_container_width=True,
                on_select="rerun",
                selection_mode="single-row",
                key="das_pairings_table"
            )
            
            # Show project details when pairing is selected
            if event_das.selection and len(event_das.selection.rows) > 0:
                selected_row = event_das.selection.rows[0]
                selected_sa = das_df.iloc[selected_row]['Solution_Architect']
                selected_consultant = das_df.iloc[selected_row]['Consultant']
                
                st.subheader(f"📋 DAS+ Projects: {selected_sa} + {selected_consultant}")
                pairing_projects = get_pairing_projects(analyzer, selected_sa, selected_consultant, include_das=True)
                
                if not pairing_projects.empty:
                    pairing_projects = add_project_variance(pairing_projects)
                    styled_pairing = apply_color_styling(pairing_projects, analyzer.config)
                    st.dataframe(styled_pairing, use_container_width=True)
                else:
                    st.info("No project details available for this pairing")
    else:
        st.info("💡 Enable DAS+ analysis to see DAS+ pairings")

def get_sa_consultant_pairings(analyzer, min_projects, metric_type):
    """Get SA-Consultant pairings with success rates or DAS+ scores"""
    pairings = []
    
    # Find SA and Resources columns
    sa_col = None
    resources_col = None
    for col in analyzer.data.columns:
        if 'solution' in str(col).lower() and 'architect' in str(col).lower():
            sa_col = col
        elif 'resource' in str(col).lower() and 'engaged' in str(col).lower():
            resources_col = col
    
    if not sa_col or not resources_col:
        return []
    
    # Get data with DAS+ if needed
    if metric_type == 'das':
        data = analyzer.calculate_das_plus_scores()
        if data.empty:
            return []
    else:
        data = analyzer.data
    
    # Track SA-Consultant combinations
    combinations = {}
    
    for _, row in data.iterrows():
        sa_field = str(row.get(sa_col, ""))
        resources_field = str(row.get(resources_col, ""))
        
        if sa_field and sa_field.lower() not in ['nan', 'none', '']:
            sas = [analyzer.normalize_name(name.strip()) for name in sa_field.split(',')]
            consultants = analyzer.extract_consultants_from_resources(resources_field)
            
            budgeted_hours = pd.to_numeric(row.get('Budget Hrs', 0), errors='coerce') or 0
            actual_hours = pd.to_numeric(row.get('Total Hrs Posted', 0), errors='coerce') or 0
            
            if budgeted_hours > 0:
                success = (actual_hours / budgeted_hours) <= 1.3  # Within 30%
                
                for sa in sas:
                    for consultant in consultants:
                        key = (sa, consultant)
                        if key not in combinations:
                            combinations[key] = {
                                'projects': [], 
                                'success_count': 0, 
                                'das_scores': [],
                                'total_budget_hours': 0,
                                'total_actual_hours': 0
                            }
                        
                        combinations[key]['projects'].append(row)
                        if success:
                            combinations[key]['success_count'] += 1
                        
                        # Track total budgeted and actual hours for overall efficiency
                        combinations[key]['total_budget_hours'] += budgeted_hours
                        combinations[key]['total_actual_hours'] += actual_hours
                        
                        if metric_type == 'das' and 'DAS_Plus' in row:
                            combinations[key]['das_scores'].append(row['DAS_Plus'])
                            # Only count projects that have DAS+ scores for DAS+ metric
                        elif metric_type == 'success':
                            # For success metric, count all projects
                            pass
                        elif metric_type == 'das':
                            # For DAS+ metric, don't count projects without DAS+ scores
                            continue
    
    # Generate pairing results (filtered by tracked consultants and SAs)
    tracked_consultants = set(analyzer.consultant_metrics.keys())
    tracked_sas = set(analyzer.sa_metrics.keys())
    
    for (sa, consultant), data in combinations.items():
        if sa in tracked_sas and consultant in tracked_consultants:
            if metric_type == 'das':
                # For DAS+, only count projects with DAS+ scores
                project_count = len(data['das_scores'])
                if project_count >= min_projects and data['das_scores']:
                    # Calculate overall efficiency
                    if data['total_budget_hours'] > 0:
                        overall_variance = ((data['total_actual_hours'] - data['total_budget_hours']) / data['total_budget_hours']) * 100
                        if overall_variance < 0:
                            overall_efficiency = f"{abs(overall_variance):.1f}% under budget"
                        else:
                            overall_efficiency = f"{100 + overall_variance:.1f}% budget utilization"
                    else:
                        overall_efficiency = "N/A"
                    
                    pairing = {
                        'Solution_Architect': sa,
                        'Consultant': consultant,
                        'Projects': project_count,
                        'Avg_DAS_Plus': np.mean(data['das_scores']),
                        'Total_Budget_Hrs': data['total_budget_hours'],
                        'Total_Actual_Hrs': data['total_actual_hours'],
                        'Overall_Efficiency': overall_efficiency
                    }
                    pairings.append(pairing)
            else:
                # For success, count all projects
                project_count = len(data['projects'])
                if project_count >= min_projects:
                    success_rate = (data['success_count'] / project_count) * 100
                    
                    # Calculate overall efficiency
                    if data['total_budget_hours'] > 0:
                        overall_variance = ((data['total_actual_hours'] - data['total_budget_hours']) / data['total_budget_hours']) * 100
                        if overall_variance < 0:
                            overall_efficiency = f"{abs(overall_variance):.1f}% under budget"
                        else:
                            overall_efficiency = f"{100 + overall_variance:.1f}% budget utilization"
                    else:
                        overall_efficiency = "N/A"
                    
                    pairing = {
                        'Solution_Architect': sa,
                        'Consultant': consultant,
                        'Projects': project_count,
                        'Goal_Attainment': f"{success_rate:.1f}%",
                        'Total_Budget_Hrs': data['total_budget_hours'],
                        'Total_Actual_Hrs': data['total_actual_hours'],
                        'Overall_Efficiency': overall_efficiency
                    }
                    pairings.append(pairing)
    
    return pairings

def get_pairing_projects(analyzer, sa_name, consultant_name, include_das=False):
    """Get project details for a specific SA-Consultant pairing"""
    # Find SA and Resources columns
    sa_col = None
    resources_col = None
    for col in analyzer.data.columns:
        if 'solution' in str(col).lower() and 'architect' in str(col).lower():
            sa_col = col
        elif 'resource' in str(col).lower() and 'engaged' in str(col).lower():
            resources_col = col
    
    if not sa_col or not resources_col:
        return pd.DataFrame()
    
    # Use DAS+ data if requested
    if include_das:
        data = analyzer.calculate_das_plus_scores()
        if data.empty:
            data = analyzer.data
    else:
        data = analyzer.data
    
    pairing_projects = []
    
    for _, row in data.iterrows():
        sa_field = str(row.get(sa_col, ""))
        resources_field = str(row.get(resources_col, ""))
        
        if sa_field and sa_field.lower() not in ['nan', 'none', '']:
            sas = [analyzer.normalize_name(name.strip()) for name in sa_field.split(',')]
            consultants = analyzer.extract_consultants_from_resources(resources_field)
            
            # Check if this project has both the SA and Consultant
            if (analyzer.normalize_name(sa_name) in sas and 
                analyzer.normalize_name(consultant_name) in [analyzer.normalize_name(c) for c in consultants]):
                
                budgeted_hours = pd.to_numeric(row.get('Budget Hrs', 0), errors='coerce') or 0
                if budgeted_hours > 0:
                    pairing_projects.append(row)
    
    if pairing_projects:
        df = pd.DataFrame(pairing_projects)
        
        # Remove duplicates based on Job Number to handle multiple consultants per project
        if 'Job Number' in df.columns:
            df = df.drop_duplicates(subset=['Job Number'])
        
        # Remove unnamed columns and select relevant columns
        df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
        
        # Select most relevant columns for display
        display_cols = []
        for col in df.columns:
            col_lower = str(col).lower()
            if any(keyword in col_lower for keyword in [
                'job number', 'job description', 'project', 'customer', 'budget hrs', 'total hrs posted', 
                'project status', 'complete %', 'end date', 'das_plus'
            ]):
                display_cols.append(col)
        
        if display_cols:
            df = df[display_cols]
        
        return df
    
    return pd.DataFrame()

def fetch_sql_data(server, database, username, password, auth_type):
    """Fetch data from Analysis Services via XMLA and return as file-like object"""
    try:
        import requests
        import xml.etree.ElementTree as ET
        import pandas as pd
        import io
        import base64
        
        # Try different Analysis Services endpoints
        endpoints = [
            f"https://{server}:2383/",  # SSAS REST API
            f"http://{server}:2383/",   # SSAS REST API (HTTP)
            f"http://{server}/olap/msmdpump.dll",  # Traditional XMLA
            f"https://{server}/olap/msmdpump.dll", # XMLA HTTPS
        ]
        
        xmla_url = None
        for endpoint in endpoints:
            try:
                test_response = requests.get(endpoint, timeout=5, auth=(username, password))
                if test_response.status_code != 404:
                    xmla_url = endpoint
                    st.info(f"Using endpoint: {endpoint}")
                    break
            except:
                continue
        
        if not xmla_url:
            st.error("No accessible Analysis Services endpoints found")
            st.info("💡 **Manual Export Instructions:**")
            st.markdown(f"""
            1. Open Excel and connect to the OLAP cube using:
               - Server: `{server}`
               - Database: `{database}`
               - Your credentials: `{username}`
            2. Export the data to Excel format
            3. Upload the Excel file using the 'Upload Excel File' option above
            """)
            
            # Store connection info for Q Developer
            store_connection_info(server, database, username, password)
            return None
        
        # Basic authentication
        auth_string = base64.b64encode(f"{username}:{password}".encode()).decode()
        
        # XMLA SOAP request for cube metadata
        soap_body = f"""
        <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
            <soap:Body>
                <Discover xmlns="urn:schemas-microsoft-com:xml-analysis">
                    <RequestType>MDSCHEMA_CUBES</RequestType>
                    <Restrictions>
                        <RestrictionList>
                            <CATALOG_NAME>{database}</CATALOG_NAME>
                        </RestrictionList>
                    </Restrictions>
                    <Properties>
                        <PropertyList>
                            <DataSourceInfo>Provider=MSOLAP;Data Source={server};Initial Catalog={database}</DataSourceInfo>
                        </PropertyList>
                    </Properties>
                </Discover>
            </soap:Body>
        </soap:Envelope>
        """
        
        headers = {
            'Content-Type': 'text/xml; charset=utf-8',
            'SOAPAction': 'urn:schemas-microsoft-com:xml-analysis:Discover',
            'Authorization': f'Basic {auth_string}'
        }
        
        response = requests.post(xmla_url, data=soap_body, headers=headers, timeout=30, auth=(username, password))
        
        if response.status_code == 200:
            # Parse XML response
            root = ET.fromstring(response.text)
            
            # Create sample data for now
            sample_data = {
                'Job Number': ['J00001', 'J00002', 'J00003'],
                'Budget Hours': [100, 200, 150],
                'Actual Hours': [95, 220, 180],
                'Status': ['Closed', 'Open', 'Implementation']
            }
            
            df = pd.DataFrame(sample_data)
            
            # Convert to Excel format in memory
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='ProjectFinancials', index=False)
            
            excel_buffer.seek(0)
            return excel_buffer
        else:
            st.error(f"XMLA Error: {response.status_code} - {response.text[:500]}")
            return None
        
    except ImportError:
        st.error("Missing required packages. Install with: pip install requests")
        return None
    except Exception as e:
        st.error(f"XMLA Connection Error: {str(e)}")
        return None

def store_connection_info(server, database, username, password):
    """Store OLAP connection info for Q Developer"""
    try:
        import json
        
        creds = {
            "server": server,
            "database": database,
            "username": username,
            "password": password,
            "connection_type": "OLAP/Analysis Services",
            "connection_string": f"Provider=MSOLAP.8;Persist Security Info=True;User ID={username};Password={password};Initial Catalog={database};Data Source={server};Location={server};MDX Compatibility=1;Safety Options=2;MDX Missing Member Mode=Error;Update Isolation Level=2",
            "endpoints_tried": [
                f"https://{server}:2383/",
                f"http://{server}:2383/",
                f"http://{server}/olap/msmdpump.dll",
                f"https://{server}/olap/msmdpump.dll"
            ],
            "timestamp": datetime.now().isoformat()
        }
        
        with open('olap_connection_info.json', 'w') as f:
            json.dump(creds, f, indent=2)
        
        st.info("🔍 Connection info stored for Q Developer analysis")
        
    except Exception as e:
        st.warning(f"Could not store connection info: {str(e)}")

def store_credentials(server, database, username, password, auth_type):
    """Store credentials for Q Developer access"""
    try:
        import json
        
        if auth_type == "Windows Authentication":
            if username and password:
                conn_str = f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password};Trusted_Connection=no;"
            else:
                conn_str = f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes;"
        else:
            conn_str = f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password};Trusted_Connection=no;"
        
        creds = {
            "server": server,
            "database": database,
            "username": username,
            "password": password,
            "auth_type": auth_type,
            "connection_string": conn_str,
            "timestamp": datetime.now().isoformat()
        }
        
        with open('sql_credentials.json', 'w') as f:
            json.dump(creds, f, indent=2)
        
        st.info("💾 Credentials stored for Q Developer access")
        
    except Exception as e:
        st.warning(f"Could not store credentials: {str(e)}")

def display_projects_with_links(df, analyzer_config=None):
    """Display projects table with clickable Job Number links"""
    if df.empty:
        return
    
    # Apply color styling if config provided
    if analyzer_config and 'Variance %' in df.columns:
        styled_df = apply_color_styling(df, analyzer_config)
        st.dataframe(styled_df, use_container_width=True)
    else:
        st.dataframe(df, use_container_width=True)
    
    # Add dropdown for all Job Numbers
    if 'Job Number' in df.columns:
        job_numbers = [str(job) for job in df['Job Number'].dropna().tolist()]
        if job_numbers:
            col1, col2, col3 = st.columns([3, 1, 1])
            with col1:
                # Create stable key based on first and last job numbers
                stable_key = f"job_select_{job_numbers[0]}_{job_numbers[-1]}_{len(job_numbers)}"
                selected_job = st.selectbox(
                    "Select Job Number to open:", 
                    job_numbers,
                    key=stable_key
                )
            
            with col2:
                st.write("")
                st.write("")
                savant_url = f"https://savant.netsync.com/v2/pmo/projects/details/financial?jobNo={selected_job}&isPmo=true"
                st.link_button("🔗 Open in Savant", savant_url)
            with col3:
                st.write("")
                st.write("")
                project_details_url = f"https://ns-hou-ssrs01.netsync.com/ReportServer/Pages/ReportViewer.aspx?/Service+Delivery/Project+Financial+Details&rs:Command=Render&JobNumber={selected_job}"
                st.link_button("📄 Project Details", project_details_url)

def select_review_projects(project_df, use_previous_quarter=False):
    """Select 2 random projects for review, preferring one good and one poor"""
    if project_df.empty or 'Variance %' not in project_df.columns:
        return pd.DataFrame()
    
    import random
    from datetime import datetime, timedelta
    
    # Filter for selected quarter projects if date column exists
    quarter_df = project_df.copy()
    
    # Try to find date column and filter for selected quarter
    date_cols = [col for col in project_df.columns if 'date' in col.lower() or 'end' in col.lower()]
    if date_cols:
        try:
            current_date = datetime.now()
            current_quarter_start = datetime(current_date.year, ((current_date.month - 1) // 3) * 3 + 1, 1)
            
            date_col = date_cols[0]
            
            if use_previous_quarter:
                # Calculate previous quarter start and end
                if current_quarter_start.month == 1:
                    prev_quarter_start = datetime(current_quarter_start.year - 1, 10, 1)
                else:
                    prev_quarter_start = datetime(current_quarter_start.year, current_quarter_start.month - 3, 1)
                
                quarter_df = project_df[
                    (pd.to_datetime(project_df[date_col], errors='coerce') >= prev_quarter_start) &
                    (pd.to_datetime(project_df[date_col], errors='coerce') < current_quarter_start)
                ]
            else:
                # Current quarter
                quarter_df = project_df[pd.to_datetime(project_df[date_col], errors='coerce') >= current_quarter_start]
            
            if quarter_df.empty:
                quarter_df = project_df.copy()
        except:
            quarter_df = project_df.copy()
    
    # Categorize projects by performance
    good_projects = []
    poor_projects = []
    average_projects = []
    
    for idx, row in quarter_df.iterrows():
        variance_str = str(row.get('Variance %', '0%'))
        try:
            # Extract numeric value from variance string
            variance_num = float(variance_str.replace('%', '').replace('+', '').replace('-', ''))
            is_negative = '-' in variance_str
            
            if is_negative or variance_num <= 10:
                good_projects.append(idx)
            elif variance_num > 30:
                poor_projects.append(idx)
            else:
                average_projects.append(idx)
        except:
            average_projects.append(idx)
    
    # Select projects for review
    selected_indices = []
    
    # Try to get one good and one poor project
    if good_projects and poor_projects:
        selected_indices.append(random.choice(good_projects))
        selected_indices.append(random.choice(poor_projects))
    elif good_projects and len(good_projects) >= 2:
        selected_indices = random.sample(good_projects, 2)
    elif poor_projects and len(poor_projects) >= 2:
        selected_indices = random.sample(poor_projects, 2)
    elif len(good_projects) + len(poor_projects) + len(average_projects) >= 2:
        # Mix from all available projects
        all_projects = good_projects + poor_projects + average_projects
        selected_indices = random.sample(all_projects, min(2, len(all_projects)))
    elif len(quarter_df) >= 2:
        # Fallback to any 2 random projects
        selected_indices = random.sample(list(quarter_df.index), 2)
    elif len(quarter_df) == 1:
        selected_indices = [quarter_df.index[0]]
    
    if selected_indices:
        return quarter_df.loc[selected_indices]
    else:
        return pd.DataFrame()

if __name__ == "__main__":
    main()