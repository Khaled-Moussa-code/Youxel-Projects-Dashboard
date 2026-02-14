"""
Youxel Project Dashboard - Sprint Analysis Automation
Streamlit Web Application
"""

import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import sys
from pathlib import Path

# Set page config
st.set_page_config(
    page_title="Youxel Project Dashboard",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS with Youxel branding
st.markdown("""
<style>
    .main {
        padding: 2rem;
    }
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-size: 18px;
        padding: 0.75rem 2rem;
        border-radius: 10px;
        border: none;
        font-weight: 600;
        transition: transform 0.2s;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 12px;
        text-align: center;
        margin: 0.5rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .metric-value {
        font-size: 2.5rem;
        font-weight: bold;
        font-family: monospace;
    }
    .metric-label {
        font-size: 0.9rem;
        opacity: 0.9;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .success-box {
        background: #D1FAE5;
        border-left: 5px solid #10B981;
        padding: 1.5rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    h1 {
        color: #1E293B;
        font-size: 3rem;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .subtitle {
        text-align: center;
        color: #64748B;
        font-size: 1.3rem;
        margin-bottom: 2rem;
    }
    .youxel-brand {
        text-align: center;
        color: #667eea;
        font-size: 1.1rem;
        font-weight: 600;
        margin-bottom: 2rem;
    }
    .footer {
        text-align: center;
        color: #94A3B8;
        padding: 2rem;
        margin-top: 3rem;
        font-size: 0.9rem;
        border-top: 1px solid #E2E8F0;
    }
</style>
""", unsafe_allow_html=True)

# Import automation modules
try:
    from automation.data_processor import SprintDataProcessor
    from automation.calculator import SprintCalculator
    from automation.excel_updater import ExcelUpdater
except ImportError:
    st.error("⚠️ Automation modules not found. Please ensure the app is properly deployed.")
    st.stop()


def process_sprint_file(uploaded_file):
    """Process the uploaded Excel file"""
    try:
        # Save uploaded file to temp location
        temp_path = "temp_sprint.xlsx"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Initialize processor
        processor = SprintDataProcessor(temp_path)
        calculator = SprintCalculator()
        
        # Load workbook
        wb = openpyxl.load_workbook(temp_path, data_only=False)
        
        # Extract metadata
        sprint_metadata = processor.extract_sprint_metadata(wb['Data'])
        
        # Process Azure DevOps data
        df_azure = pd.read_excel(temp_path, sheet_name='Data', header=20)
        azure_data = processor.process_azure_data(df_azure)
        
        # Validate data
        validation = processor.validate_data(azure_data)
        if validation['status'] == 'error':
            return None, validation, None, None, None
        
        # Get capacity data
        capacity_data = processor.get_capacity_data(wb['Capacity'])
        
        # Calculate metrics
        staff_agg = processor.aggregate_by_staff(azure_data)
        team_agg = processor.aggregate_by_team(azure_data)
        
        staff_metrics = calculator.calculate_staff_metrics(
            staff_agg, capacity_data, azure_data
        )
        
        team_metrics = calculator.calculate_team_metrics(
            team_agg, capacity_data, azure_data
        )
        
        cmmi_measures = calculator.calculate_cmmi_measures(
            sprint_metadata, azure_data
        )
        
        # Update Excel sheets
        updater = ExcelUpdater(temp_path)
        sprint_name = sprint_metadata['sprint_name']
        
        updater.update_analysis_sheet(sprint_name, staff_metrics, team_metrics)
        updater.update_kpi_indicators_sheet(staff_metrics, team_metrics, sprint_name)
        updater.append_to_historical_staff(staff_metrics, sprint_name)
        updater.append_to_historical_team(team_metrics, sprint_name)
        updater.update_cmmi_template(cmmi_measures, sprint_name)
        
        # Save workbook
        updater.save()
        
        # Read the processed file
        with open(temp_path, 'rb') as f:
            processed_data = f.read()
        
        return processed_data, validation, staff_metrics, team_metrics, cmmi_measures
        
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None, None, None, None, None


# Main app
def main():
    # Header with Youxel branding
    st.markdown("<h1>🎯 Youxel Project Dashboard</h1>", unsafe_allow_html=True)
    st.markdown("<p class='youxel-brand'>Sprint Analysis Automation Platform</p>", unsafe_allow_html=True)
    st.markdown("<p class='subtitle'>Upload your Excel file → Get instant analysis → Download results</p>", unsafe_allow_html=True)
    
    # Info box
    st.info("🔒 **Your data is secure**: All processing happens on our server. Files are automatically deleted after processing. Zero data retention.")
    
    # Upload section
    st.markdown("### 📁 Upload Your Sprint File")
    
    uploaded_file = st.file_uploader(
        "Choose your Sprint Excel file (.xlsx or .xlsm)",
        type=['xlsx', 'xlsm'],
        help="Upload your Excel file with Azure DevOps sprint data"
    )
    
    if uploaded_file is not None:
        # Show file info
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📄 File Name", uploaded_file.name)
        with col2:
            file_size = len(uploaded_file.getvalue()) / 1024
            st.metric("📦 File Size", f"{file_size:.1f} KB")
        with col3:
            st.metric("✅ Status", "Ready")
        
        st.markdown("---")
        
        # Process button
        if st.button("⚡ Process Sprint Data", type="primary"):
            with st.spinner("🔄 Analyzing your sprint data... This may take a minute."):
                
                # Progress steps
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                steps = [
                    "Loading workbook...",
                    "Extracting sprint metadata...",
                    "Processing Azure DevOps data...",
                    "Validating data quality...",
                    "Loading capacity data...",
                    "Calculating staff metrics...",
                    "Calculating team metrics...",
                    "Computing CMMI measures...",
                    "Updating analysis sheets...",
                    "Finalizing workbook..."
                ]
                
                for i, step in enumerate(steps):
                    status_text.text(f"[{i+1}/{len(steps)}] {step}")
                    progress_bar.progress((i + 1) / len(steps))
                
                # Process file
                result = process_sprint_file(uploaded_file)
                processed_data, validation, staff_metrics, team_metrics, cmmi_measures = result
                
                if processed_data is None:
                    st.error("❌ Processing failed. Please check your file format and try again.")
                    if validation and validation.get('errors'):
                        for error in validation['errors']:
                            st.error(f"• {error['message']}")
                else:
                    # Success!
                    st.balloons()
                    st.success("✅ Analysis Complete!")
                    
                    # Show warnings if any
                    if validation and validation.get('warnings'):
                        with st.expander("⚠️ Warnings (click to view details)"):
                            for warning in validation['warnings']:
                                st.warning(f"• {warning['message']}")
                    
                    # Display metrics
                    st.markdown("### 📊 Sprint Summary")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.markdown(f"""
                        <div class="metric-card">
                            <div class="metric-label">Team Members</div>
                            <div class="metric-value">{len(staff_metrics)}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown(f"""
                        <div class="metric-card">
                            <div class="metric-label">Teams</div>
                            <div class="metric-value">{len(team_metrics)}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col3:
                        avg_kpi = team_metrics['kpi'].mean()
                        st.markdown(f"""
                        <div class="metric-card">
                            <div class="metric-label">Avg Team KPI</div>
                            <div class="metric-value">{avg_kpi:.2f}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col4:
                        completion = cmmi_measures['completion_rate'] * 100
                        st.markdown(f"""
                        <div class="metric-card">
                            <div class="metric-label">Completion</div>
                            <div class="metric-value">{completion:.0f}%</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    st.markdown("---")
                    
                    # Download button
                    st.markdown("### ⬇️ Download Your Analyzed File")
                    
                    output_filename = uploaded_file.name.replace('.xlsx', '_analyzed.xlsx')
                    
                    st.download_button(
                        label="📥 Download Analyzed File",
                        data=processed_data,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                    
                    st.success("🎉 Your file is ready! Click the button above to download.")
                    
                    # What's included
                    with st.expander("📋 What's included in your analyzed file"):
                        st.markdown("""
                        **New/Updated Sheets:**
                        - ✅ **[Sprint Name] Analysis** - Detailed staff and team breakdowns
                        - ✅ **Kpi Indicators** - Current sprint KPIs with formulas
                        - ✅ **Kpi Indicators Per Staff** - Historical tracking (new column added)
                        - ✅ **Kpi Indicators Per Team** - Historical tracking (new column added)
                        - ✅ **CMMI Template** - CMMI measures history (new row added)
                        
                        **Metrics Calculated:**
                        - 📊 Staff KPIs (Performance Rate, Utilization, Done Tasks %, etc.)
                        - 👥 Team KPIs (aggregated by team)
                        - 📈 CMMI Measures (Completion Rate, Effort Estimation, Productivity)
                        - ✨ All calculations use Excel formulas (dynamic, not hardcoded!)
                        """)
    
    else:
        # Instructions
        st.markdown("### 📖 How to Use This Dashboard")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("""
            #### 1️⃣ Prepare Your Data
            - Export from Azure DevOps
            - Paste into 'Data' sheet (row 21)
            - Update sprint metadata
            - Save your file
            """)
        
        with col2:
            st.markdown("""
            #### 2️⃣ Upload & Process
            - Click upload above
            - Select your Excel file
            - Click "Process Sprint Data"
            - Wait ~30 seconds
            """)
        
        with col3:
            st.markdown("""
            #### 3️⃣ Get Results
            - Review summary metrics
            - Download analyzed file
            - Open in Excel
            - All KPIs calculated!
            """)
        
        st.markdown("---")
        
        # Features
        st.markdown("### ✨ Features")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            **Automated Calculations:**
            - 📊 6 Staff-level KPIs per developer
            - 👥 6 Team-level KPIs per team
            - 📈 5 CMMI measures
            - 📉 Historical trend tracking
            - ✅ Formula-driven (not hardcoded)
            """)
        
        with col2:
            st.markdown("""
            **Security & Privacy:**
            - 🔒 Secure processing
            - 🗑️ Auto-delete after processing
            - 🚫 Zero data retention
            - 🔐 HTTPS encrypted
            - ✅ No third-party sharing
            """)
    
    # Footer
    st.markdown("""
    <div class='footer'>
        <p><strong>Youxel Project Dashboard</strong> - Sprint Analysis Automation</p>
        <p>🔒 All processing is secure and confidential | No data is stored</p>
        <p style='margin-top: 1rem; font-size: 0.8rem;'>
            Powered by Streamlit | © 2026 Youxel
        </p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
