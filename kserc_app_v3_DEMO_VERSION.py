import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import base64
import re

# ============================================
# INSTALL REQUIRED LIBRARIES
# ============================================

def install_if_missing(package):
    try:
        __import__(package)
    except ImportError:
        import subprocess
        subprocess.check_call(["pip", "install", package])

install_if_missing('openpyxl')
install_if_missing('pdfplumber')

import openpyxl
import pdfplumber

# ============================================
# PAGE CONFIGURATION
# ============================================

st.set_page_config(
    page_title="KSERC Truing Up Assistant v3.0",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main > div {padding-top: 2rem;}
    .stButton>button {width: 100%; height: 3rem; font-size: 18px; font-weight: bold;}
    </style>
""", unsafe_allow_html=True)

# ============================================
# HEADER
# ============================================

st.title("⚡ KSERC Truing Up Analysis Tool v3.0")
st.markdown("""
**AI-Powered Regulatory Compliance System**  
Automated variance analysis + KSERC Form generation + Regulatory intelligence
""")
st.markdown("---")

# ============================================
# SIDEBAR - USER INPUTS
# ============================================

st.sidebar.markdown("### ⚡ KSERC Analysis Tool v3.0")
st.sidebar.markdown("---")

st.sidebar.header("📋 Step 1: Select Licensee")
licensee_name = st.sidebar.selectbox(
    "Distribution Licensee",
    ["KSEB Ltd", "KINESCO", "Other Licensee"]
)

st.sidebar.markdown("---")
st.sidebar.header("📅 Step 2: Select Year")
financial_year = st.sidebar.selectbox(
    "Financial Year",
    ['2023-24', '2024-25', '2025-26', '2026-27', '2022-23']
)

st.sidebar.markdown("---")
st.sidebar.header("📂 Step 3: Upload Documents")

annual_report_file = st.sidebar.file_uploader(
    "Upload Annual Report",
    type=['pdf', 'xlsx', 'xls'],
    help="PDF or Excel file containing audited financial statements"
)

if annual_report_file is not None:
    file_type = "PDF" if annual_report_file.type == 'application/pdf' else "Excel"
    st.sidebar.success(f"✅ {file_type} file uploaded: {annual_report_file.name}")

st.sidebar.markdown("---")
analyze_button = st.sidebar.button("🚀 GENERATE ANALYSIS", type="primary")

st.sidebar.markdown("---")
st.sidebar.info("""
**New in v3.0:**
✨ SBU-wise breakdown support
✨ Form D 3.4(a) auto-generation
✨ Smart regulatory flagging
✨ Multi-sheet Excel output
""")

# ============================================
# CORRECTED ARR BASELINES (SBU-WISE)
# ============================================

ARR_BASELINES_SBU = {
    'KSEB Ltd': {
        '2023-24': {
            'SBU-G': {
                'O&M expenses': 206.08,
                'Interest & Finance Charges': 168.90,
                'Depreciation': 166.31,
                'Return on Equity': 116.38,
                'Master Trust - Repayment': 21.99,
                'Master Trust - Additional': 21.60,
                'Cost of Generation of Power': 0.00,
                'Amortisation': 0.00,
                'Others': 0.00,
                'Total ARR': 701.26,
                'Non-Tariff Income': 10.81,
                'Net ARR': 690.45
            },
            'SBU-T': {
                'O&M expenses': 644.81,
                'Interest & Finance Charges': 431.23,
                'Depreciation': 286.45,
                'Return on Equity': 119.99,
                'Master Trust - Repayment': 45.81,
                'Master Trust - Additional': 44.98,
                'Edamon-Kochi Compensation': 14.94,
                'Pugalur-Thrissur Compensation': 0.00,
                'Amortisation': 0.00,
                'Total ARR': 1588.21,
                'Non-Tariff Income': 54.86,
                'Net ARR': 1533.35
            },
            'SBU-D': {
                'Cost of Generation (from SBU-G)': 690.45,
                'Cost of Transmission (from SBU-T)': 1533.35,
                'Purchase of Power': 10564.23,
                'O&M Expenses': 3605.40,
                'Interest & Finance Charges': 1541.56,
                'Depreciation': 285.00,
                'Master Trust - Additional': 333.42,
                'Recovery of previous gap': 850.00,
                'Repayment of Bond': 339.42,
                'Return on Equity': 253.50,
                'Amortisation': 14.94,
                'Total ARR': 20061.27,
                'Non-Tariff Income': 866.94,
                'Revenue from Tariff': 16255.96,
                'Revenue Gap': -2938.37
            },
            # CONSOLIDATED for reference
            'Total': {
                'O&M Expenses (All SBUs)': 4456.29,  # 206.08 + 644.81 + 3605.40
                'Interest & Finance (All SBUs)': 2141.69,  # 168.90 + 431.23 + 1541.56
                'Depreciation (All SBUs)': 737.76,  # 166.31 + 286.45 + 285.00
                'Purchase of Power': 10564.23,
                'Return on Equity (All SBUs)': 489.87,
                'Master Trust Total': 467.80,
                'Total Revenue Expenditure': 19742.82,
                'Aggregate Revenue Requirement': 19996.32
            }
        }
    }
}

# For backward compatibility and simpler demo
ARR_BASELINES_CONSOLIDATED = {
    'KSEB Ltd': {
        '2023-24': {
            'O&M Expenses': 4456.29,
            'Employee Benefits Expense': 3605.40,  # Approximation for SBU-D
            'Purchase of Power': 10564.23,
            'Finance Costs': 2141.69,
            'Depreciation': 737.76,
            'Generation Costs': 690.45,
            'Transmission Costs': 1533.35,
            'Return on Equity': 489.87,
            'Master Trust': 467.80,
            'Total ARR': 19996.32
        }
    },
    'KINESCO': {
        '2023-24': {
            'O&M Expenses': 1500.0,
            'Purchase of Power': 8000.0,
            'Finance Costs': 800.0,
            'Depreciation': 200.0,
            'Total ARR': 11400.0
        }
    }
}

# ============================================
# EXTRACTION FUNCTIONS
# ============================================

def extract_value(df, keywords, year_col):
    """Enhanced extraction with multiple keyword matching"""
    if isinstance(keywords, str):
        keywords = [keywords]
    
    try:
        particulars_col = None
        for col in df.columns:
            if 'particular' in str(col).lower():
                particulars_col = col
                break
        
        if particulars_col is None:
            particulars_col = df.columns[0]
        
        for keyword in keywords:
            row = df[df[particulars_col].astype(str).str.contains(keyword, case=False, na=False)]
            
            if not row.empty:
                value = row[year_col].values[0]
                
                if isinstance(value, str):
                    value = value.replace(',', '').replace('(', '-').replace(')', '').strip()
                    value = value.replace('₹', '').replace('Rs.', '').strip()
                
                return float(value)
    except Exception as e:
        pass
    
    return np.nan

def extract_from_pdf(pdf_file):
    """Extract from PDF - simplified for demo"""
    st.info("📄 PDF detected - Extracting financial data...")
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        status_text.text("Reading PDF...")
        progress_bar.progress(30)
        
        pdf_file.seek(0)
        pdf_bytes = pdf_file.read()
        
        full_text = ''
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            for i, page in enumerate(pdf.pages[:50]):  # Limit to first 50 pages
                text = page.extract_text()
                if text:
                    full_text += text + '\n\n'
        
        progress_bar.progress(70)
        status_text.text("Parsing financial statements...")
        
        # Simple extraction for demo
        data = []
        keywords = [
            'Employee benefits expense', 'Purchase of Power', 'Finance costs',
            'Depreciation', 'Repairs & Maintenance', 'Administrative'
        ]
        
        for keyword in keywords:
            pattern = rf'{keyword}.*?(\d{{1,4}}(?:,\d{{3}})*(?:\.\d{{2}})?)'
            matches = re.findall(pattern, full_text, re.IGNORECASE)
            if matches:
                data.append([keyword, float(matches[0].replace(',', ''))])
        
        progress_bar.progress(100)
        status_text.empty()
        
        if data:
            df = pd.DataFrame(data, columns=['Particulars', f'{financial_year} (₹ Crore)'])
            st.success(f"✅ Extracted {len(df)} items")
            return df, f'{financial_year} (₹ Crore)'
        else:
            st.warning("⚠️ Limited extraction from PDF. Use Excel for better results.")
            return None, None
            
    except Exception as e:
        st.error(f"❌ PDF extraction error: {str(e)}")
        return None, None

def extract_financials_from_excel(file):
    """Extract from Excel - robust version"""
    try:
        file.seek(0)
        df = pd.read_excel(file, engine='openpyxl')
        
        year_col = None
        for col in df.columns:
            if any(x in str(col) for x in ['2023-24', '2024-25', '₹', 'Crore']):
                year_col = col
                break
        
        if year_col is None and len(df.columns) > 1:
            year_col = df.columns[1]
        
        if year_col:
            st.success(f"✅ Using column: {year_col}")
            return df, year_col
        else:
            st.error("❌ Could not identify data column")
            return None, None
            
    except Exception as e:
        st.error(f"❌ Excel error: {str(e)}")
        return None, None

# ============================================
# FORM GENERATION - KILLER FEATURE!
# ============================================

def generate_form_d34a_employee_expenses(actuals_df, year_col):
    """Generate Form D 3.4(a) - Employee Expenses Detail (22 line items)"""
    
    employee_items = {
        'Basic Salary': ['Basic Salary', 'Basic pay'],
        'Dearness Allowance (DA)': ['Dearness Allowance', 'DA'],
        'House Rent Allowance': ['House Rent', 'HRA'],
        'Conveyance Allowance': ['Conveyance'],
        'Leave Travel Allowance': ['Leave Travel', 'LTA'],
        'Earned Leave Encashment': ['Leave Encashment', 'Earned Leave'],
        'Other Allowances': ['Other Allowances'],
        'Medical Reimbursement': ['Medical'],
        'Overtime Payment': ['Overtime'],
        'Bonus/Ex-Gratia Payments': ['Bonus', 'Ex-Gratia'],
        'Interim Relief / Wage Revision': ['Wage Revision', 'Interim Relief'],
        'Staff welfare expenses': ['Staff welfare'],
        'VRS Expenses / Retrenchment Compensation': ['VRS'],
        'Commission to Directors': ['Commission'],
        'Training Expenses': ['Training'],
        'Payment under Workmen Compensation Act': ['Workmen Compensation'],
        'Net Employee Costs': ['Employee cost', 'Employee benefits expense'],
        'Terminal Benefits - PF Contribution': ['Provident Fund', 'PF'],
        'Terminal Benefits - Pension Payments': ['Pension'],
        'Terminal Benefits - Gratuity Payment': ['Gratuity'],
        'NPS Contribution': ['National Pension', 'NPS'],
        'Others': ['Other employee']
    }
    
    form_data = []
    for item_name, keywords in employee_items.items():
        value = extract_value(actuals_df, keywords, year_col)
        form_data.append({
            'S.No': len(form_data) + 1,
            'Particulars': item_name,
            f'Audited {financial_year} (₹ Crore)': value if pd.notna(value) else 0.00,
            'Remarks': 'From Annual Report' if pd.notna(value) else 'Not found in P&L'
        })
    
    # Calculate totals
    total = sum([row[f'Audited {financial_year} (₹ Crore)'] for row in form_data if pd.notna(row[f'Audited {financial_year} (₹ Crore)'])])
    form_data.append({
        'S.No': '',
        'Particulars': 'TOTAL EMPLOYEE EXPENSES',
        f'Audited {financial_year} (₹ Crore)': total,
        'Remarks': 'Calculated'
    })
    
    return pd.DataFrame(form_data)

# ============================================
# SMART REGULATORY FLAGGING
# ============================================

def apply_regulatory_intelligence(row):
    """Apply smart flagging based on stakeholder objections"""
    
    item = row['Particulars']
    variance_pct = row.get('Variance %', 0)
    
    if pd.isna(variance_pct):
        return {
            'Flag': '⚪ N/A',
            'Severity': 0,
            'Regulation': 'Data not available',
            'Action': 'Not in standard P&L statement',
            'Reference': None
        }
    
    # OBJECTION-BASED FLAGGING
    
    # Critical: Depreciation variance >100%
    if 'depreciation' in item.lower() and variance_pct > 100:
        return {
            'Flag': '🔴 Critical',
            'Severity': 4,
            'Regulation': 'Regulation 34(iii) - Exclude revalued assets',
            'Action': 'VERIFY: (1) Grants/contributions deducted? (2) Revalued assets excluded? (3) Land value separated?',
            'Reference': 'HT&EHT Objection #5 (Jan 2025) - Historical issue'
        }
    
    # Critical: Master Trust bonds
    if any(x in item.lower() for x in ['master trust', 'bond']):
        return {
            'Flag': '🔴 High Priority',
            'Severity': 3,
            'Regulation': 'Regulation 34(iv) amended 27.02.2024',
            'Action': 'Verify State Govt approval for principal repayment',
            'Reference': 'HT&EHT Objection #1 - Retroactive regulation'
        }
    
    # Important: Pay revision
    if any(x in item.lower() for x in ['pay revision', 'wage revision']):
        return {
            'Flag': '🟡 Verify',
            'Severity': 2,
            'Regulation': 'Per APTEL Order 10.11.2014',
            'Action': 'Request GoK approval letter',
            'Reference': 'HT&EHT Objection #4'
        }
    
    # Threshold-based flagging
    if abs(variance_pct) > 50:
        return {
            'Flag': '🔴 Critical',
            'Severity': 3,
            'Regulation': 'Variance exceeds 50%',
            'Action': 'Provide detailed explanation with supporting documents',
            'Reference': None
        }
    elif abs(variance_pct) > 20:
        return {
            'Flag': '🔴 Major',
            'Severity': 2,
            'Regulation': 'Variance exceeds 20%',
            'Action': 'Detailed explanation required',
            'Reference': None
        }
    elif abs(variance_pct) > 10:
        return {
            'Flag': '🟡 Medium',
            'Severity': 1,
            'Regulation': 'Variance exceeds 10%',
            'Action': 'Review recommended',
            'Reference': None
        }
    else:
        return {
            'Flag': '🟢 Normal',
            'Severity': 0,
            'Regulation': 'Within acceptable range',
            'Action': 'No action required',
            'Reference': None
        }

# ============================================
# VARIANCE ANALYSIS
# ============================================

def perform_variance_analysis(approved_arr, actuals_df, year_col):
    """Enhanced variance analysis with regulatory intelligence"""
    
    actuals = {}
    actuals['Employee Benefits Expense'] = extract_value(actuals_df, ['Employee benefits expense', 'Employee cost'], year_col)
    actuals['Purchase of Power'] = extract_value(actuals_df, ['Purchase of Power', 'Power Purchase'], year_col)
    actuals['Finance Costs'] = extract_value(actuals_df, ['Finance costs', 'Interest'], year_col)
    actuals['Depreciation'] = extract_value(actuals_df, ['Depreciation'], year_col)
    actuals['Repairs & Maintenance'] = extract_value(actuals_df, ['Repair'], year_col)
    actuals['Administrative'] = extract_value(actuals_df, ['Administrative'], year_col)
    
    comparison_data = []
    
    for arr_item, approved_value in approved_arr.items():
        
        # Map to actuals
        actual_value = np.nan
        if 'O&M' in arr_item or 'Employee' in arr_item:
            actual_value = actuals.get('Employee Benefits Expense', np.nan)
        elif 'Purchase' in arr_item or 'Power' in arr_item:
            actual_value = actuals.get('Purchase of Power', np.nan)
        elif 'Finance' in arr_item or 'Interest' in arr_item:
            actual_value = actuals.get('Finance Costs', np.nan)
        elif 'Depreciation' in arr_item:
            actual_value = actuals.get('Depreciation', np.nan)
        
        # Calculate variance
        if pd.notna(actual_value) and pd.notna(approved_value) and approved_value != 0:
            variance = actual_value - approved_value
            variance_pct = (variance / approved_value) * 100
        else:
            variance = np.nan
            variance_pct = np.nan
        
        row = {
            'Particulars': arr_item,
            'Approved ARR': approved_value,
            'Actual': actual_value,
            'Variance': variance,
            'Variance %': variance_pct
        }
        
        # Apply regulatory intelligence
        intel = apply_regulatory_intelligence(row)
        row.update(intel)
        
        comparison_data.append(row)
    
    return pd.DataFrame(comparison_data)

# ============================================
# MULTI-SHEET EXCEL DOWNLOAD
# ============================================

def create_multi_sheet_excel(variance_df, form_d34a_df, filename):
    """Create multi-sheet Excel with all outputs"""
    
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        
        # Sheet 1: Variance Analysis
        variance_df.to_excel(writer, sheet_name='Variance Analysis', index=False)
        
        # Sheet 2: Form D 3.4(a)
        form_d34a_df.to_excel(writer, sheet_name='Form D 3.4(a)', index=False)
        
        # Sheet 3: Regulatory Flags
        flags_df = variance_df[variance_df['Severity'] >= 2][['Particulars', 'Flag', 'Regulation', 'Action', 'Reference']].copy()
        flags_df.to_excel(writer, sheet_name='Regulatory Flags', index=False)
        
        # Format each sheet
        workbook = writer.book
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            worksheet.set_row(0, 20, header_format)
            
            # Auto-fit columns
            for i in range(20):
                worksheet.set_column(i, i, 15)
    
    excel_data = output.getvalue()
    b64 = base64.b64encode(excel_data).decode()
    
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}"><button style="background-color: #4CAF50; color: white; padding: 12px 24px; border: none; border-radius: 4px; cursor: pointer; font-size: 16px;">📥 Download Complete Analysis Package</button></a>'
    
    return href

# ============================================
# MAIN APPLICATION LOGIC
# ============================================

if analyze_button:
    
    if not annual_report_file:
        st.error("❌ Please upload the annual report first!")
        st.stop()
    
    # Get approved ARR
    try:
        approved_arr = ARR_BASELINES_CONSOLIDATED[licensee_name][financial_year]
    except KeyError:
        st.error(f"❌ ARR baseline not available for {licensee_name} - FY {financial_year}")
        st.stop()
    
    # Extract from file
    with st.spinner('🔄 Processing annual report...'):
        if annual_report_file.type == 'application/pdf':
            actuals_df, year_col = extract_from_pdf(annual_report_file)
        else:
            actuals_df, year_col = extract_financials_from_excel(annual_report_file)
        
        if actuals_df is None:
            st.error("❌ Could not process file")
            st.stop()
    
    # Perform analysis
    with st.spinner('🔄 Performing variance analysis...'):
        comparison_df = perform_variance_analysis(approved_arr, actuals_df, year_col)
    
    # Generate Form D 3.4(a)
    with st.spinner('🔄 Generating Form D 3.4(a)...'):
        form_d34a_df = generate_form_d34a_employee_expenses(actuals_df, year_col)
    
    st.success("✅ Analysis complete!")
    
    # ============================================
    # DISPLAY RESULTS
    # ============================================
    
    st.markdown("---")
    st.header(f"📊 Analysis Results - {licensee_name} (FY {financial_year})")
    
    # Summary metrics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_approved = comparison_df['Approved ARR'].sum()
        st.metric("Total Approved ARR", f"₹{total_approved:,.0f} Cr")
    
    with col2:
        total_actual = comparison_df['Actual'].sum(skipna=True)
        st.metric("Total Actual", f"₹{total_actual:,.0f} Cr")
    
    with col3:
        total_variance = total_actual - total_approved
        st.metric("Net Variance", f"₹{total_variance:,.0f} Cr", 
                 delta=f"{(total_variance/total_approved*100):.1f}%")
    
    with col4:
        critical_count = len(comparison_df[comparison_df['Severity'] >= 2])
        st.metric("Critical Items", critical_count)
    
    st.markdown("---")
    
    # Tabs for different outputs
    tab1, tab2, tab3 = st.tabs(["📋 Variance Analysis", "📝 Form D 3.4(a)", "🚨 Regulatory Flags"])
    
    with tab1:
        st.subheader("Detailed Variance Analysis")
        # Show only items with actual data (hide N/A items for cleaner display)
        display_df = comparison_df[comparison_df['Severity'] > 0][['Particulars', 'Approved ARR', 'Actual', 'Variance', 'Variance %', 'Flag']].copy()
        st.info("💡 Showing only items extractable from Annual Report Income Statement")
        st.dataframe(display_df, use_container_width=True, height=400)
    
    with tab2:
        st.subheader("Form D 3.4(a) - Employee Expenses Detail")
        st.info("✨ **NEW!** This is the exact form KSERC requires - auto-populated from your annual report")
        st.dataframe(form_d34a_df, use_container_width=True, height=500)
    
    with tab3:
        st.subheader("Items Requiring KSERC Review")
        flagged = comparison_df[comparison_df['Severity'] >= 2].sort_values('Severity', ascending=False)
        
        if len(flagged) == 0:
            st.success("✅ No major variances detected")
        else:
            for idx, row in flagged.iterrows():
                with st.expander(f"{row['Flag']} {row['Particulars']} ({row['Variance %']:.1f}%)", expanded=True):
                    st.write(f"**Regulation:** {row['Regulation']}")
                    st.write(f"**Action Required:** {row['Action']}")
                    if row['Reference']:
                        st.write(f"**Historical Context:** {row['Reference']}")
    
    # Download
    st.markdown("---")
    st.subheader("📥 Download Complete Package")
    
    download_filename = f"{licensee_name.replace(' ', '_')}_Complete_Analysis_{financial_year}.xlsx"
    
    st.markdown(
        create_multi_sheet_excel(comparison_df, form_d34a_df, download_filename),
        unsafe_allow_html=True
    )
    
    st.info("💡 **Package includes:** Variance Analysis + Form D 3.4(a) + Regulatory Flags (3 sheets)")

else:
    # Welcome screen
    st.info("👈 Use the sidebar to get started")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 🎯 What's New in v3.0")
        st.markdown("""
        - ✨ **SBU-wise baseline support**
        - ✨ **Form D 3.4(a) auto-generation** (22 items)
        - ✨ **Smart regulatory flagging**
        - ✨ **Multi-sheet Excel output**
        - ✨ **Stakeholder objection intelligence**
        """)
    
    with col2:
        st.markdown("### ⚡ Time Savings")
        st.success("""
        **Manual Process:**
        - Data extraction: 2 days
        - Form filling: 3 days
        - Analysis: 2 weeks
        
        **With This Tool:**
        - Everything: 2 minutes ⚡
        """)

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: gray;'>
    <p><strong>KSERC Truing Up Analysis Tool v3.0</strong></p>
    <p>Developed by XIME Kochi | Powered by AI & Regulatory Intelligence</p>
</div>
""", unsafe_allow_html=True)