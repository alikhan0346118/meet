import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
from pathlib import Path
import time

# Configuration
EXCEL_FILE = "Meeting_Schedule_Template.xlsx"
DATE_FORMAT = "%Y-%m-%d %H:%M"

# Initialize session state
if 'meetings_df' not in st.session_state:
    st.session_state.meetings_df = pd.DataFrame()
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'current_page' not in st.session_state:
    st.session_state.current_page = "Add New Meeting"

def load_meetings():
    """Load meetings from Excel file"""
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE)
            # Ensure all template columns exist
            template_columns = [
                'Meeting ID', 'Meeting Title', 'Organization', 'Client', 'Stakeholder Name',
                'Purpose', 'Agenda', 'Meeting Date', 'Start Time', 'Time Zone',
                'Meeting Type', 'Meeting Link', 'Location', 'Status', 'Priority',
                'Attendees', 'Internal External Guests', 'Notes', 'Next Action',
                'Follow up Date', 'Reminder Sent', 'Calendar Sync', 'Calendar Event Title'
            ]
            for col in template_columns:
                if col not in df.columns:
                    df[col] = ''
            
            # Convert date and time columns if they exist
            if 'Meeting Date' in df.columns:
                df['Meeting Date'] = pd.to_datetime(df['Meeting Date'], errors='coerce')
            if 'Follow up Date' in df.columns:
                df['Follow up Date'] = pd.to_datetime(df['Follow up Date'], errors='coerce')
            
            return df
        except Exception as e:
            st.error(f"Error loading meetings: {e}")
            return pd.DataFrame(columns=[
                'Meeting ID', 'Meeting Title', 'Organization', 'Client', 'Stakeholder Name',
                'Purpose', 'Agenda', 'Meeting Date', 'Start Time', 'Time Zone',
                'Meeting Type', 'Meeting Link', 'Location', 'Status', 'Priority',
                'Attendees', 'Internal External Guests', 'Notes', 'Next Action',
                'Follow up Date', 'Reminder Sent', 'Calendar Sync', 'Calendar Event Title'
            ])
    else:
        return pd.DataFrame(columns=[
            'Meeting ID', 'Meeting Title', 'Organization', 'Client', 'Stakeholder Name',
            'Purpose', 'Agenda', 'Meeting Date', 'Start Time', 'Time Zone',
            'Meeting Type', 'Meeting Link', 'Location', 'Status', 'Priority',
            'Attendees', 'Internal External Guests', 'Notes', 'Next Action',
            'Follow up Date', 'Reminder Sent', 'Calendar Sync', 'Calendar Event Title'
        ])

def save_meetings(df):
    """Save meetings to Excel file"""
    try:
        df.to_excel(EXCEL_FILE, index=False)
        return True
    except Exception as e:
        st.error(f"Error saving meetings: {e}")
        return False

def calculate_status(row):
    """Calculate meeting status based on current time"""
    now = datetime.now()
    
    # Handle Meeting Date and Start Time
    meeting_date = pd.to_datetime(row.get('Meeting Date', pd.NaT), errors='coerce')
    start_time_str = str(row.get('Start Time', ''))
    
    if pd.isna(meeting_date) or not start_time_str or start_time_str.strip() == '':
        return "Upcoming"  # Default if date/time not available
    
    # Try to parse start time
    try:
        if isinstance(start_time_str, datetime):
            start_time = start_time_str.time()
        elif isinstance(start_time_str, str):
            # Try different time formats
            for fmt in ['%H:%M:%S', '%H:%M', '%I:%M %p', '%I:%M:%S %p']:
                try:
                    start_time = datetime.strptime(start_time_str.strip(), fmt).time()
                    break
                except:
                    continue
            else:
                return "Upcoming"  # Could not parse time
        else:
            return "Upcoming"
        
        # Combine date and time
        start_datetime = datetime.combine(meeting_date.date(), start_time)
        
        # Since template doesn't have end time, assume 1 hour duration
        end_datetime = start_datetime + timedelta(hours=1)
        
        if now < start_datetime:
            return "Upcoming"
        elif start_datetime <= now < end_datetime:
            return "Ongoing"
        else:
            return "Ended"
    except:
        return "Upcoming"

def get_next_meeting_id(df):
    """Get the next available meeting ID"""
    if df.empty or 'Meeting ID' not in df.columns:
        return 1
    meeting_ids = df['Meeting ID'].dropna()
    if meeting_ids.empty:
        return 1
    # Try to convert to int, handle if it's string
    try:
        numeric_ids = pd.to_numeric(meeting_ids, errors='coerce').dropna()
        if numeric_ids.empty:
            return 1
        return int(numeric_ids.max()) + 1
    except:
        return 1

def update_all_statuses(df):
    """Update status for all meetings and save to Excel"""
    if not df.empty:
        if 'Status' in df.columns:
            # Preserve manually set statuses
            manually_set_ids = set()
            if 'manually_set_statuses' in st.session_state:
                manually_set_ids = set(st.session_state.manually_set_statuses.keys())
            
            # Only update status for meetings that are not manually set
            if 'Meeting ID' in df.columns:
                mask = ~df['Meeting ID'].isin(manually_set_ids)
            else:
                # If no Meeting ID, only update empty statuses
                mask = df['Status'].isna() | (df['Status'].astype(str).str.strip() == '')
            
            if mask.any():
                df.loc[mask, 'Status'] = df.loc[mask].apply(calculate_status, axis=1)
            
            # Restore manually set statuses
            if 'manually_set_statuses' in st.session_state and 'Meeting ID' in df.columns:
                for meeting_id, status in st.session_state.manually_set_statuses.items():
                    mask = df['Meeting ID'] == meeting_id
                    if mask.any():
                        df.loc[mask, 'Status'] = status
        
        save_meetings(df)
    return df

def load_data():
    """Load data into session state"""
    if not st.session_state.data_loaded:
        st.session_state.meetings_df = load_meetings()
        st.session_state.data_loaded = True
    
    # Only recalculate status for empty/NaN statuses on initial load
    # Preserve all manually set statuses (they are saved to Excel)
    if not st.session_state.meetings_df.empty:
        if 'Status' in st.session_state.meetings_df.columns:
            # Only recalculate if status is empty/NaN (not set)
            mask = st.session_state.meetings_df['Status'].isna() | (st.session_state.meetings_df['Status'].astype(str).str.strip() == '')
            if mask.any():
                st.session_state.meetings_df.loc[mask, 'Status'] = st.session_state.meetings_df.loc[mask].apply(calculate_status, axis=1)
            
            # Restore manually set statuses from session state if they exist
            if 'manually_set_statuses' in st.session_state and 'Meeting ID' in st.session_state.meetings_df.columns:
                for meeting_id, status in st.session_state.manually_set_statuses.items():
                    mask = st.session_state.meetings_df['Meeting ID'] == meeting_id
                    if mask.any():
                        st.session_state.meetings_df.loc[mask, 'Status'] = status

def filter_meetings(df, status_filter, date_start, date_end, search_text):
    """Filter meetings based on criteria"""
    filtered_df = df.copy()
    
    # Status filter
    if status_filter != "All" and 'Status' in filtered_df.columns:
        filtered_df = filtered_df[filtered_df['Status'] == status_filter]
    
    # Date range filter
    if date_start and 'Meeting Date' in filtered_df.columns:
        filtered_df = filtered_df[pd.to_datetime(filtered_df['Meeting Date'], errors='coerce') >= pd.to_datetime(date_start)]
    if date_end and 'Meeting Date' in filtered_df.columns:
        # Add end of day to date_end
        date_end_datetime = pd.to_datetime(date_end) + timedelta(days=1) - timedelta(seconds=1)
        filtered_df = filtered_df[pd.to_datetime(filtered_df['Meeting Date'], errors='coerce') <= date_end_datetime]
    
    # Search filter
    if search_text:
        search_text_lower = search_text.lower()
        search_mask = pd.Series([False] * len(filtered_df))
        
        # Search in multiple columns
        search_columns = ['Meeting Title', 'Organization', 'Client', 'Stakeholder Name', 
                         'Purpose', 'Attendees', 'Internal External Guests', 'Notes']
        for col in search_columns:
            if col in filtered_df.columns:
                search_mask |= filtered_df[col].astype(str).str.lower().str.contains(search_text_lower, na=False)
        
        filtered_df = filtered_df[search_mask]
    
    return filtered_df

# Load data
load_data()

# Page configuration
st.set_page_config(
    page_title="Meeting Dashboard", 
    page_icon="📅", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Enhanced Professional CSS Styling
st.markdown("""
<style>
    /* Global Variables & Base Styling */
    :root {
        --primary-color: #2563eb;
        --primary-dark: #1e40af;
        --primary-light: #3b82f6;
        --secondary-color: #64748b;
        --success-color: #10b981;
        --warning-color: #f59e0b;
        --error-color: #ef4444;
        --info-color: #06b6d4;
        --bg-gradient: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        --shadow-sm: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
        --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
        --shadow-lg: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05);
        --shadow-xl: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
    }
    
    /* Main container styling */
    .main .block-container {
        padding-top: 3rem;
        padding-bottom: 3rem;
        max-width: 1200px;
    }
    
    /* Enhanced Header styling with gradient */
    h1 {
        background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-size: 2.5rem;
        font-weight: 700;
        padding-bottom: 0.75rem;
        margin-bottom: 2rem;
        border-bottom: 3px solid transparent;
        border-image: linear-gradient(to right, #2563eb, #7c3aed) 1;
        letter-spacing: -0.5px;
    }
    
    h2 {
        color: #1e293b;
        margin-top: 2.5rem;
        margin-bottom: 1.25rem;
        font-size: 1.75rem;
        font-weight: 600;
        position: relative;
        padding-left: 1rem;
    }
    
    h2::before {
        content: '';
        position: absolute;
        left: 0;
        top: 50%;
        transform: translateY(-50%);
        width: 4px;
        height: 80%;
        background: linear-gradient(180deg, #2563eb, #7c3aed);
        border-radius: 2px;
    }
    
    h3 {
        color: #334155;
        margin-top: 1.75rem;
        margin-bottom: 1rem;
        font-size: 1.35rem;
        font-weight: 600;
    }
    
    /* Professional Sidebar styling */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%);
        border-right: 1px solid #e2e8f0;
    }
    
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h1 {
        background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-size: 1.5rem;
        font-weight: 700;
        border: none;
        padding-bottom: 0.5rem;
        margin-bottom: 1rem;
    }
    
    /* Enhanced Button styling */
    .stButton>button {
        background: linear-gradient(135deg, #2563eb 0%, #1e40af 100%);
        color: white;
        border-radius: 10px;
        border: none;
        padding: 0.625rem 1.75rem;
        font-weight: 600;
        font-size: 0.95rem;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        box-shadow: var(--shadow-md);
        position: relative;
        overflow: hidden;
    }
    
    .stButton>button::before {
        content: '';
        position: absolute;
        top: 50%;
        left: 50%;
        width: 0;
        height: 0;
        border-radius: 50%;
        background: rgba(255, 255, 255, 0.2);
        transform: translate(-50%, -50%);
        transition: width 0.6s, height 0.6s;
    }
    
    .stButton>button:hover::before {
        width: 300px;
        height: 300px;
    }
    
    .stButton>button:hover {
        background: linear-gradient(135deg, #1e40af 0%, #1e3a8a 100%);
        transform: translateY(-2px);
        box-shadow: var(--shadow-lg);
    }
    
    .stButton>button:active {
        transform: translateY(0);
        box-shadow: var(--shadow-sm);
    }
    
    /* Secondary button with subtle styling */
    button[kind="secondary"] {
        background: linear-gradient(135deg, #64748b 0%, #475569 100%);
        color: white;
        box-shadow: var(--shadow-md);
    }
    
    button[kind="secondary"]:hover {
        background: linear-gradient(135deg, #475569 0%, #334155 100%);
        box-shadow: var(--shadow-lg);
    }
    
    /* Enhanced Alert Messages with icons */
    .stSuccess {
        background: linear-gradient(135deg, #d1fae5 0%, #a7f3d0 100%);
        border-left: 5px solid #10b981;
        color: #065f46;
        padding: 1.25rem 1.5rem;
        border-radius: 10px;
        box-shadow: var(--shadow-md);
        font-weight: 500;
    }
    
    .stError {
        background: linear-gradient(135deg, #fee2e2 0%, #fecaca 100%);
        border-left: 5px solid #ef4444;
        color: #991b1b;
        padding: 1.25rem 1.5rem;
        border-radius: 10px;
        box-shadow: var(--shadow-md);
        font-weight: 500;
    }
    
    .stInfo {
        background: linear-gradient(135deg, #cffafe 0%, #a5f3fc 100%);
        border-left: 5px solid #06b6d4;
        color: #164e63;
        padding: 1.25rem 1.5rem;
        border-radius: 10px;
        box-shadow: var(--shadow-md);
        font-weight: 500;
    }
    
    .stWarning {
        background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%);
        border-left: 5px solid #f59e0b;
        color: #92400e;
        padding: 1.25rem 1.5rem;
        border-radius: 10px;
        box-shadow: var(--shadow-md);
        font-weight: 500;
    }
    
    /* Enhanced Metric cards with gradient */
    [data-testid="stMetricValue"] {
        background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-size: 2.25rem;
        font-weight: 700;
    }
    
    [data-testid="stMetricLabel"] {
        color: #64748b;
        font-weight: 600;
        font-size: 0.9rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    /* Professional Dataframe styling */
    .dataframe {
        border-radius: 12px;
        overflow: hidden;
        box-shadow: var(--shadow-lg);
        border: 1px solid #e2e8f0;
    }
    
    .dataframe thead {
        background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
        color: white;
    }
    
    .dataframe thead th {
        font-weight: 600;
        padding: 1rem;
        text-transform: uppercase;
        font-size: 0.85rem;
        letter-spacing: 0.5px;
    }
    
    .dataframe tbody tr {
        transition: all 0.2s ease;
    }
    
    .dataframe tbody tr:hover {
        background-color: #f1f5f9 !important;
        transform: scale(1.01);
        box-shadow: var(--shadow-sm);
    }
    
    .dataframe tbody td {
        padding: 0.875rem 1rem;
        border-bottom: 1px solid #e2e8f0;
    }
    
    /* Enhanced Input fields */
    .stTextInput>div>div>input,
    .stTextArea>div>div>textarea,
    .stSelectbox>div>div>select {
        border-radius: 8px;
        border: 2px solid #e2e8f0;
        padding: 0.625rem 0.875rem;
        font-size: 0.95rem;
        transition: all 0.3s ease;
        background-color: #ffffff;
    }
    
    .stTextInput>div>div>input:focus,
    .stTextArea>div>div>textarea:focus,
    .stSelectbox>div>div>select:focus {
        border-color: #2563eb;
        box-shadow: 0 0 0 4px rgba(37, 99, 235, 0.1);
        outline: none;
    }
    
    .stTextArea>div>div>textarea {
        min-height: 100px;
        resize: vertical;
    }
    
    /* Enhanced Expander styling */
    .streamlit-expanderHeader {
        background: linear-gradient(135deg, #f1f5f9 0%, #e2e8f0 100%);
        border-radius: 8px;
        padding: 1rem;
        font-weight: 600;
        color: #1e293b;
        border: 1px solid #cbd5e1;
        transition: all 0.3s ease;
    }
    
    .streamlit-expanderHeader:hover {
        background: linear-gradient(135deg, #e2e8f0 0%, #cbd5e1 100%);
        box-shadow: var(--shadow-sm);
    }
    
    /* Enhanced Divider/HR styling */
    hr {
        border: none;
        height: 2px;
        background: linear-gradient(to right, #2563eb, #7c3aed, transparent);
        margin: 2.5rem 0;
        border-radius: 2px;
    }
    
    /* Professional Sidebar radio buttons */
    [data-testid="stSidebar"] [data-testid="stRadio"] label {
        padding: 0.75rem 1rem;
        border-radius: 8px;
        margin: 0.5rem 0;
        transition: all 0.3s ease;
        background-color: #ffffff;
        border: 2px solid #e2e8f0;
        cursor: pointer;
    }
    
    [data-testid="stSidebar"] [data-testid="stRadio"] label:hover {
        background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%);
        border-color: #2563eb;
        transform: translateX(5px);
        box-shadow: var(--shadow-sm);
    }
    
    [data-testid="stSidebar"] [data-testid="stRadio"] label[data-baseweb="radio"] {
        background: linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%);
        border-color: #2563eb;
        box-shadow: var(--shadow-md);
    }
    
    /* Enhanced File uploader */
    [data-testid="stFileUploader"] {
        border: 3px dashed #2563eb;
        border-radius: 12px;
        padding: 2.5rem;
        background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%);
        transition: all 0.3s ease;
    }
    
    [data-testid="stFileUploader"]:hover {
        border-color: #1e40af;
        background: linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%);
        box-shadow: var(--shadow-md);
    }
    
    /* Enhanced Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 12px;
        border-bottom: 2px solid #e2e8f0;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    
    .stTabs [data-baseweb="tab"]:hover {
        background-color: #f1f5f9;
    }
    
    /* Enhanced Date and time inputs */
    [data-testid="stDateInput"] {
        border-radius: 8px;
    }
    
    [data-testid="stTimeInput"] {
        border-radius: 8px;
    }
    
    /* Enhanced Checkbox styling */
    [data-testid="stCheckbox"] label {
        font-weight: 500;
        color: #1e293b;
        padding: 0.5rem;
        border-radius: 6px;
        transition: all 0.2s ease;
    }
    
    [data-testid="stCheckbox"] label:hover {
        background-color: #f1f5f9;
    }
    
    /* Enhanced Caption styling */
    .stCaption {
        color: #64748b;
        font-style: italic;
        font-size: 0.875rem;
    }
    
    /* Professional Status badges */
    .status-badge {
        padding: 0.375rem 0.875rem;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: 700;
        display: inline-block;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        box-shadow: var(--shadow-sm);
    }
    
    .status-upcoming {
        background: linear-gradient(135deg, #d1fae5 0%, #a7f3d0 100%);
        color: #065f46;
        border: 1px solid #10b981;
    }
    
    .status-ongoing {
        background: linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%);
        color: #1e40af;
        border: 1px solid #2563eb;
    }
    
    .status-ended {
        background: linear-gradient(135deg, #f1f5f9 0%, #e2e8f0 100%);
        color: #475569;
        border: 1px solid #94a3b8;
    }
    
    .status-completed {
        background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%);
        color: #92400e;
        border: 1px solid #f59e0b;
    }
    
    /* Professional Card containers */
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 16px;
        color: white;
        box-shadow: var(--shadow-xl);
        transition: all 0.3s ease;
    }
    
    .metric-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25);
    }
    
    /* Smooth transitions for all elements */
    * {
        transition: background-color 0.2s ease, color 0.2s ease, transform 0.2s ease;
    }
    
    /* Enhanced Form sections */
    .form-section {
        background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
        padding: 2rem;
        border-radius: 12px;
        box-shadow: var(--shadow-lg);
        margin-bottom: 2rem;
        border: 1px solid #e2e8f0;
    }
    
    /* Subheader styling */
    .stSubheader {
        color: #475569;
        font-weight: 600;
        font-size: 1.1rem;
        margin-bottom: 1rem;
    }
    
    /* Selectbox enhanced */
    .stSelectbox>div>div>select {
        background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%232563eb' d='M6 9L1 4h10z'/%3E%3C/svg%3E");
        background-repeat: no-repeat;
        background-position: right 0.75rem center;
        padding-right: 2.5rem;
    }
    
    /* Enhanced number input */
    [data-testid="stNumberInput"] input {
        border-radius: 8px;
        border: 2px solid #e2e8f0;
    }
    
    /* Loading spinner enhancement */
    .stSpinner>div {
        border-color: #2563eb transparent transparent transparent;
    }
    
    /* Markdown styling */
    .stMarkdown {
        line-height: 1.7;
    }
    
    /* Code block styling */
    code {
        background-color: #f1f5f9;
        padding: 0.25rem 0.5rem;
        border-radius: 4px;
        font-size: 0.9em;
        color: #7c3aed;
        border: 1px solid #e2e8f0;
    }
</style>
""", unsafe_allow_html=True)

# Enhanced Sidebar Navigation
st.sidebar.markdown("""
<div style="text-align: center; padding: 1.5rem 0; border-bottom: 2px solid #e2e8f0; margin-bottom: 1.5rem;">
    <h1 style="background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
               -webkit-background-clip: text;
               -webkit-text-fill-color: transparent;
               background-clip: text;
               font-size: 1.75rem;
               font-weight: 700;
               margin: 0;
               padding: 0;">📅 Meeting Dashboard</h1>
    <p style="color: #64748b; font-size: 0.85rem; margin: 0.5rem 0 0 0;">AI Geo Navigators</p>
</div>
""", unsafe_allow_html=True)

# Page selection with enhanced styling
st.sidebar.markdown("<h3 style='font-size: 1rem; color: #475569; margin-bottom: 0.75rem; font-weight: 600;'>📑 Navigate to:</h3>", unsafe_allow_html=True)

page = st.sidebar.radio(
    "Navigate to:",
    ["1️⃣ Add New Meeting", "2️⃣ Edit or Delete Meeting", "3️⃣ Meetings Summary & Export"],
    index=0 if st.session_state.current_page == "Add New Meeting" else 
          1 if st.session_state.current_page == "Edit or Delete Meeting" else 2,
    label_visibility="collapsed"
)

# Update current page based on selection
if "Add New Meeting" in page:
    st.session_state.current_page = "Add New Meeting"
elif "Edit or Delete" in page:
    st.session_state.current_page = "Edit or Delete Meeting"
elif "Summary" in page:
    st.session_state.current_page = "Meetings Summary & Export"

# Sidebar - Auto-refresh configuration
st.sidebar.markdown("---")
st.sidebar.markdown("<h3 style='font-size: 1rem; color: #475569; margin-bottom: 0.75rem; font-weight: 600;'>⚙️ Settings</h3>", unsafe_allow_html=True)

auto_refresh_enabled = st.sidebar.checkbox("🔄 Enable Auto-refresh (60s)", value=False)

if st.sidebar.button("🔄 Refresh Status Now", help="Manually update all meeting statuses"):
    if not st.session_state.meetings_df.empty:
        st.session_state.meetings_df = update_all_statuses(st.session_state.meetings_df)
        st.sidebar.success("✅ Status updated!")
        st.rerun()
    else:
        st.sidebar.info("No meetings to update.")

# Handle auto-refresh
if auto_refresh_enabled:
    if not st.session_state.meetings_df.empty:
        st.session_state.meetings_df = update_all_statuses(st.session_state.meetings_df)
    
    auto_refresh_js = """
    <script>
    function setupAutoRefresh() {
        setTimeout(function() {
            window.location.reload();
        }, 60000);
    }
    setupAutoRefresh();
    </script>
    """
    st.components.v1.html(auto_refresh_js, height=0)
    st.sidebar.caption("🔄 Auto-refresh enabled")

# Enhanced Main Title with better visual hierarchy
page_titles = {
    "Add New Meeting": "➕ Add New Meeting",
    "Edit or Delete Meeting": "✏️ Edit or Delete Meeting",
    "Meetings Summary & Export": "📊 Meetings Summary & Export"
}
page_icon = "➕" if st.session_state.current_page == "Add New Meeting" else ("✏️" if st.session_state.current_page == "Edit or Delete Meeting" else "📊")

st.markdown(f"""
<div style="background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%);
            padding: 2rem;
            border-radius: 16px;
            margin-bottom: 2rem;
            border-left: 5px solid #2563eb;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);">
    <h1 style="background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
               -webkit-background-clip: text;
               -webkit-text-fill-color: transparent;
               background-clip: text;
               font-size: 2.5rem;
               font-weight: 700;
               margin: 0;
               padding: 0;
               border: none;">{page_icon} {st.session_state.current_page}</h1>
    <p style="color: #64748b; margin: 0.5rem 0 0 0; font-size: 1rem;">AI Geo Navigators Meeting Management System</p>
</div>
""", unsafe_allow_html=True)

# ============================================================================
# PAGE 1: Add New Meeting
# ============================================================================
if st.session_state.current_page == "Add New Meeting":
    with st.form("add_meeting_form", clear_on_submit=True):
        # Basic Information
        st.markdown("### 📝 Basic Information")
        col1, col2 = st.columns(2)
        
        with col1:
            meeting_title = st.text_input(
                "Meeting Title *", 
                value="",
                placeholder="Enter the meeting title",
                help="Enter the title of the meeting"
            )
            organization = st.text_input(
                "Organization",
                value="",
                placeholder="Enter organization name",
                help="Enter the organization name"
            )
            client = st.text_input(
                "Client",
                value="",
                placeholder="Enter client name",
                help="Enter the client name"
            )
            stakeholder_name = st.text_input(
                "Stakeholder Name *",
                value="",
                placeholder="Enter stakeholder name(s)",
                help="Enter the name(s) of key stakeholders (Required)"
            )
        
        with col2:
            purpose = st.text_input(
                "Purpose",
                value="",
                placeholder="Enter meeting purpose",
                help="Enter the purpose of the meeting"
            )
            meeting_type = st.selectbox("Meeting Type", ["In Person", "Virtual"], index=1)
            priority = st.selectbox("Priority", ["Low", "Medium", "High", "Urgent"], index=1)
            status = st.selectbox("Status", ["Upcoming", "Ongoing", "Ended", "Completed"], index=0)
        
        # Date and Time
        st.markdown("### 🕐 Date & Time")
        col_date1, col_date2, col_date3 = st.columns(3)
        
        with col_date1:
            meeting_date = st.date_input("Meeting Date *", value=datetime.now().date())
        with col_date2:
            start_time = st.time_input("Start Time *", value=datetime.now().time())
        with col_date3:
            time_zone = st.text_input("Time Zone", value="UTC", placeholder="e.g., UTC, EST, PST")
        
        # Location and Links
        st.markdown("### 📍 Location & Links")
        col_loc1, col_loc2 = st.columns(2)
        
        with col_loc1:
            meeting_link = st.text_input(
                "Meeting Link",
                value="",
                placeholder="Enter meeting link (for virtual meetings)",
                help="Enter the meeting link for virtual meetings"
            )
        with col_loc2:
            location = st.text_input(
                "Location",
                value="",
                placeholder="Enter physical location (for in-person meetings)",
                help="Enter the physical location for in-person meetings"
            )
        
        # Attendees
        st.markdown("### 👥 Attendees")
        col_att1, col_att2 = st.columns(2)
        
        with col_att1:
            attendees = st.text_input(
                "Attendees *",
                value="",
                placeholder="Enter attendee names (comma-separated)",
                help="Enter names of all attendees (separate multiple names with commas) (Required)"
            )
        with col_att2:
            internal_external_guests = st.text_input(
                "Internal External Guests *",
                value="",
                placeholder="Enter internal/external guest names",
                help="Enter names of internal and external guests (Required)"
            )
        
        # Agenda and Notes
        st.markdown("### 📋 Agenda & Notes")
        agenda = st.text_area(
            "Agenda", 
            value="", 
            height=80,
            placeholder="Enter meeting agenda items...",
            help="Enter the agenda items for the meeting"
        )
        notes = st.text_area(
            "Notes", 
            value="", 
            height=80,
            placeholder="Enter additional notes...",
            help="Enter any additional notes about the meeting"
        )
        
        # Follow-up and Actions
        st.markdown("### ✅ Follow-up & Actions")
        col_follow1, col_follow2 = st.columns(2)
        
        with col_follow1:
            next_action = st.text_input(
                "Next Action",
                value="",
                placeholder="Enter next action items",
                help="Enter the next action items"
            )
            follow_up_date = st.date_input("Follow up Date", value=None)
        with col_follow2:
            reminder_sent = st.selectbox("Reminder Sent", ["Yes", "No"], index=1)
            calendar_sync = st.selectbox("Calendar Sync", ["Yes", "No"], index=1)
        
        calendar_event_title = st.text_input(
            "Calendar Event Title",
            value="",
            placeholder="Enter calendar event title",
            help="Enter the title for calendar sync"
        )
        
        submitted = st.form_submit_button("💾 Save Meeting", type="primary", use_container_width=True)
        
        if submitted:
            # Validation
            errors = []
            if not meeting_title.strip():
                errors.append("Meeting Title is required")
            
            if not stakeholder_name.strip():
                errors.append("Stakeholder Name is required")
            
            if not attendees.strip():
                errors.append("Attendees is required")
            
            if not internal_external_guests.strip():
                errors.append("Internal External Guests is required")
            
            if errors:
                for error in errors:
                    st.error(error)
            else:
                # Create new meeting
                new_meeting = pd.DataFrame([{
                    'Meeting ID': get_next_meeting_id(st.session_state.meetings_df),
                    'Meeting Title': meeting_title.strip(),
                    'Organization': organization.strip(),
                    'Client': client.strip(),
                    'Stakeholder Name': stakeholder_name.strip(),
                    'Purpose': purpose.strip(),
                    'Agenda': agenda.strip(),
                    'Meeting Date': meeting_date,
                    'Start Time': start_time.strftime('%H:%M:%S') if start_time else '',
                    'Time Zone': time_zone.strip(),
                    'Meeting Type': meeting_type,
                    'Meeting Link': meeting_link.strip(),
                    'Location': location.strip(),
                    'Status': status,
                    'Priority': priority,
                    'Attendees': attendees.strip(),
                    'Internal External Guests': internal_external_guests.strip(),
                    'Notes': notes.strip(),
                    'Next Action': next_action.strip(),
                    'Follow up Date': follow_up_date if follow_up_date else '',
                    'Reminder Sent': reminder_sent,
                    'Calendar Sync': calendar_sync,
                    'Calendar Event Title': calendar_event_title.strip()
                }])
                
                # Calculate status if not manually set
                if status == "Upcoming":
                    new_meeting['Status'] = new_meeting.apply(calculate_status, axis=1)
                
                # Append to existing dataframe
                if st.session_state.meetings_df.empty:
                    st.session_state.meetings_df = new_meeting
                else:
                    st.session_state.meetings_df = pd.concat([st.session_state.meetings_df, new_meeting], ignore_index=True)
                
                # Save to Excel
                if save_meetings(st.session_state.meetings_df):
                    st.success("✅ Meeting saved successfully!")
                    st.balloons()
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error("Failed to save meeting")

# ============================================================================
# PAGE 2: Edit or Delete Meeting
# ============================================================================
elif st.session_state.current_page == "Edit or Delete Meeting":
    st.markdown("### ✏️ Edit or Delete an Existing Meeting")
    st.markdown("Select a meeting from the list below to edit or delete it.")
    
    if not st.session_state.meetings_df.empty:
        # Create selection list with index tracking for reliable lookup
        meeting_options = {}
        meeting_index_map = {}  # Map label to DataFrame index
        for idx, row in st.session_state.meetings_df.iterrows():
            meeting_title = str(row.get('Meeting Title', 'N/A'))
            meeting_date = row.get('Meeting Date', '')
            if pd.notna(meeting_date):
                try:
                    date_str = pd.to_datetime(meeting_date).strftime('%Y-%m-%d')
                except:
                    date_str = str(meeting_date)
            else:
                date_str = 'N/A'
            label = f"{meeting_title} - {date_str}"
            meeting_id = row.get('Meeting ID', idx)
            meeting_options[label] = meeting_id
            meeting_index_map[label] = idx  # Store the actual DataFrame index
        
        selected_meeting_label = st.selectbox("Select Meeting to Edit/Delete", list(meeting_options.keys()))
        selected_meeting_id = meeting_options[selected_meeting_label]
        
        # Find the meeting using the stored index for reliable lookup
        selected_meeting = None
        selected_df_index = None  # Store the DataFrame index for later use
        
        if selected_meeting_label in meeting_index_map:
            df_index = meeting_index_map[selected_meeting_label]
            selected_df_index = df_index  # Store for use in update/delete
            try:
                selected_meeting = st.session_state.meetings_df.loc[df_index]
            except (KeyError, IndexError):
                # Fallback: try filtering by Meeting ID
                if 'Meeting ID' in st.session_state.meetings_df.columns:
                    try:
                        # Try exact match first
                        mask = st.session_state.meetings_df['Meeting ID'] == selected_meeting_id
                        if not mask.any():
                            # Try string conversion for type mismatch
                            mask = st.session_state.meetings_df['Meeting ID'].astype(str) == str(selected_meeting_id)
                        if mask.any():
                            filtered_df = st.session_state.meetings_df[mask]
                            selected_meeting = filtered_df.iloc[0]
                            # Get the actual index from the filtered result
                            selected_df_index = filtered_df.index[0]
                    except Exception as e:
                        st.error(f"❌ Error finding meeting: {str(e)}")
                        st.stop()
                else:
                    # Last resort: use first row
                    if not st.session_state.meetings_df.empty:
                        selected_meeting = st.session_state.meetings_df.iloc[0]
                        selected_df_index = st.session_state.meetings_df.index[0]
        
        # Final check to ensure we have a valid meeting
        if selected_meeting is None:
            st.error("❌ Selected meeting not found. The meeting may have been deleted. Please refresh the page.")
            st.stop()
        elif isinstance(selected_meeting, pd.Series) and len(selected_meeting) == 0:
            st.error("❌ Selected meeting data is empty. Please refresh the page.")
            st.stop()
        
        # Store the index in session state for use in update/delete operations
        if selected_df_index is not None:
            st.session_state.selected_meeting_index = selected_df_index
        
        # Display current meeting info
        with st.expander("📋 View Current Meeting Details", expanded=False):
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"**Meeting Title:** {selected_meeting.get('Meeting Title', 'N/A')}")
                st.write(f"**Status:** {selected_meeting.get('Status', 'N/A')}")
                st.write(f"**Type:** {selected_meeting.get('Meeting Type', 'N/A')}")
                st.write(f"**Organization:** {selected_meeting.get('Organization', 'N/A')}")
                st.write(f"**Client:** {selected_meeting.get('Client', 'N/A')}")
                if selected_meeting.get('Stakeholder Name'):
                    st.write(f"**Stakeholder:** {selected_meeting.get('Stakeholder Name', 'N/A')}")
            with col2:
                meeting_date = selected_meeting.get('Meeting Date', '')
                if pd.notna(meeting_date):
                    try:
                        date_str = pd.to_datetime(meeting_date).strftime('%Y-%m-%d')
                    except:
                        date_str = str(meeting_date)
                else:
                    date_str = 'N/A'
                st.write(f"**Meeting Date:** {date_str}")
                st.write(f"**Start Time:** {selected_meeting.get('Start Time', 'N/A')}")
                st.write(f"**Time Zone:** {selected_meeting.get('Time Zone', 'N/A')}")
                if selected_meeting.get('Meeting Link'):
                    st.write(f"**Meeting Link:** {selected_meeting.get('Meeting Link', 'N/A')}")
                if selected_meeting.get('Location'):
                    st.write(f"**Location:** {selected_meeting.get('Location', 'N/A')}")
                if selected_meeting.get('Attendees'):
                    st.write(f"**Attendees:** {selected_meeting.get('Attendees', 'N/A')}")
        
        # Edit form
        st.markdown("---")
        with st.form("edit_meeting_form"):
            # Basic Information
            st.markdown("### 📝 Basic Information")
            col1, col2 = st.columns(2)
            
            with col1:
                edit_meeting_title = st.text_input(
                    "Meeting Title *", 
                    value=str(selected_meeting.get('Meeting Title', '')),
                    placeholder="Enter the meeting title",
                    help="Enter the title of the meeting"
                )
                edit_organization = st.text_input(
                    "Organization",
                    value=str(selected_meeting.get('Organization', '')),
                    placeholder="Enter organization name"
                )
                edit_client = st.text_input(
                    "Client",
                    value=str(selected_meeting.get('Client', '')),
                    placeholder="Enter client name"
                )
                edit_stakeholder_name = st.text_input(
                    "Stakeholder Name *",
                    value=str(selected_meeting.get('Stakeholder Name', '')),
                    placeholder="Enter stakeholder name(s)",
                    help="Enter the name(s) of key stakeholders (Required)"
                )
            
            with col2:
                edit_purpose = st.text_input(
                    "Purpose",
                    value=str(selected_meeting.get('Purpose', '')),
                    placeholder="Enter meeting purpose"
                )
                current_meeting_type = str(selected_meeting.get('Meeting Type', 'Virtual'))
                edit_meeting_type = st.selectbox("Meeting Type", ["In Person", "Virtual"], 
                                               index=0 if current_meeting_type == "In Person" else 1)
                current_priority = str(selected_meeting.get('Priority', 'Medium'))
                edit_priority = st.selectbox("Priority", ["Low", "Medium", "High", "Urgent"], 
                                            index=["Low", "Medium", "High", "Urgent"].index(current_priority) if current_priority in ["Low", "Medium", "High", "Urgent"] else 1)
                current_status = str(selected_meeting.get('Status', 'Upcoming'))
                edit_status = st.selectbox("Status", ["Upcoming", "Ongoing", "Ended", "Completed"], 
                                         index=["Upcoming", "Ongoing", "Ended", "Completed"].index(current_status) if current_status in ["Upcoming", "Ongoing", "Ended", "Completed"] else 0)
            
            # Date and Time
            st.markdown("### 🕐 Date & Time")
            col_date1, col_date2, col_date3 = st.columns(3)
            
            with col_date1:
                meeting_date_val = selected_meeting.get('Meeting Date', datetime.now().date())
                if pd.notna(meeting_date_val):
                    try:
                        edit_meeting_date = st.date_input("Meeting Date *", value=pd.to_datetime(meeting_date_val).date())
                    except:
                        edit_meeting_date = st.date_input("Meeting Date *", value=datetime.now().date())
                else:
                    edit_meeting_date = st.date_input("Meeting Date *", value=datetime.now().date())
            
            with col_date2:
                start_time_str = str(selected_meeting.get('Start Time', ''))
                try:
                    if ':' in start_time_str:
                        time_parts = start_time_str.split(':')
                        edit_start_time = st.time_input("Start Time *", value=datetime.strptime(start_time_str[:5], '%H:%M').time() if len(time_parts) >= 2 else datetime.now().time())
                    else:
                        edit_start_time = st.time_input("Start Time *", value=datetime.now().time())
                except:
                    edit_start_time = st.time_input("Start Time *", value=datetime.now().time())
            
            with col_date3:
                edit_time_zone = st.text_input("Time Zone", value=str(selected_meeting.get('Time Zone', 'UTC')))
            
            # Location and Links
            st.markdown("### 📍 Location & Links")
            col_loc1, col_loc2 = st.columns(2)
            
            with col_loc1:
                edit_meeting_link = st.text_input("Meeting Link", value=str(selected_meeting.get('Meeting Link', '')))
            with col_loc2:
                edit_location = st.text_input("Location", value=str(selected_meeting.get('Location', '')))
            
            # Attendees
            st.markdown("### 👥 Attendees")
            col_att1, col_att2 = st.columns(2)
            
            with col_att1:
                edit_attendees = st.text_input("Attendees *", value=str(selected_meeting.get('Attendees', '')),
                                              help="Enter names of all attendees (Required)")
            with col_att2:
                edit_internal_external_guests = st.text_input("Internal External Guests *", 
                                                             value=str(selected_meeting.get('Internal External Guests', '')),
                                                             help="Enter names of internal and external guests (Required)")
            
            # Agenda and Notes
            st.markdown("### 📋 Agenda & Notes")
            edit_agenda = st.text_area("Agenda", value=str(selected_meeting.get('Agenda', '')), height=80)
            edit_notes = st.text_area("Notes", value=str(selected_meeting.get('Notes', '')), height=80)
            
            # Follow-up and Actions
            st.markdown("### ✅ Follow-up & Actions")
            col_follow1, col_follow2 = st.columns(2)
            
            with col_follow1:
                edit_next_action = st.text_input("Next Action", value=str(selected_meeting.get('Next Action', '')))
                follow_up_date_val = selected_meeting.get('Follow up Date', None)
                if pd.notna(follow_up_date_val) and follow_up_date_val != '':
                    try:
                        edit_follow_up_date = st.date_input("Follow up Date", value=pd.to_datetime(follow_up_date_val).date())
                    except:
                        edit_follow_up_date = st.date_input("Follow up Date", value=None)
                else:
                    edit_follow_up_date = st.date_input("Follow up Date", value=None)
            with col_follow2:
                current_reminder = str(selected_meeting.get('Reminder Sent', 'No'))
                edit_reminder_sent = st.selectbox("Reminder Sent", ["Yes", "No"], 
                                                 index=0 if current_reminder == "Yes" else 1)
                current_cal_sync = str(selected_meeting.get('Calendar Sync', 'No'))
                edit_calendar_sync = st.selectbox("Calendar Sync", ["Yes", "No"], 
                                                 index=0 if current_cal_sync == "Yes" else 1)
            
            edit_calendar_event_title = st.text_input("Calendar Event Title", 
                                                      value=str(selected_meeting.get('Calendar Event Title', '')))
            
            col_btn1, col_btn2 = st.columns([1, 1])
            with col_btn1:
                update_submitted = st.form_submit_button("💾 Update Meeting", type="primary", use_container_width=True)
            with col_btn2:
                pass  # Delete button is outside form
            
            if update_submitted:
                # Validation
                errors = []
                if not edit_meeting_title.strip():
                    errors.append("Meeting Title is required")
                
                if not edit_stakeholder_name.strip():
                    errors.append("Stakeholder Name is required")
                
                if not edit_attendees.strip():
                    errors.append("Attendees is required")
                
                if not edit_internal_external_guests.strip():
                    errors.append("Internal External Guests is required")
                
                if errors:
                    for error in errors:
                        st.error(error)
                else:
                    # Find the index - use stored index if available, otherwise lookup
                    idx = None
                    
                    # First, try to use the stored index from session state
                    if 'selected_meeting_index' in st.session_state:
                        stored_idx = st.session_state.selected_meeting_index
                        # Verify the index still exists in the DataFrame
                        if stored_idx in st.session_state.meetings_df.index:
                            idx = stored_idx
                    
                    # If stored index not available, try to find by Meeting ID
                    if idx is None and 'Meeting ID' in st.session_state.meetings_df.columns:
                        try:
                            # Try exact match first
                            mask = st.session_state.meetings_df['Meeting ID'] == selected_meeting_id
                            if not mask.any():
                                # Try string conversion for type mismatch
                                mask = st.session_state.meetings_df['Meeting ID'].astype(str) == str(selected_meeting_id)
                            
                            if mask.any():
                                idx = st.session_state.meetings_df[mask].index[0]
                        except (IndexError, KeyError):
                            pass
                    
                    # If still not found, try using the meeting_index_map
                    if idx is None and selected_meeting_label in meeting_index_map:
                        try:
                            potential_idx = meeting_index_map[selected_meeting_label]
                            if potential_idx in st.session_state.meetings_df.index:
                                idx = potential_idx
                        except (KeyError, IndexError):
                            pass
                    
                    # Final fallback
                    if idx is None:
                        st.error("❌ Could not find the meeting to update. The meeting may have been deleted. Please refresh the page.")
                        st.stop()
                    
                    # Update meeting
                    st.session_state.meetings_df.at[idx, 'Meeting Title'] = edit_meeting_title.strip()
                    st.session_state.meetings_df.at[idx, 'Organization'] = edit_organization.strip()
                    st.session_state.meetings_df.at[idx, 'Client'] = edit_client.strip()
                    st.session_state.meetings_df.at[idx, 'Stakeholder Name'] = edit_stakeholder_name.strip()
                    st.session_state.meetings_df.at[idx, 'Purpose'] = edit_purpose.strip()
                    st.session_state.meetings_df.at[idx, 'Agenda'] = edit_agenda.strip()
                    st.session_state.meetings_df.at[idx, 'Meeting Date'] = edit_meeting_date
                    st.session_state.meetings_df.at[idx, 'Start Time'] = edit_start_time.strftime('%H:%M:%S') if edit_start_time else ''
                    st.session_state.meetings_df.at[idx, 'Time Zone'] = edit_time_zone.strip()
                    st.session_state.meetings_df.at[idx, 'Meeting Type'] = edit_meeting_type
                    st.session_state.meetings_df.at[idx, 'Meeting Link'] = edit_meeting_link.strip()
                    st.session_state.meetings_df.at[idx, 'Location'] = edit_location.strip()
                    st.session_state.meetings_df.at[idx, 'Status'] = edit_status
                    st.session_state.meetings_df.at[idx, 'Priority'] = edit_priority
                    st.session_state.meetings_df.at[idx, 'Attendees'] = edit_attendees.strip()
                    st.session_state.meetings_df.at[idx, 'Internal External Guests'] = edit_internal_external_guests.strip()
                    st.session_state.meetings_df.at[idx, 'Notes'] = edit_notes.strip()
                    st.session_state.meetings_df.at[idx, 'Next Action'] = edit_next_action.strip()
                    st.session_state.meetings_df.at[idx, 'Follow up Date'] = edit_follow_up_date if edit_follow_up_date else ''
                    st.session_state.meetings_df.at[idx, 'Reminder Sent'] = edit_reminder_sent
                    st.session_state.meetings_df.at[idx, 'Calendar Sync'] = edit_calendar_sync
                    st.session_state.meetings_df.at[idx, 'Calendar Event Title'] = edit_calendar_event_title.strip()
                    
                    # Save to Excel
                    if save_meetings(st.session_state.meetings_df):
                        # Store the manually set status to preserve it after reload
                        if 'manually_set_statuses' not in st.session_state:
                            st.session_state.manually_set_statuses = {}
                        st.session_state.manually_set_statuses[selected_meeting_id] = edit_status
                        
                        # Get old status for comparison
                        old_status = selected_meeting.get('Status', 'N/A')
                        new_status = edit_status
                        
                        # Show success message with status update information
                        if old_status != new_status:
                            status_msg = f"✅ Meeting updated successfully! Status changed from '{old_status}' to '{new_status}'"
                        else:
                            status_msg = f"✅ Meeting updated successfully! Current status: '{new_status}'"
                        
                        st.success(status_msg)
                        st.balloons()
                        time.sleep(1.5)
                        st.rerun()
                    else:
                        st.error("Failed to update meeting")
        
        # Delete button (outside form)
        st.markdown("---")
        st.markdown("### 🗑️ Delete Meeting")
        
        if 'confirm_delete' not in st.session_state:
            st.session_state.confirm_delete = False
        
        if st.button("🗑️ Delete This Meeting", type="secondary", use_container_width=True):
            st.session_state.confirm_delete = True
        
        if st.session_state.confirm_delete:
            st.warning("⚠️ **Are you sure you want to delete this meeting?** This action cannot be undone.")
            
            # Use the same index lookup logic as in update
            delete_idx = None
            if 'selected_meeting_index' in st.session_state:
                stored_idx = st.session_state.selected_meeting_index
                if stored_idx in st.session_state.meetings_df.index:
                    delete_idx = stored_idx
            
            if delete_idx is None and 'Meeting ID' in st.session_state.meetings_df.columns:
                try:
                    mask = st.session_state.meetings_df['Meeting ID'] == selected_meeting_id
                    if not mask.any():
                        mask = st.session_state.meetings_df['Meeting ID'].astype(str) == str(selected_meeting_id)
                    if mask.any():
                        delete_idx = st.session_state.meetings_df[mask].index[0]
                except (IndexError, KeyError):
                    pass
            
            if delete_idx is None and selected_meeting_label in meeting_index_map:
                try:
                    potential_idx = meeting_index_map[selected_meeting_label]
                    if potential_idx in st.session_state.meetings_df.index:
                        delete_idx = potential_idx
                except (KeyError, IndexError):
                    pass
            col1, col2 = st.columns(2)
            with col1:
                if st.button("✅ Confirm Delete", type="primary", use_container_width=True):
                    # Remove meeting using the stored index
                    delete_idx = None
                    
                    # First, try to use the stored index from session state
                    if 'selected_meeting_index' in st.session_state:
                        stored_idx = st.session_state.selected_meeting_index
                        if stored_idx in st.session_state.meetings_df.index:
                            delete_idx = stored_idx
                    
                    # If stored index not available, try to find by Meeting ID
                    if delete_idx is None and 'Meeting ID' in st.session_state.meetings_df.columns:
                        try:
                            mask = st.session_state.meetings_df['Meeting ID'] == selected_meeting_id
                            if not mask.any():
                                mask = st.session_state.meetings_df['Meeting ID'].astype(str) == str(selected_meeting_id)
                            if mask.any():
                                delete_idx = st.session_state.meetings_df[mask].index[0]
                        except (IndexError, KeyError):
                            pass
                    
                    # If still not found, try using the meeting_index_map
                    if delete_idx is None and selected_meeting_label in meeting_index_map:
                        try:
                            potential_idx = meeting_index_map[selected_meeting_label]
                            if potential_idx in st.session_state.meetings_df.index:
                                delete_idx = potential_idx
                        except (KeyError, IndexError):
                            pass
                    
                    # Delete the meeting using the found index
                    if delete_idx is not None:
                        st.session_state.meetings_df = st.session_state.meetings_df.drop(delete_idx)
                    else:
                        # Fallback: filter by Meeting ID
                        if 'Meeting ID' in st.session_state.meetings_df.columns:
                            try:
                                mask = st.session_state.meetings_df['Meeting ID'] != selected_meeting_id
                                if not mask.all():  # If any match, filter them out
                                    # Try string comparison if needed
                                    mask = st.session_state.meetings_df['Meeting ID'].astype(str) != str(selected_meeting_id)
                                st.session_state.meetings_df = st.session_state.meetings_df[mask]
                            except Exception as e:
                                st.error(f"❌ Error deleting meeting: {str(e)}")
                                st.stop()
                        else:
                            # Last resort: delete first row (not ideal)
                            if not st.session_state.meetings_df.empty:
                                st.session_state.meetings_df = st.session_state.meetings_df.drop(st.session_state.meetings_df.index[0])
                            else:
                                st.error("❌ Could not find the meeting to delete.")
                                st.stop()
                    
                    # Save to Excel
                    if save_meetings(st.session_state.meetings_df):
                        st.session_state.confirm_delete = False
                        st.success("✅ Meeting deleted successfully!")
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error("Failed to delete meeting")
            with col2:
                if st.button("❌ Cancel", use_container_width=True):
                    st.session_state.confirm_delete = False
                    st.rerun()
    else:
        st.info("📭 No meetings available to edit or delete. Add a meeting first using the 'Add New Meeting' page.")

# ============================================================================
# PAGE 3: Meetings Summary & Export
# ============================================================================
elif st.session_state.current_page == "Meetings Summary & Export":
    st.markdown("### 📊 View All Meetings Summary")
    st.markdown("Filter and view all meetings, then export the data if needed.")
    
    # Filters section
    st.markdown("---")
    st.markdown("""
    <div style="background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
                padding: 1.5rem;
                border-radius: 12px;
                margin-bottom: 1.5rem;
                border-left: 4px solid #2563eb;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);">
        <h2 style="margin: 0; color: #1e293b; font-size: 1.5rem; font-weight: 600;">🔍 Filters</h2>
    </div>
    """, unsafe_allow_html=True)
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        status_options = ["All", "Upcoming", "Ongoing", "Ended", "Completed"]
        selected_status = st.selectbox("Filter by Status", status_options)
    
    with col2:
        date_start = st.date_input("Start Date", value=None)
    
    with col3:
        date_end = st.date_input("End Date", value=None)
    
    with col4:
        search_text = st.text_input("Search (Title/Organizer/Attendees)", value="", 
                                  help="Search by title, organizer, stakeholder, or attendee names")
    
    # Apply filters
    if not st.session_state.meetings_df.empty:
        filtered_meetings = filter_meetings(
            st.session_state.meetings_df,
            selected_status,
            date_start,
            date_end,
            search_text
        )
    else:
        filtered_meetings = pd.DataFrame()
    
    # Summary metrics
    st.markdown("---")
    st.markdown("""
    <div style="background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
                padding: 1.5rem;
                border-radius: 12px;
                margin-bottom: 1.5rem;
                border-left: 4px solid #10b981;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);">
        <h2 style="margin: 0; color: #1e293b; font-size: 1.5rem; font-weight: 600;">📈 Summary Statistics</h2>
    </div>
    """, unsafe_allow_html=True)
    col1, col2, col3, col4 = st.columns(4)
    
    if not filtered_meetings.empty:
        total_count = len(filtered_meetings)
        if 'Status' in filtered_meetings.columns:
            upcoming_count = len(filtered_meetings[filtered_meetings['Status'] == 'Upcoming'])
            ongoing_count = len(filtered_meetings[filtered_meetings['Status'] == 'Ongoing'])
            ended_count = len(filtered_meetings[filtered_meetings['Status'] == 'Ended'])
            completed_count = len(filtered_meetings[filtered_meetings['Status'] == 'Completed'])
        else:
            upcoming_count = ongoing_count = ended_count = completed_count = 0
    else:
        total_count = upcoming_count = ongoing_count = ended_count = completed_count = 0
    
    with col1:
        st.metric("Total Meetings", total_count)
    with col2:
        st.metric("Upcoming", upcoming_count, delta=None)
    with col3:
        st.metric("Ongoing", ongoing_count, delta=None)
    with col4:
        st.metric("Ended/Completed", ended_count + completed_count, delta=None)
    
    # Meetings table
    st.markdown("---")
    st.markdown("""
    <div style="background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
                padding: 1.5rem;
                border-radius: 12px;
                margin-bottom: 1.5rem;
                border-left: 4px solid #7c3aed;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);">
        <h2 style="margin: 0; color: #1e293b; font-size: 1.5rem; font-weight: 600;">📋 Meetings Table</h2>
    </div>
    """, unsafe_allow_html=True)
    
    if not filtered_meetings.empty:
        # Prepare display dataframe
        display_df = filtered_meetings.copy()
        
        # Format date columns if they exist
        if 'Meeting Date' in display_df.columns:
            display_df['Meeting Date'] = pd.to_datetime(display_df['Meeting Date'], errors='coerce').dt.strftime('%Y-%m-%d')
        if 'Follow up Date' in display_df.columns:
            display_df['Follow up Date'] = pd.to_datetime(display_df['Follow up Date'], errors='coerce').dt.strftime('%Y-%m-%d')
        
        # Select columns to display (show most important ones)
        display_columns = ['Meeting Title', 'Meeting Date', 'Start Time', 'Status', 
                          'Meeting Type', 'Organization', 'Client', 'Stakeholder Name', 
                          'Priority', 'Attendees', 'Location', 'Meeting Link']
        available_columns = [col for col in display_columns if col in display_df.columns]
        
        st.dataframe(
            display_df[available_columns],
            use_container_width=True,
            hide_index=True,
            height=400
        )
        
        st.caption(f"Showing {len(display_df)} meeting(s)")
    else:
        st.info("📭 No meetings found matching your filters.")
    
    # Import/Upload section
    st.markdown("---")
    st.markdown("""
    <div style="background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
                padding: 1.5rem;
                border-radius: 12px;
                margin-bottom: 1.5rem;
                border-left: 4px solid #f59e0b;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);">
        <h2 style="margin: 0; color: #1e293b; font-size: 1.5rem; font-weight: 600;">📤 Import/Update from Excel</h2>
    </div>
    """, unsafe_allow_html=True)
    
    # Template download option
    col_template1, col_template2 = st.columns([3, 1])
    with col_template1:
        st.write("Upload an Excel file to import or update meeting records.")
    with col_template2:
        # Create template dataframe with all template columns
        template_columns = [
            'Meeting ID', 'Meeting Title', 'Organization', 'Client', 'Stakeholder Name',
            'Purpose', 'Agenda', 'Meeting Date', 'Start Time', 'Time Zone',
            'Meeting Type', 'Meeting Link', 'Location', 'Status', 'Priority',
            'Attendees', 'Internal External Guests', 'Notes', 'Next Action',
            'Follow up Date', 'Reminder Sent', 'Calendar Sync', 'Calendar Event Title'
        ]
        template_df = pd.DataFrame(columns=template_columns)
        # Add sample row
        template_df = pd.concat([template_df, pd.DataFrame([{
            'Meeting ID': 1,
            'Meeting Title': 'Sample Meeting',
            'Organization': 'Sample Org',
            'Client': 'Sample Client',
            'Stakeholder Name': 'Jane Smith',
            'Purpose': 'Sample Purpose',
            'Agenda': 'Sample agenda items',
            'Meeting Date': datetime.now().date(),
            'Start Time': datetime.now().time().strftime('%H:%M:%S'),
            'Time Zone': 'UTC',
            'Meeting Type': 'Virtual',
            'Meeting Link': 'https://meet.example.com',
            'Location': '',
            'Status': 'Upcoming',
            'Priority': 'Medium',
            'Attendees': 'Team Member 1, Team Member 2',
            'Internal External Guests': 'Client A, Client B',
            'Notes': 'Sample notes',
            'Next Action': 'Follow up required',
            'Follow up Date': '',
            'Reminder Sent': 'No',
            'Calendar Sync': 'No',
            'Calendar Event Title': 'Sample Meeting'
        }])], ignore_index=True)
        
        # Save template to bytes
        import io
        template_buffer = io.BytesIO()
        template_df.to_excel(template_buffer, index=False)
        template_buffer.seek(0)
        
        st.download_button(
            label="📥 Download Template",
            data=template_buffer,
            file_name="meeting_import_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Download a template Excel file with the correct column format"
        )
    
    uploaded_file = st.file_uploader(
        "Choose an Excel file to import",
        type=['xlsx', 'xls'],
        help="Upload an Excel file with meeting data. Required columns: Meeting Title, Meeting Date, Start Time. All other columns are optional."
    )
    
    if uploaded_file is not None:
        try:
            # Read the uploaded file
            import_df = pd.read_excel(uploaded_file)
            
            # Check only critical required columns
            critical_required_columns = ['Meeting Title']
            missing_critical = [col for col in critical_required_columns if col not in import_df.columns]
            
            if missing_critical:
                st.error(f"❌ Missing critical required column: {', '.join(missing_critical)}")
                st.info("At minimum, 'Meeting Title' column is required. Other missing columns will be filled with empty values.")
            else:
                # Add missing columns with empty values
                template_columns = [
                    'Meeting ID', 'Meeting Title', 'Organization', 'Client', 'Stakeholder Name',
                    'Purpose', 'Agenda', 'Meeting Date', 'Start Time', 'Time Zone',
                    'Meeting Type', 'Meeting Link', 'Location', 'Status', 'Priority',
                    'Attendees', 'Internal External Guests', 'Notes', 'Next Action',
                    'Follow up Date', 'Reminder Sent', 'Calendar Sync', 'Calendar Event Title'
                ]
                
                missing_columns = [col for col in template_columns if col not in import_df.columns]
                if missing_columns:
                    # Add missing columns silently without showing message
                    for col in missing_columns:
                        import_df[col] = ''
                
                # Ensure datetime columns are properly formatted
                if 'Meeting Date' in import_df.columns:
                    import_df['Meeting Date'] = pd.to_datetime(import_df['Meeting Date'], errors='coerce')
                if 'Follow up Date' in import_df.columns:
                    import_df['Follow up Date'] = pd.to_datetime(import_df['Follow up Date'], errors='coerce')
                
                # Show preview
                st.markdown("**📋 Preview of Uploaded Data:**")
                st.dataframe(import_df.head(10), use_container_width=True, hide_index=True)
                st.caption(f"Total rows to import: {len(import_df)}")
                
                # Import options
                col1, col2 = st.columns(2)
                with col1:
                    import_mode = st.radio(
                        "Import Mode:",
                        ["Add New Only", "Update Existing", "Update & Add New"],
                        help="Add New: Only import records with new meeting_id\nUpdate Existing: Only update records with matching meeting_id\nUpdate & Add New: Do both"
                    )
                
                with col2:
                    overwrite_status = st.checkbox(
                        "Overwrite existing status",
                        value=False,
                        help="If checked, will overwrite status of existing meetings. If unchecked, will preserve current status for existing meetings."
                    )
                
                # Normalize empty values to null - only process rows that have a Meeting Title
                # Rows without Meeting Title are treated as empty and will be imported as-is
                for idx, row in import_df.iterrows():
                    # Check if row has a Meeting Title - only process rows with Meeting Title
                    meeting_title = row.get('Meeting Title', '')
                    has_meeting_title = (
                        pd.notna(meeting_title) and 
                        str(meeting_title).strip() != '' and 
                        str(meeting_title).strip().lower() not in ['nan', 'none', 'null', '']
                    )
                    
                    # Only normalize empty fields if row has a Meeting Title
                    # Rows without Meeting Title are treated as empty and allowed through
                    if has_meeting_title:
                        meeting_date = row.get('Meeting Date', '')
                        # Check if date is empty - handle various formats including NaT
                        is_date_empty = True
                        if pd.notna(meeting_date):
                            if isinstance(meeting_date, (pd.Timestamp, datetime)):
                                # Check if it's a valid timestamp (not NaT)
                                if isinstance(meeting_date, pd.Timestamp) and pd.isna(meeting_date):
                                    is_date_empty = True
                                else:
                                    is_date_empty = False
                            elif isinstance(meeting_date, str):
                                date_str = meeting_date.strip()
                                if date_str and date_str.lower() not in ['nan', 'none', 'null', '', 'nat']:
                                    # Try to parse it
                                    try:
                                        parsed_date = pd.to_datetime(date_str)
                                        if pd.notna(parsed_date):
                                            is_date_empty = False
                                    except:
                                        pass
                            else:
                                # Try to convert to string and check
                                date_str = str(meeting_date).strip()
                                if date_str and date_str.lower() not in ['nan', 'none', 'null', '', 'nat']:
                                    try:
                                        parsed_date = pd.to_datetime(date_str)
                                        if pd.notna(parsed_date):
                                            is_date_empty = False
                                    except:
                                        pass
                        
                        # Set missing Meeting Date to null (NaT)
                        if is_date_empty:
                            import_df.at[idx, 'Meeting Date'] = pd.NaT
                        
                        start_time = row.get('Start Time', '')
                        # Check if time is empty
                        is_time_empty = True
                        if pd.notna(start_time):
                            time_str = str(start_time).strip()
                            if time_str and time_str.lower() not in ['nan', 'none', 'null', '']:
                                is_time_empty = False
                        
                        # Set missing Start Time to null (empty string)
                        if is_time_empty:
                            import_df.at[idx, 'Start Time'] = ''
                    # If row has no Meeting Title, skip processing (treat as empty row)
                
                # Validate required fields
                validation_errors = []
                for idx, row in import_df.iterrows():
                    # Only validate rows that have a Meeting Title
                    meeting_title = row.get('Meeting Title', '')
                    has_meeting_title = (
                        pd.notna(meeting_title) and 
                        str(meeting_title).strip() != '' and 
                        str(meeting_title).strip().lower() not in ['nan', 'none', 'null', '']
                    )
                    
                    # Note: Stakeholder Name, Attendees, and Internal External Guests are allowed to be empty/null
                    # They will be filled with empty strings during import
                
                # Show info if there are validation warnings (but don't block import)
                if validation_errors:
                    st.warning("⚠️ **Data Quality Warnings (import will proceed):**")
                    for error in validation_errors[:5]:  # Show first 5 warnings
                        st.write(f"- {error}")
                    if len(validation_errors) > 5:
                        st.write(f"- ... and {len(validation_errors) - 5} more warnings")
                    st.info("💡 Missing values will be filled with empty strings. You can update them later.")
                
                # Always proceed with import (validation_errors are just warnings now)
                # Proceed with import
                if st.button("✅ Import Data", type="primary", use_container_width=True):
                    try:
                            # Ensure all template columns exist in import_df (already done above, but double-check)
                            template_columns = [
                                'Meeting ID', 'Meeting Title', 'Organization', 'Client', 'Stakeholder Name',
                                'Purpose', 'Agenda', 'Meeting Date', 'Start Time', 'Time Zone',
                                'Meeting Type', 'Meeting Link', 'Location', 'Status', 'Priority',
                                'Attendees', 'Internal External Guests', 'Notes', 'Next Action',
                                'Follow up Date', 'Reminder Sent', 'Calendar Sync', 'Calendar Event Title'
                            ]
                            for col in template_columns:
                                if col not in import_df.columns:
                                    import_df[col] = ''
                            
                            # Fill NaN values with empty strings for text columns
                            text_columns = ['Meeting ID', 'Meeting Title', 'Organization', 'Client', 'Stakeholder Name', 'Purpose', 
                                          'Agenda', 'Start Time', 'Time Zone', 'Meeting Type', 'Meeting Link', 
                                          'Location', 'Status', 'Priority', 'Attendees', 'Internal External Guests', 'Notes', 
                                          'Next Action', 'Reminder Sent', 'Calendar Sync', 'Calendar Event Title']
                            for col in text_columns:
                                if col in import_df.columns:
                                    import_df[col] = import_df[col].fillna('').astype(str)
                            # For date columns, keep NaT for missing dates (already handled above)
                            
                            # Handle Status - only calculate for rows with data, set to empty for empty rows
                            if 'Status' not in import_df.columns:
                                import_df['Status'] = ''
                            
                            # Calculate status only for rows with Meeting Date and Start Time
                            for idx, row in import_df.iterrows():
                                if pd.isna(row.get('Status')) or str(row.get('Status', '')).strip() == '':
                                    # Check if row has date and time
                                    has_date = pd.notna(row.get('Meeting Date')) and str(row.get('Meeting Date', '')).strip() != ''
                                    has_time = pd.notna(row.get('Start Time')) and str(row.get('Start Time', '')).strip() != ''
                                    if has_date and has_time:
                                        import_df.at[idx, 'Status'] = calculate_status(row)
                                    else:
                                        import_df.at[idx, 'Status'] = ''
                            
                            # Get current dataframe
                            current_df = st.session_state.meetings_df.copy()
                            
                            if current_df.empty:
                                # If no existing data, just add all
                                if 'Meeting ID' not in import_df.columns or import_df['Meeting ID'].isna().all():
                                    # Generate meeting IDs
                                    import_df['Meeting ID'] = range(1, len(import_df) + 1)
                                st.session_state.meetings_df = import_df.copy()
                                added_count = len(import_df)
                                updated_count = 0
                            else:
                                # Handle Meeting ID
                                if 'Meeting ID' not in import_df.columns or import_df['Meeting ID'].isna().all():
                                    # Generate new IDs for records without IDs
                                    if 'Meeting ID' in current_df.columns:
                                        max_id = pd.to_numeric(current_df['Meeting ID'], errors='coerce').max()
                                        if pd.isna(max_id):
                                            max_id = 0
                                    else:
                                        max_id = 0
                                    import_df['Meeting ID'] = range(int(max_id) + 1, int(max_id) + 1 + len(import_df))
                                
                                # Convert Meeting ID for comparison
                                import_df['Meeting ID'] = pd.to_numeric(import_df['Meeting ID'], errors='coerce')
                                if 'Meeting ID' in current_df.columns:
                                    current_df['Meeting ID'] = pd.to_numeric(current_df['Meeting ID'], errors='coerce')
                                
                                added_count = 0
                                updated_count = 0
                                
                                if import_mode == "Add New Only":
                                    # Only add records with new Meeting IDs
                                    if 'Meeting ID' in current_df.columns:
                                        existing_ids = set(pd.to_numeric(current_df['Meeting ID'], errors='coerce').dropna().astype(int))
                                    else:
                                        existing_ids = set()
                                    if 'Meeting ID' in import_df.columns:
                                        # Create boolean mask with same index as import_df
                                        import_df_ids = pd.to_numeric(import_df['Meeting ID'], errors='coerce')
                                        mask = ~import_df_ids.isin(existing_ids) | import_df_ids.isna()
                                        new_records = import_df[mask].copy()
                                    else:
                                        new_records = import_df.copy()
                                    if not new_records.empty:
                                        # Recalculate status for new records
                                        new_records['Status'] = new_records.apply(calculate_status, axis=1)
                                        st.session_state.meetings_df = pd.concat([current_df, new_records], ignore_index=True)
                                        added_count = len(new_records)
                                
                                elif import_mode == "Update Existing":
                                    # Only update existing records
                                    if 'Meeting ID' in current_df.columns and 'Meeting ID' in import_df.columns:
                                        existing_ids = set(pd.to_numeric(current_df['Meeting ID'], errors='coerce').dropna().astype(int))
                                        # Create boolean mask with same index as import_df
                                        import_df_ids = pd.to_numeric(import_df['Meeting ID'], errors='coerce')
                                        mask = import_df_ids.isin(existing_ids) & import_df_ids.notna()
                                        to_update = import_df[mask].copy()
                                    else:
                                        to_update = pd.DataFrame()
                                    if not to_update.empty:
                                        for _, row in to_update.iterrows():
                                            meeting_id = pd.to_numeric(row.get('Meeting ID'), errors='coerce')
                                            if pd.notna(meeting_id) and 'Meeting ID' in current_df.columns:
                                                idx = current_df[pd.to_numeric(current_df['Meeting ID'], errors='coerce') == meeting_id].index[0]
                                                # Update all fields
                                                for col in current_df.columns:
                                                    if col in row and col != 'Status':
                                                        current_df.at[idx, col] = row[col]
                                                    elif col == 'Status':
                                                        if overwrite_status:
                                                            current_df.at[idx, col] = row.get('Status', calculate_status(row))
                                                # Recalculate status if overwrite is enabled or status is missing
                                                if overwrite_status or pd.isna(row.get('Status')):
                                                    current_df.at[idx, 'Status'] = calculate_status(current_df.iloc[idx])
                                        st.session_state.meetings_df = current_df
                                        updated_count = len(to_update)
                                    else:
                                        st.warning("No records with matching Meeting ID found to update.")
                                        st.session_state.meetings_df = current_df
                                
                                else:  # Update & Add New
                                    if 'Meeting ID' in current_df.columns and 'Meeting ID' in import_df.columns:
                                        existing_ids = set(pd.to_numeric(current_df['Meeting ID'], errors='coerce').dropna().astype(int))
                                        # Create boolean mask with same index as import_df
                                        import_df_ids = pd.to_numeric(import_df['Meeting ID'], errors='coerce')
                                        mask_update = import_df_ids.isin(existing_ids) & import_df_ids.notna()
                                        mask_add = ~mask_update
                                        to_update = import_df[mask_update].copy()
                                        to_add = import_df[mask_add].copy()
                                    else:
                                        to_update = pd.DataFrame()
                                        to_add = import_df.copy()
                                    
                                    # Update existing
                                    if not to_update.empty:
                                        for _, row in to_update.iterrows():
                                            meeting_id = pd.to_numeric(row.get('Meeting ID'), errors='coerce')
                                            if pd.notna(meeting_id) and 'Meeting ID' in current_df.columns:
                                                idx = current_df[pd.to_numeric(current_df['Meeting ID'], errors='coerce') == meeting_id].index[0]
                                                for col in current_df.columns:
                                                    if col in row and col != 'Status':
                                                        current_df.at[idx, col] = row[col]
                                                    elif col == 'Status':
                                                        if overwrite_status:
                                                            current_df.at[idx, col] = row.get('Status', calculate_status(row))
                                                if overwrite_status or pd.isna(row.get('Status')):
                                                    current_df.at[idx, 'Status'] = calculate_status(current_df.iloc[idx])
                                        updated_count = len(to_update)
                                    
                                    # Add new
                                    if not to_add.empty:
                                        to_add['Status'] = to_add.apply(calculate_status, axis=1)
                                        current_df = pd.concat([current_df, to_add], ignore_index=True)
                                        added_count = len(to_add)
                                    
                                    st.session_state.meetings_df = current_df
                            
                            # Save to Excel
                            if save_meetings(st.session_state.meetings_df):
                                success_msg = "✅ Import completed successfully!"
                                if added_count > 0:
                                    success_msg += f" Added {added_count} new meeting(s)."
                                if updated_count > 0:
                                    success_msg += f" Updated {updated_count} existing meeting(s)."
                                st.success(success_msg)
                                time.sleep(1)
                                st.rerun()
                            else:
                                st.error("Failed to save imported data.")
                    
                    except Exception as e:
                            st.error(f"Error during import: {str(e)}")
                            import traceback
                            st.code(traceback.format_exc())
        
        except Exception as e:
            st.error(f"Error reading file: {str(e)}")
            st.info("Please ensure the file is a valid Excel file (.xlsx or .xls format)")
    
    # Export section
    st.markdown("---")
    st.markdown("""
    <div style="background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
                padding: 1.5rem;
                border-radius: 12px;
                margin-bottom: 1.5rem;
                border-left: 4px solid #06b6d4;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);">
        <h2 style="margin: 0; color: #1e293b; font-size: 1.5rem; font-weight: 600;">📥 Export Data</h2>
    </div>
    """, unsafe_allow_html=True)
    
    if not st.session_state.meetings_df.empty:
        col1, col2 = st.columns([1, 3])
        with col1:
            if st.button("📥 Export to Excel", type="primary", use_container_width=True):
                export_filename = f"meeting_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                try:
                    st.session_state.meetings_df.to_excel(export_filename, index=False)
                    st.success(f"✅ Data exported to {export_filename}")
                    
                    # Provide download button
                    with open(export_filename, "rb") as file:
                        st.download_button(
                            label="⬇️ Download Exported File",
                            data=file,
                            file_name=export_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                except Exception as e:
                    st.error(f"Error exporting data: {e}")
    else:
        st.info("📭 No data available to export. Add meetings first.")

