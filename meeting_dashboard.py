import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
from pathlib import Path
import time

# Configuration
EXCEL_FILE = "meeting_data.xlsx"
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
            # Ensure datetime columns are properly formatted
            if 'start_datetime' in df.columns:
                df['start_datetime'] = pd.to_datetime(df['start_datetime'])
            if 'end_datetime' in df.columns:
                df['end_datetime'] = pd.to_datetime(df['end_datetime'])
            
            # Ensure new columns exist (for backward compatibility)
            required_columns = ['stakeholder', 'attendees_internal', 'attendees_external']
            for col in required_columns:
                if col not in df.columns:
                    df[col] = ''
            
            return df
        except Exception as e:
            st.error(f"Error loading meetings: {e}")
            return pd.DataFrame(columns=[
                'meeting_id', 'title', 'start_datetime', 'end_datetime',
                'meeting_type', 'organizer', 'location_or_link', 'notes', 
                'stakeholder', 'attendees_internal', 'attendees_external', 'status'
            ])
    else:
        return pd.DataFrame(columns=[
            'meeting_id', 'title', 'start_datetime', 'end_datetime',
            'meeting_type', 'organizer', 'location_or_link', 'notes',
            'stakeholder', 'attendees_internal', 'attendees_external', 'status'
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
    start = pd.to_datetime(row['start_datetime'])
    end = pd.to_datetime(row['end_datetime'])
    
    if now < start:
        return "Upcoming"
    elif start <= now < end:
        return "Ongoing"
    else:
        return "Ended"

def get_next_meeting_id(df):
    """Get the next available meeting ID"""
    if df.empty or 'meeting_id' not in df.columns:
        return 1
    if df['meeting_id'].isna().all():
        return 1
    return int(df['meeting_id'].max()) + 1

def update_all_statuses(df):
    """Update status for all meetings and save to Excel"""
    if not df.empty:
        # Only update status for meetings that are not manually set to "Completed"
        mask = df['status'] != 'Completed'
        if mask.any():
            df.loc[mask, 'status'] = df.loc[mask].apply(calculate_status, axis=1)
        save_meetings(df)
    return df

def load_data():
    """Load data into session state"""
    if not st.session_state.data_loaded:
        st.session_state.meetings_df = load_meetings()
        st.session_state.data_loaded = True
    
    # Recalculate status on each load (for real-time updates)
    # But preserve manually set "Completed" status
    if not st.session_state.meetings_df.empty:
        # Only recalculate status for meetings that are not manually set to "Completed"
        mask = st.session_state.meetings_df['status'] != 'Completed'
        if mask.any():
            st.session_state.meetings_df.loc[mask, 'status'] = st.session_state.meetings_df.loc[mask].apply(calculate_status, axis=1)

def filter_meetings(df, status_filter, date_start, date_end, search_text):
    """Filter meetings based on criteria"""
    filtered_df = df.copy()
    
    # Status filter
    if status_filter != "All":
        filtered_df = filtered_df[filtered_df['status'] == status_filter]
    
    # Date range filter
    if date_start:
        filtered_df = filtered_df[pd.to_datetime(filtered_df['start_datetime']) >= pd.to_datetime(date_start)]
    if date_end:
        # Add end of day to date_end
        date_end_datetime = pd.to_datetime(date_end) + timedelta(days=1) - timedelta(seconds=1)
        filtered_df = filtered_df[pd.to_datetime(filtered_df['start_datetime']) <= date_end_datetime]
    
    # Search filter
    if search_text:
        search_text_lower = search_text.lower()
        search_mask = (
            filtered_df['title'].astype(str).str.lower().str.contains(search_text_lower, na=False) |
            filtered_df['organizer'].astype(str).str.lower().str.contains(search_text_lower, na=False)
        )
        
        # Add optional fields to search if they exist
        if 'stakeholder' in filtered_df.columns:
            search_mask |= filtered_df['stakeholder'].astype(str).str.lower().str.contains(search_text_lower, na=False)
        if 'attendees_internal' in filtered_df.columns:
            search_mask |= filtered_df['attendees_internal'].astype(str).str.lower().str.contains(search_text_lower, na=False)
        if 'attendees_external' in filtered_df.columns:
            search_mask |= filtered_df['attendees_external'].astype(str).str.lower().str.contains(search_text_lower, na=False)
        
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

# Custom CSS for better styling
st.markdown("""
<style>
    /* Main container styling */
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    
    /* Header styling */
    h1 {
        color: #0066CC;
        border-bottom: 3px solid #0066CC;
        padding-bottom: 0.5rem;
        margin-bottom: 1.5rem;
    }
    
    h2 {
        color: #004499;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
    
    h3 {
        color: #0055AA;
        margin-top: 1.5rem;
        margin-bottom: 0.75rem;
    }
    
    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background-color: #F0F4F8;
    }
    
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h1 {
        color: #0066CC;
        border-bottom: 2px solid #0066CC;
    }
    
    /* Button styling */
    .stButton>button {
        background-color: #0066CC;
        color: white;
        border-radius: 8px;
        border: none;
        padding: 0.5rem 1.5rem;
        font-weight: 500;
        transition: all 0.3s;
    }
    
    .stButton>button:hover {
        background-color: #0055AA;
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0, 102, 204, 0.3);
    }
    
    /* Secondary button */
    button[kind="secondary"] {
        background-color: #6C757D;
        color: white;
    }
    
    button[kind="secondary"]:hover {
        background-color: #5A6268;
    }
    
    /* Success messages */
    .stSuccess {
        background-color: #D4EDDA;
        border-left: 4px solid #28A745;
        color: #155724;
        padding: 1rem;
        border-radius: 4px;
    }
    
    /* Error messages */
    .stError {
        background-color: #F8D7DA;
        border-left: 4px solid #DC3545;
        color: #721C24;
        padding: 1rem;
        border-radius: 4px;
    }
    
    /* Info messages */
    .stInfo {
        background-color: #D1ECF1;
        border-left: 4px solid #17A2B8;
        color: #0C5460;
        padding: 1rem;
        border-radius: 4px;
    }
    
    /* Warning messages */
    .stWarning {
        background-color: #FFF3CD;
        border-left: 4px solid #FFC107;
        color: #856404;
        padding: 1rem;
        border-radius: 4px;
    }
    
    /* Metric cards */
    [data-testid="stMetricValue"] {
        color: #0066CC;
        font-size: 2rem;
        font-weight: bold;
    }
    
    [data-testid="stMetricLabel"] {
        color: #6C757D;
        font-weight: 600;
    }
    
    /* Dataframe styling */
    .dataframe {
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    
    /* Input fields */
    .stTextInput>div>div>input,
    .stTextArea>div>div>textarea,
    .stSelectbox>div>div>select {
        border-radius: 6px;
        border: 1px solid #DEE2E6;
    }
    
    .stTextInput>div>div>input:focus,
    .stTextArea>div>div>textarea:focus,
    .stSelectbox>div>div>select:focus {
        border-color: #0066CC;
        box-shadow: 0 0 0 3px rgba(0, 102, 204, 0.1);
    }
    
    /* Expander styling */
    .streamlit-expanderHeader {
        background-color: #E9ECEF;
        border-radius: 6px;
        padding: 0.75rem;
        font-weight: 600;
        color: #004499;
    }
    
    /* Divider/HR styling */
    hr {
        border: none;
        height: 2px;
        background: linear-gradient(to right, #0066CC, transparent);
        margin: 2rem 0;
    }
    
    /* Sidebar radio buttons */
    [data-testid="stSidebar"] [data-testid="stRadio"] label {
        padding: 0.5rem;
        border-radius: 6px;
        margin: 0.25rem 0;
        transition: all 0.2s;
    }
    
    [data-testid="stSidebar"] [data-testid="stRadio"] label:hover {
        background-color: #E9ECEF;
    }
    
    /* File uploader */
    [data-testid="stFileUploader"] {
        border: 2px dashed #0066CC;
        border-radius: 8px;
        padding: 2rem;
        background-color: #F8F9FA;
    }
    
    /* Tabs if used */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 6px 6px 0 0;
        padding: 0.5rem 1.5rem;
    }
    
    /* Date and time inputs */
    [data-testid="stDateInput"] {
        border-radius: 6px;
    }
    
    [data-testid="stTimeInput"] {
        border-radius: 6px;
    }
    
    /* Checkbox styling */
    [data-testid="stCheckbox"] label {
        font-weight: 500;
        color: #1A1A1A;
    }
    
    /* Caption styling */
    .stCaption {
        color: #6C757D;
        font-style: italic;
    }
    
    /* Status badges styling */
    .status-badge {
        padding: 0.25rem 0.75rem;
        border-radius: 12px;
        font-size: 0.85rem;
        font-weight: 600;
        display: inline-block;
    }
    
    .status-upcoming {
        background-color: #C8E6C9;
        color: #2E7D32;
    }
    
    .status-ongoing {
        background-color: #BBDEFB;
        color: #1565C0;
    }
    
    .status-ended {
        background-color: #E0E0E0;
        color: #424242;
    }
    
    .status-completed {
        background-color: #C5E1A5;
        color: #33691E;
    }
    
    /* Card-like containers */
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: white;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    
    /* Smooth transitions */
    * {
        transition: background-color 0.2s ease, color 0.2s ease;
    }
    
    /* Form sections */
    .form-section {
        background-color: #FFFFFF;
        padding: 1.5rem;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
        margin-bottom: 1.5rem;
    }
    
    /* Table row hover effect */
    .dataframe tbody tr:hover {
        background-color: #F0F4F8 !important;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar Navigation
st.sidebar.title("📅 Meeting Dashboard")
st.sidebar.markdown("---")

# Page selection
page = st.sidebar.radio(
    "Navigate to:",
    ["1️⃣ Add New Meeting", "2️⃣ Edit or Delete Meeting", "3️⃣ Meetings Summary & Export"],
    index=0 if st.session_state.current_page == "Add New Meeting" else 
          1 if st.session_state.current_page == "Edit or Delete Meeting" else 2
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
st.sidebar.subheader("⚙️ Settings")

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

# Main Title
st.title(f"📅 AI Geo Navigators - {st.session_state.current_page}")

# ============================================================================
# PAGE 1: Add New Meeting
# ============================================================================
if st.session_state.current_page == "Add New Meeting":
    st.markdown("### ➕ Create a New Meeting")
    st.markdown("Fill in the form below to add a new meeting to the dashboard.")
    
    with st.form("add_meeting_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            title = st.text_input(
                "Title *", 
                value="",
                placeholder="Enter the purpose of the meeting",
                help="Enter the purpose or title of the meeting"
            )
            start_date = st.date_input("Start Date *", value=datetime.now().date())
            start_time = st.time_input("Start Time *", value=datetime.now().time())
            end_date = st.date_input("End Date *", value=datetime.now().date())
            end_time = st.time_input("End Time *", value=(datetime.now() + timedelta(hours=1)).time())
        
        with col2:
            meeting_type = st.selectbox("Meeting Type *", ["In Person", "Virtual"])
            organizer = st.text_input(
                "Organizer Name *", 
                value="",
                placeholder="Enter organizer's full name",
                help="Enter the name of the person organizing this meeting"
            )
            location_or_link = st.text_input("Location or Link", value="")
        
        # New fields section
        st.markdown("**Attendees & Stakeholders**")
        col_attendees1, col_attendees2 = st.columns(2)
        
        with col_attendees1:
            stakeholder = st.text_input(
                "Stakeholder",
                value="",
                placeholder="Enter stakeholder name(s)",
                help="Enter the name(s) of key stakeholders for this meeting"
            )
            attendees_internal = st.text_input(
                "Attendees Internal",
                value="",
                placeholder="Enter internal team member names (comma-separated)",
                help="Enter names of internal team members attending (separate multiple names with commas)"
            )
        
        with col_attendees2:
            attendees_external = st.text_input(
                "External Guests",
                value="",
                placeholder="Enter external guest names (comma-separated)",
                help="Enter names of external guests or clients attending (separate multiple names with commas)"
            )
        
        notes = st.text_area(
            "Agenda or Notes", 
            value="", 
            height=100,
            placeholder="Enter the overall context of the meeting, agenda items, discussion points, or any relevant notes...",
            help="Provide the overall context of the meeting, including agenda items, discussion points, and any relevant notes"
        )
        
        submitted = st.form_submit_button("💾 Save Meeting", type="primary", use_container_width=True)
        
        if submitted:
            # Validation
            errors = []
            if not title.strip():
                errors.append("Title is required")
            
            if not organizer.strip():
                errors.append("Organizer is required")
            
            # Combine date and time
            start_datetime = datetime.combine(start_date, start_time)
            end_datetime = datetime.combine(end_date, end_time)
            
            if end_datetime <= start_datetime:
                errors.append("End date/time must be after start date/time")
            
            if errors:
                for error in errors:
                    st.error(error)
            else:
                # Create new meeting
                new_meeting = pd.DataFrame([{
                    'meeting_id': get_next_meeting_id(st.session_state.meetings_df),
                    'title': title.strip(),
                    'start_datetime': start_datetime,
                    'end_datetime': end_datetime,
                    'meeting_type': meeting_type,
                    'organizer': organizer.strip(),
                    'location_or_link': location_or_link.strip(),
                    'notes': notes.strip(),
                    'stakeholder': stakeholder.strip(),
                    'attendees_internal': attendees_internal.strip(),
                    'attendees_external': attendees_external.strip()
                }])
                
                new_meeting['status'] = new_meeting.apply(calculate_status, axis=1)
                
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
        # Create selection list
        meeting_options = {}
        for idx, row in st.session_state.meetings_df.iterrows():
            label = f"{row['title']} - {pd.to_datetime(row['start_datetime']).strftime('%Y-%m-%d %H:%M')}"
            meeting_options[label] = int(row['meeting_id'])
        
        selected_meeting_label = st.selectbox("Select Meeting to Edit/Delete", list(meeting_options.keys()))
        selected_meeting_id = meeting_options[selected_meeting_label]
        
        selected_meeting = st.session_state.meetings_df[st.session_state.meetings_df['meeting_id'] == selected_meeting_id].iloc[0]
        
        # Display current meeting info
        with st.expander("📋 View Current Meeting Details", expanded=False):
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"**Title:** {selected_meeting['title']}")
                st.write(f"**Status:** {selected_meeting.get('status', 'N/A')}")
                st.write(f"**Type:** {selected_meeting['meeting_type']}")
                st.write(f"**Organizer:** {selected_meeting['organizer']}")
                if selected_meeting.get('stakeholder'):
                    st.write(f"**Stakeholder:** {selected_meeting.get('stakeholder', 'N/A')}")
            with col2:
                st.write(f"**Start:** {pd.to_datetime(selected_meeting['start_datetime']).strftime(DATE_FORMAT)}")
                st.write(f"**End:** {pd.to_datetime(selected_meeting['end_datetime']).strftime(DATE_FORMAT)}")
                if selected_meeting.get('location_or_link'):
                    st.write(f"**Location/Link:** {selected_meeting['location_or_link']}")
                if selected_meeting.get('attendees_internal'):
                    st.write(f"**Internal Attendees:** {selected_meeting.get('attendees_internal', 'N/A')}")
                if selected_meeting.get('attendees_external'):
                    st.write(f"**External Guests:** {selected_meeting.get('attendees_external', 'N/A')}")
        
        # Edit form
        st.markdown("---")
        with st.form("edit_meeting_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                edit_title = st.text_input(
                    "Title *", 
                    value=selected_meeting['title'],
                    placeholder="Enter the purpose of the meeting",
                    help="Enter the purpose or title of the meeting"
                )
                edit_start_datetime = pd.to_datetime(selected_meeting['start_datetime'])
                edit_start_date = st.date_input("Start Date *", value=edit_start_datetime.date())
                edit_start_time = st.time_input("Start Time *", value=edit_start_datetime.time())
                edit_end_datetime = pd.to_datetime(selected_meeting['end_datetime'])
                edit_end_date = st.date_input("End Date *", value=edit_end_datetime.date())
                edit_end_time = st.time_input("End Time *", value=edit_end_datetime.time())
            
            with col2:
                edit_meeting_type = st.selectbox("Meeting Type *", ["In Person", "Virtual"], 
                                               index=0 if selected_meeting['meeting_type'] == "In Person" else 1)
                edit_organizer = st.text_input(
                    "Organizer Name *", 
                    value=selected_meeting['organizer'],
                    placeholder="Enter organizer's full name",
                    help="Enter the name of the person organizing this meeting"
                )
                edit_location_or_link = st.text_input("Location or Link", 
                                                    value=selected_meeting.get('location_or_link', ''))
            
            # Attendees & Stakeholders section
            st.markdown("**Attendees & Stakeholders**")
            col_attendees1, col_attendees2 = st.columns(2)
            
            with col_attendees1:
                edit_stakeholder = st.text_input(
                    "Stakeholder",
                    value=selected_meeting.get('stakeholder', ''),
                    placeholder="Enter stakeholder name(s)",
                    help="Enter the name(s) of key stakeholders for this meeting"
                )
                edit_attendees_internal = st.text_input(
                    "Attendees Internal",
                    value=selected_meeting.get('attendees_internal', ''),
                    placeholder="Enter internal team member names (comma-separated)",
                    help="Enter names of internal team members attending (separate multiple names with commas)"
                )
            
            with col_attendees2:
                edit_attendees_external = st.text_input(
                    "External Guests",
                    value=selected_meeting.get('attendees_external', ''),
                    placeholder="Enter external guest names (comma-separated)",
                    help="Enter names of external guests or clients attending (separate multiple names with commas)"
                )
            
            # Status selection section
            st.markdown("**Status**")
            status_col1, status_col2 = st.columns([2, 1])
            with status_col1:
                # Calculate current auto status
                current_auto_status = calculate_status(selected_meeting)
                status_options = ["Upcoming", "Ongoing", "Ended", "Completed"]
                
                # Find current status index
                current_status = selected_meeting.get('status', current_auto_status)
                try:
                    current_status_index = status_options.index(current_status) if current_status in status_options else status_options.index("Ended")
                except ValueError:
                    current_status_index = status_options.index("Ended")
                
                edit_status = st.selectbox(
                    "Meeting Status *", 
                    status_options,
                    index=current_status_index,
                    help="Manually set the meeting status. 'Completed' can be used to mark meetings as finished regardless of time."
                )
            
            with status_col2:
                use_auto_status = st.checkbox(
                    "Use Auto Status", 
                    value=False,
                    help="If checked, status will be automatically calculated based on current time"
                )
                if use_auto_status:
                    st.caption(f"Auto Status: {current_auto_status}")
            
            edit_notes = st.text_area(
                "Agenda or Notes", 
                value=selected_meeting.get('notes', ''), 
                height=100,
                placeholder="Enter the overall context of the meeting, agenda items, discussion points, or any relevant notes...",
                help="Provide the overall context of the meeting, including agenda items, discussion points, and any relevant notes"
            )
            
            col_btn1, col_btn2 = st.columns([1, 1])
            with col_btn1:
                update_submitted = st.form_submit_button("💾 Update Meeting", type="primary", use_container_width=True)
            with col_btn2:
                pass  # Delete button is outside form
            
            if update_submitted:
                # Validation
                errors = []
                if not edit_title.strip():
                    errors.append("Title is required")
                
                if not edit_organizer.strip():
                    errors.append("Organizer is required")
                
                # Combine date and time
                edit_start_dt = datetime.combine(edit_start_date, edit_start_time)
                edit_end_dt = datetime.combine(edit_end_date, edit_end_time)
                
                if edit_end_dt <= edit_start_dt:
                    errors.append("End date/time must be after start date/time")
                
                if errors:
                    for error in errors:
                        st.error(error)
                else:
                    # Update meeting
                    idx = st.session_state.meetings_df[st.session_state.meetings_df['meeting_id'] == selected_meeting_id].index[0]
                    st.session_state.meetings_df.at[idx, 'title'] = edit_title.strip()
                    st.session_state.meetings_df.at[idx, 'start_datetime'] = edit_start_dt
                    st.session_state.meetings_df.at[idx, 'end_datetime'] = edit_end_dt
                    st.session_state.meetings_df.at[idx, 'meeting_type'] = edit_meeting_type
                    st.session_state.meetings_df.at[idx, 'organizer'] = edit_organizer.strip()
                    st.session_state.meetings_df.at[idx, 'location_or_link'] = edit_location_or_link.strip()
                    st.session_state.meetings_df.at[idx, 'notes'] = edit_notes.strip()
                    st.session_state.meetings_df.at[idx, 'stakeholder'] = edit_stakeholder.strip()
                    st.session_state.meetings_df.at[idx, 'attendees_internal'] = edit_attendees_internal.strip()
                    st.session_state.meetings_df.at[idx, 'attendees_external'] = edit_attendees_external.strip()
                    
                    # Set status - use manual selection or auto-calculate based on checkbox
                    if use_auto_status:
                        # Create a temporary row with updated values to calculate status
                        temp_row = st.session_state.meetings_df.iloc[idx].copy()
                        temp_row['start_datetime'] = edit_start_dt
                        temp_row['end_datetime'] = edit_end_dt
                        st.session_state.meetings_df.at[idx, 'status'] = calculate_status(temp_row)
                    else:
                        st.session_state.meetings_df.at[idx, 'status'] = edit_status
                    
                    # Save to Excel
                    if save_meetings(st.session_state.meetings_df):
                        status_msg = f"✅ Meeting updated successfully! Status set to: {st.session_state.meetings_df.at[idx, 'status']}"
                        st.success(status_msg)
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
            col1, col2 = st.columns(2)
            with col1:
                if st.button("✅ Confirm Delete", type="primary", use_container_width=True):
                    # Remove meeting
                    st.session_state.meetings_df = st.session_state.meetings_df[
                        st.session_state.meetings_df['meeting_id'] != selected_meeting_id
                    ]
                    
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
    st.subheader("🔍 Filters")
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
    st.subheader("📈 Summary Statistics")
    col1, col2, col3, col4 = st.columns(4)
    
    if not filtered_meetings.empty:
        total_count = len(filtered_meetings)
        upcoming_count = len(filtered_meetings[filtered_meetings['status'] == 'Upcoming'])
        ongoing_count = len(filtered_meetings[filtered_meetings['status'] == 'Ongoing'])
        ended_count = len(filtered_meetings[filtered_meetings['status'] == 'Ended'])
        completed_count = len(filtered_meetings[filtered_meetings['status'] == 'Completed'])
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
    st.subheader("📋 Meetings Table")
    
    if not filtered_meetings.empty:
        # Prepare display dataframe
        display_df = filtered_meetings.copy()
        display_df['start_datetime'] = pd.to_datetime(display_df['start_datetime']).dt.strftime(DATE_FORMAT)
        display_df['end_datetime'] = pd.to_datetime(display_df['end_datetime']).dt.strftime(DATE_FORMAT)
        
        # Select columns to display
        display_columns = ['title', 'start_datetime', 'end_datetime', 'status', 
                          'meeting_type', 'organizer', 'stakeholder', 'attendees_internal', 
                          'attendees_external', 'location_or_link', 'notes']
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
    st.subheader("📤 Import/Update from Excel")
    
    # Template download option
    col_template1, col_template2 = st.columns([3, 1])
    with col_template1:
        st.write("Upload an Excel file to import or update meeting records.")
    with col_template2:
        # Create template dataframe
        template_df = pd.DataFrame(columns=[
            'meeting_id', 'title', 'start_datetime', 'end_datetime',
            'meeting_type', 'organizer', 'location_or_link', 'notes',
            'stakeholder', 'attendees_internal', 'attendees_external', 'status'
        ])
        # Add sample row
        template_df = pd.concat([template_df, pd.DataFrame([{
            'meeting_id': 1,
            'title': 'Sample Meeting',
            'start_datetime': datetime.now(),
            'end_datetime': datetime.now() + timedelta(hours=1),
            'meeting_type': 'Virtual',
            'organizer': 'John Doe',
            'location_or_link': 'https://meet.example.com',
            'notes': 'Sample agenda items',
            'stakeholder': 'Jane Smith',
            'attendees_internal': 'Team Member 1, Team Member 2',
            'attendees_external': 'Client A, Client B',
            'status': 'Upcoming'
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
        help="Upload an Excel file with meeting data. Required columns: title, start_datetime, end_datetime, meeting_type, organizer. Optional: meeting_id, location_or_link, notes, stakeholder, attendees_internal, attendees_external, status"
    )
    
    if uploaded_file is not None:
        try:
            # Read the uploaded file
            import_df = pd.read_excel(uploaded_file)
            
            # Check required columns
            required_columns = ['title', 'start_datetime', 'end_datetime', 'meeting_type', 'organizer']
            missing_columns = [col for col in required_columns if col not in import_df.columns]
            
            if missing_columns:
                st.error(f"❌ Missing required columns: {', '.join(missing_columns)}")
                st.info("Required columns: title, start_datetime, end_datetime, meeting_type, organizer")
            else:
                # Ensure datetime columns are properly formatted
                if 'start_datetime' in import_df.columns:
                    import_df['start_datetime'] = pd.to_datetime(import_df['start_datetime'], errors='coerce')
                if 'end_datetime' in import_df.columns:
                    import_df['end_datetime'] = pd.to_datetime(import_df['end_datetime'], errors='coerce')
                
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
                
                # Validate data before import
                validation_errors = []
                for idx, row in import_df.iterrows():
                    if pd.isna(row.get('title')) or str(row.get('title', '')).strip() == '':
                        validation_errors.append(f"Row {idx + 1}: Title is required")
                    if pd.isna(row.get('organizer')) or str(row.get('organizer', '')).strip() == '':
                        validation_errors.append(f"Row {idx + 1}: Organizer is required")
                    if pd.isna(row.get('start_datetime')):
                        validation_errors.append(f"Row {idx + 1}: Start datetime is required and must be valid")
                    if pd.isna(row.get('end_datetime')):
                        validation_errors.append(f"Row {idx + 1}: End datetime is required and must be valid")
                    if not pd.isna(row.get('start_datetime')) and not pd.isna(row.get('end_datetime')):
                        if row['start_datetime'] >= row['end_datetime']:
                            validation_errors.append(f"Row {idx + 1}: End datetime must be after start datetime")
                
                if validation_errors:
                    st.warning("⚠️ **Validation Errors Found:**")
                    for error in validation_errors[:10]:  # Show first 10 errors
                        st.write(f"- {error}")
                    if len(validation_errors) > 10:
                        st.write(f"- ... and {len(validation_errors) - 10} more errors")
                else:
                    # Proceed with import
                    if st.button("✅ Import Data", type="primary", use_container_width=True):
                        try:
                            # Ensure required columns exist in import_df
                            for col in ['stakeholder', 'attendees_internal', 'attendees_external', 'location_or_link', 'notes']:
                                if col not in import_df.columns:
                                    import_df[col] = ''
                            if 'status' not in import_df.columns:
                                # Calculate status for new records
                                import_df['status'] = import_df.apply(calculate_status, axis=1)
                            
                            # Get current dataframe
                            current_df = st.session_state.meetings_df.copy()
                            
                            if current_df.empty:
                                # If no existing data, just add all
                                if 'meeting_id' not in import_df.columns:
                                    # Generate meeting IDs
                                    import_df['meeting_id'] = range(1, len(import_df) + 1)
                                st.session_state.meetings_df = import_df.copy()
                                added_count = len(import_df)
                                updated_count = 0
                            else:
                                # Handle meeting_id
                                if 'meeting_id' not in import_df.columns:
                                    # Generate new IDs for records without IDs
                                    max_id = current_df['meeting_id'].max() if 'meeting_id' in current_df.columns else 0
                                    import_df['meeting_id'] = range(int(max_id) + 1, int(max_id) + 1 + len(import_df))
                                
                                # Convert meeting_id to int for comparison
                                import_df['meeting_id'] = import_df['meeting_id'].astype(int)
                                current_df['meeting_id'] = current_df['meeting_id'].astype(int)
                                
                                added_count = 0
                                updated_count = 0
                                
                                if import_mode == "Add New Only":
                                    # Only add records with new meeting_ids
                                    existing_ids = set(current_df['meeting_id'].values)
                                    new_records = import_df[~import_df['meeting_id'].isin(existing_ids)].copy()
                                    if not new_records.empty:
                                        # Recalculate status for new records
                                        new_records['status'] = new_records.apply(calculate_status, axis=1)
                                        st.session_state.meetings_df = pd.concat([current_df, new_records], ignore_index=True)
                                        added_count = len(new_records)
                                
                                elif import_mode == "Update Existing":
                                    # Only update existing records
                                    existing_ids = set(current_df['meeting_id'].values)
                                    to_update = import_df[import_df['meeting_id'].isin(existing_ids)].copy()
                                    if not to_update.empty:
                                        for _, row in to_update.iterrows():
                                            idx = current_df[current_df['meeting_id'] == row['meeting_id']].index[0]
                                            # Update all fields
                                            for col in current_df.columns:
                                                if col in row and col != 'status':
                                                    current_df.at[idx, col] = row[col]
                                                elif col == 'status':
                                                    if overwrite_status:
                                                        current_df.at[idx, col] = row.get('status', calculate_status(row))
                                                    # Otherwise keep existing status
                                            # Recalculate status if overwrite is enabled or status is missing
                                            if overwrite_status or pd.isna(row.get('status')):
                                                current_df.at[idx, 'status'] = calculate_status(current_df.iloc[idx])
                                        st.session_state.meetings_df = current_df
                                        updated_count = len(to_update)
                                    else:
                                        st.warning("No records with matching meeting_id found to update.")
                                        st.session_state.meetings_df = current_df
                                
                                else:  # Update & Add New
                                    existing_ids = set(current_df['meeting_id'].values)
                                    to_update = import_df[import_df['meeting_id'].isin(existing_ids)].copy()
                                    to_add = import_df[~import_df['meeting_id'].isin(existing_ids)].copy()
                                    
                                    # Update existing
                                    if not to_update.empty:
                                        for _, row in to_update.iterrows():
                                            idx = current_df[current_df['meeting_id'] == row['meeting_id']].index[0]
                                            for col in current_df.columns:
                                                if col in row and col != 'status':
                                                    current_df.at[idx, col] = row[col]
                                                elif col == 'status':
                                                    if overwrite_status:
                                                        current_df.at[idx, col] = row.get('status', calculate_status(row))
                                            if overwrite_status or pd.isna(row.get('status')):
                                                current_df.at[idx, 'status'] = calculate_status(current_df.iloc[idx])
                                        updated_count = len(to_update)
                                    
                                    # Add new
                                    if not to_add.empty:
                                        to_add['status'] = to_add.apply(calculate_status, axis=1)
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
    st.subheader("📥 Export Data")
    
    if not st.session_state.meetings_df.empty:
        col1, col2 = st.columns([2, 1])
        with col1:
            st.write("Export all meeting data to an Excel file.")
        
        with col2:
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

# Footer
st.markdown("---")
st.caption("💡 Tip: Use the sidebar to navigate between pages. Enable auto-refresh to see status updates in real-time.")
