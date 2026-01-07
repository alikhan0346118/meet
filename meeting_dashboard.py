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
            return df
        except Exception as e:
            st.error(f"Error loading meetings: {e}")
            return pd.DataFrame(columns=[
                'meeting_id', 'title', 'start_datetime', 'end_datetime',
                'meeting_type', 'organizer', 'location_or_link', 'notes', 'status'
            ])
    else:
        return pd.DataFrame(columns=[
            'meeting_id', 'title', 'start_datetime', 'end_datetime',
            'meeting_type', 'organizer', 'location_or_link', 'notes', 'status'
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
        mask = (
            filtered_df['title'].astype(str).str.lower().str.contains(search_text_lower, na=False) |
            filtered_df['organizer'].astype(str).str.lower().str.contains(search_text_lower, na=False)
        )
        filtered_df = filtered_df[mask]
    
    return filtered_df

# Load data
load_data()

# Page configuration
st.set_page_config(page_title="Meeting Dashboard", page_icon="📅", layout="wide")

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
                    'notes': notes.strip()
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
            with col2:
                st.write(f"**Start:** {pd.to_datetime(selected_meeting['start_datetime']).strftime(DATE_FORMAT)}")
                st.write(f"**End:** {pd.to_datetime(selected_meeting['end_datetime']).strftime(DATE_FORMAT)}")
                if selected_meeting.get('location_or_link'):
                    st.write(f"**Location/Link:** {selected_meeting['location_or_link']}")
        
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
        search_text = st.text_input("Search (Title/Organizer)", value="")
    
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
                          'meeting_type', 'organizer', 'location_or_link', 'notes']
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
