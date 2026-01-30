import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
from pathlib import Path
import time
import psycopg2
from psycopg2.extras import execute_values, RealDictCursor
from psycopg2.pool import SimpleConnectionPool
from contextlib import contextmanager

# Configuration
EXCEL_FILE = "Meeting_Schedule_Template.xlsx"
DATE_FORMAT = "%Y-%m-%d %H:%M"

# Supabase Database Configuration
def get_db_config():
    """Get database configuration from secrets or environment"""
    try:
        # Try to get database password (preferred for direct PostgreSQL connection)
        password = st.secrets.get('SUPABASE_DB_PASSWORD', os.getenv('SUPABASE_DB_PASSWORD', ''))
        # If no password, try API key (which might be used as password in some setups)
        if not password:
            password = st.secrets.get('SUPABASE_API_KEY', os.getenv('SUPABASE_API_KEY', ''))
            # If API key format (starts with sbp_ or sb-), this won't work for direct PostgreSQL
            # But we'll try it anyway in case it's actually a password
    except:
        password = os.getenv('SUPABASE_DB_PASSWORD', '') or os.getenv('SUPABASE_API_KEY', '')
    
    return {
        'host': 'aws-1-ap-south-1.pooler.supabase.com',
        'port': 6543,
        'database': 'postgres',
        'user': 'postgres.xrpmswlgatrshvgwtvjw',
        'password': password
    }

def get_use_supabase():
    """Check if Supabase should be used"""
    try:
        use_supabase = st.secrets.get('USE_SUPABASE', 'true').lower() == 'true'
    except:
        use_supabase = os.getenv('USE_SUPABASE', 'true').lower() == 'true'
    return use_supabase

# Initialize session state
if 'meetings_df' not in st.session_state:
    st.session_state.meetings_df = pd.DataFrame()
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'current_page' not in st.session_state:
    st.session_state.current_page = "Add New Meeting"  # Default to Add New Meeting
if 'selected_meetings' not in st.session_state:
    st.session_state.selected_meetings = set()
if 'db_pool' not in st.session_state:
    st.session_state.db_pool = None
if 'supabase_connected' not in st.session_state:
    st.session_state.supabase_connected = False
if 'supabase_error' not in st.session_state:
    st.session_state.supabase_error = None
# Podcast meetings session state
if 'podcast_meetings_df' not in st.session_state:
    st.session_state.podcast_meetings_df = pd.DataFrame()
if 'podcast_data_loaded' not in st.session_state:
    st.session_state.podcast_data_loaded = False
if 'selected_podcast_meetings' not in st.session_state:
    st.session_state.selected_podcast_meetings = set()

@contextmanager
def get_db_connection():
    """Get database connection from pool or create new one"""
    conn = None
    try:
        db_config = get_db_config()
        if st.session_state.db_pool is not None:
            conn = st.session_state.db_pool.getconn()
        else:
            conn = psycopg2.connect(**db_config)
        yield conn
        conn.commit()
    except Exception as e:
        if conn:
            conn.rollback()
        raise e
    finally:
        if conn:
            if st.session_state.db_pool is not None:
                st.session_state.db_pool.putconn(conn)
            else:
                conn.close()

def init_db_pool():
    """Initialize database connection pool"""
    try:
        db_config = get_db_config()
        if db_config.get('password') and not st.session_state.db_pool:
            # Test connection first
            try:
                test_conn = psycopg2.connect(**db_config)
                test_conn.close()
                
                # If test successful, create pool
                st.session_state.db_pool = SimpleConnectionPool(
                    minconn=1,
                    maxconn=5,
                    **db_config
                )
                st.session_state.supabase_connected = True
                return True
            except psycopg2.Error as e:
                # Store error for display
                st.session_state.supabase_error = str(e)
                st.session_state.supabase_connected = False
                return False
        elif not db_config.get('password'):
            st.session_state.supabase_connected = False
            st.session_state.supabase_error = "No password configured"
            return False
        return st.session_state.supabase_connected
    except Exception as e:
        st.session_state.supabase_connected = False
        st.session_state.supabase_error = str(e)
        return False

def load_meetings_from_supabase():
    """Load meetings from Supabase database"""
    db_config = get_db_config()
    if not db_config.get('password'):
        return None
    
    try:
        with get_db_connection() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    SELECT 
                        meeting_id as "Meeting ID",
                        meeting_title as "Meeting Title",
                        organization as "Organization",
                        client as "Client",
                        stakeholder_name as "Stakeholder Name",
                        purpose as "Purpose",
                        agenda as "Agenda",
                        meeting_date as "Meeting Date",
                        start_time as "Start Time",
                        time_zone as "Time Zone",
                        meeting_type as "Meeting Type",
                        meeting_link as "Meeting Link",
                        location as "Website",
                        status as "Status",
                        priority as "Priority",
                        attendees as "Attendees",
                        internal_external_guests as "Internal External Guests",
                        notes as "Notes",
                        next_action as "Next Action",
                        follow_up_date as "Follow up Date",
                        reminder_sent as "Reminder Sent",
                        calendar_sync as "Calendar Sync",
                        calendar_event_title as "Calendar Event Title"
                    FROM meetings
                    ORDER BY meeting_id
                """)
                rows = cur.fetchall()
                
                if rows:
                    df = pd.DataFrame(rows)
                    # Convert date columns
                    if 'Meeting Date' in df.columns:
                        df['Meeting Date'] = pd.to_datetime(df['Meeting Date'], errors='coerce')
                    if 'Follow up Date' in df.columns:
                        df['Follow up Date'] = pd.to_datetime(df['Follow up Date'], errors='coerce')
                    # Convert time to string format
                    if 'Start Time' in df.columns:
                        df['Start Time'] = df['Start Time'].apply(
                            lambda x: str(x).split('.')[0] if pd.notna(x) and x != '' else ''
                        )
                    return df
                else:
                    return pd.DataFrame(columns=[
                        'Meeting ID', 'Meeting Title', 'Organization', 'Client', 'Stakeholder Name',
                        'Purpose', 'Agenda', 'Meeting Date', 'Start Time', 'Time Zone',
                        'Meeting Type', 'Meeting Link', 'Website', 'Status', 'Priority',
                        'Attendees', 'Internal External Guests', 'Notes', 'Next Action',
                        'Follow up Date', 'Reminder Sent', 'Calendar Sync', 'Calendar Event Title'
                    ])
    except Exception as e:
        st.error(f"Error loading from Supabase: {e}")
        return None

def load_meetings():
    """Load meetings from Supabase (if available) or Excel file"""
    # Try Supabase first if enabled
    if get_use_supabase() and init_db_pool():
        df = load_meetings_from_supabase()
        if df is not None:
            return df
    
    # Fallback to Excel
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE)
            # Backwards compatibility: rename Location to Website if present
            if 'Location' in df.columns and 'Website' not in df.columns:
                df = df.rename(columns={'Location': 'Website'})
            # Ensure all template columns exist
            template_columns = [
                'Meeting ID', 'Meeting Title', 'Organization', 'Client', 'Stakeholder Name',
                'Purpose', 'Agenda', 'Meeting Date', 'Start Time', 'Time Zone',
                'Meeting Type', 'Meeting Link', 'Website', 'Status', 'Priority',
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
                'Meeting Type', 'Meeting Link', 'Website', 'Status', 'Priority',
                'Attendees', 'Internal External Guests', 'Notes', 'Next Action',
                'Follow up Date', 'Reminder Sent', 'Calendar Sync', 'Calendar Event Title'
            ])
    else:
        return pd.DataFrame(columns=[
            'Meeting ID', 'Meeting Title', 'Organization', 'Client', 'Stakeholder Name',
            'Purpose', 'Agenda', 'Meeting Date', 'Start Time', 'Time Zone',
            'Meeting Type', 'Meeting Link', 'Website', 'Status', 'Priority',
            'Attendees', 'Internal External Guests', 'Notes', 'Next Action',
            'Follow up Date', 'Reminder Sent', 'Calendar Sync', 'Calendar Event Title'
        ])

def sync_excel_to_supabase(df=None):
    """Sync Excel data to Supabase - used on initial load"""
    if not get_use_supabase() or not init_db_pool():
        return False
    
    # Use provided dataframe or load from Excel
    if df is None:
        if not os.path.exists(EXCEL_FILE):
            return True  # No Excel file to sync
        
        try:
            df = pd.read_excel(EXCEL_FILE)
            if df.empty:
                return True  # Empty Excel file
        except Exception as e:
            st.error(f"Error reading Excel file: {e}")
            return False
    
    if df.empty:
        return True
    
    try:
        success_count = 0
        error_count = 0
        errors = []
        
        # Ensure Meeting ID column exists
        if 'Meeting ID' not in df.columns:
            # Generate Meeting IDs if missing
            df['Meeting ID'] = range(1, len(df) + 1)
        
        # Sync each row to Supabase
        for idx, row in df.iterrows():
            try:
                # Ensure Meeting ID exists
                if pd.isna(row.get('Meeting ID')):
                    # Generate ID if missing
                    row['Meeting ID'] = get_next_meeting_id(df) + idx
                
                if save_meeting_to_supabase(row):
                    success_count += 1
                else:
                    error_count += 1
                    errors.append(f"Row {idx + 1}: Failed to save")
            except Exception as e:
                error_count += 1
                errors.append(f"Row {idx + 1}: {str(e)}")
                continue
        
        if error_count > 0 and errors:
            for error in errors[:5]:  # Show first 5 errors
                st.warning(error)
            if len(errors) > 5:
                st.warning(f"... and {len(errors) - 5} more errors")
        
        return success_count > 0
    except Exception as e:
        st.error(f"Error syncing to Supabase: {e}")
        return False

def normalize_meeting_type(meeting_type):
    """Normalize meeting type to valid database values: 'Virtual' or 'In Person'"""
    if meeting_type is None or pd.isna(meeting_type):
        return None
    
    # Convert to string and strip whitespace
    meeting_type_str = str(meeting_type).strip()
    
    if not meeting_type_str or meeting_type_str.lower() in ['nan', 'none', 'null', '']:
        return None
    
    # Normalize common variations
    meeting_type_lower = meeting_type_str.lower()
    
    # Map to valid values: "Virtual" or "In Person"
    if meeting_type_lower in ['online', 'virtual', 'email', 'zoom', 'teams', 'webex', 'google meet']:
        return 'Virtual'
    elif meeting_type_lower in ['physical', 'in person', 'in-person', 'onsite', 'on-site', 'office']:
        return 'In Person'
    elif meeting_type_str in ['Virtual', 'In Person']:
        # Already valid, return as-is
        return meeting_type_str
    else:
        # Default to Virtual for unknown types
        return 'Virtual'

def normalize_status(status):
    """Normalize status to valid database values: 'Upcoming', 'Ongoing', 'Ended', 'Completed'"""
    if status is None or pd.isna(status):
        return 'Upcoming'  # Default status
    
    # Convert to string and strip whitespace
    status_str = str(status).strip()
    
    if not status_str or status_str.lower() in ['nan', 'none', 'null', '']:
        return 'Upcoming'  # Default status
    
    # Normalize common variations
    status_lower = status_str.lower()
    
    # Map to valid values: "Upcoming", "Ongoing", "Ended", "Completed"
    if status_lower in ['upcoming', 'scheduled', 'pending', 'planned', 'future']:
        return 'Upcoming'
    elif status_lower in ['ongoing', 'in progress', 'in-progress', 'active', 'running', 'started']:
        return 'Ongoing'
    elif status_lower == 'completed':
        return 'Completed'
    elif status_lower in ['ended', 'finished', 'done', 'closed', 'past']:
        return 'Ended'
    elif status_str in ['Upcoming', 'Ongoing', 'Ended', 'Completed']:
        # Already valid, return as-is
        return status_str
    else:
        # Default to Upcoming for unknown statuses
        return 'Upcoming'

def save_meeting_to_supabase(row):
    """Save a single meeting row to Supabase"""
    try:
        # Ensure Meeting ID exists and is valid
        meeting_id_value = row.get('Meeting ID') if 'Meeting ID' in row else None
        
        # Handle empty strings, NaN, None
        if meeting_id_value is None or pd.isna(meeting_id_value) or (isinstance(meeting_id_value, str) and meeting_id_value.strip() == ''):
            return False
        
        # Convert to int, handling string values
        try:
            meeting_id = int(float(str(meeting_id_value).strip()))
        except (ValueError, TypeError):
            return False
            
        if meeting_id is None or meeting_id <= 0:
            return False
            
        with get_db_connection() as conn:
            with conn.cursor() as cur:
                # Check if meeting exists
                cur.execute("SELECT id FROM meetings WHERE meeting_id = %s", (meeting_id,))
                exists = cur.fetchone()
                
                if exists:
                    # Update existing meeting
                    cur.execute("""
                        UPDATE meetings SET
                            meeting_title = %s,
                            organization = %s,
                            client = %s,
                            stakeholder_name = %s,
                            purpose = %s,
                            agenda = %s,
                            meeting_date = %s,
                            start_time = %s,
                            time_zone = %s,
                            meeting_type = %s,
                            meeting_link = %s,
                            location = %s,
                            status = %s,
                            priority = %s,
                            attendees = %s,
                            internal_external_guests = %s,
                            notes = %s,
                            next_action = %s,
                            follow_up_date = %s,
                            reminder_sent = %s,
                            calendar_sync = %s,
                            calendar_event_title = %s
                        WHERE meeting_id = %s
                    """, (
                        row.get('Meeting Title', ''),
                        row.get('Organization') if pd.notna(row.get('Organization')) else None,
                        row.get('Client') if pd.notna(row.get('Client')) else None,
                        row.get('Stakeholder Name', ''),
                        row.get('Purpose') if pd.notna(row.get('Purpose')) else None,
                        row.get('Agenda') if pd.notna(row.get('Agenda')) else None,
                        row.get('Meeting Date') if pd.notna(row.get('Meeting Date')) else None,
                        row.get('Start Time') if (pd.notna(row.get('Start Time')) and str(row.get('Start Time', '')).strip() != '') else None,
                        row.get('Time Zone') if pd.notna(row.get('Time Zone')) else 'UTC',
                        normalize_meeting_type(row.get('Meeting Type')),
                        row.get('Meeting Link') if pd.notna(row.get('Meeting Link')) else None,
                        row.get('Website') if pd.notna(row.get('Website')) else None,
                        normalize_status(row.get('Status')),
                        row.get('Priority') if pd.notna(row.get('Priority')) else 'Medium',
                        row.get('Attendees') if pd.notna(row.get('Attendees')) else None,
                        row.get('Internal External Guests', ''),
                        row.get('Notes') if pd.notna(row.get('Notes')) else None,
                        row.get('Next Action') if pd.notna(row.get('Next Action')) else None,
                        row.get('Follow up Date') if pd.notna(row.get('Follow up Date')) else None,
                        row.get('Reminder Sent') if pd.notna(row.get('Reminder Sent')) else 'No',
                        row.get('Calendar Sync') if pd.notna(row.get('Calendar Sync')) else 'No',
                        row.get('Calendar Event Title') if pd.notna(row.get('Calendar Event Title')) else None,
                        meeting_id
                    ))
                else:
                    # Insert new meeting
                    cur.execute("""
                        INSERT INTO meetings (
                            meeting_id, meeting_title, organization, client, stakeholder_name,
                            purpose, agenda, meeting_date, start_time, time_zone,
                            meeting_type, meeting_link, location, status, priority,
                            attendees, internal_external_guests, notes, next_action,
                            follow_up_date, reminder_sent, calendar_sync, calendar_event_title
                        ) VALUES (
                            %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                            %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
                        )
                    """, (
                        meeting_id,
                        row.get('Meeting Title', ''),
                        row.get('Organization') if pd.notna(row.get('Organization')) else None,
                        row.get('Client') if pd.notna(row.get('Client')) else None,
                        row.get('Stakeholder Name', ''),
                        row.get('Purpose') if pd.notna(row.get('Purpose')) else None,
                        row.get('Agenda') if pd.notna(row.get('Agenda')) else None,
                        row.get('Meeting Date') if pd.notna(row.get('Meeting Date')) else None,
                        row.get('Start Time') if (pd.notna(row.get('Start Time')) and str(row.get('Start Time', '')).strip() != '') else None,
                        row.get('Time Zone') if pd.notna(row.get('Time Zone')) else 'UTC',
                        normalize_meeting_type(row.get('Meeting Type')),
                        row.get('Meeting Link') if pd.notna(row.get('Meeting Link')) else None,
                        row.get('Website') if pd.notna(row.get('Website')) else None,
                        normalize_status(row.get('Status')),
                        row.get('Priority') if pd.notna(row.get('Priority')) else 'Medium',
                        row.get('Attendees') if pd.notna(row.get('Attendees')) else None,
                        row.get('Internal External Guests', ''),
                        row.get('Notes') if pd.notna(row.get('Notes')) else None,
                        row.get('Next Action') if pd.notna(row.get('Next Action')) else None,
                        row.get('Follow up Date') if pd.notna(row.get('Follow up Date')) else None,
                        row.get('Reminder Sent') if pd.notna(row.get('Reminder Sent')) else 'No',
                        row.get('Calendar Sync') if pd.notna(row.get('Calendar Sync')) else 'No',
                        row.get('Calendar Event Title') if pd.notna(row.get('Calendar Event Title')) else None
                    ))
        return True
    except Exception as e:
        st.error(f"Error saving to Supabase: {e}")
        return False

def delete_meeting_from_supabase(meeting_id):
    """Delete a meeting from Supabase"""
    try:
        with get_db_connection() as conn:
            with conn.cursor() as cur:
                # Disable the trigger first to prevent it from trying to insert into audit_log
                trigger_disabled = False
                try:
                    cur.execute("ALTER TABLE meetings DISABLE TRIGGER log_meeting_changes_trigger")
                    trigger_disabled = True
                except Exception:
                    # If we can't disable trigger, continue anyway
                    pass
                
                # Delete audit log entries first to avoid foreign key constraint error
                # The FK constraint prevents deleting meetings that are referenced in audit_log
                try:
                    cur.execute("DELETE FROM meetings_audit_log WHERE meeting_id = %s", (meeting_id,))
                except Exception:
                    # If audit log doesn't exist or has no entries, that's fine
                    pass
                
                # Now delete the meeting
                cur.execute("DELETE FROM meetings WHERE meeting_id = %s", (meeting_id,))
                
                # Check if any rows were actually deleted
                if cur.rowcount == 0:
                    # No rows deleted - meeting doesn't exist
                    return False
                
                # Re-enable the trigger if we disabled it
                if trigger_disabled:
                    try:
                        cur.execute("ALTER TABLE meetings ENABLE TRIGGER log_meeting_changes_trigger")
                    except Exception:
                        pass
                        
        return True
    except Exception as e:
        error_msg = str(e)
        # Log the actual error for debugging
        st.error(f"Error deleting from Supabase: {error_msg}")
        return False

def save_meetings(df):
    """Save meetings to Supabase (if available) and/or Excel file - Real-time sync"""
    if df.empty:
        # If dataframe is empty, clear Supabase and Excel
        if get_use_supabase() and init_db_pool():
            try:
                with get_db_connection() as conn:
                    with conn.cursor() as cur:
                        cur.execute("DELETE FROM meetings")
            except Exception as e:
                pass  # Ignore errors when clearing
        
        try:
            # Create empty Excel with columns
            template_columns = [
                'Meeting ID', 'Meeting Title', 'Organization', 'Client', 'Stakeholder Name',
                'Purpose', 'Agenda', 'Meeting Date', 'Start Time', 'Time Zone',
                'Meeting Type', 'Meeting Link', 'Website', 'Status', 'Priority',
                'Attendees', 'Internal External Guests', 'Notes', 'Next Action',
                'Follow up Date', 'Reminder Sent', 'Calendar Sync', 'Calendar Event Title'
            ]
            pd.DataFrame(columns=template_columns).to_excel(EXCEL_FILE, index=False)
        except:
            pass
        return True
    
    success = True
    supabase_success = True
    
    # Save to Supabase if enabled - sync ALL rows
    if get_use_supabase() and init_db_pool():
        try:
            # First, get all existing meeting IDs from Supabase
            with get_db_connection() as conn:
                with conn.cursor() as cur:
                    cur.execute("SELECT meeting_id FROM meetings")
                    existing_ids = {row[0] for row in cur.fetchall()}
            
            # Get current meeting IDs from dataframe
            current_ids = set()
            if 'Meeting ID' in df.columns:
                current_ids = set(pd.to_numeric(df['Meeting ID'], errors='coerce').dropna().astype(int))
            
            # Delete meetings from Supabase that are not in current dataframe
            ids_to_delete = existing_ids - current_ids
            for meeting_id in ids_to_delete:
                try:
                    if not delete_meeting_from_supabase(meeting_id):
                        # Log if delete failed
                        pass
                except Exception as delete_err:
                    # Log delete errors but continue with other operations
                    pass
            
            # Sync all rows to Supabase (insert or update)
            sync_errors = []
            for idx, row in df.iterrows():
                # Check if Meeting ID exists and is valid (not empty string, not NaN)
                meeting_id_val = row.get('Meeting ID') if 'Meeting ID' in row else None
                if meeting_id_val is not None and pd.notna(meeting_id_val):
                    # Check if it's not an empty string
                    if isinstance(meeting_id_val, str) and meeting_id_val.strip() == '':
                        continue  # Skip rows with empty string Meeting IDs
                    try:
                        # Try to convert to int to validate it's a valid ID
                        int(float(str(meeting_id_val).strip()))
                        if not save_meeting_to_supabase(row):
                            supabase_success = False
                            # Error message already displayed by save_meeting_to_supabase
                    except (ValueError, TypeError):
                        # Invalid Meeting ID, skip this row
                        continue
                else:
                    # Missing Meeting ID - skip but don't fail entire sync
                    continue
            
            # Only show general error if there were sync errors and no specific errors were shown
            if not supabase_success and not sync_errors:
                # Errors were already shown by save_meeting_to_supabase, just mark as failed
                pass
        except Exception as e:
            st.error(f"Error syncing to Supabase: {e}")
            supabase_success = False
    
    # Always save to Excel as backup
    try:
        df.to_excel(EXCEL_FILE, index=False)
    except Exception as e:
        if not get_use_supabase():
            st.error(f"Error saving meetings to Excel: {e}")
            return False
        else:
            # If Supabase works but Excel fails, still return success
            pass
    
    return supabase_success if get_use_supabase() and init_db_pool() else True

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

def get_next_meeting_id_from_supabase():
    """Get next meeting ID from Supabase"""
    try:
        with get_db_connection() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT get_next_meeting_id()")
                result = cur.fetchone()
                return result[0] if result else 1
    except Exception as e:
        # Fallback to database query if function doesn't exist
        try:
            with get_db_connection() as conn:
                with conn.cursor() as cur:
                    cur.execute("SELECT COALESCE(MAX(meeting_id), 0) + 1 FROM meetings")
                    result = cur.fetchone()
                    return result[0] if result else 1
        except:
            return 1

def get_next_meeting_id(df):
    """Get the next available meeting ID from Supabase or DataFrame"""
    # Try Supabase first if enabled
    if get_use_supabase() and init_db_pool():
        return get_next_meeting_id_from_supabase()
    
    # Fallback to DataFrame
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
    """Load data into session state with automatic sync"""
    if not st.session_state.data_loaded:
        # Try to load from Supabase first
        if get_use_supabase() and init_db_pool():
            supabase_df = load_meetings_from_supabase()
            
            # If Supabase has data, use it
            if supabase_df is not None and not supabase_df.empty:
                st.session_state.meetings_df = supabase_df
                # Sync to Excel as backup
                try:
                    st.session_state.meetings_df.to_excel(EXCEL_FILE, index=False)
                except:
                    pass
            # If Supabase is empty, use Excel data (but don't auto-sync)
            elif supabase_df is not None and supabase_df.empty:
                excel_df = None
                if os.path.exists(EXCEL_FILE):
                    try:
                        excel_df = pd.read_excel(EXCEL_FILE)
                        if 'Location' in excel_df.columns and 'Website' not in excel_df.columns:
                            excel_df = excel_df.rename(columns={'Location': 'Website'})
                        # Ensure all template columns exist
                        template_columns = [
                            'Meeting ID', 'Meeting Title', 'Organization', 'Client', 'Stakeholder Name',
                            'Purpose', 'Agenda', 'Meeting Date', 'Start Time', 'Time Zone',
                            'Meeting Type', 'Meeting Link', 'Website', 'Status', 'Priority',
                            'Attendees', 'Internal External Guests', 'Notes', 'Next Action',
                            'Follow up Date', 'Reminder Sent', 'Calendar Sync', 'Calendar Event Title'
                        ]
                        for col in template_columns:
                            if col not in excel_df.columns:
                                excel_df[col] = ''
                        
                        # Convert date and time columns
                        if 'Meeting Date' in excel_df.columns:
                            excel_df['Meeting Date'] = pd.to_datetime(excel_df['Meeting Date'], errors='coerce')
                        if 'Follow up Date' in excel_df.columns:
                            excel_df['Follow up Date'] = pd.to_datetime(excel_df['Follow up Date'], errors='coerce')
                    except:
                        excel_df = None
                
                if excel_df is not None and not excel_df.empty:
                    st.session_state.meetings_df = excel_df
                else:
                    st.session_state.meetings_df = pd.DataFrame(columns=[
                        'Meeting ID', 'Meeting Title', 'Organization', 'Client', 'Stakeholder Name',
                        'Purpose', 'Agenda', 'Meeting Date', 'Start Time', 'Time Zone',
                        'Meeting Type', 'Meeting Link', 'Website', 'Status', 'Priority',
                        'Attendees', 'Internal External Guests', 'Notes', 'Next Action',
                        'Follow up Date', 'Reminder Sent', 'Calendar Sync', 'Calendar Event Title'
                    ])
            else:
                # Supabase connection failed, fall back to Excel
                st.session_state.meetings_df = load_meetings()
        else:
            # Supabase not enabled or connection failed, use Excel
            st.session_state.meetings_df = load_meetings()
        
        st.session_state.data_loaded = True
    
    # Only recalculate status for empty/NaN statuses on initial load
    # Preserve all manually set statuses (they are saved to Excel/Supabase)
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
        # Convert Meeting Date to timezone-naive for comparison
        meeting_dates = pd.to_datetime(filtered_df['Meeting Date'], errors='coerce')
        # Remove timezone if present - convert to UTC first then remove timezone
        try:
            if meeting_dates.dt.tz is not None:
                meeting_dates = meeting_dates.dt.tz_convert('UTC').dt.tz_localize(None)
        except (AttributeError, TypeError):
            # If already naive or conversion fails, use as-is
            pass
        # Convert date_start to datetime (already naive)
        date_start_dt = pd.to_datetime(date_start)
        filtered_df = filtered_df[meeting_dates >= date_start_dt]
    
    if date_end and 'Meeting Date' in filtered_df.columns:
        # Convert Meeting Date to timezone-naive for comparison
        meeting_dates = pd.to_datetime(filtered_df['Meeting Date'], errors='coerce')
        # Remove timezone if present - convert to UTC first then remove timezone
        try:
            if meeting_dates.dt.tz is not None:
                meeting_dates = meeting_dates.dt.tz_convert('UTC').dt.tz_localize(None)
        except (AttributeError, TypeError):
            # If already naive or conversion fails, use as-is
            pass
        # Add end of day to date_end (already naive)
        date_end_datetime = pd.to_datetime(date_end) + timedelta(days=1) - timedelta(seconds=1)
        filtered_df = filtered_df[meeting_dates <= date_end_datetime]
    
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

# ============================================================================
# PODCAST MEETINGS FUNCTIONS
# ============================================================================

def load_podcast_meetings_from_supabase():
    """Load podcast meetings from Supabase database"""
    db_config = get_db_config()
    if not db_config.get('password'):
        return None
    
    try:
        with get_db_connection() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    SELECT 
                        podcast_id as "Podcast ID",
                        name as "Name",
                        designation as "Designation",
                        organization as "Organization",
                        linkedin_url as "LinkedIn URL",
                        host as "Host",
                        date as "Date",
                        day as "Day",
                        time as "Time",
                        status as "Status",
                        contacted_through as "Contacted Through",
                        comments as "Comments"
                    FROM podcast_meetings
                    ORDER BY podcast_id
                """)
                rows = cur.fetchall()
                
                if rows:
                    df = pd.DataFrame(rows)
                    if 'Date' in df.columns:
                        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
                    if 'Time' in df.columns:
                        df['Time'] = df['Time'].apply(
                            lambda x: str(x).split('.')[0] if pd.notna(x) and x != '' else ''
                        )
                    return df
                else:
                    return pd.DataFrame(columns=[
                        'Podcast ID', 'Name', 'Designation', 'Organization', 'LinkedIn URL',
                        'Host', 'Date', 'Day', 'Time', 'Status', 'Contacted Through', 'Comments'
                    ])
    except Exception as e:
        st.error(f"Error loading podcast meetings from Supabase: {e}")
        return None

def load_podcast_meetings():
    """Load podcast meetings from Supabase (if available) or Excel file"""
    if get_use_supabase() and init_db_pool():
        df = load_podcast_meetings_from_supabase()
        if df is not None:
            return df
    
    EXCEL_FILE_PODCAST = "Podcast_Meetings_Template.xlsx"
    if os.path.exists(EXCEL_FILE_PODCAST):
        try:
            df = pd.read_excel(EXCEL_FILE_PODCAST)
            template_columns = [
                'Podcast ID', 'Name', 'Designation', 'Organization', 'LinkedIn URL',
                'Host', 'Date', 'Day', 'Time', 'Status', 'Contacted Through', 'Comments'
            ]
            for col in template_columns:
                if col not in df.columns:
                    df[col] = ''
            if 'Date' in df.columns:
                df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            return df
        except Exception as e:
            return pd.DataFrame(columns=[
                'Podcast ID', 'Name', 'Designation', 'Organization', 'LinkedIn URL',
                'Host', 'Date', 'Day', 'Time', 'Status', 'Contacted Through', 'Comments'
            ])
    else:
        return pd.DataFrame(columns=[
            'Podcast ID', 'Name', 'Designation', 'Organization', 'LinkedIn URL',
            'Host', 'Date', 'Day', 'Time', 'Status', 'Contacted Through', 'Comments'
        ])

def normalize_podcast_status(status):
    """Normalize podcast status to valid database values"""
    if status is None or pd.isna(status):
        return 'Upcoming'
    status_str = str(status).strip()
    if not status_str or status_str.lower() in ['nan', 'none', 'null', '']:
        return 'Upcoming'
    status_lower = status_str.lower()
    if status_lower in ['upcoming', 'scheduled', 'pending', 'planned', 'future']:
        return 'Upcoming'
    elif status_lower in ['completed', 'done', 'finished']:
        return 'Completed'
    elif status_lower in ['cancelled', 'canceled']:
        return 'Cancelled'
    elif status_str in ['Upcoming', 'Completed', 'Cancelled']:
        return status_str
    else:
        return 'Upcoming'

def save_podcast_meeting_to_supabase(row):
    """Save a single podcast meeting row to Supabase"""
    try:
        podcast_id_value = row.get('Podcast ID') if 'Podcast ID' in row else None
        if podcast_id_value is None or pd.isna(podcast_id_value) or (isinstance(podcast_id_value, str) and podcast_id_value.strip() == ''):
            return False
        try:
            podcast_id = int(float(str(podcast_id_value).strip()))
        except (ValueError, TypeError):
            return False
        if podcast_id is None or podcast_id <= 0:
            return False
        with get_db_connection() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT id FROM podcast_meetings WHERE podcast_id = %s", (podcast_id,))
                exists = cur.fetchone()
                if exists:
                    cur.execute("""
                        UPDATE podcast_meetings SET
                            name = %s, designation = %s, organization = %s, linkedin_url = %s,
                            host = %s, date = %s, day = %s, time = %s, status = %s,
                            contacted_through = %s, comments = %s
                        WHERE podcast_id = %s
                    """, (
                        row.get('Name', ''),
                        row.get('Designation') if pd.notna(row.get('Designation')) else None,
                        row.get('Organization') if pd.notna(row.get('Organization')) else None,
                        row.get('LinkedIn URL') if pd.notna(row.get('LinkedIn URL')) else None,
                        row.get('Host') if pd.notna(row.get('Host')) else None,
                        row.get('Date') if pd.notna(row.get('Date')) else None,
                        row.get('Day') if pd.notna(row.get('Day')) else None,
                        row.get('Time') if (pd.notna(row.get('Time')) and str(row.get('Time', '')).strip() != '') else None,
                        normalize_podcast_status(row.get('Status')),
                        row.get('Contacted Through') if pd.notna(row.get('Contacted Through')) else None,
                        row.get('Comments') if pd.notna(row.get('Comments')) else None,
                        podcast_id
                    ))
                else:
                    cur.execute("""
                        INSERT INTO podcast_meetings (
                            podcast_id, name, designation, organization, linkedin_url,
                            host, date, day, time, status, contacted_through, comments
                        ) VALUES (
                            %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
                        )
                    """, (
                        podcast_id, row.get('Name', ''),
                        row.get('Designation') if pd.notna(row.get('Designation')) else None,
                        row.get('Organization') if pd.notna(row.get('Organization')) else None,
                        row.get('LinkedIn URL') if pd.notna(row.get('LinkedIn URL')) else None,
                        row.get('Host') if pd.notna(row.get('Host')) else None,
                        row.get('Date') if pd.notna(row.get('Date')) else None,
                        row.get('Day') if pd.notna(row.get('Day')) else None,
                        row.get('Time') if (pd.notna(row.get('Time')) and str(row.get('Time', '')).strip() != '') else None,
                        normalize_podcast_status(row.get('Status')),
                        row.get('Contacted Through') if pd.notna(row.get('Contacted Through')) else None,
                        row.get('Comments') if pd.notna(row.get('Comments')) else None
                    ))
        return True
    except Exception as e:
        st.error(f"Error saving podcast meeting to Supabase: {e}")
        return False

def delete_podcast_meeting_from_supabase(podcast_id):
    """Delete a podcast meeting from Supabase"""
    try:
        with get_db_connection() as conn:
            with conn.cursor() as cur:
                try:
                    cur.execute("DELETE FROM podcast_meetings_audit_log WHERE podcast_id = %s", (podcast_id,))
                except Exception:
                    pass
                cur.execute("DELETE FROM podcast_meetings WHERE podcast_id = %s", (podcast_id,))
                if cur.rowcount == 0:
                    return False
        return True
    except Exception as e:
        st.error(f"Error deleting podcast meeting from Supabase: {e}")
        return False

def save_podcast_meetings(df):
    """Save podcast meetings to Supabase (if available) and/or Excel file"""
    EXCEL_FILE_PODCAST = "Podcast_Meetings_Template.xlsx"
    if df.empty:
        if get_use_supabase() and init_db_pool():
            try:
                with get_db_connection() as conn:
                    with conn.cursor() as cur:
                        cur.execute("DELETE FROM podcast_meetings")
            except Exception:
                pass
        try:
            template_columns = [
                'Podcast ID', 'Name', 'Designation', 'Organization', 'LinkedIn URL',
                'Host', 'Date', 'Day', 'Time', 'Status', 'Contacted Through', 'Comments'
            ]
            pd.DataFrame(columns=template_columns).to_excel(EXCEL_FILE_PODCAST, index=False)
        except:
            pass
        return True
    success = True
    supabase_success = True
    if get_use_supabase() and init_db_pool():
        try:
            with get_db_connection() as conn:
                with conn.cursor() as cur:
                    cur.execute("SELECT podcast_id FROM podcast_meetings")
                    existing_ids = {row[0] for row in cur.fetchall()}
            for idx, row in df.iterrows():
                try:
                    if not save_podcast_meeting_to_supabase(row):
                        supabase_success = False
                except Exception as e:
                    supabase_success = False
                    st.warning(f"Error syncing podcast meeting row {idx + 1}: {str(e)}")
            if supabase_success:
                df_ids = set()
                for idx, row in df.iterrows():
                    podcast_id = row.get('Podcast ID')
                    if pd.notna(podcast_id):
                        try:
                            df_ids.add(int(float(str(podcast_id).strip())))
                        except:
                            pass
                ids_to_delete = existing_ids - df_ids
                for podcast_id in ids_to_delete:
                    try:
                        delete_podcast_meeting_from_supabase(podcast_id)
                    except Exception as e:
                        st.warning(f"Error deleting podcast meeting {podcast_id}: {str(e)}")
        except Exception as e:
            st.error(f"Error syncing podcast meetings to Supabase: {str(e)}")
            supabase_success = False
    try:
        df.to_excel(EXCEL_FILE_PODCAST, index=False)
    except Exception as e:
        st.error(f"Error saving podcast meetings to Excel: {e}")
        success = False
    return success and supabase_success

def get_next_podcast_id_from_supabase():
    """Get next podcast ID from Supabase"""
    try:
        with get_db_connection() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT COALESCE(MAX(podcast_id), 0) + 1 FROM podcast_meetings")
                result = cur.fetchone()
                return result[0] if result else 1
    except:
        return 1

def get_next_podcast_id(df):
    """Get the next available podcast ID from Supabase or DataFrame"""
    if get_use_supabase() and init_db_pool():
        return get_next_podcast_id_from_supabase()
    if df.empty or 'Podcast ID' not in df.columns:
        return 1
    podcast_ids = df['Podcast ID'].dropna()
    if podcast_ids.empty:
        return 1
    try:
        numeric_ids = pd.to_numeric(podcast_ids, errors='coerce').dropna()
        if numeric_ids.empty:
            return 1
        return int(numeric_ids.max()) + 1
    except:
        return 1

def load_podcast_data():
    """Load podcast meetings data into session state"""
    if not st.session_state.podcast_data_loaded:
        if get_use_supabase() and init_db_pool():
            supabase_df = load_podcast_meetings_from_supabase()
            if supabase_df is not None:
                st.session_state.podcast_meetings_df = supabase_df
            else:
                st.session_state.podcast_meetings_df = load_podcast_meetings()
        else:
            st.session_state.podcast_meetings_df = load_podcast_meetings()
        st.session_state.podcast_data_loaded = True

# Load podcast data on startup
load_podcast_data()

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
    
    /* Ensure emojis render with native colors */
    span[style*="Emoji"] {
        font-family: 'Segoe UI Emoji', 'Apple Color Emoji', 'Noto Color Emoji', 'EmojiOne Color', 'Android Emoji', emoji, sans-serif !important;
        color: initial !important;
        background: none !important;
        -webkit-background-clip: initial !important;
        -webkit-text-fill-color: initial !important;
        background-clip: initial !important;
        text-rendering: optimizeLegibility;
        -webkit-font-smoothing: antialiased;
    }
    
    /* Prevent gradient from affecting emojis */
    h1 span[style*="Emoji"] {
        background: none !important;
        -webkit-background-clip: initial !important;
        -webkit-text-fill-color: initial !important;
        background-clip: initial !important;
    }
    
    /* Table cell text wrapping and truncation - prevent stacking */
    .stColumn {
        overflow-wrap: break-word;
        word-break: break-word;
        line-height: 1.5 !important;
        padding: 0.5rem 0.25rem !important;
        vertical-align: top !important;
    }
    
    /* Ensure table headers are fully visible */
    small strong {
        white-space: nowrap;
        display: block;
        line-height: 1.4;
    }
    
    /* Prevent text from stacking in table cells */
    .stColumn small {
        display: block;
        line-height: 1.5;
        word-wrap: break-word;
        overflow-wrap: break-word;
        white-space: normal;
        max-width: 100%;
    }
    
    /* Ensure proper spacing between table rows */
    hr {
        margin: 0.5rem 0 !important;
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
    <h1 style="font-size: 1.75rem;
               font-weight: 700;
               margin: 0;
               padding: 0;
               display: flex;
               align-items: center;
               gap: 0.5rem;">
        <span style="font-size: 1.75rem; line-height: 1; font-family: 'Segoe UI Emoji', 'Apple Color Emoji', 'Noto Color Emoji', 'EmojiOne Color', 'Android Emoji', emoji, sans-serif; display: inline-block; text-rendering: optimizeLegibility; -webkit-font-smoothing: antialiased;">📅</span>
        <span style="background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
                     -webkit-background-clip: text;
                     -webkit-text-fill-color: transparent;
                     background-clip: text;">Meeting Dashboard</span>
    </h1>
    <p style="color: #64748b; font-size: 0.85rem; margin: 0.5rem 0 0 0;">AI Geo Navigators</p>
</div>
""", unsafe_allow_html=True)

# Page selection with enhanced styling
st.sidebar.markdown("<h3 style='font-size: 1rem; color: #475569; margin-bottom: 0.75rem; font-weight: 600;'>📑 Navigate to:</h3>", unsafe_allow_html=True)

# All pages in a single radio group for single selection
all_pages = [
    "📊 Smart Meeting Summary",
    "➕ Add New Meeting", 
    "✏️ Edit/Update Meeting", 
    "➕ Add New Podcast Meeting", 
    "✏️ Edit/Update Podcast Meeting", 
    "📊 Podcast Meetings Summary & Export"
]

# Calculate index based on current page (default to 0 = Add New Meeting)
current_index = 0
if st.session_state.current_page == "Add New Meeting":
    current_index = 0
elif st.session_state.current_page == "Edit/Update Meeting":
    current_index = 1
elif st.session_state.current_page == "Meetings Summary & Export":
    current_index = 2
elif st.session_state.current_page == "Add New Podcast Meeting":
    current_index = 3
elif st.session_state.current_page == "Edit/Update Podcast Meeting":
    current_index = 4
elif st.session_state.current_page == "Podcast Meetings Summary & Export":
    current_index = 5

page = st.sidebar.radio(
    "Navigate to:",
    all_pages,
    index=current_index,
    label_visibility="collapsed"
)

# Update current page based on selection
if "Add New Meeting" in page and "Podcast" not in page:
    st.session_state.current_page = "Add New Meeting"
elif "Edit/Update Meeting" in page and "Podcast" not in page:
    st.session_state.current_page = "Edit/Update Meeting"
elif "Meetings Summary" in page and "Podcast" not in page:
    st.session_state.current_page = "Meetings Summary & Export"
    # Refresh button below Meetings Summary & Export
    st.sidebar.markdown("---")
    if st.sidebar.button("🔄 Refresh", help="Reload data from database", use_container_width=True, key="refresh_btn"):
        # Reset data loaded flag to force reload
        st.session_state.data_loaded = False
        # Clear and reload data
        with st.spinner("Refreshing data..."):
            load_data()
            st.sidebar.success("✅ Data refreshed!")
            st.rerun()
elif "Add New Podcast" in page:
    st.session_state.current_page = "Add New Podcast Meeting"
elif "Edit/Update Podcast" in page:
    st.session_state.current_page = "Edit/Update Podcast Meeting"
elif "Podcast Meetings Summary" in page:
    st.session_state.current_page = "Podcast Meetings Summary & Export"
    # Refresh button below Podcast Meetings Summary & Export
    st.sidebar.markdown("---")
    if st.sidebar.button("🔄 Refresh", help="Reload podcast data from database", use_container_width=True, key="refresh_podcast_btn"):
        # Reset podcast data loaded flag to force reload
        st.session_state.podcast_data_loaded = False
        # Clear and reload podcast data
        with st.spinner("Refreshing podcast data..."):
            load_podcast_data()
            st.sidebar.success("✅ Podcast data refreshed!")
            st.rerun()

# Sidebar - Manual Sync
st.sidebar.markdown("---")
st.sidebar.markdown("<h3 style='font-size: 1rem; color: #475569; margin-bottom: 0.75rem; font-weight: 600;'>🗄️ Database Sync</h3>", unsafe_allow_html=True)

# Show connection status
db_config = get_db_config()
if db_config.get('password'):
    if get_use_supabase() and init_db_pool():
        st.sidebar.success("✅ Connected to Supabase")
        
        # Manual sync button only
        if st.sidebar.button("🔄 Sync", help="Sync current data to Supabase", use_container_width=True, key="sync_btn"):
            if not st.session_state.meetings_df.empty:
                with st.spinner("Syncing to database..."):
                    if save_meetings(st.session_state.meetings_df):
                        st.sidebar.success("✅ Sync completed!")
                    else:
                        st.sidebar.error("❌ Sync failed. Check errors above.")
            else:
                st.sidebar.info("No meetings to sync.")
    else:
        st.sidebar.warning("⚠️ Supabase connection failed")
        if st.session_state.supabase_error:
            with st.sidebar.expander("🔍 Error Details"):
                st.caption(st.session_state.supabase_error[:300])
else:
    if not db_config.get('password'):
        st.sidebar.warning("⚠️ Database password not configured")
        st.sidebar.caption("Set SUPABASE_DB_PASSWORD in secrets.toml")
    else:
        st.sidebar.warning("⚠️ Supabase connection failed")
        if st.session_state.supabase_error:
            with st.sidebar.expander("🔍 Error Details"):
                st.caption(st.session_state.supabase_error[:300])

# Enhanced Main Title with better visual hierarchy
page_titles = {
    "Add New Meeting": "➕ Add New Meeting",
    "Edit/Update Meeting": "✏️ Edit/Update Meeting",
    "Meetings Summary & Export": "📊 Meetings Summary & Export",
    "Add New Podcast Meeting": "➕ Add New Podcast Meeting",
    "Edit/Update Podcast Meeting": "✏️ Edit/Update Podcast Meeting",
    "Podcast Meetings Summary & Export": "📊 Podcast Meetings Summary & Export"
}
page_icon = "➕" if "Add New" in st.session_state.current_page else ("✏️" if "Edit" in st.session_state.current_page else "📊")

st.markdown(f"""
<div style="background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%);
            padding: 2rem;
            border-radius: 16px;
            margin-bottom: 2rem;
            border-left: 5px solid #2563eb;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);">
    <h1 style="font-size: 2.5rem;
               font-weight: 700;
               margin: 0;
               padding: 0;
               border: none;
               display: flex;
               align-items: center;
               gap: 0.5rem;">
        <span style="font-size: 2.5rem; line-height: 1; font-family: 'Segoe UI Emoji', 'Apple Color Emoji', 'Noto Color Emoji', 'EmojiOne Color', 'Android Emoji', emoji, sans-serif; display: inline-block; text-rendering: optimizeLegibility; -webkit-font-smoothing: antialiased;">{page_icon}</span>
        <span style="background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
                     -webkit-background-clip: text;
                     -webkit-text-fill-color: transparent;
                     background-clip: text;">{st.session_state.current_page}</span>
    </h1>
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
        
        # Website and Links
        st.markdown("### 📍 Website & Links")
        col_loc1, col_loc2 = st.columns(2)
        
        with col_loc1:
            meeting_link = st.text_input(
                "Meeting Link",
                value="",
                placeholder="Enter meeting link (for virtual meetings)",
                help="Enter the meeting link for virtual meetings"
            )
        with col_loc2:
            website = st.text_input(
                "Website",
                value="",
                placeholder="Enter website URL",
                help="Enter the website URL"
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
                    'Meeting Title': '',
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
                    'Website': website.strip(),
                    'Status': status,
                    'Priority': priority,
                    'Attendees': attendees.strip(),
                    'Internal External Guests': internal_external_guests.strip(),
                    'Notes': notes.strip(),
                    'Next Action': next_action.strip(),
                    'Follow up Date': follow_up_date if follow_up_date else pd.NaT,
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
# PAGE 2: Edit/Update Meeting
# ============================================================================
elif st.session_state.current_page == "Edit/Update Meeting":
    st.markdown("### ✏️ Edit/Update an Existing Meeting")
    st.markdown("Select a meeting from the list below to edit or update it.")
    
    if not st.session_state.meetings_df.empty:
        # Create selection list with index tracking for reliable lookup
        meeting_options = {}
        meeting_index_map = {}  # Map label to DataFrame index
        for idx, row in st.session_state.meetings_df.iterrows():
            org = str(row.get('Organization', 'N/A'))
            stakeholder = str(row.get('Stakeholder Name', 'N/A'))
            meeting_date = row.get('Meeting Date', '')
            if pd.notna(meeting_date):
                try:
                    date_str = pd.to_datetime(meeting_date).strftime('%Y-%m-%d')
                except:
                    date_str = str(meeting_date)
            else:
                date_str = 'N/A'
            label = f"{org} - {stakeholder} - {date_str}"
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
                if selected_meeting.get('Website'):
                    st.write(f"**Website:** {selected_meeting.get('Website', 'N/A')}")
                if selected_meeting.get('Attendees'):
                    st.write(f"**Attendees:** {selected_meeting.get('Attendees', 'N/A')}")
        
        # Edit form
        st.markdown("---")
        with st.form("edit_meeting_form"):
            # Basic Information
            st.markdown("### 📝 Basic Information")
            col1, col2 = st.columns(2)
            
            with col1:
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
            
            # Website and Links
            st.markdown("### 📍 Website & Links")
            col_loc1, col_loc2 = st.columns(2)
            
            with col_loc1:
                edit_meeting_link = st.text_input("Meeting Link", value=str(selected_meeting.get('Meeting Link', '')))
            with col_loc2:
                edit_website = st.text_input("Website", value=str(selected_meeting.get('Website', '')), placeholder="Enter website URL")
            
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
            
            update_submitted = st.form_submit_button("💾 Update Meeting", type="primary", use_container_width=True)
            
            if update_submitted:
                # Auto-fill missing required fields with empty strings (null) instead of showing errors
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
                    
                    # Update meeting - automatically set missing required fields to empty string (null)
                    st.session_state.meetings_df.at[idx, 'Organization'] = edit_organization.strip() if edit_organization else ''
                    st.session_state.meetings_df.at[idx, 'Client'] = edit_client.strip() if edit_client else ''
                    st.session_state.meetings_df.at[idx, 'Stakeholder Name'] = edit_stakeholder_name.strip() if edit_stakeholder_name.strip() else ''
                    st.session_state.meetings_df.at[idx, 'Purpose'] = edit_purpose.strip() if edit_purpose else ''
                    st.session_state.meetings_df.at[idx, 'Agenda'] = edit_agenda.strip() if edit_agenda else ''
                    st.session_state.meetings_df.at[idx, 'Meeting Date'] = edit_meeting_date if edit_meeting_date else pd.NaT
                    st.session_state.meetings_df.at[idx, 'Start Time'] = edit_start_time.strftime('%H:%M:%S') if edit_start_time else ''
                    st.session_state.meetings_df.at[idx, 'Time Zone'] = edit_time_zone.strip() if edit_time_zone else ''
                    st.session_state.meetings_df.at[idx, 'Meeting Type'] = edit_meeting_type if edit_meeting_type else ''
                    st.session_state.meetings_df.at[idx, 'Meeting Link'] = edit_meeting_link.strip() if edit_meeting_link else ''
                    st.session_state.meetings_df.at[idx, 'Website'] = edit_website.strip() if edit_website else ''
                    st.session_state.meetings_df.at[idx, 'Status'] = edit_status if edit_status else ''
                    st.session_state.meetings_df.at[idx, 'Priority'] = edit_priority if edit_priority else ''
                    st.session_state.meetings_df.at[idx, 'Attendees'] = edit_attendees.strip() if edit_attendees.strip() else ''
                    st.session_state.meetings_df.at[idx, 'Internal External Guests'] = edit_internal_external_guests.strip() if edit_internal_external_guests.strip() else ''
                    st.session_state.meetings_df.at[idx, 'Notes'] = edit_notes.strip() if edit_notes else ''
                    st.session_state.meetings_df.at[idx, 'Next Action'] = edit_next_action.strip() if edit_next_action else ''
                    st.session_state.meetings_df.at[idx, 'Follow up Date'] = edit_follow_up_date if edit_follow_up_date else pd.NaT
                    st.session_state.meetings_df.at[idx, 'Reminder Sent'] = edit_reminder_sent if edit_reminder_sent else ''
                    st.session_state.meetings_df.at[idx, 'Calendar Sync'] = edit_calendar_sync if edit_calendar_sync else ''
                    st.session_state.meetings_df.at[idx, 'Calendar Event Title'] = edit_calendar_event_title.strip() if edit_calendar_event_title else ''
                    
                    # Save to Excel
                    if save_meetings(st.session_state.meetings_df):
                        # Store the manually set status to preserve it after reload
                        if 'manually_set_statuses' not in st.session_state:
                            st.session_state.manually_set_statuses = {}
                        st.session_state.manually_set_statuses[selected_meeting_id] = edit_status
                        
                        st.success("Meeting Updated Successfully")
                        time.sleep(1.5)
                        st.rerun()
                    else:
                        st.error("Failed to update meeting")
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
    
    # Import/Upload section - moved to top for better visibility
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
        st.write("Upload an Excel file to import or update meeting records. Download the template below to ensure correct format.")
    with col_template2:
        # Create template dataframe with all template columns
        template_columns = [
            'Meeting ID', 'Meeting Title', 'Organization', 'Client', 'Stakeholder Name',
            'Purpose', 'Agenda', 'Meeting Date', 'Start Time', 'Time Zone',
            'Meeting Type', 'Meeting Link', 'Website', 'Status', 'Priority',
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
            'Website': '',
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
            help="Download a template Excel file with the correct column format",
            use_container_width=True
        )
    
    uploaded_file = st.file_uploader(
        "Choose an Excel file to import",
        type=['xlsx', 'xls'],
        help="Upload an Excel file with meeting data. Required columns: Organization. All other columns are optional."
    )
    
    if uploaded_file is not None:
        try:
            # Read the uploaded file - try with header=0 first
            import_df = pd.read_excel(uploaded_file, header=0)
            
            # If we got "Unnamed" columns, try to find the header row
            if any('Unnamed' in str(col) for col in import_df.columns) or (len(import_df.columns) > 0 and str(import_df.columns[0]).startswith('Unnamed')):
                # Try reading without header first to see the data
                temp_df = pd.read_excel(uploaded_file, header=None)
                # Look for a row that contains "Organization" or "Meeting Title" (case-insensitive)
                header_row = None
                for idx in range(min(5, len(temp_df))):  # Check first 5 rows
                    row_values = [str(val).strip().lower() for val in temp_df.iloc[idx].values if pd.notna(val)]
                    if any('organization' in str(val).lower() or 'meeting title' in str(val).lower() or 'title' in str(val).lower() for val in row_values):
                        header_row = idx
                        break
                
                if header_row is not None:
                    # Re-read with the correct header row
                    import_df = pd.read_excel(uploaded_file, header=header_row)
                else:
                    # If no header row found, use first row as header
                    import_df = pd.read_excel(uploaded_file, header=0)
                    # If still unnamed, try header=None and use first row
                    if any('Unnamed' in str(col) for col in import_df.columns):
                        temp_df = pd.read_excel(uploaded_file, header=None)
                        if len(temp_df) > 0:
                            # Use first row as column names
                            import_df.columns = [str(val).strip() if pd.notna(val) else f'Unnamed_{i}' for i, val in enumerate(temp_df.iloc[0].values)]
                            import_df = temp_df.iloc[1:].copy()
                            import_df.columns = [str(val).strip() if pd.notna(val) else f'Unnamed_{i}' for i, val in enumerate(import_df.columns)]
            
            # Normalize column names (strip whitespace and make case-insensitive mapping)
            column_mapping = {}
            for col in import_df.columns:
                normalized = str(col).strip()
                if normalized and not normalized.startswith('Unnamed'):
                    column_mapping[normalized.lower()] = col
            
            # Check only critical required columns (case-insensitive)
            critical_required_columns = ['organization']
            missing_critical = []
            found_columns = {}
            
            for req_col in critical_required_columns:
                if req_col.lower() in column_mapping:
                    found_columns[req_col] = column_mapping[req_col.lower()]
                else:
                    missing_critical.append(req_col)
            
            if missing_critical:
                st.error(f"❌ Missing critical required column: {', '.join([c.title() for c in missing_critical])}")
                st.info("At minimum, 'Organization' column is required. Other missing columns will be filled with empty values.")
                st.info(f"📋 Found columns in your file: {', '.join([str(c) for c in import_df.columns[:10]])}")
            else:
                # Rename columns to standard format (case-insensitive)
                rename_dict = {}
                for std_col in ['Meeting ID', 'Meeting Title', 'Organization', 'Client', 'Stakeholder Name',
                               'Purpose', 'Agenda', 'Meeting Date', 'Start Time', 'Time Zone',
                               'Meeting Type', 'Meeting Link', 'Website', 'Status', 'Priority',
                               'Attendees', 'Internal External Guests', 'Notes', 'Next Action',
                               'Follow up Date', 'Reminder Sent', 'Calendar Sync', 'Calendar Event Title']:
                    std_col_lower = std_col.lower()
                    if std_col_lower in column_mapping:
                        original_col = column_mapping[std_col_lower]
                        if original_col != std_col:
                            rename_dict[original_col] = std_col
                # Backwards compatibility: map Location to Website
                if 'location' in column_mapping and 'website' not in column_mapping:
                    rename_dict[column_mapping['location']] = 'Website'
                
                if rename_dict:
                    import_df = import_df.rename(columns=rename_dict)
                # Add missing columns with empty values
                template_columns = [
                    'Meeting ID', 'Meeting Title', 'Organization', 'Client', 'Stakeholder Name',
                    'Purpose', 'Agenda', 'Meeting Date', 'Start Time', 'Time Zone',
                    'Meeting Type', 'Meeting Link', 'Website', 'Status', 'Priority',
                    'Attendees', 'Internal External Guests', 'Notes', 'Next Action',
                    'Follow up Date', 'Reminder Sent', 'Calendar Sync', 'Calendar Event Title'
                ]
                
                missing_columns = [col for col in template_columns if col not in import_df.columns]
                if missing_columns:
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
                
                # Import mode is now fixed to "Update & Add New"
                import_mode = "Update & Add New"
                overwrite_status = False
                
                # Normalize empty values to null
                for idx, row in import_df.iterrows():
                    organization = row.get('Organization', '')
                    has_organization = (
                        pd.notna(organization) and 
                        str(organization).strip() != '' and 
                        str(organization).strip().lower() not in ['nan', 'none', 'null', '']
                    )
                    
                    if has_organization:
                        meeting_date = row.get('Meeting Date', '')
                        is_date_empty = True
                        if pd.notna(meeting_date):
                            if isinstance(meeting_date, (pd.Timestamp, datetime)):
                                if isinstance(meeting_date, pd.Timestamp) and pd.isna(meeting_date):
                                    is_date_empty = True
                                else:
                                    is_date_empty = False
                            elif isinstance(meeting_date, str):
                                date_str = meeting_date.strip()
                                if date_str and date_str.lower() not in ['nan', 'none', 'null', '', 'nat']:
                                    try:
                                        parsed_date = pd.to_datetime(date_str)
                                        if pd.notna(parsed_date):
                                            is_date_empty = False
                                    except:
                                        pass
                            else:
                                date_str = str(meeting_date).strip()
                                if date_str and date_str.lower() not in ['nan', 'none', 'null', '', 'nat']:
                                    try:
                                        parsed_date = pd.to_datetime(date_str)
                                        if pd.notna(parsed_date):
                                            is_date_empty = False
                                    except:
                                        pass
                        
                        if is_date_empty:
                            import_df.at[idx, 'Meeting Date'] = pd.NaT
                        
                        start_time = row.get('Start Time', '')
                        is_time_empty = True
                        if pd.notna(start_time):
                            time_str = str(start_time).strip()
                            if time_str and time_str.lower() not in ['nan', 'none', 'null', '']:
                                is_time_empty = False
                        
                        if is_time_empty:
                            import_df.at[idx, 'Start Time'] = ''
                
                # Proceed with import
                if st.button("✅ Import Data", type="primary", use_container_width=True, key="import_btn_top"):
                    try:
                        # Ensure all template columns exist
                        template_columns = [
                            'Meeting ID', 'Meeting Title', 'Organization', 'Client', 'Stakeholder Name',
                            'Purpose', 'Agenda', 'Meeting Date', 'Start Time', 'Time Zone',
                            'Meeting Type', 'Meeting Link', 'Website', 'Status', 'Priority',
                            'Attendees', 'Internal External Guests', 'Notes', 'Next Action',
                            'Follow up Date', 'Reminder Sent', 'Calendar Sync', 'Calendar Event Title'
                        ]
                        for col in template_columns:
                            if col not in import_df.columns:
                                import_df[col] = ''
                        
                        # Fill NaN values with empty strings for text columns
                        text_columns = ['Meeting ID', 'Meeting Title', 'Organization', 'Client', 'Stakeholder Name', 'Purpose', 
                                      'Agenda', 'Start Time', 'Time Zone', 'Meeting Type', 'Meeting Link', 
                                      'Website', 'Status', 'Priority', 'Attendees', 'Internal External Guests', 'Notes', 
                                      'Next Action', 'Reminder Sent', 'Calendar Sync', 'Calendar Event Title']
                        for col in text_columns:
                            if col in import_df.columns:
                                import_df[col] = import_df[col].fillna('').astype(str)
                        
                        # Handle Status
                        if 'Status' not in import_df.columns:
                            import_df['Status'] = ''
                        
                        # Calculate status only for rows with Meeting Date and Start Time
                        for idx, row in import_df.iterrows():
                            if pd.isna(row.get('Status')) or str(row.get('Status', '')).strip() == '':
                                has_date = pd.notna(row.get('Meeting Date')) and str(row.get('Meeting Date', '')).strip() != ''
                                has_time = pd.notna(row.get('Start Time')) and str(row.get('Start Time', '')).strip() != ''
                                if has_date and has_time:
                                    import_df.at[idx, 'Status'] = calculate_status(row)
                                else:
                                    import_df.at[idx, 'Status'] = ''
                        
                        # Get current dataframe
                        current_df = st.session_state.meetings_df.copy()
                        
                        # Clean up Meeting ID column
                        if 'Meeting ID' in import_df.columns:
                            import_df['Meeting ID'] = import_df['Meeting ID'].replace('', pd.NA)
                            import_df['Meeting ID'] = import_df['Meeting ID'].replace(' ', pd.NA)
                        
                        if current_df.empty:
                            if 'Meeting ID' not in import_df.columns or import_df['Meeting ID'].isna().all():
                                import_df['Meeting ID'] = range(1, len(import_df) + 1)
                            else:
                                missing_mask = import_df['Meeting ID'].isna()
                                if missing_mask.any():
                                    max_id = pd.to_numeric(import_df['Meeting ID'], errors='coerce').max()
                                    if pd.isna(max_id):
                                        max_id = 0
                                    next_id = int(max_id) + 1
                                    import_df.loc[missing_mask, 'Meeting ID'] = range(next_id, next_id + missing_mask.sum())
                            st.session_state.meetings_df = import_df.copy()
                            added_count = len(import_df)
                            updated_count = 0
                        else:
                            if 'Meeting ID' not in import_df.columns or import_df['Meeting ID'].isna().all():
                                if 'Meeting ID' in current_df.columns:
                                    max_id = pd.to_numeric(current_df['Meeting ID'], errors='coerce').max()
                                    if pd.isna(max_id):
                                        max_id = 0
                                else:
                                    max_id = 0
                                import_df['Meeting ID'] = range(int(max_id) + 1, int(max_id) + 1 + len(import_df))
                            else:
                                missing_mask = import_df['Meeting ID'].isna()
                                if missing_mask.any():
                                    max_current = pd.to_numeric(current_df['Meeting ID'], errors='coerce').max() if 'Meeting ID' in current_df.columns else 0
                                    max_import = pd.to_numeric(import_df['Meeting ID'], errors='coerce').max()
                                    max_id = max(max_current if not pd.isna(max_current) else 0, max_import if not pd.isna(max_import) else 0)
                                    next_id = int(max_id) + 1
                                    import_df.loc[missing_mask, 'Meeting ID'] = range(next_id, next_id + missing_mask.sum())
                            
                            import_df['Meeting ID'] = pd.to_numeric(import_df['Meeting ID'], errors='coerce')
                            if 'Meeting ID' in current_df.columns:
                                current_df['Meeting ID'] = pd.to_numeric(current_df['Meeting ID'], errors='coerce')
                            
                            added_count = 0
                            updated_count = 0
                            
                            if 'Meeting ID' in current_df.columns and 'Meeting ID' in import_df.columns:
                                existing_ids = set(pd.to_numeric(current_df['Meeting ID'], errors='coerce').dropna().astype(int))
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
                        
                        # Save to database and Excel
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
        
        # Select columns to display (show most important ones; Meeting Title removed)
        display_columns = ['Organization', 'Meeting Date', 'Start Time', 'Status', 
                          'Meeting Type', 'Client', 'Stakeholder Name', 
                          'Priority', 'Attendees', 'Website', 'Meeting Link']
        available_columns = [col for col in display_columns if col in display_df.columns]
        
        # Define column widths based on content importance and typical size
        # Increased widths to prevent text stacking
        column_width_map = {
            'Organization': 1.8,      # Wider - important and can be long
            'Meeting Date': 1.0,       # Medium - dates are standard width
            'Start Time': 0.8,         # Narrow - time format is short
            'Status': 0.9,             # Medium - status values are short
            'Meeting Type': 1.0,       # Medium - "Virtual" or "In Person"
            'Organization': 1.8,       # Wider - can vary in length significantly
            'Client': 1.5,             # Medium-wide - can vary in length
            'Stakeholder Name': 2.0,   # Wider - names can be long
            'Priority': 0.7,           # Narrow - "High", "Medium", "Low"
            'Attendees': 1.2,          # Medium - can be empty or long
            'Website': 1.2,           # Medium - can be empty or long
            'Meeting Link': 1.5        # Medium-wide - URLs can be long
        }
        
        # Create width list based on available columns
        col_widths = [0.3]  # Select checkbox column - increased to prevent overlap
        for col in available_columns:
            col_widths.append(column_width_map.get(col, 1.0))  # Default to 1.0 if not in map
        col_widths.extend([0.35, 0.35])  # Edit and Delete button columns
        
        # Multi-select and delete section
        col_select1, col_select2 = st.columns([3, 1])
        with col_select1:
            select_all = st.checkbox("Select All", key="select_all_meetings", help="Select/deselect all meetings")
            if select_all:
                # Select all meeting IDs
                if 'Meeting ID' in filtered_meetings.columns:
                    st.session_state.selected_meetings = set(
                        int(meeting_id) for meeting_id in filtered_meetings['Meeting ID'].dropna() 
                        if pd.notna(meeting_id)
                    )
            else:
                # Clear selection if "Select All" is unchecked
                if 'Meeting ID' in filtered_meetings.columns:
                    filtered_ids = set(
                        int(meeting_id) for meeting_id in filtered_meetings['Meeting ID'].dropna() 
                        if pd.notna(meeting_id)
                    )
                    # Only clear if all filtered meetings were selected
                    if st.session_state.selected_meetings == filtered_ids:
                        st.session_state.selected_meetings = set()
        
        with col_select2:
            if st.session_state.selected_meetings:
                if st.button("🗑️ Delete Selected", type="primary", use_container_width=True, 
                           help=f"Delete {len(st.session_state.selected_meetings)} selected meeting(s)"):
                    # Delete selected meetings
                    deleted_count = 0
                    failed_count = 0
                    ids_to_delete = list(st.session_state.selected_meetings.copy())
                    
                    for meeting_id_int in ids_to_delete:
                        try:
                            delete_success = True
                            
                            # Delete from Supabase first
                            if get_use_supabase() and init_db_pool():
                                delete_success = delete_meeting_from_supabase(meeting_id_int)
                            
                            # Only delete from dataframe if database delete succeeded
                            if delete_success:
                                # Delete from dataframe
                                if 'Meeting ID' in st.session_state.meetings_df.columns:
                                    st.session_state.meetings_df = st.session_state.meetings_df[
                                        st.session_state.meetings_df['Meeting ID'] != meeting_id_int
                                    ]
                                # Remove from selected set
                                st.session_state.selected_meetings.discard(meeting_id_int)
                                deleted_count += 1
                            else:
                                failed_count += 1
                        except Exception as e:
                            failed_count += 1
                            st.session_state.selected_meetings.discard(meeting_id_int)
                    
                    # Save updated dataframe
                    if deleted_count > 0:
                        save_meetings(st.session_state.meetings_df)
                        if failed_count == 0:
                            st.success(f"{deleted_count} Meeting(s) Deleted Successfully")
                        else:
                            st.warning(f"{deleted_count} Meeting(s) Deleted Successfully, {failed_count} Failed")
                        st.rerun()
                    elif failed_count > 0:
                        st.error(f"Failed to delete {failed_count} meeting(s)")
        
        # Scrollable table container: marker + CSS so the table body scrolls
        st.markdown("<div id='meetings-table-scroll-marker'></div>", unsafe_allow_html=True)
        st.markdown("""
        <style>
        div[data-testid="stVerticalBlock"]:has(#meetings-table-scroll-marker) {
            max-height: 65vh !important;
            overflow-y: auto !important;
            overflow-x: auto !important;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 0.5rem;
            margin-bottom: 0.5rem;
        }
        </style>
        """, unsafe_allow_html=True)
        # Create a custom table with Edit and Delete buttons (optimized column widths)
        header_cols = st.columns(col_widths)
        
        # Display header with smaller font - ensure full visibility and no overlap
        header_cols[0].markdown("<div style='line-height: 1.4; white-space: nowrap;'><small><strong>Select</strong></small></div>", unsafe_allow_html=True)
        for idx, col_name in enumerate(available_columns):
            # Show full column name with proper wrapping - ensure no overlap
            header_cols[idx + 1].markdown(
                f"<div style='line-height: 1.4; word-wrap: break-word; white-space: normal; overflow: hidden;'><small><strong title='{col_name}'>{col_name}</strong></small></div>", 
                unsafe_allow_html=True
            )
        header_cols[-2].markdown("<div style='line-height: 1.4; white-space: nowrap;'><small><strong>Edit</strong></small></div>", unsafe_allow_html=True)
        header_cols[-1].markdown("<div style='line-height: 1.4; white-space: nowrap;'><small><strong>Delete</strong></small></div>", unsafe_allow_html=True)
        
        st.markdown("<hr style='margin: 0.5rem 0;'>", unsafe_allow_html=True)
        
        # Display each row with buttons
        for pos, (idx, row) in enumerate(display_df.iterrows()):
            row_cols = st.columns(col_widths)
            
            # Get meeting ID for this row (use position since display_df is a copy of filtered_meetings)
            meeting_id = None
            if 'Meeting ID' in filtered_meetings.columns and pos < len(filtered_meetings):
                meeting_id = filtered_meetings.iloc[pos].get('Meeting ID')
            
            # Checkbox for selection
            if meeting_id is not None and pd.notna(meeting_id):
                meeting_id_int = int(meeting_id)
                is_selected = meeting_id_int in st.session_state.selected_meetings
                checkbox_key = f"select_checkbox_{meeting_id_int}_{idx}"
                
                # Use checkbox and update selected_meetings based on its state
                checkbox_state = row_cols[0].checkbox("", value=is_selected, key=checkbox_key, 
                                                     label_visibility="collapsed", help="Select meeting")
                
                # Update selected_meetings set based on checkbox state
                if checkbox_state:
                    st.session_state.selected_meetings.add(meeting_id_int)
                else:
                    st.session_state.selected_meetings.discard(meeting_id_int)
            
            # Display row data with smaller font - prevent text stacking
            for col_idx, col_name in enumerate(available_columns):
                value = row.get(col_name, '')
                if pd.isna(value):
                    value = ''
                # Truncate very long values to prevent excessive wrapping
                value_str = str(value) if value else ''
                if len(value_str) > 50 and col_name in ['Organization', 'Stakeholder Name']:
                    value_str = value_str[:47] + "..."
                row_cols[col_idx + 1].markdown(
                    f"<div style='line-height: 1.5; word-wrap: break-word; overflow-wrap: break-word; white-space: normal;'><small>{value_str}</small></div>", 
                    unsafe_allow_html=True
                )
            
            # Edit button - small button
            if meeting_id is not None and pd.notna(meeting_id):
                if row_cols[-2].button("✏️", key=f"edit_{meeting_id}_{idx}", use_container_width=False, help="Edit meeting"):
                    st.session_state.current_page = "Edit/Update Meeting"
                    st.session_state.edit_meeting_id = int(meeting_id)
                    st.rerun()
            
            # Delete button - small button
            if meeting_id is not None and pd.notna(meeting_id):
                if row_cols[-1].button("🗑️", key=f"delete_{meeting_id}_{idx}", use_container_width=False, type="secondary", help="Delete meeting"):
                    try:
                        meeting_id_int = int(meeting_id)
                        delete_success = True
                        
                        # Delete from Supabase first
                        if get_use_supabase() and init_db_pool():
                            delete_success = delete_meeting_from_supabase(meeting_id_int)
                        
                        # Only delete from dataframe if database delete succeeded
                        if delete_success:
                            # Delete from dataframe
                            if 'Meeting ID' in st.session_state.meetings_df.columns:
                                st.session_state.meetings_df = st.session_state.meetings_df[
                                    st.session_state.meetings_df['Meeting ID'] != meeting_id_int
                                ]
                            
                            # Save updated dataframe
                            save_meetings(st.session_state.meetings_df)
                            st.success("Meeting Deleted Successfully")
                            st.rerun()
                        else:
                            st.error(f"❌ Failed to delete meeting {meeting_id_int} from database. Please try again.")
                    except Exception as e:
                        st.error(f"Error deleting meeting: {e}")
            
            st.markdown("<hr style='margin: 0.3rem 0;'>", unsafe_allow_html=True)
        
        st.caption(f"Showing {len(display_df)} meeting(s)")
    else:
        st.info("📭 No meetings found matching your filters.")
    
    # Export section (Import section moved to top after filters)
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

# ============================================================================
# PODCAST MEETING PAGES
# ============================================================================

# PAGE 4: Add New Podcast Meeting
elif st.session_state.current_page == "Add New Podcast Meeting":
    with st.form("add_podcast_meeting_form", clear_on_submit=True):
        st.markdown("### 📝 Basic Information")
        col1, col2 = st.columns(2)
        
        with col1:
            name = st.text_input("Name *", value="", placeholder="Enter guest/speaker name", help="Enter the name of the guest/speaker (Required)")
            designation = st.text_input("Designation", value="", placeholder="Enter job title or designation", help="Enter the job title or designation of the guest")
            organization = st.text_input("Organization", value="", placeholder="Enter organization name", help="Enter the organization the guest belongs to")
            linkedin_url = st.text_input("LinkedIn URL", value="", placeholder="Enter LinkedIn profile URL", help="Enter the LinkedIn profile URL of the guest")
        
        with col2:
            host = st.text_input("Host", value="", placeholder="Enter podcast host name", help="Enter the name of the podcast host")
            status = st.selectbox("Status", ["Upcoming", "Completed", "Cancelled"], index=0)
            contacted_through = st.text_input("Contacted Through (Platform)", value="", placeholder="e.g., LinkedIn, Email, etc.", help="Enter the platform used to contact the guest")
        
        st.markdown("### 🕐 Date & Time")
        col_date1, col_date2, col_date3 = st.columns(3)
        with col_date1:
            date = st.date_input("Date", value=None)
        with col_date2:
            day = st.text_input("Day", value="", placeholder="e.g., Monday, Tuesday")
        with col_date3:
            time_val = st.time_input("Time", value=None)
        
        st.markdown("### 💬 Comments")
        comments = st.text_area("Comments", value="", height=100, placeholder="Enter additional notes or comments...", help="Enter any additional notes or comments about the podcast meeting")
        
        submitted = st.form_submit_button("💾 Save Podcast Meeting", type="primary", use_container_width=True)
        
        if submitted:
            if not name.strip():
                st.error("Name is required")
            else:
                new_podcast_meeting = pd.DataFrame([{
                    'Podcast ID': get_next_podcast_id(st.session_state.podcast_meetings_df),
                    'Name': name.strip(),
                    'Designation': designation.strip() if designation else '',
                    'Organization': organization.strip() if organization else '',
                    'LinkedIn URL': linkedin_url.strip() if linkedin_url else '',
                    'Host': host.strip() if host else '',
                    'Date': date if date else pd.NaT,
                    'Day': day.strip() if day else '',
                    'Time': time_val.strftime('%H:%M:%S') if time_val else '',
                    'Status': status,
                    'Contacted Through': contacted_through.strip() if contacted_through else '',
                    'Comments': comments.strip() if comments else ''
                }])
                
                if st.session_state.podcast_meetings_df.empty:
                    st.session_state.podcast_meetings_df = new_podcast_meeting
                else:
                    st.session_state.podcast_meetings_df = pd.concat([st.session_state.podcast_meetings_df, new_podcast_meeting], ignore_index=True)
                
                if save_podcast_meetings(st.session_state.podcast_meetings_df):
                    st.success("✅ Podcast Meeting Saved Successfully!")
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error("❌ Failed to save podcast meeting. Please try again.")

# PAGE 5: Edit/Update Podcast Meeting
elif st.session_state.current_page == "Edit/Update Podcast Meeting":
    if st.session_state.podcast_meetings_df.empty:
        st.info("📭 No podcast meetings found. Please add a podcast meeting first.")
    else:
        podcast_meetings_list = []
        meeting_index_map = {}
        
        for idx, row in st.session_state.podcast_meetings_df.iterrows():
            podcast_id = row.get('Podcast ID', 'N/A')
            name = row.get('Name', 'N/A')
            date = row.get('Date', '')
            if pd.notna(date):
                try:
                    date_str = pd.to_datetime(date).strftime('%Y-%m-%d')
                except:
                    date_str = str(date)
            else:
                date_str = 'N/A'
            label = f"ID: {podcast_id} - {name} ({date_str})"
            podcast_meetings_list.append(label)
            meeting_index_map[label] = idx
        
        if 'edit_podcast_meeting_id' in st.session_state:
            selected_meeting_id = st.session_state.edit_podcast_meeting_id
            mask = st.session_state.podcast_meetings_df['Podcast ID'] == selected_meeting_id
            if mask.any():
                selected_idx = st.session_state.podcast_meetings_df[mask].index[0]
                selected_meeting = st.session_state.podcast_meetings_df.iloc[selected_idx]
                selected_meeting_label = f"ID: {selected_meeting_id} - {selected_meeting.get('Name', 'N/A')} ({pd.to_datetime(selected_meeting.get('Date', '')).strftime('%Y-%m-%d') if pd.notna(selected_meeting.get('Date', '')) else 'N/A'})"
                default_index = podcast_meetings_list.index(selected_meeting_label) if selected_meeting_label in podcast_meetings_list else 0
                del st.session_state.edit_podcast_meeting_id
            else:
                default_index = 0
        else:
            default_index = 0
        
        selected_meeting_label = st.selectbox("Select Podcast Meeting to Edit:", podcast_meetings_list, index=default_index)
        
        if selected_meeting_label in meeting_index_map:
            selected_idx = meeting_index_map[selected_meeting_label]
            selected_meeting = st.session_state.podcast_meetings_df.iloc[selected_idx]
            selected_podcast_id = selected_meeting.get('Podcast ID')
            st.session_state.selected_podcast_meeting_index = selected_idx
            
            st.markdown("### 📋 Current Podcast Meeting Information")
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"**Podcast ID:** {selected_podcast_id}")
                st.write(f"**Name:** {selected_meeting.get('Name', 'N/A')}")
                if selected_meeting.get('Designation'):
                    st.write(f"**Designation:** {selected_meeting.get('Designation', 'N/A')}")
                if selected_meeting.get('Organization'):
                    st.write(f"**Organization:** {selected_meeting.get('Organization', 'N/A')}")
            with col2:
                date_val = selected_meeting.get('Date', '')
                if pd.notna(date_val):
                    try:
                        date_str = pd.to_datetime(date_val).strftime('%Y-%m-%d')
                    except:
                        date_str = str(date_val)
                else:
                    date_str = 'N/A'
                st.write(f"**Date:** {date_str}")
                st.write(f"**Time:** {selected_meeting.get('Time', 'N/A')}")
                st.write(f"**Status:** {selected_meeting.get('Status', 'N/A')}")
        
            st.markdown("---")
            with st.form("edit_podcast_meeting_form"):
                st.markdown("### 📝 Basic Information")
                col1, col2 = st.columns(2)
                
                with col1:
                    edit_name = st.text_input("Name *", value=str(selected_meeting.get('Name', '')), placeholder="Enter guest/speaker name", help="Enter the name of the guest/speaker (Required)")
                    edit_designation = st.text_input("Designation", value=str(selected_meeting.get('Designation', '')), placeholder="Enter job title or designation")
                    edit_organization = st.text_input("Organization", value=str(selected_meeting.get('Organization', '')), placeholder="Enter organization name")
                    edit_linkedin_url = st.text_input("LinkedIn URL", value=str(selected_meeting.get('LinkedIn URL', '')), placeholder="Enter LinkedIn profile URL")
                
                with col2:
                    edit_host = st.text_input("Host", value=str(selected_meeting.get('Host', '')), placeholder="Enter podcast host name")
                    current_status = str(selected_meeting.get('Status', 'Upcoming'))
                    edit_status = st.selectbox("Status", ["Upcoming", "Completed", "Cancelled"], index=["Upcoming", "Completed", "Cancelled"].index(current_status) if current_status in ["Upcoming", "Completed", "Cancelled"] else 0)
                    edit_contacted_through = st.text_input("Contacted Through (Platform)", value=str(selected_meeting.get('Contacted Through', '')), placeholder="e.g., LinkedIn, Email, etc.")
                
                st.markdown("### 🕐 Date & Time")
                col_date1, col_date2, col_date3 = st.columns(3)
                
                with col_date1:
                    date_val = selected_meeting.get('Date', None)
                    if pd.notna(date_val):
                        try:
                            edit_date = st.date_input("Date", value=pd.to_datetime(date_val).date())
                        except:
                            edit_date = st.date_input("Date", value=None)
                    else:
                        edit_date = st.date_input("Date", value=None)
                
                with col_date2:
                    edit_day = st.text_input("Day", value=str(selected_meeting.get('Day', '')), placeholder="e.g., Monday, Tuesday")
                
                with col_date3:
                    time_str = str(selected_meeting.get('Time', ''))
                    try:
                        if ':' in time_str and time_str.strip():
                            time_parts = time_str.split(':')
                            edit_time = st.time_input("Time", value=datetime.strptime(time_str[:5], '%H:%M').time() if len(time_parts) >= 2 else None)
                        else:
                            edit_time = st.time_input("Time", value=None)
                    except:
                        edit_time = st.time_input("Time", value=None)
                
                st.markdown("### 💬 Comments")
                edit_comments = st.text_area("Comments", value=str(selected_meeting.get('Comments', '')), height=100)
                
                update_submitted = st.form_submit_button("💾 Update Podcast Meeting", type="primary", use_container_width=True)
                
                if update_submitted:
                    if not edit_name.strip():
                        st.error("Name is required")
                    else:
                        idx = st.session_state.selected_podcast_meeting_index
                        st.session_state.podcast_meetings_df.at[idx, 'Name'] = edit_name.strip()
                        st.session_state.podcast_meetings_df.at[idx, 'Designation'] = edit_designation.strip() if edit_designation else ''
                        st.session_state.podcast_meetings_df.at[idx, 'Organization'] = edit_organization.strip() if edit_organization else ''
                        st.session_state.podcast_meetings_df.at[idx, 'LinkedIn URL'] = edit_linkedin_url.strip() if edit_linkedin_url else ''
                        st.session_state.podcast_meetings_df.at[idx, 'Host'] = edit_host.strip() if edit_host else ''
                        st.session_state.podcast_meetings_df.at[idx, 'Date'] = edit_date if edit_date else pd.NaT
                        st.session_state.podcast_meetings_df.at[idx, 'Day'] = edit_day.strip() if edit_day else ''
                        st.session_state.podcast_meetings_df.at[idx, 'Time'] = edit_time.strftime('%H:%M:%S') if edit_time else ''
                        st.session_state.podcast_meetings_df.at[idx, 'Status'] = edit_status
                        st.session_state.podcast_meetings_df.at[idx, 'Contacted Through'] = edit_contacted_through.strip() if edit_contacted_through else ''
                        st.session_state.podcast_meetings_df.at[idx, 'Comments'] = edit_comments.strip() if edit_comments else ''
                        
                        if save_podcast_meetings(st.session_state.podcast_meetings_df):
                            st.success("✅ Podcast Meeting Updated Successfully!")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error("❌ Failed to update podcast meeting. Please try again.")

# PAGE 6: Podcast Meetings Summary & Export
elif st.session_state.current_page == "Podcast Meetings Summary & Export":
    if st.session_state.podcast_meetings_df.empty:
        st.info("📭 No podcast meetings found. Please add a podcast meeting first.")
    else:
        st.markdown("### 🔍 Filters")
        col_filter1, col_filter2, col_filter3 = st.columns(3)
        
        with col_filter1:
            status_filter = st.multiselect("Filter by Status", options=["Upcoming", "Completed", "Cancelled"], default=[])
        with col_filter2:
            organization_filter = st.text_input("Filter by Organization", placeholder="Enter organization name")
        with col_filter3:
            host_filter = st.text_input("Filter by Host", placeholder="Enter host name")
        
        filtered_meetings = st.session_state.podcast_meetings_df.copy()
        
        if status_filter:
            filtered_meetings = filtered_meetings[filtered_meetings['Status'].isin(status_filter)]
        if organization_filter:
            filtered_meetings = filtered_meetings[filtered_meetings['Organization'].astype(str).str.contains(organization_filter, case=False, na=False)]
        if host_filter:
            filtered_meetings = filtered_meetings[filtered_meetings['Host'].astype(str).str.contains(host_filter, case=False, na=False)]
        
        # Import/Upload section - moved to top for better visibility
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
            st.write("Upload an Excel file to import or update podcast meeting records. Download the template below to ensure correct format.")
        with col_template2:
            # Create template dataframe with all template columns
            template_columns = [
                'Podcast ID', 'Name', 'Designation', 'Organization', 'LinkedIn URL',
                'Host', 'Date', 'Day', 'Time', 'Status', 'Contacted Through', 'Comments'
            ]
            template_df = pd.DataFrame(columns=template_columns)
            # Add sample row
            template_df = pd.concat([template_df, pd.DataFrame([{
                'Podcast ID': 1,
                'Name': 'Sample Guest',
                'Designation': 'CEO',
                'Organization': 'Sample Org',
                'LinkedIn URL': 'https://linkedin.com/in/sample',
                'Host': 'John Doe',
                'Date': datetime.now().date(),
                'Day': 'Monday',
                'Time': datetime.now().time().strftime('%H:%M:%S'),
                'Status': 'Upcoming',
                'Contacted Through': 'LinkedIn',
                'Comments': 'Sample comments'
            }])], ignore_index=True)
            
            # Save template to bytes
            import io
            template_buffer = io.BytesIO()
            template_df.to_excel(template_buffer, index=False)
            template_buffer.seek(0)
            
            st.download_button(
                label="📥 Download Template",
                data=template_buffer,
                file_name="podcast_meetings_import_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Download a template Excel file with the correct column format",
                use_container_width=True
            )
        
        uploaded_file = st.file_uploader(
            "Choose an Excel file to import",
            type=['xlsx', 'xls'],
            help="Upload an Excel file with podcast meeting data. Required columns: Name. All other columns are optional.",
            key="podcast_uploader"
        )
        
        if uploaded_file is not None:
            try:
                # Read the uploaded file without header first to inspect
                temp_df = pd.read_excel(uploaded_file, header=None)
                
                # Look for a row that contains "Name" (case-insensitive) - check first 10 rows
                header_row = None
                for idx in range(min(10, len(temp_df))):
                    row_values = [str(val).strip() if pd.notna(val) else '' for val in temp_df.iloc[idx].values]
                    row_lower = [val.lower() for val in row_values]
                    # Check if this row contains "name" as a column header
                    if any('name' in val and len(val) < 20 for val in row_lower):  # "name" should be a short word, not part of a long text
                        header_row = idx
                        break
                
                if header_row is not None:
                    # Re-read with the correct header row
                    import_df = pd.read_excel(uploaded_file, header=header_row)
                else:
                    # Try with header=0 first
                    import_df = pd.read_excel(uploaded_file, header=0)
                    # If we got "Unnamed" columns, use first non-empty row as header
                    if any('Unnamed' in str(col) for col in import_df.columns) or len([c for c in import_df.columns if str(c).strip()]) == 0:
                        # Find first row with actual data/headers
                        for idx in range(min(5, len(temp_df))):
                            row_values = [str(val).strip() if pd.notna(val) else '' for val in temp_df.iloc[idx].values]
                            if any(val for val in row_values):  # If row has any non-empty values
                                import_df = pd.read_excel(uploaded_file, header=idx)
                                break
                
                # Normalize column names (strip whitespace and make case-insensitive mapping)
                column_mapping = {}
                for col in import_df.columns:
                    normalized = str(col).strip()
                    # Skip unnamed columns and empty columns
                    if normalized and not normalized.startswith('Unnamed') and normalized != '':
                        column_mapping[normalized.lower()] = col
                
                # Check only critical required columns (case-insensitive)
                critical_required_columns = ['name']
                missing_critical = []
                found_columns = {}
                
                for req_col in critical_required_columns:
                    if req_col.lower() in column_mapping:
                        found_columns[req_col] = column_mapping[req_col.lower()]
                    else:
                        missing_critical.append(req_col)
                
                if missing_critical:
                    st.error(f"❌ Missing critical required column: {', '.join([c.capitalize() for c in missing_critical])}")
                    st.info("At minimum, 'Name' column is required. Other missing columns will be filled with empty values.")
                    st.info(f"📋 Found columns in your file: {', '.join([str(c) for c in import_df.columns[:10]])}")
                else:
                    # Rename columns to standard format (case-insensitive)
                    rename_dict = {}
                    for std_col in ['Podcast ID', 'Name', 'Designation', 'Organization', 'LinkedIn URL',
                                   'Host', 'Date', 'Day', 'Time', 'Status', 'Contacted Through', 'Comments']:
                        std_col_lower = std_col.lower()
                        if std_col_lower in column_mapping:
                            original_col = column_mapping[std_col_lower]
                            if original_col != std_col:
                                rename_dict[original_col] = std_col
                    
                    if rename_dict:
                        import_df = import_df.rename(columns=rename_dict)
                    # Add missing columns with empty values
                    template_columns = [
                        'Podcast ID', 'Name', 'Designation', 'Organization', 'LinkedIn URL',
                        'Host', 'Date', 'Day', 'Time', 'Status', 'Contacted Through', 'Comments'
                    ]
                    
                    missing_columns = [col for col in template_columns if col not in import_df.columns]
                    if missing_columns:
                        for col in missing_columns:
                            import_df[col] = ''
                    
                    # Ensure datetime columns are properly formatted
                    if 'Date' in import_df.columns:
                        import_df['Date'] = pd.to_datetime(import_df['Date'], errors='coerce')
                    
                    # Show preview
                    st.markdown("**📋 Preview of Uploaded Data:**")
                    st.dataframe(import_df.head(10), use_container_width=True, hide_index=True)
                    st.caption(f"Total rows to import: {len(import_df)}")
                    
                    # Proceed with import
                    if st.button("✅ Import Data", type="primary", use_container_width=True, key="import_podcast_btn"):
                        try:
                            # Ensure all template columns exist
                            template_columns = [
                                'Podcast ID', 'Name', 'Designation', 'Organization', 'LinkedIn URL',
                                'Host', 'Date', 'Day', 'Time', 'Status', 'Contacted Through', 'Comments'
                            ]
                            for col in template_columns:
                                if col not in import_df.columns:
                                    import_df[col] = ''
                            
                            # Fill NaN values with empty strings for text columns
                            text_columns = ['Podcast ID', 'Name', 'Designation', 'Organization', 'LinkedIn URL',
                                          'Host', 'Day', 'Time', 'Status', 'Contacted Through', 'Comments']
                            for col in text_columns:
                                if col in import_df.columns:
                                    import_df[col] = import_df[col].fillna('').astype(str)
                            
                            # Handle Status - normalize to valid values
                            if 'Status' not in import_df.columns:
                                import_df['Status'] = 'Upcoming'
                            else:
                                import_df['Status'] = import_df['Status'].apply(normalize_podcast_status)
                            
                            # Get current dataframe
                            current_df = st.session_state.podcast_meetings_df.copy()
                            
                            # Clean up Podcast ID column
                            if 'Podcast ID' in import_df.columns:
                                import_df['Podcast ID'] = import_df['Podcast ID'].replace('', pd.NA)
                                import_df['Podcast ID'] = import_df['Podcast ID'].replace(' ', pd.NA)
                            
                            if current_df.empty:
                                # If no existing data, just add all
                                if 'Podcast ID' not in import_df.columns or import_df['Podcast ID'].isna().all():
                                    import_df['Podcast ID'] = range(1, len(import_df) + 1)
                                else:
                                    missing_mask = import_df['Podcast ID'].isna()
                                    if missing_mask.any():
                                        max_id = pd.to_numeric(import_df['Podcast ID'], errors='coerce').max()
                                        if pd.isna(max_id):
                                            max_id = 0
                                        next_id = int(max_id) + 1
                                        import_df.loc[missing_mask, 'Podcast ID'] = range(next_id, next_id + missing_mask.sum())
                                st.session_state.podcast_meetings_df = import_df.copy()
                                added_count = len(import_df)
                                updated_count = 0
                            else:
                                # Handle Podcast ID
                                if 'Podcast ID' not in import_df.columns or import_df['Podcast ID'].isna().all():
                                    if 'Podcast ID' in current_df.columns:
                                        max_id = pd.to_numeric(current_df['Podcast ID'], errors='coerce').max()
                                        if pd.isna(max_id):
                                            max_id = 0
                                    else:
                                        max_id = 0
                                    import_df['Podcast ID'] = range(int(max_id) + 1, int(max_id) + 1 + len(import_df))
                                else:
                                    missing_mask = import_df['Podcast ID'].isna()
                                    if missing_mask.any():
                                        max_current = pd.to_numeric(current_df['Podcast ID'], errors='coerce').max() if 'Podcast ID' in current_df.columns else 0
                                        max_import = pd.to_numeric(import_df['Podcast ID'], errors='coerce').max()
                                        max_id = max(max_current if not pd.isna(max_current) else 0, max_import if not pd.isna(max_import) else 0)
                                        next_id = int(max_id) + 1
                                        import_df.loc[missing_mask, 'Podcast ID'] = range(next_id, next_id + missing_mask.sum())
                                
                                import_df['Podcast ID'] = pd.to_numeric(import_df['Podcast ID'], errors='coerce')
                                if 'Podcast ID' in current_df.columns:
                                    current_df['Podcast ID'] = pd.to_numeric(current_df['Podcast ID'], errors='coerce')
                                
                                added_count = 0
                                updated_count = 0
                                
                                # Update & Add New mode
                                if 'Podcast ID' in current_df.columns and 'Podcast ID' in import_df.columns:
                                    existing_ids = set(pd.to_numeric(current_df['Podcast ID'], errors='coerce').dropna().astype(int))
                                    import_df_ids = pd.to_numeric(import_df['Podcast ID'], errors='coerce')
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
                                        podcast_id = pd.to_numeric(row.get('Podcast ID'), errors='coerce')
                                        if pd.notna(podcast_id) and 'Podcast ID' in current_df.columns:
                                            idx = current_df[pd.to_numeric(current_df['Podcast ID'], errors='coerce') == podcast_id].index[0]
                                            for col in current_df.columns:
                                                if col in row:
                                                    current_df.at[idx, col] = row[col]
                                    updated_count = len(to_update)
                                
                                # Add new
                                if not to_add.empty:
                                    current_df = pd.concat([current_df, to_add], ignore_index=True)
                                    added_count = len(to_add)
                                
                                st.session_state.podcast_meetings_df = current_df
                            
                            # Save to database and Excel
                            if save_podcast_meetings(st.session_state.podcast_meetings_df):
                                success_msg = "✅ Import completed successfully!"
                                if added_count > 0:
                                    success_msg += f" Added {added_count} new podcast meeting(s)."
                                if updated_count > 0:
                                    success_msg += f" Updated {updated_count} existing podcast meeting(s)."
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
        
        st.markdown("### 📊 Summary Statistics")
        col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
        with col_stat1:
            st.metric("Total Podcast Meetings", len(st.session_state.podcast_meetings_df))
        with col_stat2:
            upcoming_count = len(st.session_state.podcast_meetings_df[st.session_state.podcast_meetings_df['Status'] == 'Upcoming'])
            st.metric("Upcoming", upcoming_count)
        with col_stat3:
            completed_count = len(st.session_state.podcast_meetings_df[st.session_state.podcast_meetings_df['Status'] == 'Completed'])
            st.metric("Completed", completed_count)
        with col_stat4:
            cancelled_count = len(st.session_state.podcast_meetings_df[st.session_state.podcast_meetings_df['Status'] == 'Cancelled'])
            st.metric("Cancelled", cancelled_count)
        
        st.markdown("### 📋 Podcast Meetings List")
        
        if not filtered_meetings.empty:
            available_columns = ['Podcast ID', 'Name', 'Designation', 'Organization', 'Host', 'Date', 'Day', 'Time', 'Status', 'Contacted Through']
            display_columns = [col for col in available_columns if col in filtered_meetings.columns]
            display_df = filtered_meetings[display_columns].copy()
            
            if 'Date' in display_df.columns:
                display_df['Date'] = display_df['Date'].apply(lambda x: pd.to_datetime(x).strftime('%Y-%m-%d') if pd.notna(x) else 'N/A')
            
            # Multi-select and delete section
            col_select1, col_select2 = st.columns([3, 1])
            with col_select1:
                select_all = st.checkbox("Select All", key="select_all_podcast_meetings", help="Select/deselect all podcast meetings")
                if select_all:
                    # Select all podcast IDs
                    if 'Podcast ID' in filtered_meetings.columns:
                        st.session_state.selected_podcast_meetings = set(
                            int(podcast_id) for podcast_id in filtered_meetings['Podcast ID'].dropna() 
                            if pd.notna(podcast_id)
                        )
                else:
                    # Clear selection if "Select All" is unchecked
                    if 'Podcast ID' in filtered_meetings.columns:
                        filtered_ids = set(
                            int(podcast_id) for podcast_id in filtered_meetings['Podcast ID'].dropna() 
                            if pd.notna(podcast_id)
                        )
                        # Only clear if all filtered podcast meetings were selected
                        if st.session_state.selected_podcast_meetings == filtered_ids:
                            st.session_state.selected_podcast_meetings = set()
            
            with col_select2:
                if st.session_state.selected_podcast_meetings:
                    if st.button("🗑️ Delete Selected", type="primary", use_container_width=True, 
                               help=f"Delete {len(st.session_state.selected_podcast_meetings)} selected podcast meeting(s)", key="bulk_delete_podcast"):
                        # Delete selected podcast meetings
                        deleted_count = 0
                        failed_count = 0
                        ids_to_delete = list(st.session_state.selected_podcast_meetings.copy())
                        
                        for podcast_id_int in ids_to_delete:
                            try:
                                delete_success = True
                                if get_use_supabase() and init_db_pool():
                                    delete_success = delete_podcast_meeting_from_supabase(podcast_id_int)
                                if delete_success:
                                    if 'Podcast ID' in st.session_state.podcast_meetings_df.columns:
                                        st.session_state.podcast_meetings_df = st.session_state.podcast_meetings_df[
                                            st.session_state.podcast_meetings_df['Podcast ID'] != podcast_id_int
                                        ]
                                    st.session_state.selected_podcast_meetings.discard(podcast_id_int)
                                    deleted_count += 1
                                else:
                                    failed_count += 1
                            except Exception as e:
                                failed_count += 1
                                st.session_state.selected_podcast_meetings.discard(podcast_id_int)
                        
                        if deleted_count > 0:
                            save_podcast_meetings(st.session_state.podcast_meetings_df)
                            if failed_count == 0:
                                st.success(f"{deleted_count} Podcast Meeting(s) Deleted Successfully")
                            else:
                                st.warning(f"{deleted_count} Podcast Meeting(s) Deleted Successfully, {failed_count} Failed")
                            st.rerun()
                        elif failed_count > 0:
                            st.error(f"Failed to delete {failed_count} podcast meeting(s)")
            
            num_data_cols = len(display_columns)
            col_widths = [1] + [3] * num_data_cols + [1, 1]
            
            # Scrollable podcast table container
            st.markdown("<div id='podcast-table-scroll-marker'></div>", unsafe_allow_html=True)
            st.markdown("""
            <style>
            div[data-testid="stVerticalBlock"]:has(#podcast-table-scroll-marker) {
                max-height: 65vh !important;
                overflow-y: auto !important;
                overflow-x: auto !important;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                padding: 0.5rem;
                margin-bottom: 0.5rem;
            }
            </style>
            """, unsafe_allow_html=True)
            
            header_cols = st.columns(col_widths)
            header_cols[0].markdown("<div style='line-height: 1.4; white-space: nowrap;'><small><strong>Select</strong></small></div>", unsafe_allow_html=True)
            for idx, col_name in enumerate(display_columns):
                header_cols[idx + 1].markdown(f"<div style='line-height: 1.4; word-wrap: break-word; white-space: normal; overflow: hidden;'><small><strong title='{col_name}'>{col_name}</strong></small></div>", unsafe_allow_html=True)
            header_cols[-2].markdown("<div style='line-height: 1.4; white-space: nowrap;'><small><strong>Edit</strong></small></div>", unsafe_allow_html=True)
            header_cols[-1].markdown("<div style='line-height: 1.4; white-space: nowrap;'><small><strong>Delete</strong></small></div>", unsafe_allow_html=True)
            
            st.markdown("<hr style='margin: 0.5rem 0;'>", unsafe_allow_html=True)
            
            for pos, (idx, row) in enumerate(display_df.iterrows()):
                row_cols = st.columns(col_widths)
                podcast_id = None
                if 'Podcast ID' in filtered_meetings.columns and pos < len(filtered_meetings):
                    podcast_id = filtered_meetings.iloc[pos].get('Podcast ID')
                
                if podcast_id is not None and pd.notna(podcast_id):
                    podcast_id_int = int(podcast_id)
                    is_selected = podcast_id_int in st.session_state.selected_podcast_meetings
                    checkbox_key = f"select_podcast_checkbox_{podcast_id_int}_{idx}"
                    checkbox_state = row_cols[0].checkbox("", value=is_selected, key=checkbox_key, label_visibility="collapsed", help="Select podcast meeting")
                    if checkbox_state:
                        st.session_state.selected_podcast_meetings.add(podcast_id_int)
                    else:
                        st.session_state.selected_podcast_meetings.discard(podcast_id_int)
                
                for col_idx, col_name in enumerate(display_columns):
                    value = row.get(col_name, '')
                    if pd.isna(value):
                        value = ''
                    value_str = str(value) if value else ''
                    if len(value_str) > 50 and col_name in ['Name', 'Organization']:
                        value_str = value_str[:47] + "..."
                    row_cols[col_idx + 1].markdown(f"<div style='line-height: 1.5; word-wrap: break-word; overflow-wrap: break-word; white-space: normal;'><small>{value_str}</small></div>", unsafe_allow_html=True)
                
                if podcast_id is not None and pd.notna(podcast_id):
                    if row_cols[-2].button("✏️", key=f"edit_podcast_{podcast_id}_{idx}", use_container_width=False, help="Edit podcast meeting"):
                        st.session_state.current_page = "Edit/Update Podcast Meeting"
                        st.session_state.edit_podcast_meeting_id = int(podcast_id)
                        st.rerun()
                
                if podcast_id is not None and pd.notna(podcast_id):
                    if row_cols[-1].button("🗑️", key=f"delete_podcast_{podcast_id}_{idx}", use_container_width=False, type="secondary", help="Delete podcast meeting"):
                        try:
                            podcast_id_int = int(podcast_id)
                            delete_success = True
                            if get_use_supabase() and init_db_pool():
                                delete_success = delete_podcast_meeting_from_supabase(podcast_id_int)
                            if delete_success:
                                if 'Podcast ID' in st.session_state.podcast_meetings_df.columns:
                                    st.session_state.podcast_meetings_df = st.session_state.podcast_meetings_df[
                                        st.session_state.podcast_meetings_df['Podcast ID'] != podcast_id_int
                                    ]
                                save_podcast_meetings(st.session_state.podcast_meetings_df)
                                st.success("✅ Podcast Meeting Deleted Successfully")
                                st.rerun()
                            else:
                                st.error(f"❌ Failed to delete podcast meeting {podcast_id_int} from database. Please try again.")
                        except Exception as e:
                            st.error(f"Error deleting podcast meeting: {e}")
                
                st.markdown("<hr style='margin: 0.3rem 0;'>", unsafe_allow_html=True)
            
            st.caption(f"Showing {len(display_df)} podcast meeting(s)")
        else:
            st.info("📭 No podcast meetings found matching your filters.")
        
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
        
        if not st.session_state.podcast_meetings_df.empty:
            if st.button("📥 Export to Excel", type="primary", use_container_width=True, key="export_podcast"):
                export_filename = f"podcast_meetings_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                try:
                    st.session_state.podcast_meetings_df.to_excel(export_filename, index=False)
                    st.success(f"✅ Data exported to {export_filename}")
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
            st.info("📭 No data available to export. Add podcast meetings first.")

