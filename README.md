# 📅 AI Geo Navigators - Meeting Dashboard

A simple and efficient Streamlit application for managing and tracking team meetings.

## Features

- ✅ **Add New Meetings** - Create meetings with all necessary details
- ✏️ **Edit Meetings** - Update meeting information including manual status control
- 🗑️ **Delete Meetings** - Remove meetings with confirmation
- 📊 **Meeting Summary** - View all meetings with filters and search
- 📥 **Export Data** - Download meeting data as Excel files
- 🔄 **Auto Status Updates** - Automatic status calculation (Upcoming, Ongoing, Ended, Completed)
- 🎯 **Real-time Updates** - Auto-refresh option for live status updates

## Applications

### Meeting Dashboard (`meeting_dashboard.py`)
1. **Add New Meeting** - Create new meetings with form validation
2. **Edit or Delete Meeting** - Modify existing meetings or remove them
3. **Meetings Summary & Export** - View filtered meetings and export data

### Podcast Dashboard (`podcast_dashboard.py`) - NEW
A separate dashboard for managing podcast meetings with:
1. **Add New Podcast Meeting** - Create podcast meeting records with guest information
2. **Edit/Update Podcast Meeting** - Modify existing podcast meetings
3. **Podcast Meetings Summary & Export** - View, filter, and export podcast meeting data

**Note:** The podcast dashboard uses a separate database table (`podcast_meetings`) and does not interfere with regular meeting data.

## Installation

1. Clone this repository:
```bash
git clone <your-repo-url>
cd Meeting
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
# For regular meetings:
streamlit run meeting_dashboard.py

# For podcast meetings (NEW):
streamlit run podcast_dashboard.py
```

## Usage

1. **Add a Meeting**: Navigate to "Add New Meeting" page and fill in the form
2. **Edit a Meeting**: Go to "Edit or Delete Meeting" page, select a meeting, and update details
3. **View Summary**: Use "Meetings Summary & Export" to filter and view all meetings
4. **Export Data**: Export your meeting data to Excel for backup or analysis

## Data Storage

The app uses Excel files (`meeting_data.xlsx`) for data storage. This file is automatically created when you add your first meeting.

## Status Types

- **Upcoming**: Meeting start time is in the future
- **Ongoing**: Current time is between start and end time
- **Ended**: Meeting end time has passed
- **Completed**: Manually marked as completed (overrides automatic calculation)

## Deployment

### Streamlit Cloud (Recommended)

1. Push this repository to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Sign in with your GitHub account
4. Click "New app"
5. Select your repository and branch
6. Set the main file path: `meeting_dashboard.py`
7. Click "Deploy"

### Other Platforms

This app can also be deployed on:
- Heroku
- AWS EC2
- Google Cloud Run
- Azure App Service

Make sure to:
- Set up the required Python environment
- Install dependencies from `requirements.txt`
- Ensure file write permissions for Excel storage

## Requirements

- Python 3.7+
- Streamlit >= 1.28.0
- Pandas >= 2.0.0
- OpenPyXL >= 3.1.0

## License

This project is for internal use by AI Geo Navigators.

## Support

For issues or questions, please contact the development team.

