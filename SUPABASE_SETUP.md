# Supabase Database Setup Guide

This application can connect to Supabase for storing meeting data in a PostgreSQL database.

## Prerequisites

1. Supabase account with a project created
2. Database password from your Supabase project settings

## Configuration

### Option 1: Using Streamlit Secrets (Recommended for Streamlit Cloud)

1. Create a `.streamlit/secrets.toml` file in your project root (or use Streamlit Cloud's secrets manager)
2. Add the following configuration:

```toml
SUPABASE_DB_PASSWORD = "your_database_password_here"
USE_SUPABASE = "true"
```

**Important**: You need the **Database Password**, NOT the API key!

- ❌ API Key format: `sbp_...` or `sb-...` (used for REST API)
- ✅ Database Password: Usually a long random string (used for direct PostgreSQL connection)

**How to find your Database Password:**
1. Go to your Supabase project dashboard
2. Click **Settings** (gear icon in sidebar)
3. Click **Database**
4. Scroll to **Connection string** or **Connection pooling**
5. Look for the **Password** field - this is what you need!
6. Copy that password (NOT the API key)

### Option 2: Using Environment Variables (Recommended for local development)

Set the following environment variables:

```bash
export SUPABASE_DB_PASSWORD="your_database_password_here"
export USE_SUPABASE="true"
```

On Windows PowerShell:
```powershell
$env:SUPABASE_DB_PASSWORD="your_database_password_here"
$env:USE_SUPABASE="true"
```

## Database Setup

Run the SQL script provided earlier to create the `meetings` table and related functions in your Supabase database.

### Steps:

1. Go to your Supabase project dashboard
2. Navigate to SQL Editor
3. Run the SQL script to create:
   - `meetings` table
   - `meetings_audit_log` table (optional, for tracking changes)
   - Required indexes and triggers
   - Helper functions

## Connection Details

The application uses the following connection settings (configured in code):
- **Host**: aws-1-ap-south-1.pooler.supabase.com
- **Port**: 6543
- **Database**: postgres
- **User**: postgres.xrpmswlgatrshvgwtvjw
- **Password**: Set in secrets/environment variable

## Features

### Automatic Syncing
- All meetings are automatically saved to Supabase when added, updated, or deleted
- Excel file is also maintained as a backup

### Data Migration
If you have existing Excel data:
1. The app will load from Supabase first (if available)
2. Falls back to Excel if Supabase is not available
3. To migrate Excel data to Supabase:
   - Enable Supabase connection
   - Use the "Import from Excel" feature in the app
   - All data will be synced to Supabase

### Fallback Mode
If Supabase connection fails, the app automatically falls back to Excel-only mode, ensuring your data is never lost.

## Troubleshooting

### Connection Issues

1. **Check Password**: Ensure the password in secrets.toml matches your Supabase database password
2. **Check Network**: Ensure your network allows connections to Supabase
3. **Check Supabase Status**: Verify your Supabase project is active

### Data Not Syncing

1. Check if `USE_SUPABASE` is set to `"true"`
2. Verify database password is correct
3. Check Streamlit logs for error messages
4. Ensure the `meetings` table exists in your Supabase database

## Security Notes

- Never commit `.streamlit/secrets.toml` to version control
- Use environment variables in production
- Keep your database password secure
- Consider using Supabase Row Level Security (RLS) for additional security
