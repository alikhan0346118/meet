# Deployment Guide

## Step 1: Configure Git (First Time Only)

If you haven't configured Git before, run these commands:

```bash
git config --global user.name "Your Name"
git config --global user.email "your.email@example.com"
```

Or for this repository only:
```bash
git config user.name "Your Name"
git config user.email "your.email@example.com"
```

## Step 2: Create GitHub Repository

1. Go to [GitHub.com](https://github.com) and sign in
2. Click the "+" icon in the top right corner
3. Select "New repository"
4. Repository name: `meeting-dashboard` (or any name you prefer)
5. Description: "AI Geo Navigators Meeting Dashboard"
6. Choose Public or Private (Private recommended for internal use)
7. **DO NOT** initialize with README, .gitignore, or license (we already have these)
8. Click "Create repository"

## Step 3: Push to GitHub

After creating the repository, GitHub will show you commands. Run these in your terminal:

```bash
# Add the remote repository (replace YOUR_USERNAME with your GitHub username)
git remote add origin https://github.com/YOUR_USERNAME/meeting-dashboard.git

# Rename branch to main (if needed)
git branch -M main

# Push to GitHub
git push -u origin main
```

## Step 4: Deploy to Streamlit Cloud

### Option A: Quick Deploy (Recommended)

1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Click "Sign in" and authorize with your GitHub account
3. Click "New app"
4. Fill in the details:
   - **Repository**: Select `your-username/meeting-dashboard`
   - **Branch**: `main`
   - **Main file path**: `meeting_dashboard.py`
   - **App URL**: (auto-generated, you can customize)
5. Click "Deploy"
6. Wait for deployment (usually 1-2 minutes)
7. Your app will be live at: `https://your-app-name.streamlit.app`

### Option B: Using Streamlit CLI

```bash
streamlit deploy
```

## Step 5: Post-Deployment Setup

### File Storage on Streamlit Cloud

**Important**: Streamlit Cloud has ephemeral file storage. Changes to files are lost when the app restarts.

**Solution Options:**

1. **Use Streamlit Secrets for database connection** (Recommended for production)
   - Store connection strings in Streamlit Secrets
   - Use SQLite or external database instead of Excel

2. **Use GitHub to persist data** (Quick fix)
   - Commit `meeting_data.xlsx` to GitHub
   - Note: Not ideal for multiple users

3. **Use external storage** (Best for production)
   - Google Sheets API
   - AWS S3
   - PostgreSQL/MySQL database

### Access Your Deployed App

Once deployed, you can:
- Share the URL with your team
- Bookmark it for easy access
- Update the code by pushing to GitHub (auto-redeploys)

## Troubleshooting

### If deployment fails:
1. Check that `requirements.txt` has all dependencies
2. Verify `meeting_dashboard.py` is in the root directory
3. Check Streamlit Cloud logs for errors

### If file writes don't persist:
- This is expected behavior on Streamlit Cloud
- Consider using a database or external storage solution

## Next Steps

For production use, consider:
- Adding authentication
- Using a proper database (SQLite, PostgreSQL)
- Implementing user roles and permissions
- Setting up automated backups

## Need Help?

Check the [Streamlit Cloud documentation](https://docs.streamlit.io/streamlit-community-cloud) for more details.

