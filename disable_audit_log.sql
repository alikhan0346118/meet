-- Quick fix: Disable the problematic audit log trigger temporarily
-- Run this in Supabase SQL Editor to allow syncing to work

DROP TRIGGER IF EXISTS log_meeting_changes_trigger ON meetings;

-- Optional: If you want to fix the trigger function later, use the fix_audit_log_function.sql file
