-- URGENT FIX: Run this NOW in Supabase SQL Editor
-- This will fix the audit log trigger error immediately

-- Step 1: Drop the broken trigger
DROP TRIGGER IF EXISTS log_meeting_changes_trigger ON meetings;

-- Step 2: Drop the broken function
DROP FUNCTION IF EXISTS log_meeting_changes();

-- Step 3: Recreate the function with correct JSONB handling
CREATE OR REPLACE FUNCTION log_meeting_changes()
RETURNS TRIGGER AS $$
DECLARE
    changed_fields_array TEXT[];
    field_name TEXT;
    old_val TEXT;
    new_val TEXT;
    old_json JSONB;
    new_json JSONB;
BEGIN
    IF TG_OP = 'DELETE' THEN
        INSERT INTO meetings_audit_log (
            meeting_id,
            operation,
            old_values,
            changed_by,
            changed_from
        ) VALUES (
            OLD.meeting_id,
            'DELETE',
            to_jsonb(row_to_json(OLD)),
            current_setting('app.user', true),
            COALESCE(current_setting('app.source', true), 'Streamlit App')
        );
        RETURN OLD;
    ELSIF TG_OP = 'INSERT' THEN
        INSERT INTO meetings_audit_log (
            meeting_id,
            operation,
            new_values,
            changed_by,
            changed_from
        ) VALUES (
            NEW.meeting_id,
            'INSERT',
            to_jsonb(row_to_json(NEW)),
            current_setting('app.user', true),
            COALESCE(current_setting('app.source', true), 'Streamlit App')
        );
        RETURN NEW;
    ELSIF TG_OP = 'UPDATE' THEN
        old_json := to_jsonb(row_to_json(OLD));
        new_json := to_jsonb(row_to_json(NEW));
        changed_fields_array := ARRAY[]::TEXT[];
        
        FOR field_name IN 
            SELECT column_name 
            FROM information_schema.columns 
            WHERE table_name = 'meetings' 
            AND column_name NOT IN ('id', 'created_at', 'updated_at', 'last_synced_at', 'sync_version')
        LOOP
            old_val := (old_json->>field_name);
            new_val := (new_json->>field_name);
            
            IF old_val IS DISTINCT FROM new_val THEN
                changed_fields_array := array_append(changed_fields_array, field_name);
            END IF;
        END LOOP;
        
        IF array_length(changed_fields_array, 1) > 0 THEN
            INSERT INTO meetings_audit_log (
                meeting_id,
                operation,
                old_values,
                new_values,
                changed_fields,
                changed_by,
                changed_from
            ) VALUES (
                NEW.meeting_id,
                'UPDATE',
                old_json,
                new_json,
                changed_fields_array,
                current_setting('app.user', true),
                COALESCE(current_setting('app.source', true), 'Streamlit App')
            );
        END IF;
        
        RETURN NEW;
    END IF;
    
    RETURN NULL;
END;
$$ LANGUAGE plpgsql;

-- Step 4: Recreate the trigger
CREATE TRIGGER log_meeting_changes_trigger
    AFTER INSERT OR UPDATE OR DELETE ON meetings
    FOR EACH ROW
    EXECUTE FUNCTION log_meeting_changes();

-- Done! Now try syncing your data again.
