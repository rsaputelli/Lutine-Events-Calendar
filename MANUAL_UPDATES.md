
# Manual Updates: Clients & Meeting Managers

For occasional updates (new clients or meeting managers), we recommend making changes directly in Supabase rather than using SQL scripts.

## Access the Supabase Table Editor
1. Log in to [Supabase Dashboard](https://app.supabase.com/).
2. Select the **Master Calendar** project.
3. In the left menu, go to **Table Editor**.
4. Open the relevant table:
   - `clients`
   - `meeting_managers`

## Adding a New Entry
1. Click **Insert Row** (➕ button).
2. Fill in required fields:
   - **Clients table**:  
     - `name` → the full client name (case-insensitive uniqueness is recommended; avoid duplicates).  
   - **Meeting Managers table**:  
     - `name` → the person’s name.  
     - `email` → their work email (must be unique; check before adding).
3. Click **Save**.

## Editing an Existing Entry
1. Locate the row.  
   - Use search/filter in the table editor if needed.
2. Click into the cell you want to update (e.g., email or name).
3. Make the change and **Save**.

⚠️ **Tip:** Changing a manager’s email here does *not* change their user login (Supabase Auth). This table is just for dropdowns in the app.

## Deleting an Entry
1. Select the row you want to remove.
2. Click the **trash can icon** or **Delete row**.
3. Confirm.

⚠️ **Caution:**  
- Deleting a client or manager row does **not** automatically update past events that reference it. Past events will keep the old value.  
- That’s usually fine for historical accuracy, but double-check if you’re renaming/replacing someone.

## How Dropdowns Use These Tables
- When creating or editing an event in the Master Calendar app:
  - The **Client** dropdown is populated from the `clients` table.
  - The **Meeting Manager** dropdown is populated from the `meeting_managers` table.
- If you add a new client or manager in Supabase, they will immediately appear in the dropdowns the next time the app is loaded.
- If you remove a client/manager, they will no longer appear in the dropdowns. Past events that reference them will still show their name/email.

✅ **Best Practice**:  
- Keep `clients` list short and consistent — use canonical names like “DAFP” or “WAPA” (avoid variations).  
- For `meeting_managers`, always verify the email before adding to prevent duplicates.  
