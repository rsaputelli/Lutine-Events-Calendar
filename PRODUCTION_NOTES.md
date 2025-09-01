# Production Notes — Master Calendar Intake (Streamlit + Supabase + Microsoft Graph)

**App:** Streamlit Master Calendar Intake (`app.py`)  
**DB:** Supabase (tables: `events`, `clients`, `meeting_managers`, `notifications`, `profiles`, `graph_state`)  
**Calendar:** Microsoft Graph (shared mailbox)  
**Mail (optional):** SMTP for manager + accreditation emails  
**Auth:** Supabase email/password (no self-signup UI), password reset via Site URL

---

## Core behaviors

### Create
- All-day vs. timed events handled correctly (Graph end is **exclusive**; ≥24h for all-day).
- `showAs = "free"` for all events.
- Outlook body contains:
  - `<p><b>Client:</b> …</p>` (if provided)
  - `<p><b>Accreditation:</b> Yes|No</p>`
  - Red **Meeting Manager** block (11pt table) + `[App Outlook Event ID: …]`.
- Virtual events:
  - Teams link auto (if chosen).
  - Zoom/Other: optional link. If blank, user confirms; app seeds a **7-day “missing_link”** reminder in `notifications`.

### Edit
- Patches Outlook core fields (subject/time/location/reminder).
- Preserves and **upserts** the red Manager block and the Client/Accreditation lines in body.
- Optional date-certain reminders via `notifications` (email sent by your worker/crontask).

### Export
- Word (DOCX) export grouped by month (sidebar panel + main button).  
- Shows ET times; all-day shows only dates.

### Admin tools (sidebar)
- **Bulk Sync (delta)**: Outlook ➜ App for app-owned events (by `outlook_event_id`).
- **Per-event Refresh**: paste an Outlook Event ID to pull updated fields.
- **Recent IDs**: quick list for copy/paste.

### Cleanup script (local)
- `cleanup_non_dafp.py`: deletes non-DAFP events in **Supabase + Outlook**.
  - `--dry-run` previews.
  - Requires `.env` with `SUPABASE_URL`, `SUPABASE_SERVICE_KEY`, Graph creds, `GRAPH_SHARED_MAILBOX_UPN`.

---

## Security & roles

### Auth gate
- Sign-in form only (no public sign-up).
- Password reset flows back via `_SITE_URL`.

### Roles (table: `profiles`)
- `role ∈ {admin, editor, viewer}` with default `viewer`.
- App reads role: `ROLE = _get_user_role(user["email"])` and stores it in `st.session_state["role"]`.
- **Admin sidebar tools** are **hidden** unless `role == "admin"` (guarded by UI containment).
- (Optional future) Add code-level checks before any admin action—pattern shown in thread.

> **Viewer edit limits**: With the current RLS plan, the intent is:
> - **Viewers** can only edit events they **created** (or where they are the **meeting manager**), once RLS is finalized in SQL.  
> - You verified the **UI gating** works; ensure RLS policies are applied in Supabase when you enable them fully (see “RLS notes” below).

---

## Config in use

### Streamlit `secrets.toml`
```toml
[graph]
tenant_id = "..."
client_id = "..."
client_secret = "..."
shared_mailbox_upn = "master-calendar@njafp.org"

[supabase]
url = "https://xxxxx.supabase.co"
key = "SUPABASE_SERVICE_KEY"        # service key for app DB ops
anon_key = "SUPABASE_ANON_KEY"      # used only for auth client
site_url = "https://<your-streamlit-app>.streamlit.app"

[smtp]  # optional
host = "..."
port = 587
user = "..."
password = "..."
from_addr = "no-reply@lutinemanagement.com"  # (future change, easy swap)
from_name = "Lutine Calendar Bot"
```

### Local `.env` (for cleanup script)
```env
SUPABASE_URL=...
SUPABASE_SERVICE_KEY=...

GRAPH_TENANT_ID=...
GRAPH_CLIENT_ID=...
GRAPH_CLIENT_SECRET=...
GRAPH_SHARED_MAILBOX_UPN=master-calendar@njafp.org
GRAPH_SLEEP=0.3
```

---

## Operational runbook

### Invite users
1. Add user via **Supabase Auth** (or have them sign in the first time if enabled internally).
2. Give a role (default viewer, then promote):
   ```sql
   insert into public.profiles (user_id, email, role)
   values ('<auth.users.id>', 'user@org.org', 'viewer')
   on conflict (user_id) do update set role = excluded.role;
   ```
   or later:
   ```sql
   update public.profiles set role = 'editor' where email = 'user@org.org';
   update public.profiles set role = 'admin'  where email = 'user@org.org';
   ```

### Known safe tests
- Create: all-day single day; multi-day; timed event crossing midnight; Virtual with no link (confirm) + reminder seed.
- Edit: change times/TZ; toggle all-day; change manager; verify Manager block & Client/Accreditation update.
- Export: verify ET times; all-day formatting; red manager label on card.
- Admin: bulk delta sync and per-event refresh with a known ID.

### Orphan fix (event without `created_by`)
- Manually set:
  ```sql
  update public.events
     set created_by = '<auth.users.id>'
   where id = '<event_id>';
  ```
- Or leave for backlog (see below) to auto-backfill on first edit by the creator/manager.

---

## RLS notes (when you enable/lock them down)

**High-level policy intent** (apply as SQL policies in Supabase):
- **SELECT**: all authenticated users can read `events`.
- **INSERT**: anyone authenticated can insert; set `created_by = auth.uid()`.
- **UPDATE/DELETE**:
  - allowed if `created_by = auth.uid()` **OR**
  - the row’s `meeting_manager_email = auth.jwt().email` **OR**
  - user’s role is `admin`.
- **Admin tools**: UI-gated already; add server-side checks (SQL policies) to enforce.

> Until all RLS policies are in place and **enabled**, owner privileges in the Supabase UI can bypass restrictions. Test in a normal user session to confirm behavior.

---

## Future backlog
1. **Export restrictions**: show export buttons only for `role in ('admin','editor')`.
2. **From address switch**: change `smtp.from_addr` to `no-reply@lutinemanagement.com` in `secrets.toml`.
3. **Orphan auto-claim**: when a logged-in user edits a row with `created_by is null`, set `created_by = auth.uid()` **only if** user is current meeting manager (and not otherwise).
4. **Bulk list admin**: small admin page to add/remove clients/managers with case-insensitive de-dupe.
5. **Periodic worker**: process `notifications` table for scheduled emails (if not already wired in your infra).

---

## Resume later
- Start a new chat and paste any error + the relevant function or block from `app.py`. I’ll diff and patch.
- Include: which tab you used, event type, TZ, Graph error text (if any), and whether any emails sent.

## Rollback
- Keep last good `app.py` as a **tag/release** (e.g., `v1.0.0-prod`) in the repo.
- If a deploy misbehaves, redeploy the previous commit/tag.
- DB changes committed here were additive and safe; no destructive migrations left pending.
