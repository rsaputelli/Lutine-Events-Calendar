# Future Updates – Master Calendar App

_Last updated: 2025-09-01 04:50 UTC_

This document tracks agreed next steps and potential enhancements for the Lutine Master Calendar application.

---

## 1) Restrict Export to Word
- [ ] Hide the **Admin: Export Events to Word** sidebar expander from non-admin roles (optionally allow `editor`).
- [ ] Gate download/build buttons with a role check.
- [ ] Add a small “why not visible?” hint for viewers (optional).

**Notes**
- Minimal code: wrap the expander and its buttons in `if ROLE in ('admin', 'editor'):`.
- Consider a `feature_flags` dict if we need per-env behavior.

---

## 2) Update “From” Email to no-reply@lutinemanagement.com
- [ ] Create mailbox/alias and credentials.
- [ ] Update `st.secrets["smtp"]` keys: `user`, `password`, and optionally `from_addr` / `from_name`.
- [ ] Send a test email from the app and confirm DKIM/SPF pass.

**Secrets block example**
```toml
[smtp]
host = "smtp.office365.com"
port = 587
user = "no-reply@lutinemanagement.com"
password = "********"
from_addr = "no-reply@lutinemanagement.com"
from_name = "Lutine Calendar Bot"
```

---

## 3) Admin-Only “Assign Owner” (Fix Orphaned Events)
- [ ] Add an **admin-only** control on the Edit tab: “Assign Owner to Event”.
- [ ] Inputs: `event_id` (prefilled from current selection), `user_email`.
- [ ] Server action: Look up user by email in `auth.users`/`profiles`, then set `events.created_by` **only if it is currently NULL**.
- [ ] Error if email not found, or `created_by` already set.
- [ ] Log `updated_by` and `updated_at`.

**Rationale**
Prevents viewers from editing events they shouldn’t, while allowing admins to resolve orphaned rows (e.g., created before RLS rules or by non-user meeting managers).

---

## 4) Nice-to-Haves / Backlog
- [ ] Per-client role scoping (future): allow editors for specific clients.
- [ ] Bulk owner-assignment tool (admin): query for `created_by IS NULL` and batch-assign.
- [ ] Audit trail table for sensitive updates (owner changes, deletes).
- [ ] Optional: SSO with Microsoft Entra ID for streamlined login.
- [ ] Optional: Ingest external Outlook events (toggle-able), with a flag to mark non-app events read-only in the app.

---

## 5) Testing Checklist (when items ship)
- [ ] Viewer cannot see admin/exports.
- [ ] Editor can export (if allowed) but cannot see other admin tools (if so configured).
- [ ] Admin can assign owner and the change takes effect immediately (viewer can no longer edit if not owner/manager).
- [ ] Emails send from `no-reply@…`, and pass SPF/DKIM.
- [ ] RLS still blocks direct SQL writes from non-owners.

---

## 6) Operational Notes
- Keep a small runbook for environment variables and secrets (`graph`, `supabase`, `smtp`).
- Any schema changes should be paired with a quick rollback note.
- Consider a staging app URL for testing role/RLS changes before production.
