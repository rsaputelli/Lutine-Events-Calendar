# pages/01_AV_Intake.py
import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timezone

from lib.supabase_client import get_supabase

st.title("AV Request Intake")
st.caption("Submit an AV request for a scheduled meeting (or as a standalone request).")

sb = get_supabase()

# ===============================
# Helpers
# ===============================
def _fmt_dt(ts) -> str:
    if not ts:
        return ""
    try:
        # Supabase returns ISO strings for timestamptz
        dt = pd.to_datetime(ts, utc=True)
        dt_local = dt.tz_convert("America/New_York")
        return dt_local.strftime("%Y-%m-%d %I:%M %p ET")
    except Exception:
        return str(ts)


def _venue_label(v: dict) -> str:
    name = v.get("name") or "Unnamed Venue"
    city = v.get("city") or ""
    state = v.get("state") or ""
    parts = [p for p in [city, state] if p]
    if parts:
        return f"{name} ({', '.join(parts)})"
    return name


def _build_venue_ship_to_text(v: dict) -> str:
    # Single formatted text block for shipping labels
    lines = []
    if v.get("name"):
        lines.append(v["name"])
    if v.get("address_line1"):
        lines.append(v["address_line1"])
    if v.get("address_line2"):
        lines.append(v["address_line2"])

    city = v.get("city") or ""
    state = v.get("state") or ""
    postal = v.get("postal_code") or ""
    city_line = " ".join([p for p in [city + ("," if city and state else ""), state, postal] if p]).strip()
    if city_line:
        lines.append(city_line)

    country = v.get("country") or "US"
    if country and country.upper() != "US":
        lines.append(country)

    return "\n".join([ln for ln in lines if ln]).strip()


def _now_utc_iso():
    return datetime.now(timezone.utc).isoformat()


def _date_to_iso(d):
    # Converts a Python date to YYYY-MM-DD string, or None
    if not d:
        return None
    try:
        return d.isoformat()
    except Exception:
        return str(d)

def _get_graph_token() -> str:
    g = st.secrets["graph"]
    token_url = f"https://login.microsoftonline.com/{g['tenant_id']}/oauth2/v2.0/token"
    data = {
        "client_id": g["client_id"],
        "client_secret": g["client_secret"],
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    r = requests.post(token_url, data=data, timeout=20)
    r.raise_for_status()
    return r.json()["access_token"]


def _graph_send_mail(token: str, shared_mailbox_upn: str, *, to_emails: list[str], cc_emails: list[str], subject: str, body_html: str):
    url = f"https://graph.microsoft.com/v1.0/users/{shared_mailbox_upn}/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    def _mk(email: str):
        return {"emailAddress": {"address": email}}

    to_recipients = [_mk(e) for e in to_emails if e]
    cc_recipients = [_mk(e) for e in cc_emails if e]

    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body_html},
            "toRecipients": to_recipients,
            "ccRecipients": cc_recipients,
        },
        "saveToSentItems": True,
    }

    r = requests.post(url, headers=headers, json=payload, timeout=20)
    if r.status_code >= 400:
        raise RuntimeError(f"Graph sendMail {r.status_code}: {r.text}")

# ===============================
# Dropdown loaders (reused concept from calendar app)
# ===============================
@st.cache_data(ttl=120)
def load_clients_dropdown():
    try:
        res = sb.table("clients").select("name").order("name", desc=False).execute()
        rows = res.data or []
        return [r.get("name") for r in rows if r.get("name")]
    except Exception:
        return []


@st.cache_data(ttl=120)
def load_meeting_managers_dropdown():
    try:
        res = (
            sb.table("meeting_managers")
            .select("name,email")
            .order("name", desc=False)
            .execute()
        )
        rows = res.data or []
        out = []
        for r in rows:
            name = (r.get("name") or "").strip()
            email = (r.get("email") or "").strip()
            if name:
                out.append((name, email))
        return out
    except Exception:
        return []

# ===============================
# Load Events (future only)
# ===============================
@st.cache_data(ttl=60)
def load_future_events():
    try:
        # Pull a limited set of fields used for intake
        res = (
            sb.table("events")
            .select("id, subject, client, start_dt_utc, end_dt_utc, meeting_manager_name, meeting_manager_email, location")
            .order("start_dt_utc", desc=False)
            .execute()
        )
        rows = res.data or []

        # Only include future events
        now = datetime.now(timezone.utc)
        filtered = []
        for r in rows:
            try:
                start = pd.to_datetime(r.get("start_dt_utc"), utc=True)
                if start >= now:
                    filtered.append(r)
            except Exception:
                # If parsing fails, keep it out to avoid garbage
                pass

        return filtered
    except Exception:
        return []


future_events = load_future_events()

event_options = [{"id": None, "label": "— Standalone request (no event yet) —"}]
for e in future_events:
    label = f"{_fmt_dt(e.get('start_dt_utc'))} | {e.get('subject','(No subject)')}"
    if e.get("client"):
        label += f" | {e['client']}"
    event_options.append({"id": e.get("id"), "label": label, "row": e})


# ===============================
# Load Venues
# ===============================
@st.cache_data(ttl=60)
def load_venues():
    try:
        res = (
            sb.table("venues")
            .select("id, name, address_line1, address_line2, city, state, postal_code, country, main_contact_name, main_contact_email, main_contact_phone, notes")
            .order("name", desc=False)
            .execute()
        )
        return res.data or []
    except Exception:
        return []


venues = load_venues()

# ---- Session defaults ----
if "venue_mode" not in st.session_state:
    st.session_state["venue_mode"] = "Select existing venue"

if "selected_venue_id" not in st.session_state:
    st.session_state["selected_venue_id"] = None

venue_map = {v["id"]: v for v in venues if v.get("id")}
venue_labels = []
venue_ids = []
for v in venues:
    if v.get("id"):
        venue_ids.append(v["id"])
        venue_labels.append(_venue_label(v))


# ===============================
# Layout
# ===============================
left, right = st.columns([1.1, 1])

with left:
    st.subheader("1) Event (optional)")
    selected_event_label = st.selectbox(
        "Link to an existing event (optional)",
        options=[o["label"] for o in event_options],
        index=0,
    )
    selected_event_obj = next(o for o in event_options if o["label"] == selected_event_label)
    selected_event_id = selected_event_obj.get("id")

    # Pre-fill meeting fields if event chosen
    event_row = selected_event_obj.get("row", {}) if selected_event_id else {}

    meeting_name_default = event_row.get("subject") if selected_event_id else ""
    client_name_default = event_row.get("client") if selected_event_id else ""
    mm_name_default = event_row.get("meeting_manager_name") if selected_event_id else ""
    mm_email_default = event_row.get("meeting_manager_email") if selected_event_id else ""

    st.subheader("2) Meeting info")

    meeting_name = st.text_input("Meeting Name *", value=meeting_name_default or "")

    # ---- Client dropdown (with Other...) ----
    client_options = load_clients_dropdown()
    client_pick = st.selectbox(
        "Client (optional)",
        options=(client_options + ["Other…"]) if client_options else ["Other…"],
        index=(client_options.index(client_name_default) if client_name_default in client_options else 0),
        key="av_client_pick",
    )

    client_other = ""
    if client_pick == "Other…":
        client_other = st.text_input("Enter new client name", value=client_name_default or "", key="av_client_other")

    client_name = (client_other.strip() if client_pick == "Other…" else (client_pick or "").strip()) or None


    # ---- Meeting Manager dropdown (with Other...) ----
    managers = load_meeting_managers_dropdown()  # list of (name, email)
    manager_labels = [f"{n} <{e}>" if e else n for n, e in managers]

    # if event prefilled a manager, try to match it
    default_manager_idx = 0
    if selected_event_id and (mm_name_default or mm_email_default):
        for i, (n, e) in enumerate(managers):
            if mm_email_default and e and (e.lower() == mm_email_default.lower()):
                default_manager_idx = i
                break
            if mm_name_default and n and (n.lower() == mm_name_default.lower()):
                default_manager_idx = i
                break

    mm_pick = st.selectbox(
        "Meeting Manager (optional)",
        options=(manager_labels + ["Other…"]) if manager_labels else ["Other…"],
        index=default_manager_idx if manager_labels else 0,
        key="av_mm_pick",
    )

    if mm_pick == "Other…":
        c_mm1, c_mm2 = st.columns(2)
        meeting_manager_name = c_mm1.text_input("Meeting Manager Name", value=mm_name_default or "", key="av_mm_name_other")
        meeting_manager_email = c_mm2.text_input("Meeting Manager Email", value=mm_email_default or "", key="av_mm_email_other")
        meeting_manager_name = meeting_manager_name.strip() or None
        meeting_manager_email = meeting_manager_email.strip() or None
    else:
        picked_idx = manager_labels.index(mm_pick)
        picked_name, picked_email = managers[picked_idx]
        meeting_manager_name = (picked_name or "").strip() or None
        meeting_manager_email = (picked_email or "").strip() or None

    # ✅ New: Required operational date
    st.subheader("2a) Dates")
    deliver_by_date = st.date_input("Equipment needed on-site by *")

    # ✅ Optional event date range (not required to submit)
    use_event_dates = st.checkbox("Add event date range (optional)", value=False)
    event_start_date = None
    event_end_date = None

    if use_event_dates:
        c_ed1, c_ed2 = st.columns(2)
        with c_ed1:
            event_start_date = st.date_input("Event start date")
        with c_ed2:
            event_end_date = st.date_input("Event end date")

    st.divider()

    st.subheader("3) Venue (required)")

    venue_mode = st.radio(
        "Select venue option",
        options=["Select existing venue", "Add new venue"],
        horizontal=True,
        key="venue_mode",
    )

    selected_venue_id = None
    venue_name = None

    if venue_mode == "Select existing venue":
        if not venue_ids:
            st.warning("No venues found. Please add a venue.")
        else:
            # Default to whatever we just created (if any)
            default_idx = 0
            if st.session_state.get("selected_venue_id") in venue_ids:
                default_idx = venue_ids.index(st.session_state["selected_venue_id"])

            venue_choice = st.selectbox(
                "Venue *",
                options=venue_labels,
                index=default_idx,
            )

            selected_venue_id = venue_ids[venue_labels.index(venue_choice)]
            st.session_state["selected_venue_id"] = selected_venue_id
            
            venue_name = venue_choice 

            # ✅ Clear any leftover "new venue" ship-to prefill so it doesn't carry over
            st.session_state.pop("ship_to_address_prefill", None)

    else:
        st.caption("Add the venue once — all apps can re-use it later.")
        with st.form("add_venue_form"):
            v_name = st.text_input("Venue Name *")
            col1, col2 = st.columns(2)
            with col1:
                v_addr1 = st.text_input("Address Line 1")
                v_city = st.text_input("City")
                v_postal = st.text_input("Postal Code")
            with col2:
                v_addr2 = st.text_input("Address Line 2")
                v_state = st.text_input("State")
                v_country = st.text_input("Country", value="US")

            col3, col4 = st.columns(2)
            with col3:
                v_contact_name = st.text_input("Main Contact Name")
                v_contact_email = st.text_input("Main Contact Email")
            with col4:
                v_contact_phone = st.text_input("Main Contact Phone")
                v_notes = st.text_area("Venue notes")

            create_venue_btn = st.form_submit_button("Create venue")

        if create_venue_btn:
            if not v_name.strip():
                st.error("Venue Name is required.")
            else:
                payload = {
                    "name": v_name.strip(),
                    "address_line1": v_addr1.strip() or None,
                    "address_line2": v_addr2.strip() or None,
                    "city": v_city.strip() or None,
                    "state": v_state.strip() or None,
                    "postal_code": v_postal.strip() or None,
                    "country": v_country.strip() or "US",
                    "main_contact_name": v_contact_name.strip() or None,
                    "main_contact_email": v_contact_email.strip() or None,
                    "main_contact_phone": v_contact_phone.strip() or None,
                    "notes": v_notes.strip() or None,
                }

                ins = sb.table("venues").insert(payload).execute()
                new_venue_created = (ins.data or [None])[0]

                if not (new_venue_created and new_venue_created.get("id")):
                    st.error("Venue creation failed (no venue returned).")
                else:
                    # ✅ Auto-select the newly created venue and switch back to selection mode
                    st.session_state["selected_venue_id"] = new_venue_created["id"]
                    
                    venue_name = new_venue_created.get("name") or v_name.strip()

                    # Pre-fill ship-to from this venue immediately
                    st.session_state["ship_to_address_prefill"] = _build_venue_ship_to_text(new_venue_created)

                    st.success("Venue created. Ship-to address prefilled.")
                    st.cache_data.clear()
                    st.rerun()

if not selected_venue_id and st.session_state.get("selected_venue_id"):
    selected_venue_id = st.session_state["selected_venue_id"]

with right:
    st.subheader("4) Shipping + Instructions")

    # If a venue is selected, allow "Ship to venue address"
    ship_to_venue = False
    venue_ship_to_text = ""

    if venue_mode == "Select existing venue" and selected_venue_id and selected_venue_id in venue_map:
        v = venue_map[selected_venue_id]
        venue_ship_to_text = _build_venue_ship_to_text(v)

        ship_to_venue = st.checkbox("Ship to venue address", value=True)

    default_ship_to = st.session_state.get("ship_to_address_prefill", "")

    ship_to_value = venue_ship_to_text if ship_to_venue else (default_ship_to if venue_mode == "Add new venue" else "")

    ship_to_address = st.text_area(
        "Ship To Address *",
        value=ship_to_value,
        height=110,
        placeholder="Enter the full shipping address (hotel receiving dock, staff, etc.)",
    )

    recipient_name = st.text_input("Recipient Name (optional)", value="")
    hotel_contact_email = st.text_input("Hotel / Receiving Contact Email (optional)")
    hotel_contact_phone = st.text_input("Hotel / Receiving Contact Phone (optional)")
    special_instructions = st.text_area("Special Instructions (optional)", height=110)

    st.divider()

    st.subheader("5) Equipment Requested")

    # Minimal equipment types for first pass — adjust as needed
    equipment_types = ["laptop", "projector", "owl"]

    eq_rows = []
    for et in equipment_types:
        c1, c2 = st.columns([2, 1])
        with c1:
            st.write(et.capitalize())
        with c2:
            qty = st.number_input(
                f"Qty ({et})",
                min_value=0,
                step=1,
                value=0,
                key=f"qty_{et}",
            )
        eq_rows.append({"equipment_type": et, "quantity": int(qty)})

    st.divider()

    st.subheader("6) Submit")

    # Basic validation summary (helps users understand why submit is disabled)
    problems = []
    if not meeting_name.strip():
        problems.append("Meeting Name is required.")
    if not deliver_by_date:
        problems.append("Equipment needed on-site by date is required.")
    if use_event_dates and event_start_date and event_end_date and event_end_date < event_start_date:
        problems.append("Event end date cannot be before event start date.")
    if not selected_venue_id:
        problems.append("Venue is required.")
    if not ship_to_address.strip():
        problems.append("Ship To Address is required.")
    any_qty = any(r["quantity"] > 0 for r in eq_rows)
    if not any_qty:
        problems.append("At least one equipment quantity must be > 0.")

    if problems:
        st.warning("Fix before submitting:\n- " + "\n- ".join(problems))

    submit_disabled = len(problems) > 0

    if st.button("Submit AV Request", disabled=submit_disabled, type="primary"):
        try:
            av_request_payload = {
                "event_id": selected_event_id,  # may be None
                "meeting_name": meeting_name.strip(),
                "client_name": (client_name.strip() if client_name else None),
                "meeting_manager_name": (meeting_manager_name.strip() if meeting_manager_name else None),
                "meeting_manager_email": (meeting_manager_email.strip() if meeting_manager_email else None),
                "recipient_name": recipient_name.strip() or None,
                "ship_to_address": ship_to_address.strip(),
                "hotel_contact_email": hotel_contact_email.strip() or None,
                "hotel_contact_phone": hotel_contact_phone.strip() or None,
                "special_instructions": special_instructions.strip() or None,
                "status": "requested",
                "created_at": _now_utc_iso(),
                "updated_at": _now_utc_iso(),
                "venue_id": selected_venue_id,
                # ✅ New required + optional date fields
                "deliver_by_date": _date_to_iso(deliver_by_date),
                "event_start_date": _date_to_iso(event_start_date) if use_event_dates else None,
                "event_end_date": _date_to_iso(event_end_date) if use_event_dates else None,
            }

            ins_req = sb.table("av_requests").insert(av_request_payload).execute()
            av_req_row = (ins_req.data or [None])[0]
            if not av_req_row or not av_req_row.get("id"):
                st.error("AV request insert failed (no id returned).")
                st.stop()

            av_request_id = av_req_row["id"]

            # Insert line items
            items_to_insert = [
                {"av_request_id": av_request_id, "equipment_type": r["equipment_type"], "quantity": r["quantity"]}
                for r in eq_rows
                if r["quantity"] > 0
            ]

            if items_to_insert:
                sb.table("av_request_items").insert(items_to_insert).execute()

            # Persist new dropdown values (best-effort)
            try:
                if client_pick == "Other…" and client_name:
                    sb.table("clients").upsert({"name": client_name}).execute()
            except Exception:
                pass

            try:
                if mm_pick == "Other…" and meeting_manager_name and meeting_manager_email:
                    sb.table("meeting_managers").upsert(
                        {"name": meeting_manager_name, "email": meeting_manager_email},
                        on_conflict="email"
                    ).execute()
            except Exception:
                pass

            st.cache_data.clear()

            # ===============================
            # Phase 1 Step 1 — Email routing + confirmation send
            # ===============================
            RAY_EMAIL = "ray@lutinemanagement.com"

            # LIVE (later): support@redeye.tech
            REDEYE_SUPPORT_EMAIL = "rjs2119@gmail.com"  # TESTING ONLY

            qty_map = {r["equipment_type"]: int(r["quantity"]) for r in eq_rows}
            laptop_qty = qty_map.get("laptop", 0)
            owl_qty = qty_map.get("owl", 0)
            projector_qty = qty_map.get("projector", 0)

            # Route internal owner
            to_emails = []
            cc_emails = []

            if laptop_qty > 0:
                to_emails = [REDEYE_SUPPORT_EMAIL]
                cc_emails = [RAY_EMAIL]  # only CC Ray when Redeye is TO
            elif (owl_qty > 0) or (projector_qty > 0):
                to_emails = [RAY_EMAIL]
            else:
                to_emails = [RAY_EMAIL]

            # Always CC meeting manager/requester if available
            if meeting_manager_email:
                cc_emails.append(meeting_manager_email)

            # Build items list
            items_lines = []
            for et in ["laptop", "projector", "owl"]:
                q = qty_map.get(et, 0)
                if q > 0:
                    items_lines.append(f"<li><b>{et.capitalize()}</b>: {q}</li>")
            items_html = "<ul>" + "".join(items_lines) + "</ul>" if items_lines else "<i>No items</i>"

            ship_to_html = "<br>".join([ln for ln in (ship_to_address or "").splitlines() if ln.strip()])

            subject = f"LUTINE AV Request Submitted — {meeting_name.strip()} (Deliver by {_date_to_iso(deliver_by_date)})"

            body_html = f"""
            <div style="font-family:Segoe UI, Arial, sans-serif; font-size:11pt;">
              <h3 style="margin:0 0 8px 0;">New AV Request Submitted</h3>

              <p style="margin:0 0 8px 0;"><b>Meeting:</b> {meeting_name.strip()}</p>
              <p style="margin:0 0 8px 0;"><b>Client:</b> {(client_name or '—')}</p>
              <p style="margin:0 0 8px 0;"><b>Deliver-by:</b> {deliver_by_date}</p>

              <p style="margin:0 0 8px 0;"><b>Venue:</b> {(venue_name or '—')}</p>
              <p style="margin:0 0 8px 0;"><b>Ship To:</b><br>{ship_to_html}</p>

              <p style="margin:0 0 8px 0;"><b>Requested items:</b></p>
              {items_html}

              <p style="margin:8px 0 0 0;"><b>Request ID:</b> {av_request_id}</p>
              <p style="margin:8px 0 0 0;"><b>Request detail:</b> (link coming soon)</p>
            </div>
            """

            try:
                g = st.secrets["graph"]
                token = _get_graph_token()
                _graph_send_mail(
                    token,
                    g["shared_mailbox_upn"],
                    to_emails=to_emails,
                    cc_emails=cc_emails,
                    subject=subject,
                    body_html=body_html,
                )
                st.info("Confirmation email sent.")
            except Exception as e:
                st.warning(f"Request saved, but email failed to send: {e}")

            st.success("AV request submitted.")
            st.info(f"Request ID: {av_request_id}")

        except Exception as e:
            st.error(f"Submit failed: {e}")
