# pages/01_AV_Intake.py
import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timezone, date, time, timedelta
from zoneinfo import ZoneInfo
import html

from lib.supabase_client import get_supabase

# ---- header/logo ----
logo_col, title_col = st.columns([1.4, 5.6])
with logo_col:
    st.image("assets/lutine-logo.png", width=170)
with title_col:
    st.title("AV Request Intake")
    st.caption("Submit an AV request for a scheduled meeting.")

st.info(
    "Best Practice: AV requests should normally be tied to an existing Calendar event. "
    "If the meeting is not yet on the calendar, please create it in the Calendar app first. "
    "If you submit this form without selecting an event, the system will automatically create a basic calendar event and notify the meeting manager to complete it."
)
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
# Calendar auto-create helpers
# ===============================
TZ_WINDOWS = "Eastern Standard Time"
TZ_IANA = "America/New_York"

def _graph_create_event(token: str, shared_mailbox_upn: str, payload: dict) -> dict:
    url = f"https://graph.microsoft.com/v1.0/users/{shared_mailbox_upn}/calendar/events"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    r = requests.post(url, headers=headers, json=payload, timeout=20)
    if r.status_code >= 400:
        raise RuntimeError(f"Graph create event {r.status_code}: {r.text}")
    return r.json()

def _build_graph_event_payload(
    *,
    subject: str,
    body_html: str,
    start_dt,
    end_dt,
    is_all_day: bool,
    location_str: str | None,
    reminder_minutes: int = 30,
) -> dict:
    payload = {
        "subject": subject,
        "isReminderOn": True,
        "reminderMinutesBeforeStart": int(reminder_minutes),
        "body": {"contentType": "HTML", "content": body_html},
        "showAs": "free",
    }

    if is_all_day:
        start_date = start_dt if isinstance(start_dt, date) and not isinstance(start_dt, datetime) else start_dt.date()
        end_date = end_dt if isinstance(end_dt, date) and not isinstance(end_dt, datetime) else end_dt.date()
        end_exclusive = max(end_date, start_date) + timedelta(days=1)

        payload.update({
            "isAllDay": True,
            "start": {"dateTime": start_date.isoformat(), "timeZone": TZ_WINDOWS},
            "end": {"dateTime": end_exclusive.isoformat(), "timeZone": TZ_WINDOWS},
        })
    else:
        if getattr(start_dt, "tzinfo", None) is None or getattr(end_dt, "tzinfo", None) is None:
            raise ValueError("Timed event datetimes must be timezone-aware.")
        payload.update({
            "start": {"dateTime": start_dt.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": TZ_WINDOWS},
            "end": {"dateTime": end_dt.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": TZ_WINDOWS},
        })

    if location_str:
        payload["location"] = {"displayName": location_str}

    return payload

def _create_basic_calendar_event_from_av(
    *,
    meeting_name: str,
    client_name: str | None,
    meeting_manager_name: str | None,
    meeting_manager_email: str | None,
    venue_name: str | None,
    special_instructions: str | None,
    deliver_by_date,
    event_start_date,
    event_end_date,
):
    """
    Creates the Outlook event first, then saves the linked row in public.events.
    Returns: (local_event_id, outlook_event_id, auto_created_is_all_day)
    """
    g = st.secrets["graph"]
    tz = ZoneInfo(TZ_IANA)

    # If event dates were entered, use them as an all-day event span.
    # Otherwise create a 1-hour placeholder on the deliver-by date at 9:00 AM ET.
    if event_start_date:
        is_all_day = True
        start_local = event_start_date
        end_local = event_end_date or event_start_date
        start_dt_utc = datetime.combine(start_local, time(0, 0)).replace(tzinfo=tz).astimezone(ZoneInfo("UTC"))
        end_dt_utc = datetime.combine(max(end_local, start_local) + timedelta(days=1), time(0, 0)).replace(tzinfo=tz).astimezone(ZoneInfo("UTC"))
    else:
        is_all_day = False
        start_local = datetime.combine(deliver_by_date, time(9, 0)).replace(tzinfo=tz)
        end_local = start_local + timedelta(hours=1)
        start_dt_utc = start_local.astimezone(ZoneInfo("UTC"))
        end_dt_utc = end_local.astimezone(ZoneInfo("UTC"))

    safe_meeting = html.escape(meeting_name)
    safe_client = html.escape(client_name or "")
    safe_manager = html.escape(meeting_manager_name or "")
    safe_manager_email = html.escape(meeting_manager_email or "")
    safe_venue = html.escape(venue_name or "")
    safe_notes = html.escape(special_instructions or "").replace("\n", "<br>")

    body_html = f"""
    <div style="font-family:Segoe UI, Arial, sans-serif; font-size:11pt;">
      <p><b>This event was auto-created from AV Intake.</b></p>
      <p>Please review and complete the event details in the Calendar app.</p>
      <p><b>Meeting:</b> {safe_meeting}</p>
      {f"<p><b>Client:</b> {safe_client}</p>" if safe_client else ""}
      {f"<p><b>Location:</b> {safe_venue}</p>" if safe_venue else ""}
      {f"<p><b>Meeting Manager:</b> {safe_manager} &lt;{safe_manager_email}&gt;</p>" if safe_manager or safe_manager_email else ""}
      {f"<p><b>AV Notes:</b><br>{safe_notes}</p>" if safe_notes else ""}
    </div>
    """

    payload = _build_graph_event_payload(
        subject=meeting_name,
        body_html=body_html,
        start_dt=start_local,
        end_dt=end_local,
        is_all_day=is_all_day,
        location_str=venue_name,
        reminder_minutes=30,
    )

    token = _get_graph_token()
    created = _graph_create_event(token, g["shared_mailbox_upn"], payload)
    outlook_event_id = (created or {}).get("id")
    if not outlook_event_id:
        raise RuntimeError("Outlook event was not created successfully.")

    event_row = {
        "subject": meeting_name,
        "client": client_name or None,
        "start_dt_utc": start_dt_utc.isoformat(),
        "end_dt_utc": end_dt_utc.isoformat(),
        "timezone_display": TZ_IANA,
        "is_all_day": bool(is_all_day),
        "location": venue_name or None,
        "event_type": "in_person",
        "virtual_provider": None,
        "virtual_link": None,
        "meeting_manager_name": meeting_manager_name or None,
        "meeting_manager_email": meeting_manager_email or None,
        "reminder_minutes": 30,
        "outlook_event_id": outlook_event_id,
        "accreditation_required": False,
        "created_at": _now_utc_iso(),
        "outlook_body_html": body_html,
    }

    ins = sb.table("events").insert(event_row).execute()
    inserted = (ins.data or [None])[0]
    if not inserted or not inserted.get("id"):
        raise RuntimeError("Local events row insert failed after Outlook event creation.")

    return inserted["id"], outlook_event_id, is_all_day
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

event_options = [{"id": None, "label": "— No event selected (system will create a basic calendar event) —"}]
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

# Prevent duplicate submissions
if "av_submit_lock" not in st.session_state:
    st.session_state["av_submit_lock"] = False
    
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
    if selected_event_id is None:
        st.warning(
            "No calendar event is currently linked. "
            "If you submit this request as-is, the system will create a basic master-calendar event automatically."
        )

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
    equipment_types = ["laptop", "projector", "owl", "badge printer"]

    eq_rows = []
    for et in equipment_types:
        c1, c2 = st.columns([2, 1])
        with c1:
            display_name = {
                "laptop": "Laptop",
                "projector": "Projector",
                "owl": "Meeting Owl",
                "badge_printer": "Badge Printer",
            }

            st.write(display_name.get(et, et.capitalize()))
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
    st.caption("Please click submit once and wait for confirmation.")
    if st.button("Submit AV Request", disabled=submit_disabled, type="primary"):

        if st.session_state.get("av_submit_lock"):
            st.warning("Submission already in progress. Please wait.")
            st.stop()

        st.session_state["av_submit_lock"] = True

        try:
            auto_created_event = False
            auto_created_outlook_event_id = None

            if selected_event_id is None:
                selected_event_id, auto_created_outlook_event_id, auto_created_is_all_day = _create_basic_calendar_event_from_av(
                    meeting_name=meeting_name.strip(),
                    client_name=(client_name.strip() if client_name else None),
                    meeting_manager_name=(meeting_manager_name.strip() if meeting_manager_name else None),
                    meeting_manager_email=(meeting_manager_email.strip() if meeting_manager_email else None),
                    venue_name=(venue_name or None),
                    special_instructions=special_instructions.strip() or None,
                    deliver_by_date=deliver_by_date,
                    event_start_date=event_start_date if use_event_dates else None,
                    event_end_date=event_end_date if use_event_dates else None,
                )
                auto_created_event = True
            av_request_payload = {
                "event_id": selected_event_id,
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
                raise RuntimeError("AV request insert failed (no id returned).")

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
            if auto_created_event and meeting_manager_email:
                try:
                    g = st.secrets["graph"]
                    token = _get_graph_token()

                    mm_subject = f"Calendar event needs completion — {meeting_name.strip()}"
                    mm_body_html = f"""
                    <div style="font-family:Segoe UI, Arial, sans-serif; font-size:11pt;">
                      <p>A calendar event was automatically created from AV Intake because no event was selected.</p>
                      <p><b>Meeting:</b> {meeting_name.strip()}</p>
                      <p><b>AV Request ID:</b> {av_request_id}</p>
                      <p>Please open the Calendar app and complete the event details.</p>
                    </div>
                    """

                    _graph_send_mail(
                        token,
                        g["shared_mailbox_upn"],
                        to_emails=[meeting_manager_email],
                        cc_emails=[],
                        subject=mm_subject,
                        body_html=mm_body_html,
                    )
                except Exception as e:
                    st.warning(f"Event created but meeting manager notice failed: {e}")
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
            badge_printer_qty = qty_map.get("badge_printer", 0)

            # Route internal owner
            to_emails = []
            cc_emails = []

            if laptop_qty > 0:
                to_emails = [REDEYE_SUPPORT_EMAIL]
                cc_emails = [RAY_EMAIL]  # only CC Ray when Redeye is TO
            elif (owl_qty > 0) or (projector_qty > 0) or (badge_printer_qty > 0):
                to_emails = [RAY_EMAIL]
            else:
                to_emails = [RAY_EMAIL]

            # Always CC meeting manager/requester if available
            if meeting_manager_email:
                cc_emails.append(meeting_manager_email)

            # Build items list
            items_lines = []
            for et in ["laptop", "projector", "owl", "badge printer"]:
                q = qty_map.get(et, 0)
                if q > 0:
                    email_display_name = {
                        "laptop": "Laptop",
                        "projector": "Projector",
                        "owl": "Meeting Owl",
                        "badge_printer": "Badge Printer",
                    }
                    items_lines.append(f"<li><b>{email_display_name.get(et, et.capitalize())}</b>: {q}</li>")
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
              <p style="margin:0 0 8px 0; font-size:10pt; color:#555;">
              <i>Note: Laptop requests are fulfilled by Redeye. 
              All other AV equipment (projectors, Owls, badge printers) are fulfilled by Lutine.</i>
              </p>
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

            if auto_created_event:
                st.success("AV request submitted and a basic calendar event was created automatically.")
            else:
                st.success("AV request submitted.")

            st.info(f"Request ID: {av_request_id}")
            st.session_state["av_submit_lock"] = False

        except Exception as e:
            st.session_state["av_submit_lock"] = False
            st.error(f"Submit failed: {e}")
