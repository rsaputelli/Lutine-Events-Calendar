# pages/02_AV_Request_Detail.py

import streamlit as st
import pandas as pd

from lib.supabase_client import get_supabase

st.title("AV Request Detail")
st.caption("Read-only view of an AV request (single source of truth).")

sb = get_supabase()

# -----------------------------
# Helpers
# -----------------------------
def _safe(v, fallback="—"):
    if v is None:
        return fallback
    if isinstance(v, str) and not v.strip():
        return fallback
    return v


def _fmt_date(d):
    if not d:
        return "—"
    try:
        return pd.to_datetime(d).date().isoformat()
    except Exception:
        return str(d)


def _fmt_dt(ts):
    if not ts:
        return "—"
    try:
        dt = pd.to_datetime(ts, utc=True)
        return dt.strftime("%Y-%m-%d %I:%M %p UTC")
    except Exception:
        return str(ts)


def _venue_label(v: dict) -> str:
    if not v:
        return "—"
    name = v.get("name") or "Unnamed Venue"
    city = v.get("city") or ""
    state = v.get("state") or ""
    parts = [p for p in [city, state] if p]
    if parts:
        return f"{name} ({', '.join(parts)})"
    return name


# -----------------------------
# Get request_id (query param or manual entry)
# -----------------------------
qp = st.query_params
qp_request_id = qp.get("request_id")

request_id = st.text_input(
    "Request ID",
    value=(qp_request_id or ""),
    placeholder="Paste a request UUID (e.g., f3e3e33b-17bf-41ef-939b-c90eefe62ce8)",
)

if not request_id.strip():
    st.info("Enter a Request ID to view details.")
    st.stop()

request_id = request_id.strip()

# -----------------------------
# Load request
# -----------------------------
try:
    req_res = (
        sb.table("av_requests")
        .select("*")
        .eq("id", request_id)
        .limit(1)
        .execute()
    )
    req = (req_res.data or [None])[0]
except Exception as e:
    st.error(f"Could not load request: {e}")
    st.stop()

if not req:
    st.warning("No request found with that ID.")
    st.stop()

# -----------------------------
# Load related items
# -----------------------------
try:
    items_res = (
        sb.table("av_request_items")
        .select("*")
        .eq("av_request_id", request_id)
        .order("equipment_type", desc=False)
        .execute()
    )
    items = items_res.data or []
except Exception as e:
    items = []
    st.warning(f"Could not load request items: {e}")

# -----------------------------
# Load shipments (if table exists)
# -----------------------------
shipments = []
try:
    ship_res = (
        sb.table("av_shipments")
        .select("*")
        .eq("av_request_id", request_id)
        .order("ship_date", desc=False)
        .execute()
    )
    shipments = ship_res.data or []
except Exception:
    shipments = []

# -----------------------------
# Load allocations (if table exists)
# -----------------------------
allocs = []
try:
    alloc_res = (
        sb.table("av_allocations")
        .select("*")
        .eq("av_request_id", request_id)
        .order("allocated_at", desc=False)
        .execute()
    )
    allocs = alloc_res.data or []
except Exception:
    allocs = []

# -----------------------------
# Try to enrich allocations with equipment_inventory details (optional)
# -----------------------------
equipment_by_id = {}
if allocs:
    equipment_ids = sorted({a.get("equipment_id") for a in allocs if a.get("equipment_id")})
    if equipment_ids:
        try:
            eq_res = (
                sb.table("equipment_inventory")
                .select("*")
                .in_("id", equipment_ids)
                .execute()
            )
            eq_rows = eq_res.data or []
            equipment_by_id = {r.get("id"): r for r in eq_rows if r.get("id")}
        except Exception:
            equipment_by_id = {}

# -----------------------------
# Load venue (optional)
# -----------------------------
venue = None
venue_id = req.get("venue_id")
if venue_id:
    try:
        venue_res = (
            sb.table("venues")
            .select("*")
            .eq("id", venue_id)
            .limit(1)
            .execute()
        )
        venue = (venue_res.data or [None])[0]
    except Exception:
        venue = None

# -----------------------------
# Header summary
# -----------------------------
top_left, top_right = st.columns([1.4, 1])

with top_left:
    st.subheader(_safe(req.get("meeting_name"), "Untitled Request"))

    st.write(f"**Status:** {_safe(req.get('status'))}")
    st.write(f"**Client:** {_safe(req.get('client_name'))}")

    mm_name = req.get("meeting_manager_name")
    mm_email = req.get("meeting_manager_email")
    if mm_name or mm_email:
        mm_line = f"{_safe(mm_name)}"
        if mm_email:
            mm_line += f"  <{mm_email}>"
        st.write(f"**Meeting Manager:** {mm_line}")

with top_right:
    st.write(f"**Request ID:** `{request_id}`")
    st.write(f"**Created:** {_fmt_dt(req.get('created_at'))}")
    st.write(f"**Updated:** {_fmt_dt(req.get('updated_at'))}")

st.divider()

# -----------------------------
# Dates / Shipping / Venue
# -----------------------------
c1, c2 = st.columns([1, 1])

with c1:
    st.subheader("Dates")
    st.write(f"**Deliver-by:** {_fmt_date(req.get('deliver_by_date'))}")

    es = req.get("event_start_date")
    ee = req.get("event_end_date")
    if es or ee:
        st.write(f"**Event range:** {_fmt_date(es)} → {_fmt_date(ee)}")
    else:
        st.write("**Event range:** —")

    if req.get("event_id"):
        st.write(f"**Linked Event ID:** `{req.get('event_id')}`")

with c2:
    st.subheader("Venue + Ship To")

    st.write(f"**Venue:** {_venue_label(venue) if venue else '—'}")

    st.write("**Ship To Address:**")
    st.code(_safe(req.get("ship_to_address")), language="text")

    recipient_name = req.get("recipient_name")
    if recipient_name:
        st.write(f"**Recipient:** {recipient_name}")

    h_email = req.get("hotel_contact_email")
    h_phone = req.get("hotel_contact_phone")
    if h_email or h_phone:
        st.write("**Hotel / Receiving Contact:**")
        if h_email:
            st.write(f"- Email: {h_email}")
        if h_phone:
            st.write(f"- Phone: {h_phone}")

st.divider()

# -----------------------------
# Requested Items
# -----------------------------
st.subheader("Requested Items")

if not items:
    st.info("No requested items found.")
else:
    df_items = pd.DataFrame(items)

    preferred = ["equipment_type", "quantity"]
    cols = [c for c in preferred if c in df_items.columns] + [c for c in df_items.columns if c not in preferred]
    df_items = df_items[cols] if cols else df_items

    st.dataframe(df_items, use_container_width=True, hide_index=True)

st.divider()

# -----------------------------
# Shipments / Tracking
# -----------------------------
st.subheader("Shipments / Tracking (if any)")

if not shipments:
    st.caption("No shipments recorded yet.")
else:
    df_ship = pd.DataFrame(shipments)

    preferred = [
        "ship_direction",
        "carrier",
        "tracking_number",
        "ship_date",
        "shipped_by",
        "ship_from_location",
        "ship_to_location",
        "received_confirmed",
        "received_at",
        "notes",
    ]
    cols = [c for c in preferred if c in df_ship.columns] + [c for c in df_ship.columns if c not in preferred]
    df_ship = df_ship[cols] if cols else df_ship

    st.dataframe(df_ship, use_container_width=True, hide_index=True)

st.divider()

# -----------------------------
# Special Instructions
# -----------------------------
st.subheader("Special Instructions")
special = req.get("special_instructions")
if special and str(special).strip():
    st.write(special)
else:
    st.caption("—")

st.divider()

# -----------------------------
# Allocations (read-only)
# -----------------------------
st.subheader("Assigned Equipment (if any)")

if not allocs:
    st.caption("No allocations recorded yet.")
else:
    # Enrich allocations for readability if possible
    pretty_rows = []
    for a in allocs:
        eq = equipment_by_id.get(a.get("equipment_id")) or {}
        pretty_rows.append(
            {
                "allocated_at": a.get("allocated_at"),
                "equipment_id": a.get("equipment_id"),
                "equipment_name": eq.get("name") or eq.get("equipment_name") or "",
                "asset_tag": eq.get("asset_tag") or "",
                "serial_number": eq.get("serial_number") or "",
                "allocated_from": a.get("allocated_from"),
                "allocated_to": a.get("allocated_to"),
            }
        )

    df_alloc = pd.DataFrame(pretty_rows)

    # Format timestamps (soft)
    for col in ["allocated_at", "allocated_from", "allocated_to"]:
        if col in df_alloc.columns:
            df_alloc[col] = df_alloc[col].apply(_fmt_dt)

    st.dataframe(df_alloc, use_container_width=True, hide_index=True)
