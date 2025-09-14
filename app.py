
import streamlit as st
import pandas as pd
from dateutil.parser import parse as dparse
from datetime import datetime, timedelta, date, time
from pathlib import Path

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "data" / "db.xlsx"

st.set_page_config(page_title="Men's Shed Scheduler", page_icon="🪚", layout="wide")

# === Minimal header/logo injection (no theme override) ===
def _inject_min_css():
    css_path = BASE_DIR / "assets" / "styles.css"
    if css_path.exists():
        st.markdown(f"<style>{css_path.read_text()}</style>", unsafe_allow_html=True)

def _brand_header():
    # Find a logo even if case or number differs
    candidates = [
        BASE_DIR / "assets" / "logo.png", BASE_DIR / "assets" / "Logo.png",
        BASE_DIR / "assets" / "logo1.png", BASE_DIR / "assets" / "logo2.png", BASE_DIR / "assets" / "logo3.png",
        BASE_DIR / "Assets" / "logo1.png", BASE_DIR / "Assets" / "logo2.png", BASE_DIR / "Assets" / "logo3.png",
        BASE_DIR / "assets" / "logo.svg", BASE_DIR / "assets" / "Logo.svg",
        BASE_DIR / "Assets" / "logo.svg", BASE_DIR / "Assets" / "Logo.svg",
        BASE_DIR / "assets" / "logo.jpg", BASE_DIR / "assets" / "Logo.jpg",
        BASE_DIR / "Assets" / "logo.jpg", BASE_DIR / "Assets" / "Logo.jpg",
    ]
    logo_path = next((p for p in candidates if p.exists()), None)
    if logo_path:
        st.markdown(f'<div class="header"><img src="{logo_path.as_posix()}" alt="Woodturners of the Hunter"></div>', unsafe_allow_html=True)
    else:
        st.markdown("### Woodturners of the Hunter")

_inject_min_css()
_brand_header()
# === End minimal injection ===

# ---------------- Utility & settings ----------------
def load_db():
    try:
        import openpyxl
    except ImportError:
        st.error("Missing dependency 'openpyxl'. Add it to requirements.txt and redeploy.")
        st.stop()
    if not DB_PATH.exists():
        st.error("Database not found at data/db.xlsx")
        st.stop()
    xls = pd.ExcelFile(DB_PATH, engine="openpyxl")
    sheets = {name: pd.read_excel(DB_PATH, engine="openpyxl", sheet_name=name) for name in xls.sheet_names}
    for e in ["Users","Licences","UserLicences","Machines","Bookings","OperatingLog","Issues","ServiceLog","OperatingHours","ClosedDates","Settings"]:
        sheets.setdefault(e, pd.DataFrame())
    # Schema defaults
    if "max_duration_hours" not in sheets["Machines"].columns:
        sheets["Machines"]["max_duration_hours"] = 4.0
    if "category" not in sheets["Bookings"].columns and not sheets["Bookings"].empty:
        sheets["Bookings"]["category"] = "Usage"
    for col in ["phone","email","address"]:
        if col not in sheets["Users"].columns:
            sheets["Users"][col] = ""
    return sheets

def save_db(sheets):
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(DB_PATH, engine="openpyxl", mode="w") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)

def next_id(df, id_col):
    if df.empty or id_col not in df.columns:
        return 1
    return int(pd.to_numeric(df[id_col], errors="coerce").fillna(0).max()) + 1

def overlap(a_start, a_end, b_start, b_end):
    return (a_start < b_end) and (a_end > b_start)

def licence_valid_for_user(sheets, user_id:int, licence_id:int, on_date=None):
    if on_date is None:
        on_date = pd.Timestamp.today().normalize()
    UL = sheets["UserLicences"].copy()
    if UL.empty: return False
    UL["start_date"] = pd.to_datetime(UL.get("start_date", pd.NaT), errors="coerce")
    UL["end_date"] = pd.to_datetime(UL.get("end_date", pd.NaT), errors="coerce")
    rows = UL[(UL["user_id"]==user_id) & (UL["licence_id"]==licence_id)]
    if rows.empty: return False
    def _ok(r):
        s = r.get("start_date"); e = r.get("end_date")
        cond_s = (pd.isna(s) or on_date >= s.normalize())
        cond_e = (pd.isna(e) or on_date <= e.normalize())
        return bool(cond_s and cond_e)
    return any(_ok(r) for _, r in rows.iterrows())

def user_licence_ids(sheets, user_id:int):
    ul = sheets["UserLicences"]
    if ul.empty: return set()
    return set(ul.loc[ul["user_id"]==user_id, "licence_id"].astype(int).tolist())

def machine_required_licence_id(sheets, machine_id:int):
    m = sheets["Machines"]
    row = m.loc[m["machine_id"]==machine_id]
    if row.empty: return None
    return int(row.iloc[0]["required_licence_id"]) if not pd.isna(row.iloc[0]["required_licence_id"]) else None

def licence_name(sheets, licence_id):
    if licence_id is None or pd.isna(licence_id): return "(none)"
    L = sheets["Licences"]
    row = L.loc[L["licence_id"]==licence_id]
    if row.empty: return "(unknown)"
    return str(row.iloc[0]["licence_name"])

def machine_lookup(sheets, machine_id:int):
    M = sheets["Machines"]
    row = M.loc[M["machine_id"]==machine_id]
    if row.empty: return {}
    return row.iloc[0].to_dict()

def machine_max_duration_hours(sheets, machine_id:int, default=4.0):
    M = sheets["Machines"]
    row = M.loc[M["machine_id"]==machine_id]
    if row.empty: return default
    try:
        val = float(row.iloc[0].get("max_duration_hours", default))
        if val <= 0: return default
        return min(val, 12.0)
    except Exception:
        return default

def current_hours_since_last_service(sheets, machine_id:int):
    ol = sheets["OperatingLog"].copy()
    sl = sheets["ServiceLog"]
    ol["start"] = pd.to_datetime(ol.get("start"), errors="coerce")
    ol["hours"] = pd.to_numeric(ol.get("hours"), errors="coerce").fillna(0.0)
    last_service_dt = None
    if not sl.empty:
        s_m = sl.loc[sl["machine_id"]==machine_id].copy()
        if not s_m.empty:
            s_m["date"] = pd.to_datetime(s_m["date"], errors="coerce")
            s_m = s_m.sort_values("date")
            last_service_dt = s_m.iloc[-1]["date"]
    if last_service_dt is not None:
        used = ol[(ol["machine_id"]==machine_id) & (ol["start"] >= last_service_dt)]["hours"].sum()
    else:
        used = ol[ol["machine_id"]==machine_id]["hours"].sum()
    return float(used)

def hours_until_service(sheets, machine_id:int):
    M = sheets["Machines"]
    row = M.loc[M["machine_id"]==machine_id]
    if row.empty: return None
    interval = float(row.iloc[0].get("service_interval_hours", float("nan")))
    if pd.isna(interval): return None
    used = current_hours_since_last_service(sheets, machine_id)
    return float(interval - used)

def human_hours(x):
    if x is None: return "—"
    return f"{x:.1f} h"

def get_operating_hours(sheets):
    OH = sheets.get("OperatingHours", pd.DataFrame())
    if OH.empty:
        default = []
        for i, name in enumerate(["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]):
            if i < 5:
                default.append({"weekday":i,"name":name,"is_open":True,"open":"08:00","close":"17:00"})
            elif i==5:
                default.append({"weekday":i,"name":name,"is_open":True,"open":"09:00","close":"13:00"})
            else:
                default.append({"weekday":i,"name":name,"is_open":False,"open":"00:00","close":"00:00"})
        OH = pd.DataFrame(default); sheets["OperatingHours"] = OH; save_db(sheets)
    return {int(r["weekday"]):(bool(r["is_open"]), str(r["open"]), str(r["close"])) for _, r in OH.iterrows()}

def get_closed_dates(sheets):
    CD = sheets.get("ClosedDates", pd.DataFrame())
    if CD.empty:
        return set()
    CD["date"] = pd.to_datetime(CD["date"], errors="coerce").dt.date
    return set(CD["date"].dropna().tolist())

def is_open_on(sheets, day:date):
    oh = get_operating_hours(sheets); cd = get_closed_dates(sheets)
    if day in cd: return False
    w = day.weekday(); is_open, _, _ = oh.get(w, (False,"00:00","00:00"))
    return is_open

def day_open_close_times(sheets, day:date):
    oh = get_operating_hours(sheets); w = day.weekday()
    is_open, open_s, close_s = oh.get(w, (False,"00:00","00:00"))
    return is_open, open_s, close_s

def timeslots_for_day(sheets, day:date, step_mins=30):
    is_open, open_s, close_s = day_open_close_times(sheets, day)
    if not is_open: return []
    o_h, o_m = map(int, open_s.split(":")); c_h, c_m = map(int, close_s.split(":"))
    start_minutes = o_h*60 + o_m; end_minutes = c_h*60 + c_m
    slots = []; m = start_minutes
    while m < end_minutes:
        h, mm = divmod(m, 60); slots.append(time(h, mm)); m += step_mins
    return slots

def get_setting_bool(sheets, key:str, default:bool=False):
    S = sheets.get("Settings", pd.DataFrame())
    if S.empty or "key" not in S.columns: return default
    m = S[S["key"]==key]
    if m.empty: return default
    val = str(m.iloc[0].get("value","")).strip().lower()
    if val in ("1","true","yes","on"): return True
    if val in ("0","false","no","off"): return False
    return default

# ---------------- Admin Auth ----------------
def _get_admin_password_from_settings(sheets):
    S = sheets.get("Settings", pd.DataFrame())
    if S.empty: return None
    row = S.set_index("key") if "key" in S.columns else None
    if row is None or "admin_password" not in row.index: return None
    val = row.loc["admin_password", "value"]
    return str(val).strip() if isinstance(val, (str, int)) else None

def require_admin(sheets):
    if st.session_state.get("is_admin_authed", False): return True
    pw_secret = None
    try: pw_secret = st.secrets.get("ADMIN_PASSWORD", None)
    except Exception: pw_secret = None
    pw_settings = _get_admin_password_from_settings(sheets)
    expected = pw_secret or pw_settings
    if not expected:
        st.info("No admin password set. Add **ADMIN_PASSWORD** in Secrets (Manage app → Settings → Secrets) or set it in the Settings sheet.")
        st.session_state["is_admin_authed"] = True
        return True
    with st.expander("Admin login", expanded=True):
        pw = st.text_input("Password", type="password", key="adm_pw_input")
        if st.button("Unlock admin", key="adm_pw_btn"):
            if pw == expected:
                st.session_state["is_admin_authed"] = True; st.success("Admin unlocked.")
            else:
                st.error("Incorrect password.")
    return st.session_state.get("is_admin_authed", False)

# ---------------- Business helpers ----------------
def bookings_for_machine_on(sheets, machine_id:int, day:date):
    B = sheets["Bookings"].copy()
    if B.empty: return pd.DataFrame(columns=B.columns)
    B["start"] = pd.to_datetime(B["start"], errors="coerce")
    B["end"] = pd.to_datetime(B["end"], errors="coerce")
    start_day = pd.to_datetime(day); end_day = start_day + pd.Timedelta(days=1)
    return B[(B["machine_id"]==machine_id) & (B["start"]>=start_day) & (B["start"]<end_day)].sort_values("start")

def machine_choices_for_user(sheets, user_id:int):
    lic_ids = user_licence_ids(sheets, user_id)
    M = sheets["Machines"]; L = sheets["Licences"]; choices = []
    for _, row in M.iterrows():
        if row.get("status","Active") != "Active": continue
        req = row.get("required_licence_id", None)
        if pd.isna(req):
            choices.append((f'{row["machine_name"]} (No licence required)', int(row["machine_id"])))
        else:
            req = int(req)
            if req in lic_ids and licence_valid_for_user(sheets, user_id, req):
                lic_row = L.loc[L["licence_id"]==req]; lic_n = lic_row.iloc[0]["licence_name"] if not lic_row.empty else "Unknown Licence"
                choices.append((f'{row["machine_name"]} — requires: {lic_n}', int(row["machine_id"])))
    return choices

def prevent_overlap(sheets, machine_id:int, new_start:datetime, new_end:datetime):
    B = sheets["Bookings"].copy()
    if B.empty: return True, None
    B["start"] = pd.to_datetime(B["start"], errors="coerce"); B["end"] = pd.to_datetime(B["end"], errors="coerce")
    B = B[B["machine_id"]==machine_id]
    for _, r in B.iterrows():
        if overlap(new_start, new_end, r["start"], r["end"]): return False, r
    return True, None

def add_booking_and_log(sheets, user_id:int, machine_id:int, start_dt:datetime, end_dt:datetime, category:str="Usage"):
    B = sheets["Bookings"]; b_id = next_id(B, "booking_id")
    new_b = pd.DataFrame([{"booking_id": b_id, "user_id": user_id, "machine_id": machine_id, "start": start_dt, "end": end_dt, "status": "Confirmed", "category": category}])
    sheets["Bookings"] = pd.concat([B, new_b], ignore_index=True)
    if category == "Usage":
        dur_hours = (end_dt - start_dt).total_seconds()/3600.0
        OL = sheets["OperatingLog"]; op_id = next_id(OL, "op_id")
        new_ol = pd.DataFrame([{"op_id": op_id, "booking_id": b_id, "user_id": user_id, "machine_id": machine_id, "start": start_dt, "end": end_dt, "hours": dur_hours}])
        sheets["OperatingLog"] = pd.concat([OL, new_ol], ignore_index=True)
    save_db(sheets); return b_id

# ---------------- App ----------------
sheets = load_db()

tabs = st.tabs(["Book a Machine", "Calendar", "Issues & Maintenance", "Admin"])

# ---------- Book a Machine ----------
with tabs[0]:
    st.subheader("Book a Machine")
    U = sheets["Users"]; user_map = {row["name"]: int(row["user_id"]) for _, row in U.sort_values("name").iterrows()}
    user_name = st.selectbox("Your name", list(user_map.keys()), key="book_user"); user_id = user_map[user_name]

    st.caption("Your licences:")
    lic_ids = user_licence_ids(sheets, user_id); L = sheets["Licences"]
    your_licences = L[L["licence_id"].isin(list(lic_ids))]["licence_name"].tolist()
    st.info(", ".join(your_licences) if your_licences else "No licences on file.")

    choices = machine_choices_for_user(sheets, user_id); M_all = sheets["Machines"]
    if not choices:
        st.warning("No machines available to you based on your current licences.")
    else:
        label_to_id = {lbl: mid for (lbl, mid) in choices}
        chosen_label = st.selectbox("Choose a machine", list(label_to_id.keys()), key="book_machine")
        machine_id = label_to_id[chosen_label]; mi = machine_lookup(sheets, machine_id)

        cols = st.columns([3,2,2,2])
        with cols[0]:
            st.markdown(f"**Machine:** {mi.get('machine_name')}  \n**Type:** {mi.get('machine_type')}  \n**Serial:** `{mi.get('serial_number')}`  \n**Status:** {mi.get('status')}")
        with cols[1]:
            req_id = mi.get("required_licence_id"); st.markdown(f"**Required licence:** {licence_name(sheets, req_id)}")
        with cols[2]:
            hrs_left = hours_until_service(sheets, machine_id); st.markdown(f"**Hours until service:** {human_hours(hrs_left)}")
        with cols[3]:
            used_since = current_hours_since_last_service(sheets, machine_id); st.markdown(f"**Hours since last service:** {human_hours(used_since)}")

        st.divider()
        day = st.date_input("Day", value=date.today(), key="book_day")

        if not is_open_on(sheets, day):
            st.warning("The shed is **closed** on this date. Please choose another day.")
            st.stop()

        slots = timeslots_for_day(sheets, day, 30)
        if not slots:
            st.info("No slots (closed day).")
        else:
            start_time = st.selectbox("Start time", slots, index=0, key="book_start")
            max_hours = machine_max_duration_hours(sheets, machine_id); duration_hours = st.slider("Duration (hours)", min_value=0.5, max_value=float(max_hours), step=0.5, value=min(1.0,float(max_hours)), key="book_hours")
            start_dt = datetime.combine(day, start_time); end_dt = start_dt + timedelta(hours=float(duration_hours))

            st.markdown("**Existing bookings on this day:**"); show_contacts = get_setting_bool(sheets, "show_contact_on_bookings", True)
            day_bookings = bookings_for_machine_on(sheets, machine_id, day).copy()
            if day_bookings.empty:
                st.info("No bookings yet.")
            else:
                day_bookings["User"] = day_bookings["user_id"].map(lambda x: sheets["Users"].loc[sheets["Users"]["user_id"]==x, "name"].values[0])
                if show_contacts:
                    Utab = sheets["Users"].set_index("user_id")
                    day_bookings["Phone"] = day_bookings["user_id"].map(lambda x: Utab.loc[x,"phone"] if x in Utab.index else "")
                    day_bookings["Email"] = day_bookings["user_id"].map(lambda x: Utab.loc[x,"email"] if x in Utab.index else "")
                day_bookings["Category"] = day_bookings.get("category","Usage").fillna("Usage")
                cols2 = ["User","start","end","status","Category"] + (["Phone","Email"] if show_contacts else [])
                disp = day_bookings[cols2].rename(columns={"start":"Start", "end":"End", "status":"Status"})
                st.dataframe(disp, hide_index=True, use_container_width=True)

            ok, conflict = prevent_overlap(sheets, machine_id, start_dt, end_dt)
            if not ok:
                st.error(f"Selected time overlaps with an existing booking from {conflict['start']} to {conflict['end']}.")
            else:
                if st.button("Confirm Booking", key="confirm_booking"):
                    b_id = add_booking_and_log(sheets, user_id, machine_id, start_dt, end_dt, category="Usage")
                    st.success(f"Booking confirmed. Reference #{b_id}.")

    lic_allowed_ids = [mid for (_, mid) in choices] if choices else []
    blocked = M_all[(M_all["status"]=="Active") & (~M_all["machine_id"].isin(lic_allowed_ids))]
    if not blocked.empty:
        st.caption("Machines you aren't licensed for:")
        st.dataframe(blocked[["machine_name","machine_type"]].rename(columns={"machine_name":"Machine","machine_type":"Type"}), hide_index=True, use_container_width=True)

# ---------- Calendar ----------
with tabs[1]:
    st.subheader("Calendar (by machine)")
    M = sheets["Machines"]; m_map = {row["machine_name"]: int(row["machine_id"]) for _, row in M.sort_values("machine_name").iterrows()}
    if m_map:
        m_name = st.selectbox("Machine", list(m_map.keys()), key="cal_machine"); m_id = m_map[m_name]
        cal_day = st.date_input("Day", value=date.today(), key="cal_day")
        day_b = bookings_for_machine_on(sheets, m_id, cal_day).copy()
        if day_b.empty:
            st.info("No bookings for this day.")
        else:
            show_contacts_cal = get_setting_bool(sheets, "show_contact_on_bookings", True)
            day_b["User"] = day_b["user_id"].map(lambda x: sheets["Users"].loc[sheets["Users"]["user_id"]==x, "name"].values[0])
            if show_contacts_cal:
                Utab = sheets["Users"].set_index("user_id")
                day_b["Phone"] = day_b["user_id"].map(lambda x: Utab.loc[x,"phone"] if x in Utab.index else "")
                day_b["Email"] = day_b["user_id"].map(lambda x: Utab.loc[x,"email"] if x in Utab.index else "")
            day_b["Category"] = day_b.get("category","Usage").fillna("Usage")
            cols3 = ["User","start","end","status","Category"] + (["Phone","Email"] if show_contacts_cal else [])
            disp = day_b[cols3].rename(columns={"start":"Start","end":"End","status":"Status"})
            st.dataframe(disp, hide_index=True, use_container_width=True)
    else:
        st.warning("No machines in the system yet. Add some in Admin.")

# ---------- Issues & Maintenance ----------
with tabs[2]:
    st.subheader("Report an Issue")
    u_map = {row["name"]: int(row["user_id"]) for _, row in sheets["Users"].sort_values("name").iterrows()}
    issue_user_name = st.selectbox("Your name", list(u_map.keys()), key="issue_user"); issue_user_id = u_map[issue_user_name]
    m_map2 = {row["machine_name"]: int(row["machine_id"]) for _, row in sheets["Machines"].sort_values("machine_name").iterrows()}
    issue_m_name = st.selectbox("Machine", list(m_map2.keys()), key="issue_machine"); issue_m_id = m_map2[issue_m_name]
    issue_text = st.text_area("Describe the issue (e.g., vibration, sharpening needed)"); severity = st.selectbox("Severity", ["Low","Medium","High"])
    if st.button("Submit Issue", key="issue_submit"):
        Issues = sheets["Issues"]; issue_id = next_id(Issues, "issue_id")
        new_i = pd.DataFrame([{"issue_id": issue_id, "machine_id": issue_m_id, "user_id": issue_user_id, "date_reported": pd.Timestamp.now(), "issue_text": issue_text.strip(), "severity": severity, "status": "Open", "resolution_notes": "", "date_resolved": pd.NaT}])
        sheets["Issues"] = pd.concat([Issues, new_i], ignore_index=True); save_db(sheets); st.success(f"Issue logged. Reference #{issue_id}.")

    st.divider(); st.subheader("Open Issues")
    open_issues = sheets["Issues"]
    if open_issues.empty or not (open_issues["status"]=="Open").any():
        st.info("No open issues.")
    else:
        open_issues = open_issues[open_issues["status"]=="Open"].copy()
        open_issues["Machine"] = open_issues["machine_id"].map(lambda x: sheets["Machines"].loc[sheets["Machines"]["machine_id"]==x, "machine_name"].values[0])
        open_issues["Reported By"] = open_issues["user_id"].map(lambda x: sheets["Users"].loc[sheets["Users"]["user_id"]==x, "name"].values[0])
        disp = open_issues[["issue_id","Machine","Reported By","date_reported","severity","issue_text"]].rename(columns={"issue_id":"Issue #","date_reported":"Reported","severity":"Severity","issue_text":"Issue"})
        st.dataframe(disp, hide_index=True, use_container_width=True)
        st.caption("Resolve issues in Admin → Maintenance.")

# ---------- Admin ----------
with tabs[3]:
    st.subheader("Admin")
    if not require_admin(sheets): st.stop()
    at = st.tabs(["Users & Licences", "Machines", "Schedule", "Maintenance", "Data & Settings"])

    # Users & Licences
    with at[0]:
        st.markdown("### Add User")
        new_name = st.text_input("Name", key="adm_new_user")
        new_phone = st.text_input("Phone", key="adm_new_phone")
        new_email = st.text_input("Email", key="adm_new_email")
        new_addr = st.text_area("Postal address (street, suburb, state, postcode)", key="adm_new_addr")
        if st.button("Add User", key="adm_add_user"):
            if not new_name.strip(): st.error("Name required.")
            else:
                U = sheets["Users"]; uid = next_id(U, "user_id")
                sheets["Users"] = pd.concat([U, pd.DataFrame([{"user_id": uid, "name": new_name.strip(), "phone": new_phone.strip(), "email": new_email.strip(), "address": new_addr.strip()}])], ignore_index=True)
                save_db(sheets); st.success(f"User '{new_name}' added with ID {uid}.")

        st.markdown("---"); st.markdown("### Assign / Update Licence for a User")
        all_lics = sheets["Licences"].sort_values("licence_name"); lic_map = {row["licence_name"]: int(row["licence_id"]) for _, row in all_lics.iterrows()}
        u_map2 = {row["name"]: int(row["user_id"]) for _, row in sheets["Users"].sort_values("name").iterrows()}
        if not u_map2: st.info("No users yet.")
        else:
            sel_user = st.selectbox("User", list(u_map2.keys()), key="adm_user_pick"); sel_user_id = u_map2[sel_user]
            sel_lic = st.selectbox("Licence", list(lic_map.keys()), key="adm_lic_pick"); sel_lic_id = lic_map[sel_lic]
            start_d = st.date_input("Start date", value=date.today(), key="adm_lic_start")
            end_d = st.date_input("End date (optional)", key="adm_lic_end")
            if st.button("Assign/Update Licence", key="adm_assign"):
                UL = sheets["UserLicences"].copy(); mask = (UL["user_id"]==sel_user_id) & (UL["licence_id"]==sel_lic_id)
                if mask.any():
                    UL.loc[mask, "start_date"] = pd.Timestamp(start_d); UL.loc[mask, "end_date"] = pd.Timestamp(end_d) if end_d else pd.NaT
                else:
                    UL = pd.concat([UL, pd.DataFrame([{"user_id": sel_user_id, "licence_id": sel_lic_id, "start_date": pd.Timestamp(start_d), "end_date": pd.Timestamp(end_d) if end_d else pd.NaT}])], ignore_index=True)
                sheets["UserLicences"] = UL; save_db(sheets); st.success("Licence saved.")

            st.markdown("#### Current licences for user"); ULv = sheets["UserLicences"].copy()
            if ULv.empty: st.info("No licences assigned yet.")
            else:
                ULv = ULv[ULv["user_id"]==sel_user_id].copy(); ULv["Licence"] = ULv["licence_id"].map(lambda x: sheets["Licences"].loc[sheets["Licences"]["licence_id"]==x, "licence_name"].values[0])
                st.dataframe(ULv[["Licence","start_date","end_date"]].rename(columns={"start_date":"Start","end_date":"End"}), hide_index=True, use_container_width=True)

            st.markdown("#### Remove a licence")
            if not sheets["UserLicences"].empty:
                UL_user = sheets["UserLicences"][sheets["UserLicences"]["user_id"]==sel_user_id]
                if not UL_user.empty:
                    lic_names_for_user = UL_user["licence_id"].map(lambda x: sheets["Licences"].loc[sheets["Licences"]["licence_id"]==x, "licence_name"].values[0]).tolist()
                    rem_choice = st.selectbox("Licence to remove", lic_names_for_user, key="adm_remove_pick")
                    if st.button("Remove Licence", key="adm_remove"):
                        lic_id_rm = lic_map[rem_choice]; UL2 = sheets["UserLicences"]
                        sheets["UserLicences"] = UL2[~((UL2["user_id"]==sel_user_id) & (UL2["licence_id"]==lic_id_rm))]; save_db(sheets); st.success("Licence removed.")

        st.markdown("---"); st.markdown("### Existing Users")
        Udisp = sheets["Users"].copy()
        for col in ["phone","email","address"]:
            if col not in Udisp.columns: Udisp[col] = ""
        Udisp = Udisp.rename(columns={"user_id":"User ID","name":"Name","phone":"Phone","email":"Email","address":"Address"})
        st.dataframe(Udisp, hide_index=True, use_container_width=True)

    # Machines
    with at[1]:
        st.markdown("### Add Machine")
        m_name = st.text_input("Machine name", placeholder="e.g., Lathe #1", key="adm_m_name")
        m_type = st.text_input("Machine type", placeholder="e.g., Lathe", key="adm_m_type")
        serial = st.text_input("Serial number", key="adm_m_serial")
        lic_map2 = {row["licence_name"]: int(row["licence_id"]) for _, row in sheets["Licences"].sort_values("licence_name").iterrows()}
        req_lic = st.selectbox("Required licence", ["(none)"] + list(lic_map2.keys()), key="adm_m_req")
        status = st.selectbox("Status", ["Active","OutOfService"], key="adm_m_status")
        service_interval = st.number_input("Service interval (hours)", min_value=1.0, step=1.0, value=50.0, key="adm_m_serv")
        max_dur = st.number_input("Max booking duration (hours)", min_value=0.5, step=0.5, value=4.0, key="adm_maxdur")
        last_service_date = st.date_input("Last service date", value=date.today(), key="adm_m_lastsvc")

        if st.button("Add Machine", key="adm_m_add"):
            if not m_name.strip() or not m_type.strip() or not serial.strip():
                st.error("Please complete name, type, and serial.")
            else:
                M = sheets["Machines"]; mid = next_id(M, "machine_id")
                req_id = pd.NA if req_lic=="(none)" else lic_map2[req_lic]
                new_m = pd.DataFrame([{"machine_id": mid, "machine_name": m_name.strip(), "machine_type": m_type.strip(), "serial_number": serial.strip(), "required_licence_id": req_id, "status": status, "service_interval_hours": float(service_interval), "last_service_date": pd.Timestamp(last_service_date), "max_duration_hours": float(max_dur)}])
                sheets["Machines"] = pd.concat([M, new_m], ignore_index=True); save_db(sheets); st.success(f"Machine '{m_name}' added with ID {mid}.")

        st.markdown("### Machines")
        M = sheets["Machines"].copy()
        if M.empty: st.info("No machines yet.")
        else:
            M["Required Licence"] = M["required_licence_id"].map(lambda x: licence_name(sheets, int(x)) if not pd.isna(x) else "(none)")
            if "max_duration_hours" not in M.columns: M["max_duration_hours"] = 4.0
            Mdisp = M[["machine_id","machine_name","machine_type","serial_number","Required Licence","status","service_interval_hours","last_service_date","max_duration_hours"]].rename(columns={"machine_id":"ID","machine_name":"Name","machine_type":"Type","serial_number":"Serial","status":"Status","service_interval_hours":"Service Interval (h)","last_service_date":"Last Service","max_duration_hours":"Max Duration (h)"})
            st.dataframe(Mdisp, hide_index=True, use_container_width=True)

    # Schedule (Day/Week views + rescheduler)
    with at[2]:
        st.markdown("### Day view")
        day_pick = st.date_input("Day", value=date.today(), key="adm_sched_day")
        B = sheets["Bookings"].copy()
        if B.empty:
            st.info("No bookings yet.")
        else:
            B["start"] = pd.to_datetime(B["start"], errors="coerce"); B["end"] = pd.to_datetime(B["end"], errors="coerce")
            start_day = pd.to_datetime(day_pick); end_day = start_day + pd.Timedelta(days=1)
            D = B[(B["start"]>=start_day) & (B["start"]<end_day)].copy()
            if D.empty: st.info("No bookings on this day.")
            else:
                show_contacts_adm = get_setting_bool(sheets, "show_contact_on_bookings", True)
                D["User"] = D["user_id"].map(lambda x: sheets["Users"].loc[sheets["Users"]["user_id"]==x, "name"].values[0] if x in sheets["Users"]["user_id"].values else "(system)")
                D["Machine"] = D["machine_id"].map(lambda x: sheets["Machines"].loc[sheets["Machines"]["machine_id"]==x, "machine_name"].values[0])
                if show_contacts_adm:
                    Utab = sheets["Users"].set_index("user_id")
                    D["Phone"] = D["user_id"].map(lambda x: Utab.loc[x,"phone"] if x in Utab.index else "")
                    D["Email"] = D["user_id"].map(lambda x: Utab.loc[x,"email"] if x in Utab.index else "")
                D["Category"] = D.get("category","Usage").fillna("Usage")
                cols = ["Machine","User","start","end","Category","status"] + (["Phone","Email"] if show_contacts_adm else [])
                st.dataframe(D[cols].rename(columns={"start":"Start","end":"End","status":"Status"}), hide_index=True, use_container_width=True)

                hours = list(range(6, 22)); counts = []
                for h in hours:
                    h0 = start_day + pd.Timedelta(hours=h); h1 = h0 + pd.Timedelta(hours=1); cnt=0
                    for _, r in D.iterrows():
                        if (r["start"] < h1) and (r["end"] > h0) and r.get("category","Usage")=="Usage": cnt += 1
                    counts.append(cnt)
                st.bar_chart(pd.DataFrame({"hour":hours, "bookings":counts}).set_index("hour"))

        st.markdown("---"); st.markdown("### Week view")
        week_start = st.date_input("Week starting (Monday)", value=(date.today() - timedelta(days=date.today().weekday())), key="adm_week_start")
        W = pd.DataFrame()
        if not B.empty:
            B2 = B.copy(); B2["start_date"] = pd.to_datetime(B2["start"]).dt.date
            mon = week_start; week_days = [mon + timedelta(days=i) for i in range(7)]; rows = []
            for d in week_days: rows.append({"Date": d, "Total bookings": int((B2["start_date"]==d).sum())})
            W = pd.DataFrame(rows); st.dataframe(W, hide_index=True, use_container_width=True)

        st.markdown("---"); st.markdown("### Reschedule a booking")
        if B.empty: st.info("No bookings to reschedule.")
        else:
            B["label"] = B.apply(lambda r: f'#{int(r["booking_id"])} — {sheets["Machines"].loc[sheets["Machines"]["machine_id"]==r["machine_id"], "machine_name"].values[0]} for {sheets["Users"].loc[sheets["Users"]["user_id"]==r["user_id"], "name"].values[0] if r["user_id"] in sheets["Users"]["user_id"].values else "(system)"} on {r["start"]}', axis=1)
            pick = st.selectbox("Choose a booking", B["label"].tolist(), key="adm_res_pick"); bid = int(pick.split("—")[0].strip().strip("#"))
            row = B.loc[B["booking_id"]==bid].iloc[0]
            nm_map = {row["machine_name"]: int(row["machine_id"]) for _, row in sheets["Machines"].sort_values("machine_name").iterrows()}
            new_m = st.selectbox("New machine", list(nm_map.keys()), key="adm_res_mach"); new_mid = nm_map[new_m]
            new_day = st.date_input("New day", value=pd.to_datetime(row["start"]).date(), key="adm_res_day")
            if not is_open_on(sheets, new_day): st.warning("Closed on that day.")
            new_slots = timeslots_for_day(sheets, new_day, 30)
            if new_slots:
                new_start = st.selectbox("New start time", new_slots, key="adm_res_start")
                max_hours = machine_max_duration_hours(sheets, new_mid)
                new_hours = st.slider("New duration (hours)", 0.5, float(max_hours), step=0.5, value=min(1.0,float(max_hours)), key="adm_res_hours")
                ns = datetime.combine(new_day, new_start); ne = ns + timedelta(hours=float(new_hours))
                ok, conflict = prevent_overlap(sheets, new_mid, ns, ne)
                if not ok and int(row["machine_id"])!=new_mid: st.error("Time overlaps with another booking.")
                else:
                    if st.button("Apply reschedule", key="adm_res_apply"):
                        B2 = sheets["Bookings"]; idx = B2.index[B2["booking_id"]==bid]
                        if len(idx)>0:
                            B2.loc[idx, "machine_id"] = new_mid; B2.loc[idx, "start"] = ns; B2.loc[idx, "end"] = ne; sheets["Bookings"] = B2
                            if str(row.get("category","Usage"))=="Usage":
                                OL = sheets["OperatingLog"]; ol_idx = OL.index[OL["booking_id"]==bid]
                                if len(ol_idx)>0:
                                    OL.loc[ol_idx, "machine_id"] = new_mid; OL.loc[ol_idx, "start"] = ns; OL.loc[ol_idx, "end"] = ne; OL.loc[ol_idx, "hours"] = (ne - ns).total_seconds()/3600.0; sheets["OperatingLog"] = OL
                            save_db(sheets); st.success("Rescheduled.")
            else:
                st.info("No slots for that day (closed).")

    # Maintenance
    with at[3]:
        st.markdown("### Log Service")
        m_map3 = {row["machine_name"]: int(row["machine_id"]) for _, row in sheets["Machines"].sort_values("machine_name").iterrows()}
        if not m_map3: st.info("No machines to service.")
        else:
            s_m_name = st.selectbox("Machine", list(m_map3.keys()), key="svc_machine"); s_m_id = m_map3[s_m_name]
            cur_used = current_hours_since_last_service(sheets, s_m_id); st.caption(f"Hours since last service: {cur_used:.1f} h")
            s_date = st.date_input("Service date", value=date.today(), key="svc_date"); notes = st.text_input("Notes", placeholder="e.g., blades sharpened, bearings checked", key="svc_notes")
            if st.button("Record Service", key="svc_record"):
                SL = sheets["ServiceLog"]; sid = next_id(SL, "service_id")
                new_s = pd.DataFrame([{"service_id": sid, "machine_id": s_m_id, "date": pd.Timestamp(s_date), "hours_at_service": float(cur_used), "notes": notes.strip()}])
                sheets["ServiceLog"] = pd.concat([SL, new_s], ignore_index=True)
                M = sheets["Machines"]; idx = M.index[M["machine_id"]==s_m_id]
                if len(idx)>0: M.loc[idx, "last_service_date"] = pd.Timestamp(s_date); sheets["Machines"] = M
                save_db(sheets); st.success(f"Service recorded for {s_m_name}.")

        st.markdown("### Schedule maintenance downtime")
        m_map_sched = {row["machine_name"]: int(row["machine_id"]) for _, row in sheets["Machines"].sort_values("machine_name").iterrows()}
        if not m_map_sched: st.info("No machines available.")
        else:
            sched_m_name = st.selectbox("Machine", list(m_map_sched.keys()), key="mnt_sched_machine"); sched_m_id = m_map_sched[sched_m_name]
            sched_day = st.date_input("Day", value=date.today(), key="mnt_sched_day")
            if not is_open_on(sheets, sched_day): st.warning("That day is closed per operating schedule.")
            sched_slots = timeslots_for_day(sheets, sched_day, 30)
            if not sched_slots: st.info("No timeslots (closed day).")
            else:
                sched_start = st.selectbox("Start time", sched_slots, key="mnt_sched_start")
                max_hours = machine_max_duration_hours(sheets, sched_m_id, default=8.0)
                maint_hours = st.slider("Duration (hours)", min_value=0.5, max_value=float(max_hours), step=0.5, value=1.0, key="mnt_sched_hours")
                ms = datetime.combine(sched_day, sched_start); me = ms + timedelta(hours=float(maint_hours))
                ok, conflict = prevent_overlap(sheets, sched_m_id, ms, me)
                if not ok: st.error(f"Overlaps with existing booking from {conflict['start']} to {conflict['end']}.")
                else:
                    if st.button("Block out for maintenance", key="mnt_sched_btn"):
                        b_id = add_booking_and_log(sheets, 0, sched_m_id, ms, me, category="Maintenance")
                        st.success(f"Maintenance block created (booking #{b_id}).")

    # Data & Settings
    with at[4]:
        st.markdown("### Admin & privacy settings")
        S = sheets.get("Settings", pd.DataFrame(columns=["key","value"]))
        cur_pw = S.loc[S["key"]=="admin_password", "value"].iloc[0] if (not S.empty and (S["key"]=="admin_password").any()) else ""
        new_pw = st.text_input("Admin password", value=str(cur_pw), type="password")
        cur_toggle = False
        if (not S.empty) and (S["key"]=="show_contact_on_bookings").any():
            cur_toggle = str(S.loc[S["key"]=="show_contact_on_bookings", "value"].iloc[0]).strip().lower() in ("1","true","yes","on")
        show_contacts_new = st.checkbox("Show member phone/email on bookings (admin & calendar)", value=cur_toggle)
        if st.button("Save settings"):
            def upsert(key,val):
                nonlocal S
                if S.empty or not (S["key"]==key).any():
                    S = pd.concat([S, pd.DataFrame([[key,str(val)]], columns=["key","value"])], ignore_index=True)
                else:
                    S.loc[S["key"]==key, "value"] = str(val)
            upsert("admin_password", new_pw); upsert("show_contact_on_bookings", show_contacts_new)
            sheets["Settings"] = S; save_db(sheets); st.success("Settings saved.")

        st.markdown("### Operating hours")
        OH = sheets.get("OperatingHours", pd.DataFrame())
        if OH.empty: _ = get_operating_hours(sheets); OH = sheets.get("OperatingHours")
        for i, name in enumerate(["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]):
            row = OH[OH["weekday"]==i]
            is_open = bool(row.iloc[0]["is_open"]) if not row.empty else (i<5)
            c1,c2,c3,c4 = st.columns([1,1,1,5])
            with c1: st.write(name)
            with c2: is_open_new = st.checkbox("Open", value=is_open, key=f"oh_open_{i}")
            with c3: open_time = st.text_input("Open", value=str(row.iloc[0]["open"]) if not row.empty else "08:00", key=f"oh_open_time_{i}")
            with c4: close_time = st.text_input("Close", value=str(row.iloc[0]["close"]) if not row.empty else "17:00", key=f"oh_close_time_{i}")
            if row.empty:
                OH = pd.concat([OH, pd.DataFrame([{"weekday":i,"name":name,"is_open":is_open_new,"open":open_time,"close":close_time}])], ignore_index=True)
            else:
                idx = OH.index[OH["weekday"]==i]; OH.loc[idx, "is_open"] = is_open_new; OH.loc[idx, "open"] = open_time; OH.loc[idx, "close"] = close_time
        sheets["OperatingHours"] = OH

        st.markdown("### Closed dates")
        CD = sheets.get("ClosedDates", pd.DataFrame(columns=["date","reason"]))
        add_cd = st.date_input("Add closed date")
        reason = st.text_input("Reason (optional)")
        if st.button("Add closed date"):
            CD = pd.concat([CD, pd.DataFrame([[pd.Timestamp(add_cd), reason]], columns=["date","reason"])], ignore_index=True)
            sheets["ClosedDates"] = CD; save_db(sheets); st.success("Closed date added.")
        if not CD.empty: st.dataframe(CD, hide_index=True, use_container_width=True)

        st.markdown("---"); st.markdown("### Replace/Backup Database")
        up = st.file_uploader("Upload a replacement Excel DB (must match schema)", type=["xlsx"], key="db_upload")
        if st.button("Replace DB from upload", key="db_replace") and up is not None:
            (BASE_DIR / "data" / "db.xlsx").write_bytes(up.read()); st.success("Database replaced. Please refresh.")
        st.download_button("Download current DB.xlsx", data=open(DB_PATH,"rb").read(), file_name="db.xlsx")
