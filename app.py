
import streamlit as st
import pandas as pd
from dateutil.parser import parse as dparse
from datetime import datetime, timedelta, date, time
from pathlib import Path

DB_PATH = Path("data/db.xlsx")

st.set_page_config(page_title="Men's Shed Scheduler", page_icon="🪚", layout="wide")

# === Minimal header/logo injection (no theme override) ===
def _inject_min_css():
    css_path = Path("assets/styles.css")
    if css_path.exists():
        st.markdown(f"<style>{css_path.read_text()}</style>", unsafe_allow_html=True)


def _brand_header():
    # Look for logo in common paths / cases to help non-technical setup
    candidates = [
        Path("assets/logo.png"), Path("assets/Logo.png"),
        Path("Assets/logo.png"), Path("Assets/Logo.png"),
        Path("assets/logo.svg"), Path("assets/Logo.svg"),
        Path("Assets/logo.svg"), Path("Assets/Logo.svg"),
        Path("assets/logo.jpg"), Path("assets/Logo.jpg"),
        Path("Assets/logo.jpg"), Path("Assets/Logo.jpg"),
    ]
    logo_path = next((p for p in candidates if p.exists()), None)
    if logo_path:
        st.markdown(f'<div class="header"><img src="{logo_path.as_posix()}" alt="Woodturners of the Hunter"></div>', unsafe_allow_html=True)
    else:
        st.markdown("### Woodturners of the Hunter")
_brand_header()
# === End minimal injection ===


# ---------------- Utility functions ----------------
def load_db():
    if not DB_PATH.exists():
        st.error("Database not found at data/db.xlsx")
        st.stop()
    xls = pd.ExcelFile(DB_PATH)
    sheets = {name: pd.read_excel(DB_PATH, sheet_name=name) for name in xls.sheet_names}
    for e in ["Users","Licences","UserLicences","Machines","Bookings","OperatingLog","Issues","ServiceLog"]:
        sheets.setdefault(e, pd.DataFrame())
    # Defaults / schema upgrades
    M = sheets["Machines"]
    if "status" in M.columns:
        M["status"] = M["status"].fillna("Active")
    if "max_duration_hours" not in M.columns:
        M["max_duration_hours"] = 4.0
    UL = sheets["UserLicences"]
    if "start_date" not in UL.columns:
        UL["start_date"] = pd.NaT
    if "end_date" not in UL.columns:
        UL["end_date"] = pd.NaT
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

def get_timeslots(start_hour=7, end_hour=20, step_mins=30):
    return [time(hour=h, minute=m) for h in range(start_hour, end_hour) for m in range(0,60,step_mins)]

def bookings_for_machine_on(sheets, machine_id:int, day:date):
    B = sheets["Bookings"].copy()
    if B.empty: return pd.DataFrame(columns=B.columns)
    B["start"] = pd.to_datetime(B["start"], errors="coerce")
    B["end"] = pd.to_datetime(B["end"], errors="coerce")
    start_day = pd.to_datetime(day)
    end_day = start_day + pd.Timedelta(days=1)
    return B[(B["machine_id"]==machine_id) & (B["start"]>=start_day) & (B["start"]<end_day)].sort_values("start")

def machine_choices_for_user(sheets, user_id:int):
    lic_ids = user_licence_ids(sheets, user_id)
    M = sheets["Machines"]
    L = sheets["Licences"]
    choices = []
    for _, row in M.iterrows():
        if row.get("status","Active") != "Active":
            continue
        req = row.get("required_licence_id", None)
        if pd.isna(req):
            label = f'{row["machine_name"]} (No licence required)'
            choices.append((label, int(row["machine_id"])))
        else:
            req = int(req)
            if req in lic_ids and licence_valid_for_user(sheets, user_id, req):
                lic_row = L.loc[L["licence_id"]==req]
                lic_n = lic_row.iloc[0]["licence_name"] if not lic_row.empty else "Unknown Licence"
                label = f'{row["machine_name"]} — requires: {lic_n}'
                choices.append((label, int(row["machine_id"])))
    return choices

def prevent_overlap(sheets, machine_id:int, new_start:datetime, new_end:datetime):
    B = sheets["Bookings"].copy()
    if B.empty: 
        return True, None
    B["start"] = pd.to_datetime(B["start"], errors="coerce")
    B["end"] = pd.to_datetime(B["end"], errors="coerce")
    B = B[B["machine_id"]==machine_id]
    for _, r in B.iterrows():
        if overlap(new_start, new_end, r["start"], r["end"]):
            return False, r
    return True, None

def add_booking_and_log(sheets, user_id:int, machine_id:int, start_dt:datetime, end_dt:datetime):
    B = sheets["Bookings"]
    b_id = next_id(B, "booking_id")
    new_b = pd.DataFrame([{
        "booking_id": b_id, "user_id": user_id, "machine_id": machine_id,
        "start": start_dt, "end": end_dt, "status": "Confirmed"
    }])
    sheets["Bookings"] = pd.concat([B, new_b], ignore_index=True)
    # Operating log
    dur_hours = (end_dt - start_dt).total_seconds()/3600.0
    OL = sheets["OperatingLog"]
    op_id = next_id(OL, "op_id")
    new_ol = pd.DataFrame([{
        "op_id": op_id, "booking_id": b_id, "user_id": user_id, "machine_id": machine_id,
        "start": start_dt, "end": end_dt, "hours": dur_hours
    }])
    sheets["OperatingLog"] = pd.concat([OL, new_ol], ignore_index=True)
    save_db(sheets)
    return b_id

# ---------------- App ----------------
sheets = load_db()

tabs = st.tabs(["Book a Machine", "Calendar", "Issues & Maintenance", "Admin"])

# ---------- Book a Machine ----------
with tabs[0]:
    st.subheader("Book a Machine")
    # Choose user
    U = sheets["Users"]
    user_map = {row["name"]: int(row["user_id"]) for _, row in U.sort_values("name").iterrows()}
    user_name = st.selectbox("Your name", list(user_map.keys()), key="book_user")
    user_id = user_map[user_name]

    # Show user's licences
    st.caption("Your licences:")
    lic_ids = user_licence_ids(sheets, user_id)
    L = sheets["Licences"]
    your_licences = L[L["licence_id"].isin(list(lic_ids))]["licence_name"].tolist()
    st.info(", ".join(your_licences) if your_licences else "No licences on file.")

    # Machines allowed for this user
    choices = machine_choices_for_user(sheets, user_id)
    M_all = sheets["Machines"]
    if not choices:
        st.warning("No machines available to you based on your current licences.")
    else:
        label_to_id = {lbl: mid for (lbl, mid) in choices}
        chosen_label = st.selectbox("Choose a machine", list(label_to_id.keys()), key="book_machine")
        machine_id = label_to_id[chosen_label]
        mi = machine_lookup(sheets, machine_id)

        # Details + service info
        cols = st.columns([3,2,2,2])
        with cols[0]:
            st.markdown(f"**Machine:** {mi.get('machine_name')}  \n**Type:** {mi.get('machine_type')}  \n**Serial:** `{mi.get('serial_number')}`  \n**Status:** {mi.get('status')}")
        with cols[1]:
            req_id = mi.get("required_licence_id")
            st.markdown(f"**Required licence:** {licence_name(sheets, req_id)}")
        with cols[2]:
            hrs_left = hours_until_service(sheets, machine_id)
            st.markdown(f"**Hours until service:** {human_hours(hrs_left)}")
        with cols[3]:
            used_since = current_hours_since_last_service(sheets, machine_id)
            st.markdown(f"**Hours since last service:** {human_hours(used_since)}")

        st.divider()

        # Day defaults to today; show existing bookings
        day = st.date_input("Day", value=date.today(), key="book_day")
        slots = get_timeslots(7, 20, 30)
        start_time = st.selectbox("Start time", slots, index=slots.index(time(9,0)) if time(9,0) in slots else 0, key="book_start")
        max_hours = machine_max_duration_hours(sheets, machine_id)
        duration_hours = st.slider("Duration (hours)", min_value=0.5, max_value=float(max_hours), step=0.5, value=min(1.0, float(max_hours)), key="book_hours")

        start_dt = datetime.combine(day, start_time)
        end_dt = start_dt + timedelta(hours=float(duration_hours))

        st.markdown("**Existing bookings on this day:**")
        day_bookings = bookings_for_machine_on(sheets, machine_id, day).copy()
        if day_bookings.empty:
            st.info("No bookings yet.")
        else:
            day_bookings["User"] = day_bookings["user_id"].map(lambda x: sheets["Users"].loc[sheets["Users"]["user_id"]==x, "name"].values[0])
            disp = day_bookings[["User","start","end","status"]].rename(columns={"start":"Start", "end":"End", "status":"Status"})
            st.dataframe(disp, hide_index=True, use_container_width=True)

        ok, conflict = prevent_overlap(sheets, machine_id, start_dt, end_dt)
        if not ok:
            st.error(f"Selected time overlaps with an existing booking from {conflict['start']} to {conflict['end']}.")
        else:
            if st.button("Confirm Booking", key="confirm_booking"):
                b_id = add_booking_and_log(sheets, user_id, machine_id, start_dt, end_dt)
                st.success(f"Booking confirmed. Reference #{b_id}.")

    # Show machines not licensed for (informational)
    lic_allowed_ids = [mid for (_, mid) in choices] if choices else []
    blocked = M_all[(M_all["status"]=="Active") & (~M_all["machine_id"].isin(lic_allowed_ids))]
    if not blocked.empty:
        st.caption("Machines you aren't licensed for:")
        st.dataframe(blocked[["machine_name","machine_type"]].rename(columns={"machine_name":"Machine","machine_type":"Type"}), hide_index=True, use_container_width=True)

# ---------- Calendar ----------
with tabs[1]:
    st.subheader("Calendar (by machine)")
    M = sheets["Machines"]
    m_map = {row["machine_name"]: int(row["machine_id"]) for _, row in M.sort_values("machine_name").iterrows()}
    if m_map:
        m_name = st.selectbox("Machine", list(m_map.keys()), key="cal_machine")
        m_id = m_map[m_name]
        cal_day = st.date_input("Day", value=date.today(), key="cal_day")
        day_b = bookings_for_machine_on(sheets, m_id, cal_day).copy()
        if day_b.empty:
            st.info("No bookings for this day.")
        else:
            day_b["User"] = day_b["user_id"].map(lambda x: sheets["Users"].loc[sheets["Users"]["user_id"]==x, "name"].values[0])
            disp = day_b[["User","start","end","status"]].rename(columns={"start":"Start","end":"End","status":"Status"})
            st.dataframe(disp, hide_index=True, use_container_width=True)
    else:
        st.warning("No machines in the system yet. Add some in Admin.")

# ---------- Issues & Maintenance ----------
with tabs[2]:
    st.subheader("Report an Issue")
    u_map = {row["name"]: int(row["user_id"]) for _, row in sheets["Users"].sort_values("name").iterrows()}
    issue_user_name = st.selectbox("Your name", list(u_map.keys()), key="issue_user")
    issue_user_id = u_map[issue_user_name]
    m_map2 = {row["machine_name"]: int(row["machine_id"]) for _, row in sheets["Machines"].sort_values("machine_name").iterrows()}
    issue_m_name = st.selectbox("Machine", list(m_map2.keys()), key="issue_machine")
    issue_m_id = m_map2[issue_m_name]
    issue_text = st.text_area("Describe the issue (e.g., vibration, sharpening needed)")
    severity = st.selectbox("Severity", ["Low","Medium","High"])
    if st.button("Submit Issue", key="issue_submit"):
        Issues = sheets["Issues"]
        issue_id = next_id(Issues, "issue_id")
        new_i = pd.DataFrame([{
            "issue_id": issue_id, "machine_id": issue_m_id, "user_id": issue_user_id,
            "date_reported": pd.Timestamp.now(), "issue_text": issue_text.strip(),
            "severity": severity, "status": "Open", "resolution_notes": "", "date_resolved": pd.NaT
        }])
        sheets["Issues"] = pd.concat([Issues, new_i], ignore_index=True)
        save_db(sheets)
        st.success(f"Issue logged. Reference #{issue_id}.")

    st.divider()
    st.subheader("Open Issues")
    open_issues = sheets["Issues"]
    if open_issues.empty or not (open_issues["status"]=="Open").any():
        st.info("No open issues.")
    else:
        open_issues = open_issues[open_issues["status"]=="Open"].copy()
        open_issues["Machine"] = open_issues["machine_id"].map(lambda x: sheets["Machines"].loc[sheets["Machines"]["machine_id"]==x, "machine_name"].values[0])
        open_issues["Reported By"] = open_issues["user_id"].map(lambda x: sheets["Users"].loc[sheets["Users"]["user_id"]==x, "name"].values[0])
        disp = open_issues[["issue_id","Machine","Reported By","date_reported","severity","issue_text"]].rename(columns={
            "issue_id":"Issue #","date_reported":"Reported","severity":"Severity","issue_text":"Issue"
        })
        st.dataframe(disp, hide_index=True, use_container_width=True)
        st.caption("Resolve issues in Admin → Maintenance.")

# ---------- Admin ----------
with tabs[3]:
    st.subheader("Admin")
    at = st.tabs(["Users & Licences", "Machines", "Maintenance", "Data"])

    # Users & Licences
    with at[0]:
        st.markdown("### Add User")
        new_name = st.text_input("Name", key="adm_new_user")
        if st.button("Add User", key="adm_add_user"):
            if not new_name.strip():
                st.error("Name required.")
            else:
                U = sheets["Users"]
                uid = next_id(U, "user_id")
                sheets["Users"] = pd.concat([U, pd.DataFrame([{"user_id": uid, "name": new_name.strip()}])], ignore_index=True)
                save_db(sheets)
                st.success(f"User '{new_name}' added with ID {uid}.")

        st.markdown("---")
        st.markdown("### Assign / Update Licence for a User")
        all_lics = sheets["Licences"].sort_values("licence_name")
        lic_map = {row["licence_name"]: int(row["licence_id"]) for _, row in all_lics.iterrows()}
        u_map2 = {row["name"]: int(row["user_id"]) for _, row in sheets["Users"].sort_values("name").iterrows()}
        if not u_map2:
            st.info("No users yet.")
        else:
            sel_user = st.selectbox("User", list(u_map2.keys()), key="adm_user_pick")
            sel_user_id = u_map2[sel_user]
            sel_lic = st.selectbox("Licence", list(lic_map.keys()), key="adm_lic_pick")
            sel_lic_id = lic_map[sel_lic]
            start_d = st.date_input("Start date", value=date.today(), key="adm_lic_start")
            end_d = st.date_input("End date (optional)", key="adm_lic_end")
            if st.button("Assign/Update Licence", key="adm_assign"):
                UL = sheets["UserLicences"].copy()
                mask = (UL["user_id"]==sel_user_id) & (UL["licence_id"]==sel_lic_id)
                if mask.any():
                    UL.loc[mask, "start_date"] = pd.Timestamp(start_d)
                    UL.loc[mask, "end_date"] = pd.Timestamp(end_d) if end_d else pd.NaT
                else:
                    UL = pd.concat([UL, pd.DataFrame([{
                        "user_id": sel_user_id, "licence_id": sel_lic_id,
                        "start_date": pd.Timestamp(start_d),
                        "end_date": pd.Timestamp(end_d) if end_d else pd.NaT
                    }])], ignore_index=True)
                sheets["UserLicences"] = UL
                save_db(sheets)
                st.success("Licence saved.")

            st.markdown("#### Current licences for user")
            ULv = sheets["UserLicences"].copy()
            if ULv.empty:
                st.info("No licences assigned yet.")
            else:
                ULv = ULv[ULv["user_id"]==sel_user_id].copy()
                ULv["Licence"] = ULv["licence_id"].map(lambda x: sheets["Licences"].loc[sheets["Licences"]["licence_id"]==x, "licence_name"].values[0])
                st.dataframe(ULv[["Licence","start_date","end_date"]].rename(columns={"start_date":"Start","end_date":"End"}), hide_index=True, use_container_width=True)

            st.markdown("#### Remove a licence")
            if not sheets["UserLicences"].empty:
                UL_user = sheets["UserLicences"][sheets["UserLicences"]["user_id"]==sel_user_id]
                if not UL_user.empty:
                    lic_names_for_user = UL_user["licence_id"].map(lambda x: sheets["Licences"].loc[sheets["Licences"]["licence_id"]==x, "licence_name"].values[0]).tolist()
                    rem_choice = st.selectbox("Licence to remove", lic_names_for_user, key="adm_remove_pick")
                    if st.button("Remove Licence", key="adm_remove"):
                        lic_id_rm = lic_map[rem_choice]
                        UL2 = sheets["UserLicences"]
                        sheets["UserLicences"] = UL2[~((UL2["user_id"]==sel_user_id) & (UL2["licence_id"]==lic_id_rm))]
                        save_db(sheets)
                        st.success("Licence removed.")

        st.markdown("---")
        st.markdown("### Existing Users")
        Udisp = sheets["Users"].copy().rename(columns={"user_id":"User ID","name":"Name"})
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
                M = sheets["Machines"]
                mid = next_id(M, "machine_id")
                req_id = pd.NA if req_lic=="(none)" else lic_map2[req_lic]
                new_m = pd.DataFrame([{
                    "machine_id": mid, "machine_name": m_name.strip(), "machine_type": m_type.strip(),
                    "serial_number": serial.strip(), "required_licence_id": req_id, "status": status,
                    "service_interval_hours": float(service_interval), "last_service_date": pd.Timestamp(last_service_date),
                    "max_duration_hours": float(max_dur)
                }])
                sheets["Machines"] = pd.concat([M, new_m], ignore_index=True)
                save_db(sheets)
                st.success(f"Machine '{m_name}' added with ID {mid}.")

        st.markdown("### Machines")
        M = sheets["Machines"].copy()
        if M.empty:
            st.info("No machines yet.")
        else:
            M["Required Licence"] = M["required_licence_id"].map(lambda x: licence_name(sheets, int(x)) if not pd.isna(x) else "(none)")
            if "max_duration_hours" not in M.columns: M["max_duration_hours"] = 4.0
            Mdisp = M[["machine_id","machine_name","machine_type","serial_number","Required Licence","status","service_interval_hours","last_service_date","max_duration_hours"]].rename(columns={
                "machine_id":"ID","machine_name":"Name","machine_type":"Type","serial_number":"Serial","status":"Status","service_interval_hours":"Service Interval (h)","last_service_date":"Last Service","max_duration_hours":"Max Duration (h)"
            })
            st.dataframe(Mdisp, hide_index=True, use_container_width=True)

    # Maintenance
    with at[2]:
        st.markdown("### Log Service")
        m_map3 = {row["machine_name"]: int(row["machine_id"]) for _, row in sheets["Machines"].sort_values("machine_name").iterrows()}
        if not m_map3:
            st.info("No machines to service.")
        else:
            s_m_name = st.selectbox("Machine", list(m_map3.keys()), key="svc_machine")
            s_m_id = m_map3[s_m_name]
            cur_used = current_hours_since_last_service(sheets, s_m_id)
            st.caption(f"Hours since last service: {cur_used:.1f} h")
            s_date = st.date_input("Service date", value=date.today(), key="svc_date")
            notes = st.text_input("Notes", placeholder="e.g., blades sharpened, bearings checked", key="svc_notes")
            if st.button("Record Service", key="svc_record"):
                SL = sheets["ServiceLog"]
                sid = next_id(SL, "service_id")
                new_s = pd.DataFrame([{"service_id": sid, "machine_id": s_m_id, "date": pd.Timestamp(s_date), "hours_at_service": float(cur_used), "notes": notes.strip()}])
                sheets["ServiceLog"] = pd.concat([SL, new_s], ignore_index=True)
                M = sheets["Machines"]
                idx = M.index[M["machine_id"]==s_m_id]
                if len(idx)>0:
                    M.loc[idx, "last_service_date"] = pd.Timestamp(s_date)
                    sheets["Machines"] = M
                save_db(sheets)
                st.success(f"Service recorded for {s_m_name}.")

        st.divider()
        st.markdown("### Resolve Issues")
        open_issues = sheets["Issues"]
        if open_issues.empty or not (open_issues["status"]=="Open").any():
            st.info("No open issues.")
        else:
            oi = open_issues[open_issues["status"]=="Open"].copy()
            oi["label"] = oi.apply(lambda r: f'#{int(r["issue_id"])} — {sheets["Machines"].loc[sheets["Machines"]["machine_id"]==r["machine_id"], "machine_name"].values[0]} ({r["severity"]})', axis=1)
            chosen_issue = st.selectbox("Select issue to resolve", oi["label"].tolist(), key="res_issue")
            chosen_id = int(chosen_issue.split("—")[0].strip().strip("#"))
            res_notes = st.text_input("Resolution notes", key="res_notes")
            if st.button("Mark Resolved", key="res_btn"):
                I = sheets["Issues"]
                idx = I.index[I["issue_id"]==chosen_id]
                if len(idx)>0:
                    I.loc[idx, "status"] = "Resolved"
                    I.loc[idx, "resolution_notes"] = res_notes.strip()
                    I.loc[idx, "date_resolved"] = pd.Timestamp.now()
                    sheets["Issues"] = I
                    save_db(sheets)
                    st.success(f"Issue #{chosen_id} marked resolved.")

        st.divider()
        st.markdown("### Service Due Soon")
        warn_rows = []
        for _, m in sheets["Machines"].iterrows():
            h_left = hours_until_service(sheets, int(m["machine_id"]))
            if h_left is not None and h_left <= 5.0:
                warn_rows.append({"Machine": m["machine_name"], "Serial": m["serial_number"], "Hours until service": round(h_left,1)})
        if warn_rows:
            st.warning("Some machines are due for service soon (≤ 5.0 hours remaining):")
            st.dataframe(pd.DataFrame(warn_rows), hide_index=True, use_container_width=True)
        else:
            st.success("No machines are close to service.")

    # Data
    with at[3]:
        st.markdown("### Replace/Backup Database")
        up = st.file_uploader("Upload a replacement Excel DB (must match schema)", type=["xlsx"], key="db_upload")
        if st.button("Replace DB from upload", key="db_replace") and up is not None:
            with open(DB_PATH, "wb") as f:
                f.write(up.read())
            st.success("Database replaced. Please refresh the page.")
        st.download_button("Download current DB.xlsx", data=open(DB_PATH,"rb").read(), file_name="db.xlsx")
