
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date, time, timedelta
from pathlib import Path
import json, smtplib, ssl
from email.message import EmailMessage

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "data" / "db.xlsx"
ASSETS = BASE_DIR / "assets"

st.set_page_config(page_title="Woodturners Scheduler", page_icon="🪵", layout="wide")

# ========== Performance helpers ==========
@st.cache_data(show_spinner=False)
def load_db():
    xls = pd.ExcelFile(DB_PATH, engine="openpyxl")
    sheets = {name: pd.read_excel(DB_PATH, engine="openpyxl", sheet_name=name) for name in xls.sheet_names}
    # Ensure expected sheets/columns exist (minimal guards)
    if "Users" not in sheets: sheets["Users"] = pd.DataFrame(columns=["user_id","name","email","phone","address","role","password","birth_date","joined_date","newsletter_opt_in"])
    if "Licences" not in sheets: sheets["Licences"] = pd.DataFrame(columns=["licence_id","licence_name"])
    if "UserLicences" not in sheets: sheets["UserLicences"] = pd.DataFrame(columns=["user_id","licence_id","valid_from","valid_to"])
    if "Machines" not in sheets: sheets["Machines"] = pd.DataFrame(columns=["machine_id","machine_name","licence_id","max_duration_minutes","serial_no","next_service_due","hours_used"])
    if "Bookings" not in sheets: sheets["Bookings"] = pd.DataFrame(columns=["booking_id","user_id","machine_id","start","end","purpose","notes","status"])
    if "OperatingHours" not in sheets: sheets["OperatingHours"] = pd.DataFrame(columns=["day_of_week","open_time","close_time"])
    if "ClosedDates" not in sheets: sheets["ClosedDates"] = pd.DataFrame(columns=["date","reason"])
    if "Issues" not in sheets: sheets["Issues"] = pd.DataFrame(columns=["issue_id","machine_id","user_id","created","status","text"])
    if "ServiceLog" not in sheets: sheets["ServiceLog"] = pd.DataFrame(columns=["service_id","machine_id","date","notes"])
    if "OperatingLog" not in sheets: sheets["OperatingLog"] = pd.DataFrame(columns=["machine_id","date","hours"])
    if "AssistanceRequests" not in sheets: sheets["AssistanceRequests"] = pd.DataFrame(columns=["request_id","requester_user_id","licence_id","message","created"])
    if "MaintenanceRequests" not in sheets: sheets["MaintenanceRequests"] = pd.DataFrame(columns=["request_id","requester_user_id","machine_id","message","created"])
    if "Subscriptions" not in sheets: sheets["Subscriptions"] = pd.DataFrame(columns=["user_id","start_date","end_date","amount","paid","discount_reason","discount_pct"])
    if "DiscountReasons" not in sheets: sheets["DiscountReasons"] = pd.DataFrame(columns=["reason"])
    if "UserEvents" not in sheets: sheets["UserEvents"] = pd.DataFrame(columns=["event_id","user_id","event_name","event_date","notes"])
    if "Newsletters" not in sheets: sheets["Newsletters"] = pd.DataFrame(columns=["newsletter_id","title","date","filename"])
    if "Templates" not in sheets: sheets["Templates"] = pd.DataFrame(columns=["key","text"])
    if "ClubUpdates" not in sheets: sheets["ClubUpdates"] = pd.DataFrame(columns=["update_id","title","text","link","date"])
    if "Notices" not in sheets: sheets["Notices"] = pd.DataFrame(columns=["notice_id","title","text","link","date"])
    if "MeetingInfo" not in sheets: sheets["MeetingInfo"] = pd.DataFrame(columns=["meeting_id","title","date","location","agenda_link","rsvp_link"])
    if "SpotlightSubmissions" not in sheets: sheets["SpotlightSubmissions"] = pd.DataFrame(columns=["submission_id","user_id","title","text","image_file","date","approved"])
    if "ProjectSubmissions" not in sheets: sheets["ProjectSubmissions"] = pd.DataFrame(columns=["submission_id","user_id","title","description","image_file","date"])
    if "Settings" not in sheets: sheets["Settings"] = pd.DataFrame(columns=["key","value"])
    if "NotificationsLog" not in sheets: sheets["NotificationsLog"] = pd.DataFrame(columns=["when","type","message","sent_to"])
    # Coerce dates
    for col in ["birth_date","joined_date"]:
        if col in sheets["Users"].columns:
            sheets["Users"][col] = pd.to_datetime(sheets["Users"][col], errors="coerce")
    for s in ["Bookings","ClosedDates","ServiceLog","OperatingLog","UserEvents","Newsletters","ClubUpdates","Notices","MeetingInfo","Subscriptions"]:
        if s in sheets:
            for c in sheets[s].columns:
                if "date" in c or c in ("start","end","created"):
                    sheets[s][c] = pd.to_datetime(sheets[s][c], errors="coerce")
    return sheets

def save_db(sheets: dict):
    # Write all sheets at once; low frequency only (on explicit actions)
    with pd.ExcelWriter(DB_PATH, engine="openpyxl", mode="w") as w:
        for name, df in sheets.items():
            # Avoid Streamlit crying on empty frames with no columns
            if isinstance(df, pd.DataFrame) and len(df.columns)==0:
                df = pd.DataFrame({"_": []})
            df.to_excel(w, sheet_name=name, index=False)

def get_setting(sheets, key, default=None):
    S = sheets.get("Settings", pd.DataFrame(columns=["key","value"]))
    if S.empty: return default
    m = S[S["key"]==key]
    if m.empty: return default
    return m.iloc[0]["value"]

def set_setting(sheets, key, value):
    S = sheets.get("Settings", pd.DataFrame(columns=["key","value"]))
    if (S["key"]==key).any():
        S.loc[S["key"]==key,"value"] = str(value)
    else:
        S = pd.concat([S, pd.DataFrame([[key, str(value)]], columns=["key","value"])], ignore_index=True)
    sheets["Settings"]=S

def send_email(subject, body, to_email, attachment=None):
    # Will try SMTP secrets if present; otherwise pretend success (so UI isn't blocked)
    host = st.secrets.get("SMTP_HOST", None)
    if not host:
        return True, "SMTP not configured"
    port = int(st.secrets.get("SMTP_PORT", 587))
    user = st.secrets.get("SMTP_USER", None)
    pwd  = st.secrets.get("SMTP_PASSWORD", None)
    from_email = st.secrets.get("FROM_EMAIL", user or "no-reply@example.com")
    try:
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = from_email
        msg["To"] = to_email
        msg.set_content(body)
        if attachment is not None:
            fname, data, mime = attachment
            maintype, subtype = mime.split("/",1)
            msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=fname)
        ctx = ssl.create_default_context()
        with smtplib.SMTP(host, port, timeout=10) as server:
            server.starttls(context=ctx)
            if user and pwd:
                server.login(user, pwd)
            server.send_message(msg)
        return True, "sent"
    except Exception as e:
        return False, str(e)

# ========== Load DB then Header ==========
sheets = load_db()
# ========== Header ==========
col1, col2 = st.columns([1,4])
with col1:
    logo_file = get_setting(sheets, "active_logo", "logo1.png")
    st.image(str(ASSETS / logo_file), use_column_width=True)
with col2:
    st.title("Woodturners Scheduler")

# ========== Auth (simple) ==========
st.sidebar.header("Sign in")
users = sheets["Users"]
name = st.sidebar.selectbox("Your name", [""] + users["name"].tolist(), index=0)
me = None
if name:
    row = users[users["name"]==name].iloc[0]
    if row["role"] in ("admin","superuser") and str(row.get("password","")).strip():
        pwd = st.sidebar.text_input("Password", type="password")
        if st.sidebar.button("Sign in"):
            if pwd == str(row["password"]):
                st.session_state["me_id"] = int(row["user_id"])
                st.sidebar.success("Signed in")
            else:
                st.sidebar.error("Wrong password")
    else:
        if st.sidebar.button("Continue"):
            st.session_state["me_id"] = int(row["user_id"])

if "me_id" in st.session_state:
    me = users[users["user_id"]==st.session_state["me_id"]].iloc[0].to_dict()
    st.sidebar.info(f"Signed in as: {me['name']} ({me['role']})")
else:
    st.sidebar.warning("Select your name and sign in to continue.")

# ========== Tabs ==========
tabs = st.tabs(["Book a Machine","Calendar","My Profile","Assistance","Issues & Maintenance","Admin"])

# ---------- Helpers ----------
def user_licence_ids(uid):
    UL = sheets["UserLicences"]
    today = pd.Timestamp.today().normalize()
    L = UL[(UL["user_id"]==uid) & (UL["valid_from"]<=today) & (UL["valid_to"]>=today)]
    return set(L["licence_id"].astype(int).tolist())

def machine_options_for(uid):
    lids = user_licence_ids(uid)
    M = sheets["Machines"]
    allowed = M[M["licence_id"].isin(lids)].copy()
    return allowed

def day_bookings(machine_id, day):
    B = sheets["Bookings"]
    day_start = pd.Timestamp.combine(day, time(0,0))
    day_end   = day_start + timedelta(days=1)
    m = B[(B["machine_id"]==machine_id) & (B["start"]<day_end) & (B["end"]>day_start)].copy()
    return m.sort_values("start")

def is_open(day, start_t, end_t):
    # Check closed dates and operating hours
    CD = sheets["ClosedDates"]
    if not CD.empty and (pd.to_datetime(day) == pd.to_datetime(CD["date"]).dt.normalize()).any():
        return False, "Closed date"
    OH = sheets["OperatingHours"]
    dow = pd.Timestamp(day).dayofweek
    row = OH[OH["day_of_week"]==dow]
    if row.empty: return False, "Closed"
    open_s = str(row.iloc[0]["open_time"]); close_s = str(row.iloc[0]["close_time"])
    if not open_s or not close_s: return False, "Closed"
    o_h, o_m = map(int, open_s.split(":")); c_h, c_m = map(int, close_s.split(":"))
    return (o_h*60+o_m) <= (start_t.hour*60+start_t.minute) and (end_t.hour*60+end_t.minute) <= (c_h*60+c_m), f"{open_s}–{close_s}"

def next_booking_id():
    B = sheets["Bookings"]
    if B.empty: return 1
    return int(pd.to_numeric(B["booking_id"], errors="coerce").fillna(0).max()) + 1

# ---------- Book a Machine ----------
with tabs[0]:
    st.subheader("Book a Machine")
    if me is None:
        st.info("Sign in to book.")
    else:
        colA, colB = st.columns(2)
        with colA:
            st.write(f"**Member:** {me['name']}")
            # Machines filtered by licence
            allowed = machine_options_for(int(me["user_id"]))
            M = sheets["Machines"]
            # Render list with disabled (greyed) for not licensed
            all_opts = M[["machine_id","machine_name","licence_id","max_duration_minutes"]].copy()
            all_opts["label"] = all_opts.apply(lambda r: f"{r['machine_name']}", axis=1)
            allowed_ids = set(allowed["machine_id"].tolist())
            options = [f"{r.machine_id} - {r.label}" for r in all_opts.itertuples()]
            idx_map = {f"{r.machine_id} - {r.label}": r.machine_id for r in all_opts.itertuples()}
            choice = st.selectbox("Machine", options, index=0, key="book_machine")
            machine_id = idx_map[choice]
            machine_row = M[M["machine_id"]==machine_id].iloc[0]
            if machine_id not in allowed_ids:
                st.warning("You are not licensed for this machine. (Booking will be blocked.)")
        with colB:
            day = st.date_input("Day", value=date.today(), key="book_day")
            st.caption("Availability below updates as you change day/machine.")

        # Show day's bookings
        day_df = day_bookings(machine_id, day)
        if day_df.empty:
            st.success("No bookings yet — all day is available during opening hours.")
        else:
            st.dataframe(day_df[["start","end","purpose","notes","status"]], use_container_width=True, hide_index=True)

        # Duration slider bounded by machine max
        max_dur = int(machine_row.get("max_duration_minutes", 120) or 120)
        dur = st.slider("Duration (minutes)", min_value=30, max_value=max_dur, step=30, value=min(60,max_dur), key="book_dur")

        # Start time picker
        start_time = st.time_input("Start time", value=time(9,0), key="book_start")

        # Validate against opening hours & overlaps
        start_dt = datetime.combine(day, start_time)
        end_dt = start_dt + timedelta(minutes=int(dur))
        ok_hours, hours_msg = is_open(day, start_time, end_dt.time())
        if not ok_hours:
            st.error(f"Outside operating hours ({hours_msg}).")

        overlap = False
        for r in day_df.itertuples():
            if not (end_dt <= r.start or start_dt >= r.end):
                overlap = True; break
        if overlap:
            st.error("Overlaps an existing booking.")

        can_book = (me is not None) and (machine_id in allowed_ids) and ok_hours and (not overlap)

        if st.button("Confirm booking", key="book_confirm"):
            if not can_book:
                st.stop()
            B = sheets["Bookings"]
            new_id = next_booking_id()
            new = pd.DataFrame([[new_id, int(me["user_id"]), int(machine_id), pd.Timestamp(start_dt), pd.Timestamp(end_dt), "use", "", "confirmed"]],
                               columns=["booking_id","user_id","machine_id","start","end","purpose","notes","status"])
            sheets["Bookings"] = pd.concat([B, new], ignore_index=True)
            # Update hours_used rough add
            M = sheets["Machines"]
            idx = M.index[M["machine_id"]==int(machine_id)]
            if len(idx): M.loc[idx, "hours_used"] = (M.loc[idx, "hours_used"].fillna(0) + (dur/60)).values
            sheets["Machines"] = M
            save_db(sheets)
            st.success("Booking confirmed.")
            st.experimental_rerun()

# ---------- Calendar ----------
with tabs[1]:
    st.subheader("Calendar")
    col1, col2 = st.columns([1,2])
    with col1:
        mopts = sheets["Machines"][["machine_id","machine_name"]].copy()
        sel = st.selectbox("Machine", [f"{r.machine_id} - {r.machine_name}" for r in mopts.itertuples()], index=0, key="cal_machine")
        mid = int(sel.split(" - ")[0])
        view = st.radio("View", ["Day","Week"], horizontal=True, key="cal_view")
        base_day = st.date_input("Day", value=date.today(), key="cal_day")
    with col2:
        if view=="Day":
            df = day_bookings(mid, base_day)
            st.dataframe(df[["start","end","purpose","status","notes"]], use_container_width=True, hide_index=True)
        else:
            # Week view
            start_w = base_day - timedelta(days=base_day.weekday())
            rows = []
            for d in range(7):
                day = start_w + timedelta(days=d)
                for r in day_bookings(mid, day).itertuples():
                    rows.append([day, r.start.time(), r.end.time(), r.purpose, r.status])
            W = pd.DataFrame(rows, columns=["day","start","end","purpose","status"])
            st.dataframe(W, use_container_width=True, hide_index=True)

# ---------- My Profile ----------
with tabs[2]:
    st.subheader("My Profile")
    if me is None:
        st.info("Sign in to view your profile.")
    else:
        U = sheets["Users"]
        row = U[U["user_id"]==int(me["user_id"])].iloc[0]
        st.markdown(f"**Name:** {row['name']}  ")
        st.markdown(f"**Phone:** {row.get('phone','')}  ")
        st.markdown(f"**Email:** {row.get('email','')}  ")
        st.markdown(f"**Address:** {row.get('address','')}  ")
        st.markdown(f"**Birth date:** {str(row.get('birth_date',''))}")
        st.divider()
        flag = bool(row.get("newsletter_opt_in", True))
        new_flag = st.checkbox("Subscribed to newsletter", value=flag, key="prof_sub")
        if st.button("Save subscription", key="prof_save_sub"):
            idx = U.index[U["user_id"]==int(me["user_id"])]
            U.loc[idx, "newsletter_opt_in"] = bool(new_flag)
            sheets["Users"] = U; save_db(sheets); st.success("Saved.")
        st.divider()
        st.markdown("### Significant events")
        ev_name = st.text_input("Event name", key="prof_ev_name")
        ev_date = st.date_input("Event date", value=date.today(), key="prof_ev_date")
        ev_notes = st.text_input("Notes", key="prof_ev_notes")
        if st.button("Add event", key="prof_add_ev"):
            UE = sheets.get("UserEvents", pd.DataFrame(columns=["event_id","user_id","event_name","event_date","notes"]))
            eid = int(pd.to_numeric(UE.get("event_id"), errors="coerce").fillna(0).max()) + 1 if not UE.empty else 1
            new = pd.DataFrame([[eid, int(me['user_id']), ev_name.strip(), pd.Timestamp(ev_date), ev_notes.strip()]], columns=["event_id","user_id","event_name","event_date","notes"])
            sheets["UserEvents"] = pd.concat([UE, new], ignore_index=True); save_db(sheets); st.success("Event added.")
        UE = sheets.get("UserEvents", pd.DataFrame())
        mine = UE[UE["user_id"]==int(me["user_id"])]
        if not mine.empty:
            st.dataframe(mine.sort_values("event_date", ascending=False), use_container_width=True, hide_index=True)

# ---------- Assistance ----------
with tabs[3]:
    st.subheader("Ask for assistance / mentorship")
    if me is None:
        st.info("Sign in to send a request.")
    else:
        L = sheets["Licences"]
        lic_name_to_id = {r.licence_name: r.licence_id for r in L.itertuples()}
        sel = st.selectbox("Which skill/machine do you want help with?", ["(general)"] + list(lic_name_to_id.keys()), key="assist_skill")
        msg = st.text_area("Describe what you need help with", key="assist_msg")
        if st.button("Send request", key="assist_send"):
            AR = sheets["AssistanceRequests"]
            rid = int(pd.to_numeric(AR.get("request_id"), errors="coerce").fillna(0).max()) + 1 if not AR.empty else 1
            lic_id = lic_name_to_id.get(sel, None)
            new = pd.DataFrame([[rid, int(me["user_id"]), lic_id, msg, pd.Timestamp.today()]], columns=["request_id","requester_user_id","licence_id","message","created"])
            sheets["AssistanceRequests"] = pd.concat([AR,new], ignore_index=True); save_db(sheets); st.success("Request sent.")

# ---------- Issues & Maintenance ----------
with tabs[4]:
    st.subheader("Issues & Maintenance")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### Log an issue")
        msel = st.selectbox("Machine", [f"{r.machine_id} - {r.machine_name}" for r in sheets["Machines"].itertuples()], key="iss_machine")
        mid = int(msel.split(" - ")[0])
        text = st.text_area("Describe the issue", key="iss_text")
        if me and st.button("Submit issue", key="iss_submit"):
            I = sheets["Issues"]
            iid = int(pd.to_numeric(I.get("issue_id"), errors="coerce").fillna(0).max()) + 1 if not I.empty else 1
            new = pd.DataFrame([[iid, mid, int(me["user_id"]), pd.Timestamp.today(), "open", text]], columns=["issue_id","machine_id","user_id","created","status","text"])
            sheets["Issues"] = pd.concat([I,new], ignore_index=True); save_db(sheets); st.success("Issue logged.")
    with col2:
        st.markdown("### Recent issues")
        I = sheets["Issues"].copy()
        if I.empty:
            st.info("No issues logged.")
        else:
            st.dataframe(I.sort_values("created", ascending=False).head(30), use_container_width=True, hide_index=True)

# ---------- Admin ----------
with tabs[5]:
    if me is None or me["role"] not in ("admin","superuser"):
        st.info("Admins only. Sign in as an admin.")
    else:
        at = st.tabs(["Users","Licences","Machines","Schedule","Newsletter","Notifications"])

        # Users
        with at[0]:
            st.markdown("### Users")
            st.dataframe(sheets["Users"][["user_id","name","email","phone","role","birth_date","joined_date","newsletter_opt_in"]], use_container_width=True, hide_index=True)
            st.markdown("Add a user")
            name = st.text_input("Full name")
            email = st.text_input("Email")
            phone = st.text_input("Phone")
            addr = st.text_input("Address")
            role = st.selectbox("Role", ["user","superuser","admin"])
            pwd = st.text_input("Password (only needed for admin/superuser)")
            bdate = st.date_input("Birth date", value=date(1970,1,1))
            jdate = st.date_input("Joined date", value=date.today())
            if st.button("Create user"):
                U = sheets["Users"]
                new_id = int(pd.to_numeric(U["user_id"], errors="coerce").fillna(0).max()) + 1 if not U.empty else 1
                new = pd.DataFrame([[new_id, name, email, phone, addr, role, pwd if role in ("admin","superuser") else "", bdate, jdate, True]], columns=U.columns)
                sheets["Users"] = pd.concat([U,new], ignore_index=True); save_db(sheets); st.success(f"User {name} added."); st.experimental_rerun()

        # Licences
        with at[1]:
            st.markdown("### Licences")
            st.dataframe(sheets["Licences"], use_container_width=True, hide_index=True)
            st.markdown("Assign licence to user")
            u = st.selectbox("User", [f"{r.user_id} - {r.name}" for r in sheets["Users"].itertuples()])
            l = st.selectbox("Licence", [f"{r.licence_id} - {r.licence_name}" for r in sheets["Licences"].itertuples()])
            vf = st.date_input("Valid from", value=date.today())
            vt = st.date_input("Valid to", value=date.today() + timedelta(days=365))
            if st.button("Assign"):
                UL = sheets["UserLicences"]
                uid = int(u.split(" - ")[0]); lid = int(l.split(" - ")[0])
                new = pd.DataFrame([[uid,lid,vf,vt]], columns=UL.columns)
                sheets["UserLicences"] = pd.concat([UL,new], ignore_index=True); save_db(sheets); st.success("Assigned.")

        # Machines
        with at[2]:
            st.markdown("### Machines")
            st.dataframe(sheets["Machines"], use_container_width=True, hide_index=True)
            st.markdown("Add maintenance booking")
            msel = st.selectbox("Machine", [f"{r.machine_id} - {r.machine_name}" for r in sheets["Machines"].itertuples()], key="adm_msel")
            mid = int(msel.split(" - ")[0])
            mday = st.date_input("Day", value=date.today(), key="adm_mday")
            stime = st.time_input("Start", value=time(9,0), key="adm_mtime")
            dur = st.slider("Duration (minutes)", 30, 240, 60, step=30, key="adm_mdur")
            if st.button("Schedule maintenance"):
                B = sheets["Bookings"]; new_id = int(pd.to_numeric(B["booking_id"], errors="coerce").fillna(0).max()) + 1 if not B.empty else 1
                st_dt = datetime.combine(mday, stime); en_dt = st_dt + timedelta(minutes=int(dur))
                new = pd.DataFrame([[new_id, int(me["user_id"]), mid, st_dt, en_dt, "maintenance", "Scheduled by admin", "confirmed"]], columns=B.columns)
                sheets["Bookings"] = pd.concat([B,new], ignore_index=True); save_db(sheets); st.success("Maintenance scheduled.")

        # Schedule
        with at[3]:
            st.markdown("### Day roster & week utilisation")
            d = st.date_input("Day", value=date.today(), key="adm_roster_day")
            rows = []
            for m in sheets["Machines"].itertuples():
                for r in day_bookings(m.machine_id, d).itertuples():
                    rows.append([m.machine_name, r.start.time(), r.end.time(), r.purpose])
            roster = pd.DataFrame(rows, columns=["Machine","Start","End","Purpose"])
            st.dataframe(roster, use_container_width=True, hide_index=True)
            st.markdown("Week view (booked hours per machine)")
            start_w = d - timedelta(days=d.weekday())
            util = []
            for m in sheets["Machines"].itertuples():
                hours = 0.0
                for i in range(7):
                    dd = start_w + timedelta(days=i)
                    for r in day_bookings(m.machine_id, dd).itertuples():
                        hours += (r.end - r.start).total_seconds()/3600.0
                util.append([m.machine_name, round(hours,1)])
            st.dataframe(pd.DataFrame(util, columns=["Machine","Booked hours (week)"]), use_container_width=True, hide_index=True)

        # Newsletter
        with at[4]:
            st.markdown("### Newsletter")
            c1,c2 = st.columns(2)
            with c1:
                editor_name = st.text_input("Editor name", value=str(get_setting(sheets,"newsletter_editor_name","")))
                editor_email = st.text_input("Editor email", value=str(get_setting(sheets,"newsletter_editor_email","")))
                issue_day = int(get_setting(sheets,"newsletter_issue_day",1) or 1)
                issue_day_new = st.number_input("Issue day (1–28)", 1, 28, issue_day)
            with c2:
                app_url = st.text_input("App public URL", value=str(get_setting(sheets,"app_public_url","")))
                org_name = st.text_input("Organisation name", value=str(get_setting(sheets,"org_name","Woodturners of the Hunter")))
                website_url = st.text_input("Website URL", value=str(get_setting(sheets,"website_url","")))
                postal_address = st.text_input("Postal address", value=str(get_setting(sheets,"postal_address","")))
                logo_url = st.text_input("Logo URL", value=str(get_setting(sheets,"logo_url","")))
            
            st.markdown("Branding")
            logos = ["logo1.png", "logo2.png", "logo3.png"]
            current_logo = get_setting(sheets,"active_logo","logo1.png")
            pick = st.selectbox("Active logo (assets/)", logos, index=logos.index(current_logo) if current_logo in logos else 0)
            if st.button("Use selected logo"):
                set_setting(sheets,"active_logo", pick); save_db(sheets); st.success(f"Logo set to {pick}"); st.experimental_rerun()

            st.markdown("Links")
            cL1, cL2 = st.columns(2)
            with cL1:
                approve_url = st.text_input("Approve URL", value=str(get_setting(sheets,"approve_url","")))
                edit_url = st.text_input("Edit URL", value=str(get_setting(sheets,"edit_url","")))
                market_stall_eoi_link = st.text_input("Market stall EOI link", value=str(get_setting(sheets,"market_stall_eoi_link","")))
            with cL2:
                upload_link = st.text_input("Upload photos link", value=str(get_setting(sheets,"link_upload","")))
                mentorship_link = st.text_input("Mentorship link", value=str(get_setting(sheets,"link_mentorship","")))
                join_link = st.text_input("Join link", value=str(get_setting(sheets,"link_join","")))
                rsvp_link = st.text_input("RSVP link", value=str(get_setting(sheets,"link_rsvp","")))

            if st.button("Save settings"):
                set_setting(sheets,"newsletter_editor_name", editor_name)
                set_setting(sheets,"newsletter_editor_email", editor_email)
                set_setting(sheets,"newsletter_issue_day", issue_day_new)
                set_setting(sheets,"app_public_url", app_url)
                for k,v in [("org_name",org_name),("website_url",website_url),("postal_address",postal_address),("logo_url",logo_url),
                            ("approve_url",approve_url),("edit_url",edit_url),("market_stall_eoi_link",market_stall_eoi_link),
                            ("link_upload",upload_link),("link_mentorship",mentorship_link),("link_join",join_link),("link_rsvp",rsvp_link)]:
                    set_setting(sheets,k,v)
                save_db(sheets); st.success("Saved.")

            st.divider()
            st.markdown("#### Prompt template")
            T = sheets.get("Templates", pd.DataFrame(columns=["key","text"]))
            row = T[T["key"]=="newsletter_prompt"]
            default_prompt = "(set in app)"
            current = row.iloc[0]["text"] if not row.empty else default_prompt
            new_prompt = st.text_area("Template", value=str(current), height=300)
            if st.button("Save prompt"):
                if row.empty:
                    T = pd.concat([T, pd.DataFrame([["newsletter_prompt", new_prompt]], columns=["key","text"])], ignore_index=True)
                else:
                    T.loc[T["key"]=="newsletter_prompt","text"]=new_prompt
                sheets["Templates"]=T; save_db(sheets); st.success("Prompt saved.")

            st.markdown("#### Build DATA.json")
            # Build a compact dataset for the prompt
            U = sheets["Users"].copy()
            U["opted_in"] = U.get("newsletter_opt_in", True)
            members = []
            for r in U.itertuples():
                if bool(r.opted_in) and str(r.email).strip():
                    parts = r.name.split()
                    members.append({
                        "first_name": parts[0],
                        "last_name": parts[-1] if len(parts)>1 else "",
                        "email": r.email,
                        "birth_date": str(r.birth_date.date()) if pd.notna(r.birth_date) else None,
                        "joined_date": str(r.joined_date.date()) if pd.notna(r.joined_date) else None,
                        "suburb": (str(r.address).split(",")[-1].strip() if "," in str(r.address) else ""),
                        "opted_in": True
                    })
            UE = sheets.get("UserEvents", pd.DataFrame())
            events = []
            if not UE.empty:
                for r in UE.itertuples():
                    em = U.loc[U["user_id"]==r.user_id,"email"].values
                    member_email = em[0] if len(em) else ""
                    events.append({
                        "member_email": member_email,
                        "date": str(pd.to_datetime(r.event_date).date()) if pd.notna(r.event_date) else None,
                        "type": "other",
                        "title": r.event_name,
                        "detail": r.notes or ""
                    })
            data = {
                "members": members,
                "significant_events": events,
                "club_updates": sheets.get("ClubUpdates", pd.DataFrame()).to_dict(orient="records"),
                "notices": sheets.get("Notices", pd.DataFrame()).to_dict(orient="records"),
                "spotlight_submissions": sheets.get("SpotlightSubmissions", pd.DataFrame()).to_dict(orient="records"),
                "project_submissions": sheets.get("ProjectSubmissions", pd.DataFrame()).to_dict(orient="records"),
                "mentors_offering": [],
                "mentorship_requests": sheets.get("AssistanceRequests", pd.DataFrame()).to_dict(orient="records"),
                "meeting_info": sheets.get("MeetingInfo", pd.DataFrame()).iloc[-1].to_dict() if not sheets.get("MeetingInfo", pd.DataFrame()).empty else {},
                "links": {
                    "upload_link": get_setting(sheets,"link_upload",""),
                    "mentorship_link": get_setting(sheets,"link_mentorship",""),
                    "join_link": get_setting(sheets,"link_join",""),
                    "rsvp_link": get_setting(sheets,"link_rsvp",""),
                    "unsubscribe_link": (get_setting(sheets,"app_public_url","") + "?unsubscribe=1&uid={{user_id}}") if get_setting(sheets,"app_public_url","") else "{{unsubscribe_link}}"
                },
                "last_issue_date": str(get_setting(sheets,"last_issue_date",""))
            }
            data_json = json.dumps(data, indent=2, default=str)
            st.code(data_json, language="json")
            st.download_button("Download DATA.json", data=data_json.encode("utf-8"), file_name="newsletter_data.json")

            st.markdown("#### Compile prompt")
            tmpl = new_prompt or default_prompt
            compiled = (tmpl.replace("🔧ORG_NAME", get_setting(sheets,"org_name","Woodturners of the Hunter"))
                             .replace("{DATA_JSON}", data_json))
            st.text_area("Compiled prompt", value=compiled, height=300)
            st.download_button("Download compiled_prompt.txt", data=compiled.encode("utf-8"), file_name="compiled_prompt.txt")

        # Notifications
        with at[5]:
            st.markdown("### Notifications")
            msgs = []
            # Newsletter reminder
            try:
                issue_day = int(get_setting(sheets, "newsletter_issue_day", 1) or 1)
            except:
                issue_day = 1
            today = pd.Timestamp.today().normalize()
            this_issue = pd.Timestamp(year=today.year, month=today.month, day=min(issue_day,28))
            next_issue = this_issue if this_issue >= today else (this_issue + pd.DateOffset(months=1))
            days = (next_issue - today).days
            if 0 <= days <= 7:
                msgs.append(f"Newsletter due on {next_issue.date()}")
            # Subscriptions expiring within 14 days
            S = sheets.get("Subscriptions", pd.DataFrame())
            if not S.empty:
                soon = S[pd.to_datetime(S["end_date"], errors="coerce") <= (today + pd.Timedelta(days=14))]
                if not soon.empty:
                    msgs.append(f"{len(soon)} subscription(s) expiring within 14 days.")
            if msgs:
                st.write("\n".join([f"• {m}" for m in msgs]))
            else:
                st.info("No notifications due.")

