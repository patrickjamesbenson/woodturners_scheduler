
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, date, time
from pathlib import Path
import smtplib, ssl
from email.message import EmailMessage

st.set_page_config(page_title="Men's Shed Scheduler", page_icon="🪚", layout="wide")

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "data" / "db.xlsx"

# --------- UI helpers ---------
def _inject_css():
    p = BASE_DIR / "assets" / "styles.css"
    if p.exists():
        st.markdown(f"<style>{p.read_text()}</style>", unsafe_allow_html=True)

def _brand_header():
    for candidate in ["logo.png","logo1.png","logo2.png","logo3.png"]:
        p = BASE_DIR / "assets" / candidate
        if p.exists():
            st.markdown(f'<div class="header"><img src="{p.as_posix()}" alt="Woodturners of the Hunter"></div>', unsafe_allow_html=True)
            return
    st.title("Woodturners of the Hunter")

_inject_css()
_brand_header()

# --------- DB helpers ---------
def load_db():
    import openpyxl  # noqa: F401
    if not DB_PATH.exists():
        st.error("Database missing (data/db.xlsx)."); st.stop()
    x = pd.ExcelFile(DB_PATH, engine="openpyxl")
    sheets = {n: pd.read_excel(DB_PATH, engine="openpyxl", sheet_name=n) for n in x.sheet_names}
    needed = [
        "Users","Licences","UserLicences","Machines","Bookings","OperatingLog",
        "Issues","ServiceLog","OperatingHours","ClosedDates","Settings",
        "AssistanceRequests","MaintenanceRequests",
        "Subscriptions","DiscountReasons","NotificationsLog"
    ]
    for n in needed: sheets.setdefault(n, pd.DataFrame())
    U = sheets["Users"]
    for c in ["phone","email","address","role","password"]:
        if c not in U.columns: U[c] = "" if c != "role" else "user"
    sheets["Users"] = U
    B = sheets["Bookings"]
    if "category" not in B.columns and not B.empty: B["category"] = "Usage"
    # Ensure discount reasons default
    DR = sheets["DiscountReasons"]
    if DR.empty or "reason" not in DR.columns:
        DR = pd.DataFrame([
            {"reason":"Mentor","default_pct":50},
            {"reason":"Lifetime","default_pct":100},
            {"reason":"Workshop-only","default_pct":30},
            {"reason":"Other","default_pct":0},
        ])
    sheets["DiscountReasons"] = DR
    return sheets

def save_db(sheets):
    with pd.ExcelWriter(DB_PATH, engine="openpyxl", mode="w") as w:
        for n, df in sheets.items():
            df.to_excel(w, sheet_name=n, index=False)

def next_id(df, id_col):
    if df.empty or id_col not in df.columns: return 1
    return int(pd.to_numeric(df[id_col], errors="coerce").fillna(0).max()) + 1

def get_setting(sheets, key, default=None):
    S = sheets.get("Settings", pd.DataFrame())
    if S.empty or "key" not in S.columns: return default
    m = S[S["key"]==key]
    if m.empty: return default
    return m.iloc[0]["value"]

def get_setting_bool(sheets, key, default=False):
    v = str(get_setting(sheets, key, default)).strip().lower()
    return v in ("1","true","yes","on")

# --------- Hours / timeslots ---------
def get_operating_hours(sheets):
    OH = sheets.get("OperatingHours", pd.DataFrame())
    if OH.empty:
        default = []
        for i, name in enumerate(["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]):
            if i<5: default.append({"weekday":i,"name":name,"is_open":True,"open":"08:00","close":"17:00"})
            elif i==5: default.append({"weekday":i,"name":name,"is_open":True,"open":"09:00","close":"13:00"})
            else: default.append({"weekday":i,"name":name,"is_open":False,"open":"00:00","close":"00:00"})
        sheets["OperatingHours"] = pd.DataFrame(default); save_db(sheets); OH = sheets["OperatingHours"]
    return {int(r["weekday"]):(bool(r["is_open"]), str(r["open"]), str(r["close"])) for _, r in OH.iterrows()}

def get_closed_dates(sheets):
    CD = sheets.get("ClosedDates", pd.DataFrame())
    if CD.empty: return set()
    CD["date"] = pd.to_datetime(CD["date"], errors="coerce").dt.date
    return set(CD["date"].dropna())

def is_open_on(sheets, d:date):
    oh = get_operating_hours(sheets); cd = get_closed_dates(sheets)
    if d in cd: return False
    is_open, _, _ = oh.get(d.weekday(), (False,"",""))
    return is_open

def timeslots_for_day(sheets, d:date, step=30):
    if not is_open_on(sheets, d): return []
    _, o, c = get_operating_hours(sheets)[d.weekday()]
    oh, om = map(int, o.split(":")); ch, cm = map(int, c.split(":"))
    start = oh*60+om; end = ch*60+cm
    out = []
    m = start
    while m < end:
        h, mm = divmod(m, 60); out.append(time(h,mm)); m += step
    return out

# --------- Roles & auth ---------
def display_name(row):
    role = str(row.get("role","user")).lower()
    suffix = " (admin)" if role=="admin" else (" (superuser)" if role=="superuser" else "")
    return f"{row['name']}{suffix}"

def sign_in_bar(sheets):
    st.markdown('<div class="signin"></div>', unsafe_allow_html=True)
    U = sheets["Users"].copy()
    if U.empty:
        st.warning("No users yet. Add one in Admin → Users."); return
    U["label"] = U.apply(display_name, axis=1)
    names = U["label"].tolist()
    c1,c2,c3,c4,_ = st.columns([2,2,1,1,1])
    with c1:
        chosen = st.selectbox("Sign in: choose your name", names, key="signin_name")
    with c2:
        pw = st.text_input("Password (if required)", type="password", key="signin_pw", value="")
    with c3:
        if st.button("Sign in", key="do_signin"):
            row = U.loc[U["label"]==chosen].iloc[0]
            required = str(row.get("password","")).strip()
            if required and pw != required:
                st.error("Incorrect password for this user.")
            else:
                st.session_state["auth_user_id"] = int(row["user_id"])
                st.session_state["auth_role"] = str(row.get("role","user")).lower()
                st.success(f"Signed in as {row['name']} ({st.session_state['auth_role']}).")
    with c4:
        if st.button("Sign out", key="do_signout"):
            for k in ["auth_user_id","auth_role"]: st.session_state.pop(k, None)
            st.info("Signed out.")

def current_user(sheets):
    uid = st.session_state.get("auth_user_id")
    if uid is None: return None
    U = sheets["Users"]; r = U.loc[U["user_id"]==uid]
    if r.empty: return None
    out = r.iloc[0].to_dict(); out["role"] = str(out.get("role","user")).lower(); return out

def require_role(role:str):
    cur = st.session_state.get("auth_role", None)
    order = {"user":0,"superuser":1,"admin":2}
    return order.get(cur,-1) >= order.get(role,0)

# --------- Booking helpers ---------
def user_licence_ids(sheets, uid:int):
    UL = sheets["UserLicences"]
    if UL.empty: return set()
    return set(UL.loc[UL["user_id"]==uid, "licence_id"].astype(int).tolist())

def licence_name(sheets, lid):
    L = sheets["Licences"]; r = L.loc[L["licence_id"]==lid]
    return r.iloc[0]["licence_name"] if not r.empty else "(unknown)"

def machine_required_licence(sheets, mid):
    M = sheets["Machines"]; r = M.loc[M["machine_id"]==mid]
    if r.empty: return None
    v = r.iloc[0].get("required_licence_id", None)
    return int(v) if pd.notna(v) else None

def machine_max_duration_hours(sheets, mid, default=4.0):
    M = sheets["Machines"]; r = M.loc[M["machine_id"]==mid]
    if r.empty: return default
    try: return float(r.iloc[0].get("max_duration_hours", default)) or default
    except: return default

def hours_since_last_service(sheets, mid):
    OL = sheets["OperatingLog"].copy(); SL = sheets["ServiceLog"].copy()
    if not OL.empty:
        OL["start"] = pd.to_datetime(OL.get("start"), errors="coerce")
        OL["hours"] = pd.to_numeric(OL.get("hours"), errors="coerce").fillna(0.0)
    last = None
    if not SL.empty:
        s = SL[SL["machine_id"]==mid].copy()
        if not s.empty:
            s["date"] = pd.to_datetime(s["date"], errors="coerce")
            s = s.sort_values("date"); last = s.iloc[-1]["date"]
    if last is None:
        used = OL[OL["machine_id"]==mid]["hours"].sum() if not OL.empty else 0.0
    else:
        used = OL[(OL["machine_id"]==mid) & (OL["start"]>=last)]["hours"].sum() if not OL.empty else 0.0
    return float(used)

def hours_until_service(sheets, mid):
    M = sheets["Machines"]; r = M.loc[M["machine_id"]==mid]
    if r.empty: return None
    interval = float(r.iloc[0].get("service_interval_hours", float("nan")))
    if pd.isna(interval): return None
    return interval - hours_since_last_service(sheets, mid)

def bookings_for_machine_on(sheets, mid, d:date):
    B = sheets["Bookings"].copy()
    if B.empty: return pd.DataFrame(columns=B.columns)
    B["start"] = pd.to_datetime(B["start"], errors="coerce"); B["end"] = pd.to_datetime(B["end"], errors="coerce")
    s = pd.to_datetime(d); e = s + pd.Timedelta(days=1)
    return B[(B["machine_id"]==mid) & (B["start"]>=s) & (B["start"]<e)].sort_values("start")

def overlap(a1,a2,b1,b2): return (a1<b2) and (a2>b1)

def prevent_overlap(sheets, mid, sdt, edt):
    B = sheets["Bookings"].copy()
    if B.empty: return True, None
    B["start"] = pd.to_datetime(B["start"]); B["end"] = pd.to_datetime(B["end"])
    B = B[B["machine_id"]==mid]
    for _, r in B.iterrows():
        if overlap(sdt, edt, r["start"], r["end"]):
            return False, r
    return True, None

def add_booking(sheets, uid, mid, sdt, edt, category="Usage"):
    B = sheets["Bookings"]; bid = next_id(B, "booking_id")
    new = pd.DataFrame([{"booking_id":bid,"user_id":uid,"machine_id":mid,"start":sdt,"end":edt,"status":"Confirmed","category":category}])
    sheets["Bookings"] = pd.concat([B, new], ignore_index=True)
    if category=="Usage":
        OL = sheets["OperatingLog"]; op_id = next_id(OL,"op_id")
        dur = (edt-sdt).total_seconds()/3600.0
        new_ol = pd.DataFrame([{"op_id":op_id,"booking_id":bid,"user_id":uid,"machine_id":mid,"start":sdt,"end":edt,"hours":dur}])
        sheets["OperatingLog"] = pd.concat([OL,new_ol], ignore_index=True)
    save_db(sheets); return bid

# --------- Email ---------
def get_admin_email(sheets):
    U = sheets["Users"]
    jb = U[(U["role"].str.lower()=="admin") & (U["name"].str.lower()=="john benson")]
    if not jb.empty: return jb.iloc[0]["email"]
    admins = U[U["role"].str.lower()=="admin"]
    return admins.iloc[0]["email"] if not admins.empty else None

def send_email(subject, body, to_email):
    # Uses Streamlit Secrets for SMTP; otherwise returns False
    try:
        host = st.secrets["SMTP_HOST"]; port = int(st.secrets.get("SMTP_PORT", 587))
        user = st.secrets["SMTP_USER"]; pw = st.secrets["SMTP_PASSWORD"]; from_addr = st.secrets.get("FROM_EMAIL", user)
    except Exception:
        return False, "SMTP secrets not configured"
    try:
        msg = EmailMessage()
        msg["Subject"] = subject; msg["From"] = from_addr; msg["To"] = to_email
        msg.set_content(body)
        ctx = ssl.create_default_context()
        with smtplib.SMTP(host, port) as s:
            s.starttls(context=ctx)
            s.login(user, pw)
            s.send_message(msg)
        return True, "sent"
    except Exception as e:
        return False, str(e)

# --------- Load DB & sign-in ---------
sheets = load_db()
sign_in_bar(sheets)
me = current_user(sheets)

tabs = st.tabs(["Book a Machine","Calendar","Assistance","Issues & Maintenance","Admin"])

# ==== Book ====
with tabs[0]:
    st.subheader("Book a Machine")
    U = sheets["Users"].copy(); U["label"] = U.apply(display_name, axis=1)
    user_map = {row["label"]: int(row["user_id"]) for _, row in U.sort_values("label").iterrows()}
    your_label = st.selectbox("Your name", list(user_map.keys()), key="book_name")
    uid = user_map[your_label]

    lic_ids = user_licence_ids(sheets, uid)
    M = sheets["Machines"]; L = sheets["Licences"]
    allowed = []
    for _, m in M.iterrows():
        req = m.get("required_licence_id")
        if pd.isna(req): allowed.append((f'{m["machine_name"]} (No licence)', int(m["machine_id"])))
        elif int(req) in lic_ids: allowed.append((f'{m["machine_name"]} — requires: {licence_name(sheets,int(req))}', int(m["machine_id"])))
    if not allowed:
        st.warning("No machines available to you with current licences."); st.stop()
    label_to_mid = {lbl: mid for lbl, mid in allowed}
    chosen = st.selectbox("Choose a machine", list(label_to_mid.keys()), key="book_machine")
    mid = label_to_mid[chosen]

    day = st.date_input("Day", value=date.today(), key="book_day")
    if not is_open_on(sheets, day): st.warning("The shed is closed that day."); st.stop()
    slots = timeslots_for_day(sheets, day, 30)
    start_t = st.selectbox("Start time", slots, key="book_start")
    max_h = float(machine_max_duration_hours(sheets, mid)); dur_h = st.slider("Duration (hours)", 0.5, max_h, value=min(1.0,max_h), step=0.5, key="book_dur")
    sdt = datetime.combine(day, start_t); edt = sdt + timedelta(hours=float(dur_h))

    st.markdown("**Existing bookings on this day:**")
    show_contacts = get_setting_bool(sheets, "show_contact_on_bookings", True)
    day_b = bookings_for_machine_on(sheets, mid, day).copy()
    if day_b.empty: st.info("No bookings yet.")
    else:
        day_b["User"] = day_b["user_id"].map(lambda x: sheets["Users"].loc[sheets["Users"]["user_id"]==x, "name"].values[0] if x in sheets["Users"]["user_id"].values else "(system)")
        if show_contacts:
            Utab = sheets["Users"].set_index("user_id")
            day_b["Phone"] = day_b["user_id"].map(lambda x: Utab.loc[x,"phone"] if x in Utab.index else "")
            day_b["Email"] = day_b["user_id"].map(lambda x: Utab.loc[x,"email"] if x in Utab.index else "")
        st.dataframe(day_b[["User","start","end","status","category"] + (["Phone","Email"] if show_contacts else [])].rename(columns={"start":"Start","end":"End","status":"Status"}), hide_index=True, use_container_width=True)

    ok, conflict = prevent_overlap(sheets, mid, sdt, edt)
    if not ok: st.error(f"Overlaps with existing booking from {conflict['start']} to {conflict['end']}.")
    else:
        if st.button("Confirm booking", key="book_confirm"):
            bid = add_booking(sheets, uid, mid, sdt, edt, category="Usage")
            st.success(f"Booking confirmed. Reference #{bid}.")

# ==== Calendar ====
with tabs[1]:
    st.subheader("Calendar (by machine)")
    if me is None: st.info("Please sign in to view the calendar."); st.stop()
    M = sheets["Machines"]; m_map = {row["machine_name"]: int(row["machine_id"]) for _, row in M.sort_values("machine_name").iterrows()}
    m_name = st.selectbox("Machine", list(m_map.keys()), key="cal_m")
    mid = m_map[m_name]
    d = st.date_input("Day", value=date.today(), key="cal_d")
    show_contacts = get_setting_bool(sheets, "show_contact_on_bookings", True)
    day_b = bookings_for_machine_on(sheets, mid, d).copy()
    if day_b.empty: st.info("No bookings for this day.")
    else:
        day_b["User"] = day_b["user_id"].map(lambda x: sheets["Users"].loc[sheets["Users"]["user_id"]==x, "name"].values[0] if x in sheets["Users"]["user_id"].values else "(system)")
        if show_contacts:
            Utab = sheets["Users"].set_index("user_id")
            day_b["Phone"] = day_b["user_id"].map(lambda x: Utab.loc[x,"phone"] if x in Utab.index else "")
            day_b["Email"] = day_b["user_id"].map(lambda x: Utab.loc[x,"email"] if x in Utab.index else "")
        st.dataframe(day_b[["User","start","end","status","category"] + (["Phone","Email"] if show_contacts else [])].rename(columns={"start":"Start","end":"End","status":"Status"}), hide_index=True, use_container_width=True)

# ==== Assistance ====
with tabs[2]:
    st.subheader("Ask for assistance / mentorship")
    if me is None: st.info("Please sign in to request assistance."); st.stop()
    L = sheets["Licences"].sort_values("licence_name")
    lic_map = {row["licence_name"]: int(row["licence_id"]) for _, row in L.iterrows()}
    lic_name = st.selectbox("Area you want help with", list(lic_map.keys()), key="asst_lic")
    lid = lic_map[lic_name]
    UL = sheets["UserLicences"]; U = sheets["Users"]
    mentor_ids = set(UL.loc[UL["licence_id"]==lid, "user_id"].astype(int).tolist())
    mentors = U[(U["user_id"].isin(mentor_ids)) & (U["role"].str.lower()=="superuser")]
    if mentors.empty:
        st.info("No superusers marked for this area yet. Admin can update roles.")
    else:
        mentors = mentors[["name","phone","email"]].rename(columns={"name":"Super user","phone":"Phone","email":"Email"})
        st.dataframe(mentors, hide_index=True, use_container_width=True)
        recipients = ",".join(mentors["Email"].tolist())
        msg = st.text_area("Message", placeholder="Describe what help you need…")
        mailto = f"mailto:{recipients}?subject=Assistance%20request%20for%20{lic_name}&body={msg.replace(' ','%20')}"
        st.markdown(f"[Open email to superusers]({mailto})")
        if st.button("Record request", key="asst_record"):
            AR = sheets["AssistanceRequests"]; rid = next_id(AR, "request_id")
            new = pd.DataFrame([{"request_id":rid,"requester_user_id":int(me['user_id']),"licence_id":int(lid),"message":msg,"created":pd.Timestamp.now(),"status":"Open"}])
            sheets["AssistanceRequests"] = pd.concat([AR,new], ignore_index=True); save_db(sheets); st.success("Request recorded.")

# ==== Issues & Maintenance ====
with tabs[3]:
    st.subheader("Report an issue")
    if me is None: st.info("Please sign in to report issues."); st.stop()
    M = sheets["Machines"]; m_map = {row["machine_name"]: int(row["machine_id"]) for _, row in M.sort_values("machine_name").iterrows()}
    m_name = st.selectbox("Machine", list(m_map.keys()), key="iss_m")
    mid = m_map[m_name]
    txt = st.text_area("Issue description", placeholder="Vibration, sharpening needed, etc.")
    sev = st.selectbox("Severity", ["Low","Medium","High"], key="iss_sev")
    if st.button("Submit issue", key="iss_submit"):
        Issues = sheets["Issues"]; iid = next_id(Issues, "issue_id")
        new = pd.DataFrame([{"issue_id":iid,"machine_id":mid,"user_id":int(me['user_id']),"date_reported":pd.Timestamp.now(),"issue_text":txt.strip(),"severity":sev,"status":"Open","resolution_notes":"","date_resolved":pd.NaT}])
        sheets["Issues"] = pd.concat([Issues,new], ignore_index=True); save_db(sheets); st.success(f"Issue #{iid} logged.")

    st.divider()
    st.subheader("Request maintenance block")
    day = st.date_input("Day", value=date.today(), key="mr_day")
    slots = timeslots_for_day(sheets, day, 30)
    if not slots: st.info("Closed day.")
    else:
        start = st.selectbox("Start", slots, key="mr_start")
        hours = st.slider("Duration (hours)", 0.5, 4.0, step=0.5, key="mr_hours")
        note = st.text_input("Reason/notes", key="mr_notes")
        if st.button("Send request", key="mr_send"):
            MR = sheets["MaintenanceRequests"]; rid = next_id(MR, "request_id")
            new = pd.DataFrame([{"request_id":rid,"user_id":int(me['user_id']),"machine_id":mid,"start":datetime.combine(day,start),"hours":float(hours),"note":note,"status":"Pending"}])
            sheets["MaintenanceRequests"] = pd.concat([MR,new], ignore_index=True); save_db(sheets); st.success("Request sent.")

# ==== Admin ====
with tabs[4]:
    st.subheader("Admin")
    if not require_role("admin"):
        st.info("Admin access only. Sign in as an admin to continue."); st.stop()
    at = st.tabs(["Users & Licences","Machines","Schedule","Maintenance","Subscriptions & Discounts","Notifications","Data & Settings"])

    # Users & Licences
    with at[0]:
        st.markdown("### Add user")
        name = st.text_input("Name", key="adm_new_name")
        phone = st.text_input("Phone", key="adm_new_phone")
        email = st.text_input("Email", key="adm_new_email")
        addr = st.text_area("Address", key="adm_new_addr")
        role = st.selectbox("Role", ["user","superuser","admin"], key="adm_new_role")
        password = st.text_input("Set password (optional)", type="password", key="adm_new_pw")
        if st.button("Add user", key="adm_add_user"):
            U = sheets["Users"]; uid_new = next_id(U, "user_id")
            row = {"user_id":uid_new,"name":name.strip(),"phone":phone.strip(),"email":email.strip(),"address":addr.strip(),"role":role,"password":password.strip()}
            sheets["Users"] = pd.concat([U, pd.DataFrame([row])], ignore_index=True); save_db(sheets); st.success(f"Added {name} (ID {uid_new}).")

        st.markdown("---")
        st.markdown("### Set / change role & password")
        U = sheets["Users"].copy()
        if U.empty: st.info("No users yet.")
        else:
            u_map = {row["name"]: int(row["user_id"]) for _, row in U.sort_values("name").iterrows()}
            uname = st.selectbox("User", list(u_map.keys()), key="adm_edit_user")
            uid_sel = u_map[uname]
            cur_role = str(U.loc[U["user_id"]==uid_sel,"role"].iloc[0] or "user")
            new_role = st.selectbox("Role", ["user","superuser","admin"], index=["user","superuser","admin"].index(cur_role), key="adm_set_role")
            new_pw = st.text_input("New password (blank to keep)", type="password", key="adm_set_pw")
            if st.button("Save user settings", key="adm_save_role"):
                U2 = sheets["Users"]; idx = U2.index[U2["user_id"]==uid_sel]
                if len(idx)>0:
                    U2.loc[idx,"role"] = new_role
                    if new_pw.strip(): U2.loc[idx,"password"] = new_pw.strip()
                    sheets["Users"] = U2; save_db(sheets); st.success("Saved.")

        st.markdown("---")
        st.markdown("### Licences")
        all_lics = sheets["Licences"].sort_values("licence_name"); lic_map = {row["licence_name"]: int(row["licence_id"]) for _, row in all_lics.iterrows()}
        uname2 = st.selectbox("Assign to user", list(u_map.keys()) if 'u_map' in locals() else list({row['name']: int(row['user_id']) for _, row in sheets['Users'].iterrows()}.keys()), key="adm_ul_user")
        uid2 = u_map.get(uname2, sheets["Users"].loc[sheets["Users"]["name"]==uname2, "user_id"].iloc[0])
        licname = st.selectbox("Licence", list(lic_map.keys()), key="adm_ul_lic")
        lid = lic_map[licname]
        start_d = st.date_input("Start date", value=date.today(), key="adm_ul_start")
        end_d = st.date_input("End date (optional)", key="adm_ul_end")
        if st.button("Assign/Update licence", key="adm_ul_add"):
            UL = sheets["UserLicences"].copy(); mask = (UL["user_id"]==uid2) & (UL["licence_id"]==lid)
            if mask.any():
                UL.loc[mask,"start_date"] = pd.Timestamp(start_d); UL.loc[mask,"end_date"] = pd.Timestamp(end_d) if end_d else pd.NaT
            else:
                UL = pd.concat([UL, pd.DataFrame([{"user_id":uid2,"licence_id":lid,"start_date":pd.Timestamp(start_d),"end_date":pd.Timestamp(end_d) if end_d else pd.NaT}])], ignore_index=True)
            sheets["UserLicences"] = UL; save_db(sheets); st.success("Licence saved.")

        st.markdown("#### Current users")
        Udisp = sheets["Users"][["user_id","name","role","phone","email","address"]].rename(columns={"user_id":"ID","name":"Name","role":"Role"})
        st.dataframe(Udisp, hide_index=True, use_container_width=True)

    # Machines
    with at[1]:
        st.markdown("### Add machine")
        m_name = st.text_input("Machine name", key="adm_m_name")
        m_type = st.text_input("Machine type", key="adm_m_type")
        serial = st.text_input("Serial", key="adm_m_serial")
        lic_map2 = {row["licence_name"]: int(row["licence_id"]) for _, row in sheets["Licences"].sort_values("licence_name").iterrows()}
        req = st.selectbox("Required licence", ["(none)"] + list(lic_map2.keys()), key="adm_m_req")
        req_id = pd.NA if req=="(none)" else lic_map2[req]
        status = st.selectbox("Status", ["Active","OutOfService"], key="adm_m_status")
        svc = st.number_input("Service interval (hours)", min_value=1.0, value=50.0, step=1.0, key="adm_m_svc")
        maxd = st.number_input("Max booking duration (hours)", min_value=0.5, value=4.0, step=0.5, key="adm_m_maxd")
        lastsvc = st.date_input("Last service date", value=date.today(), key="adm_m_lsvc")
        if st.button("Add machine", key="adm_m_add"):
            M = sheets["Machines"]; mid_new = next_id(M, "machine_id")
            row = {"machine_id":mid_new,"machine_name":m_name.strip(),"machine_type":m_type.strip(),"serial_number":serial.strip(),"required_licence_id":req_id,"status":status,"service_interval_hours":float(svc),"last_service_date":pd.Timestamp(lastsvc),"max_duration_hours":float(maxd)}
            sheets["Machines"] = pd.concat([M, pd.DataFrame([row])], ignore_index=True); save_db(sheets); st.success("Machine added.")

        Mdisp = sheets["Machines"].copy()
        Mdisp["Required"] = Mdisp["required_licence_id"].map(lambda x: licence_name(sheets,int(x)) if pd.notna(x) else "(none)")
        st.dataframe(Mdisp.rename(columns={"machine_id":"ID","machine_name":"Name","machine_type":"Type","serial_number":"Serial","status":"Status","service_interval_hours":"Service Interval (h)","last_service_date":"Last Service","max_duration_hours":"Max Duration (h)","Required":"Required Licence"}), hide_index=True, use_container_width=True)

    # Schedule with Day/Week + rescheduler
    with at[2]:
        st.markdown("### Day view")
        day_pick = st.date_input("Day", value=date.today(), key="adm_sched_day")
        B = sheets["Bookings"].copy()
        if B.empty: st.info("No bookings yet.")
        else:
            B["start"] = pd.to_datetime(B["start"], errors="coerce"); B["end"] = pd.to_datetime(B["end"], errors="coerce")
            s = pd.to_datetime(day_pick); e = s + pd.Timedelta(days=1)
            D = B[(B["start"]>=s) & (B["start"]<e)].copy()
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

                hours = list(range(6,22))
                counts = []
                for h in hours:
                    h0 = s + pd.Timedelta(hours=h); h1 = h0 + pd.Timedelta(hours=1); cnt = 0
                    for _, r in D.iterrows():
                        if (r["start"]<h1) and (r["end"]>h0) and r.get("category","Usage")=="Usage": cnt += 1
                    counts.append(cnt)
                st.bar_chart(pd.DataFrame({"hour":hours,"bookings":counts}).set_index("hour"))

                # Export day roster CSV
                roster = D[["Machine","User","start","end","Category","status"]]
                csv = roster.to_csv(index=False).encode("utf-8")
                st.download_button("Download day roster (CSV)", data=csv, file_name=f"roster_{day_pick}.csv", mime="text/csv")

        st.markdown("---")
        st.markdown("### Week view")
        ws = st.date_input("Week starting (Monday)", value=(date.today()-timedelta(days=date.today().weekday())), key="adm_week_start")
        if not B.empty:
            B2 = B.copy(); B2["start"] = pd.to_datetime(B2["start"]); B2["end"] = pd.to_datetime(B2["end"]); B2["start_date"] = B2["start"].dt.date

            week_days = [ws + timedelta(days=i) for i in range(7)]
            rows = [{"Date":d, "Total bookings": int((B2["start_date"]==d).sum())} for d in week_days]
            st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

            # Utilization per machine (booked hours / open hours)
            util_rows = []
            M = sheets["Machines"]
            for _, m in M.iterrows():
                mid = int(m["machine_id"]); name = m["machine_name"]
                total_booked = 0.0; total_open = 0.0
                for d in week_days:
                    if is_open_on(sheets, d):
                        _, o, c = get_operating_hours(sheets)[d.weekday()]
                        oh, om = map(int, o.split(":")); ch, cm = map(int, c.split(":"))
                        total_open += ((ch*60+cm)-(oh*60+om))/60.0
                        day_b = B2[(B2["machine_id"]==mid) & (B2["start_date"]==d)]
                        for _, r in day_b.iterrows():
                            hrs = (r["end"] - r["start"]).total_seconds()/3600.0
                            total_booked += hrs if str(r.get("category","Usage"))=="Usage" else 0.0
                util = (total_booked/total_open*100.0) if total_open>0 else 0.0
                util_rows.append({"Machine":name,"Booked hours":round(total_booked,2),"Open hours":round(total_open,2),"Utilization %":round(util,1)})
            util_df = pd.DataFrame(util_rows).sort_values("Utilization %", ascending=False)
            st.dataframe(util_df, hide_index=True, use_container_width=True)
            st.download_button("Download week utilization (CSV)", data=util_df.to_csv(index=False).encode("utf-8"), file_name=f"utilization_{ws}.csv", mime="text/csv")

        st.markdown("---")
        st.markdown("### Reschedule a booking")
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
                max_h = machine_max_duration_hours(sheets, new_mid); new_hours = st.slider("New duration (hours)", 0.5, float(max_h), step=0.5, value=min(1.0,float(max_h)), key="adm_res_hours")
                ns = datetime.combine(new_day, new_start); ne = ns + timedelta(hours=float(new_hours))
                ok, conflict = prevent_overlap(sheets, new_mid, ns, ne)
                if not ok and int(row["machine_id"])!=new_mid: st.error("Overlap with another booking.")
                else:
                    if st.button("Apply reschedule", key="adm_res_apply"):
                        B2 = sheets["Bookings"]; idx = B2.index[B2["booking_id"]==bid]
                        if len(idx)>0:
                            B2.loc[idx,"machine_id"]=new_mid; B2.loc[idx,"start"]=ns; B2.loc[idx,"end"]=ne; sheets["Bookings"]=B2
                            if str(row.get("category","Usage"))=="Usage":
                                OL = sheets["OperatingLog"]; oidx = OL.index[OL["booking_id"]==bid]
                                if len(oidx)>0:
                                    OL.loc[oidx,"machine_id"]=new_mid; OL.loc[oidx,"start"]=ns; OL.loc[oidx,"end"]=ne; OL.loc[oidx,"hours"]=(ne-ns).total_seconds()/3600.0; sheets["OperatingLog"]=OL
                            save_db(sheets); st.success("Rescheduled.")

    # Maintenance admin
    with at[3]:
        st.markdown("### Approve maintenance requests")
        MR = sheets["MaintenanceRequests"]
        if MR.empty or not (MR["status"]=="Pending").any():
            st.info("No pending requests.")
        else:
            MRp = MR[MR["status"]=="Pending"].copy()
            MRp["Machine"] = MRp["machine_id"].map(lambda x: sheets["Machines"].loc[sheets["Machines"]["machine_id"]==x, "machine_name"].values[0])
            MRp["Requester"] = MRp["user_id"].map(lambda x: sheets["Users"].loc[sheets["Users"]["user_id"]==x, "name"].values[0])
            st.dataframe(MRp[["request_id","Machine","Requester","start","hours","note"]], hide_index=True, use_container_width=True)
            rid = st.number_input("Approve request id", min_value=1, step=1, key="adm_appr_rid")
            if st.button("Approve & create maintenance block", key="adm_appr_btn"):
                row = MR.loc[MR["request_id"]==int(rid)]
                if row.empty: st.error("Request id not found.")
                else:
                    r = row.iloc[0]; start = pd.to_datetime(r["start"]); end = start + pd.Timedelta(hours=float(r["hours"]))
                    ok, conflict = prevent_overlap(sheets, int(r["machine_id"]), start, end)
                    if not ok: st.error("Overlaps with an existing booking.")
                    else:
                        bid = add_booking(sheets, 0, int(r["machine_id"]), start, end, category="Maintenance")
                        MR.loc[MR["request_id"]==int(rid),"status"] = "Approved"; sheets["MaintenanceRequests"]=MR; save_db(sheets); st.success(f"Created maintenance booking #{bid}.")

    # Subscriptions & Discounts
    with at[4]:
        st.markdown("### Discount reasons")
        DR = sheets["DiscountReasons"].copy()
        st.dataframe(DR.rename(columns={"reason":"Reason","default_pct":"Default %"}), hide_index=True, use_container_width=True)
        new_reason = st.text_input("Add/Update reason", key="disc_reason")
        new_pct = st.number_input("Default %", min_value=0, max_value=100, value=0, key="disc_pct")
        if st.button("Save reason", key="disc_save"):
            if DR.empty or "reason" not in DR.columns:
                DR = pd.DataFrame([{"reason":new_reason.strip(),"default_pct":int(new_pct)}])
            else:
                if (DR["reason"].str.lower()==new_reason.strip().lower()).any():
                    DR.loc[DR["reason"].str.lower()==new_reason.strip().lower(),"default_pct"]=int(new_pct)
                else:
                    DR = pd.concat([DR, pd.DataFrame([{"reason":new_reason.strip(),"default_pct":int(new_pct)}])], ignore_index=True)
            sheets["DiscountReasons"]=DR; save_db(sheets); st.success("Saved.")

        st.markdown("---")
        st.markdown("### Create / edit subscriptions")
        U = sheets["Users"].sort_values("name"); u_map = {row["name"]: int(row["user_id"]) for _, row in U.iterrows()}
        uname = st.selectbox("Member", list(u_map.keys()), key="sub_user"); uid_sub = u_map[uname]
        amount = st.number_input("Amount ($)", min_value=0.0, step=1.0, value=120.0, key="sub_amount")
        start_d = st.date_input("Start date", value=date.today(), key="sub_start")
        end_d = st.date_input("End date", value=date.today()+timedelta(days=365), key="sub_end")
        reasons = sheets["DiscountReasons"]["reason"].tolist()
        reason = st.selectbox("Discount reason", ["(none)"]+reasons, key="sub_reason")
        pct_default = 0
        if reason!="(none)":
            pct_row = sheets["DiscountReasons"].loc[sheets["DiscountReasons"]["reason"]==reason]
            if not pct_row.empty: pct_default = int(pct_row.iloc[0]["default_pct"])
        pct = st.number_input("Discount %", min_value=0, max_value=100, value=pct_default, key="sub_pct")
        auto_months = st.number_input("Auto-renew (months, 0 = off)", min_value=0, max_value=36, value=0, step=1, key="sub_auto")
        paid = st.checkbox("Paid", value=False, key="sub_paid")
        notes = st.text_input("Notes", key="sub_notes")
        if st.button("Add subscription", key="sub_add"):
            S = sheets["Subscriptions"]; sid = next_id(S, "subscription_id")
            pay_date = pd.Timestamp.today() if bool(paid) else pd.NaT
            new = pd.DataFrame([{"subscription_id":sid,"user_id":uid_sub,"start_date":pd.Timestamp(start_d),"end_date":pd.Timestamp(end_d),"amount":float(amount),"discount_reason":(None if reason=='(none)' else reason),"discount_pct":int(pct),"paid":bool(paid),"payment_date":pay_date,"auto_renew_months":int(auto_months),"notes":notes}])
            sheets["Subscriptions"] = pd.concat([S,new], ignore_index=True); save_db(sheets); st.success(f"Subscription #{sid} added.")
        st.markdown("#### All subscriptions")
        Sdisp = sheets["Subscriptions"].copy()
        if not Sdisp.empty:
            Sdisp["Member"] = Sdisp["user_id"].map(lambda x: sheets["Users"].loc[sheets["Users"]["user_id"]==x, "name"].values[0])
            today = pd.Timestamp.today().normalize()
            Sdisp["Status"] = Sdisp.apply(lambda r: ("Expired" if pd.to_datetime(r["end_date"])<today else "Active"), axis=1)
            cols = ["subscription_id","Member","start_date","end_date","amount","discount_reason","discount_pct","paid","payment_date","auto_renew_months","Status","notes"]
            st.dataframe(Sdisp[cols].rename(columns={"subscription_id":"ID","payment_date":"Paid on","auto_renew_months":"Auto-renew (mo)"}), hide_index=True, use_container_width=True)

            # Controls: mark paid & renew
            st.markdown("##### Actions")
            sid_action = st.number_input("Subscription ID", min_value=1, step=1, key="sub_action_id")
            c1,c2,c3 = st.columns([1,1,2])
            with c1:
                if st.button("Mark paid today", key="sub_mark_paid"):
                    S = sheets["Subscriptions"]; idx = S.index[S["subscription_id"]==int(sid_action)]
                    if len(idx)>0:
                        S.loc[idx,"paid"] = True; S.loc[idx,"payment_date"] = pd.Timestamp.today(); sheets["Subscriptions"]=S; save_db(sheets); st.success("Marked paid.")
                    else: st.error("ID not found.")
            with c2:
                if st.button("Create renewal", key="sub_renew"):
                    S = sheets["Subscriptions"]; row = S.loc[S["subscription_id"]==int(sid_action)]
                    if row.empty: st.error("ID not found.")
                    else:
                        r = row.iloc[0]
                        months = int(r.get("auto_renew_months",0) or 0)
                        if months<=0: st.warning("Auto-renew months is 0; set a value on the subscription first.")
                        else:
                            start = (pd.to_datetime(r["end_date"]) + pd.Timedelta(days=1)).normalize()
                            try:
                                # month arithmetic via pandas DateOffset
                                end = start + pd.DateOffset(months=months) - pd.Timedelta(days=1)
                            except Exception:
                                end = start + pd.Timedelta(days=30*months) - pd.Timedelta(days=1)
                            sid_new = next_id(S, "subscription_id")
                            new = pd.DataFrame([{
                                "subscription_id":sid_new,"user_id":int(r["user_id"]),"start_date":start,"end_date":end,
                                "amount":float(r["amount"]),"discount_reason":r.get("discount_reason",None),"discount_pct":int(r.get("discount_pct",0)),
                                "paid":False,"payment_date":pd.NaT,"auto_renew_months":months,"notes":f"Renewal of #{int(r['subscription_id'])}"
                            }])
                            sheets["Subscriptions"] = pd.concat([S,new], ignore_index=True); save_db(sheets); st.success(f"Created renewal #{sid_new}.")
            # Export CSV
            st.download_button("Download subscriptions (CSV)", data=Sdisp.to_csv(index=False).encode("utf-8"), file_name="subscriptions.csv", mime="text/csv")
        else:
            st.info("No subscriptions yet.")

    # Notifications
    with at[5]:
        st.markdown("### Notification settings")
        admin_email = get_admin_email(sheets)
        st.write(f"Admin email: **{admin_email or '(not set!)'}**")
        days_to_expiry = int(get_setting(sheets, "notify_days_before_subscription_expiry", 30) or 30)
        hours_thresh = float(get_setting(sheets, "notify_hours_before_service", 5) or 5)
        days_upcoming_maint = int(get_setting(sheets, "notify_days_maintenance_window", 7) or 7)
        days_to_expiry_new = st.number_input("Days before subscription expiry", min_value=1, value=days_to_expiry, step=1, key="notif_days_sub")
        hours_thresh_new = st.number_input("Hours remaining threshold for service", min_value=1.0, value=float(hours_thresh), step=1.0, key="notif_hours_serv")
        days_upcoming_new = st.number_input("Days ahead to check for scheduled maintenance", min_value=1, value=days_upcoming_maint, step=1, key="notif_days_maint")
        email_members_toggle = st.checkbox("Also email members whose subscriptions are expiring", value=bool(str(get_setting(sheets, "notify_members_on_subscription_expiry", "false")).lower() in ("1","true","yes","on")), key="notif_email_members")
        if st.button("Save thresholds", key="notif_save"):
            S = sheets.get("Settings", pd.DataFrame(columns=["key","value"]))
            # inline upserts (no nonlocal)
            if S.empty or not (S["key"]=="notify_days_before_subscription_expiry").any():
                S = pd.concat([S, pd.DataFrame([["notify_days_before_subscription_expiry", str(days_to_expiry_new)]], columns=["key","value"])], ignore_index=True)
            else:
                S.loc[S["key"]=="notify_days_before_subscription_expiry","value"] = str(days_to_expiry_new)
            if S.empty or not (S["key"]=="notify_hours_before_service").any():
                S = pd.concat([S, pd.DataFrame([["notify_hours_before_service", str(hours_thresh_new)]], columns=["key","value"])], ignore_index=True)
            else:
                S.loc[S["key"]=="notify_hours_before_service","value"] = str(hours_thresh_new)
            if S.empty or not (S["key"]=="notify_days_maintenance_window").any():
                S = pd.concat([S, pd.DataFrame([["notify_days_maintenance_window", str(days_upcoming_new)]], columns=["key","value"])], ignore_index=True)
            else:
                S.loc[S["key"]=="notify_days_maintenance_window","value"] = str(days_upcoming_new)
                        # upsert member email flag
            if S.empty or not (S["key"]=="notify_members_on_subscription_expiry").any():
                S = pd.concat([S, pd.DataFrame([["notify_members_on_subscription_expiry", str(email_members_toggle)]], columns=["key","value"])], ignore_index=True)
            else:
                S.loc[S["key"]=="notify_members_on_subscription_expiry","value"] = str(email_members_toggle)
            sheets["Settings"]=S; save_db(sheets); st.success("Saved.")

        st.markdown("---")
        st.markdown("### Run notification check now")
        msgs = []

        # Subs expiring soon
        S = sheets["Subscriptions"].copy()
        if not S.empty:
            S["end_date"] = pd.to_datetime(S["end_date"], errors="coerce")
            soon = S[S["end_date"].between(pd.Timestamp.today().normalize(), pd.Timestamp.today().normalize()+pd.Timedelta(days=days_to_expiry_new)) | (S["end_date"]<pd.Timestamp.today().normalize())]
            if not soon.empty:
                soon["Member"] = soon["user_id"].map(lambda x: sheets["Users"].loc[sheets["Users"]["user_id"]==x, "name"].values[0])
                for _, r in soon.iterrows():
                    msgs.append(f"Subscription for {r['Member']} ends {r['end_date'].date()} (paid={r.get('paid',False)})")

        # Machines near service hours
        M = sheets["Machines"]
        for _, row in M.iterrows():
            mid = int(row["machine_id"]); left = hours_until_service(sheets, mid)
            if left is not None and left <= float(hours_thresh_new):
                msgs.append(f"Service due soon for {row['machine_name']} — {left:.1f} hours remaining")

        # Scheduled maintenance in the next X days
        B = sheets["Bookings"].copy()
        if not B.empty:
            B["start"] = pd.to_datetime(B["start"], errors="coerce")
            window_end = pd.Timestamp.today().normalize() + pd.Timedelta(days=days_upcoming_new)
            upcoming = B[(B["category"]=="Maintenance") & (B["start"].between(pd.Timestamp.today().normalize(), window_end))]
            if not upcoming.empty:
                for _, r in upcoming.iterrows():
                    mname = sheets["Machines"].loc[sheets["Machines"]["machine_id"]==r["machine_id"], "machine_name"].values[0]
                    msgs.append(f"Maintenance booking upcoming: {mname} on {r['start']}")

        if st.button("Send email to admin", key="notif_send"):
            if not msgs:
                st.info("No notifications to send.")
            else:
                subject = "Men's Shed — Notifications summary"
                body = "\n".join(f"- {m}" for m in msgs)
                sent="none"
                if admin_email:
                    ok, info = send_email(subject, body, admin_email)
                    if ok: st.success("Email sent to admin."); sent="sent-admin"
                    else:
                        st.warning(f"Could not send admin email ({info}). I'll create a mailto link below.")
                        st.markdown(f"[Open email draft]({'mailto:'+ (admin_email or '') +'?subject='+subject.replace(' ','%20')+'&body='+body.replace(' ','%20')})")
                        sent="mailto-admin"
                else:
                    st.warning("Admin email not found. Please set an admin with a valid email.")
                    sent="no-admin"

                # Optionally email members with expiring subs
                if get_setting_bool(sheets, "notify_members_on_subscription_expiry", False):
                    try:
                        S = sheets["Subscriptions"].copy(); U = sheets["Users"].set_index("user_id")
                        S["end_date"] = pd.to_datetime(S["end_date"], errors="coerce")
                        soon = S[S["end_date"].between(pd.Timestamp.today().normalize(), pd.Timestamp.today().normalize()+pd.Timedelta(days=days_to_expiry_new)) | (S["end_date"]<pd.Timestamp.today().normalize())]
                        for _, r in soon.iterrows():
                            email = U.loc[r["user_id"], "email"] if r["user_id"] in U.index else None
                            if email:
                                subj = "Your membership subscription is expiring"
                                msg = f"Hello {U.loc[r['user_id'],'name']},\n\nYour membership subscription ends on {r['end_date'].date()}. Please renew.\n\nThanks,\nWoodturners of the Hunter"
                                send_email(subj, msg, email)
                    except Exception as e:
                        st.warning(f"Member emails could not be sent ({e}).")

                NL = sheets.get("NotificationsLog", pd.DataFrame(columns=["timestamp","messages","status"]))
                NL = pd.concat([NL, pd.DataFrame([[pd.Timestamp.now(), "\n".join(msgs), sent]], columns=["timestamp","messages","status"])], ignore_index=True)
                sheets["NotificationsLog"]=NL; save_db(sheets)

        st.markdown("#### Log")
        NL = sheets.get("NotificationsLog", pd.DataFrame())
        if NL.empty: st.info("No notifications logged yet.")
        else: st.dataframe(NL, hide_index=True, use_container_width=True)

    # Data & Settings
    with at[6]:
        st.markdown("### Privacy")
        S = sheets.get("Settings", pd.DataFrame(columns=["key","value"]))
        cur = str(get_setting(sheets, "show_contact_on_bookings", "true")).strip().lower()
        toggle = st.checkbox("Show member phone/email on bookings", value=(cur in ("1","true","yes","on")))
        if st.button("Save privacy setting", key="save_priv"):
            if S.empty or not (S["key"]=="show_contact_on_bookings").any():
                S = pd.concat([S, pd.DataFrame([["show_contact_on_bookings", str(toggle)]], columns=["key","value"])], ignore_index=True)
            else:
                S.loc[S["key"]=="show_contact_on_bookings","value"]=str(toggle)
                        # upsert member email flag
            if S.empty or not (S["key"]=="notify_members_on_subscription_expiry").any():
                S = pd.concat([S, pd.DataFrame([["notify_members_on_subscription_expiry", str(email_members_toggle)]], columns=["key","value"])], ignore_index=True)
            else:
                S.loc[S["key"]=="notify_members_on_subscription_expiry","value"] = str(email_members_toggle)
            sheets["Settings"]=S; save_db(sheets); st.success("Saved.")

        st.markdown("### Operating hours")
        OH = sheets.get("OperatingHours", pd.DataFrame())
        if OH.empty:
            # create defaults
            for i, name in enumerate(["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]):
                if i<5:
                    OH = pd.concat([OH, pd.DataFrame([{"weekday":i,"name":name,"is_open":True,"open":"08:00","close":"17:00"}])], ignore_index=True)
                elif i==5:
                    OH = pd.concat([OH, pd.DataFrame([{"weekday":i,"name":name,"is_open":True,"open":"09:00","close":"13:00"}])], ignore_index=True)
                else:
                    OH = pd.concat([OH, pd.DataFrame([{"weekday":i,"name":name,"is_open":False,"open":"00:00","close":"00:00"}])], ignore_index=True)
        for i, name in enumerate(["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]):
            row = OH[OH["weekday"]==i]; is_open = bool(row.iloc[0]["is_open"]) if not row.empty else (i<5)
            c1,c2,c3,c4 = st.columns([1,1,1,5])
            with c1: st.write(name)
            with c2: is_open_new = st.checkbox("Open", value=is_open, key=f"oh_open_{i}")
            with c3: open_time = st.text_input("Open", value=str(row.iloc[0]["open"]) if not row.empty else "08:00", key=f"oh_open_time_{i}")
            with c4: close_time = st.text_input("Close", value=str(row.iloc[0]["close"]) if not row.empty else "17:00", key=f"oh_close_time_{i}")
            if row.empty:
                OH = pd.concat([OH, pd.DataFrame([{"weekday":i,"name":name,"is_open":is_open_new,"open":open_time,"close":close_time}])], ignore_index=True)
            else:
                idx = OH.index[OH["weekday"]==i]; OH.loc[idx,"is_open"]=is_open_new; OH.loc[idx,"open"]=open_time; OH.loc[idx,"close"]=close_time
        sheets["OperatingHours"]=OH

        st.markdown("### Closed dates")
        CD = sheets.get("ClosedDates", pd.DataFrame(columns=["date","reason"]))
        add_cd = st.date_input("Add closed date")
        reason = st.text_input("Reason (optional)")
        if st.button("Add closed date"):
            CD = pd.concat([CD, pd.DataFrame([[pd.Timestamp(add_cd), reason]], columns=["date","reason"])], ignore_index=True)
            sheets["ClosedDates"]=CD; save_db(sheets); st.success("Closed date added.")
        if not CD.empty: st.dataframe(CD, hide_index=True, use_container_width=True)

        st.markdown("---")
        st.markdown("### Replace/Backup Database")
        up = st.file_uploader("Upload a replacement Excel DB (must match schema)", type=["xlsx"], key="db_upload")
        if st.button("Replace DB from upload", key="db_replace") and up is not None:
            (BASE_DIR / "data" / "db.xlsx").write_bytes(up.read()); st.success("Database replaced. Please refresh.")
        st.download_button("Download current DB.xlsx", data=open(DB_PATH,"rb").read(), file_name="db.xlsx")
