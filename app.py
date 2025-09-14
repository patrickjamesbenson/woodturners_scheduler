
import streamlit as st
import pandas as pd
from datetime import datetime, date, time, timedelta
from pathlib import Path
import re, json

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "data" / "db.xlsx"
ASSETS = BASE_DIR / "assets"

st.set_page_config(page_title="Woodturners Scheduler", page_icon="🪵", layout="wide")

@st.cache_data(show_spinner=False)
def load_db():
    xls = pd.ExcelFile(DB_PATH, engine="openpyxl")
    sheets = {name: pd.read_excel(DB_PATH, engine="openpyxl", sheet_name=name) for name in xls.sheet_names}
    expected = {
        "Users":["user_id","name","email","phone","address","role","password","birth_date","joined_date","newsletter_opt_in"],
        "Licences":["licence_id","licence_name"],
        "UserLicences":["user_id","licence_id","valid_from","valid_to"],
        "Machines":["machine_id","machine_name","licence_id","max_duration_minutes","serial_no","next_service_due","hours_used"],
        "Bookings":["booking_id","user_id","machine_id","start","end","purpose","notes","status"],
        "OperatingHours":["day_of_week","open_time","close_time"],
        "ClosedDates":["date","reason"],
        "Issues":["issue_id","machine_id","user_id","created","status","text"],
        "ServiceLog":["service_id","machine_id","date","notes"],
        "OperatingLog":["machine_id","date","hours"],
        "AssistanceRequests":["request_id","requester_user_id","licence_id","message","created"],
        "Subscriptions":["user_id","start_date","end_date","amount","paid","discount_reason","discount_pct"],
        "DiscountReasons":["reason"],
        "UserEvents":["event_id","user_id","event_name","event_date","notes"],
        "Settings":["key","value"],
    }
    for s, cols in expected.items():
        if s not in sheets: sheets[s] = pd.DataFrame(columns=cols)
    # Coerce dates
    for col in ["birth_date","joined_date"]:
        if col in sheets["Users"].columns:
            sheets["Users"][col] = pd.to_datetime(sheets["Users"][col], errors="coerce")
    for s in ["Bookings","ClosedDates","ServiceLog","OperatingLog","UserEvents","Subscriptions"]:
        if s in sheets:
            for c in sheets[s].columns:
                if "date" in c or c in ("start","end","created"):
                    sheets[s][c] = pd.to_datetime(sheets[s][c], errors="coerce")
    return sheets

def save_db(sheets: dict):
    with pd.ExcelWriter(DB_PATH, engine="openpyxl", mode="w") as w:
        for name, df in sheets.items():
            if isinstance(df, pd.DataFrame) and len(df.columns)==0:
                df = pd.DataFrame({"_": []})
            df.to_excel(w, sheet_name=name, index=False)

def get_setting(sheets, key, default=None):
    S = sheets.get("Settings", pd.DataFrame(columns=["key","value"]))
    if S.empty: return default
    m = S[S["key"]==key]
    if m.empty: return default
    v = m.iloc[0]["value"]
    try:
        if pd.isna(v): return default
    except Exception:
        pass
    if v is None: return default
    s = str(v).strip()
    if s.lower() in ("nan","none","null",""): return default
    return s

sheets = load_db()

# Header (logo centered)
c1,c2,c3 = st.columns([1,2,1])
with c2:
    st.image(str(ASSETS / get_setting(sheets,"active_logo","logo1.png")), use_column_width=True)

# Auth (simple)
U = sheets["Users"]
labels = [f"{r.name} ({r.role})" for r in U.itertuples()]
id_by_label = {f"{r.name} ({r.role})": int(r.user_id) for r in U.itertuples()}
st.sidebar.header("Sign in")
label = st.sidebar.selectbox("Your name", [""] + labels, index=0, key="signin_name")
me = None
if label:
    me_id = id_by_label.get(label)
    row = U[U["user_id"]==me_id].iloc[0]
    if row["role"] in ("admin","superuser") and str(row.get("password","")).strip():
        pwd = st.sidebar.text_input("Password", type="password", key="signin_pwd")
        if st.sidebar.button("Sign in", key="signin_btn"):
            if pwd == str(row["password"]):
                st.session_state["me_id"] = int(row["user_id"])
                st.sidebar.success("Signed in")
            else:
                st.sidebar.error("Wrong password")
    else:
        if st.sidebar.button("Continue", key="signin_continue"):
            st.session_state["me_id"] = int(row["user_id"])

if "me_id" in st.session_state:
    me = U[U["user_id"]==st.session_state["me_id"]].iloc[0].to_dict()
    st.sidebar.info(f"Signed in as: {me['name']} ({me['role']})")
else:
    st.sidebar.warning("Select your name and sign in to continue.")

tabs = st.tabs(["Book a Machine","Calendar","Issues & Maintenance","Admin"])

def user_licence_ids(uid):
    UL = sheets["UserLicences"]
    today = pd.Timestamp.today().normalize()
    L = UL[(UL["user_id"]==uid) & (UL["valid_from"]<=today) & (UL["valid_to"]>=today)]
    return set(L["licence_id"].astype(int).tolist())

def machine_options_for(uid):
    lids = user_licence_ids(uid)
    M = sheets["Machines"]
    return M[M["licence_id"].isin(lids)], M[~M["licence_id"].isin(lids)]

def day_bookings(mid, day):
    B = sheets["Bookings"]
    day_start = pd.Timestamp.combine(day, time(0,0))
    day_end   = day_start + timedelta(days=1)
    m = B[(B["machine_id"]==mid) & (B["start"]<day_end) & (B["end"]>day_start)].copy()
    return m.sort_values("start")

def is_open(day, start_t, end_t):
    CD = sheets["ClosedDates"]
    day_norm = pd.Timestamp(day).normalize()
    if not CD.empty:
        cdn = pd.to_datetime(CD["date"], errors="coerce").dt.normalize()
        if (cdn == day_norm).any(): return False, "Closed date"
    OH = sheets["OperatingHours"]; dow = pd.Timestamp(day).dayofweek
    row = OH[OH["day_of_week"]==dow]
    if row.empty: return False, "Closed"
    def parse_t(val):
        try:
            if val is None: return None
            if isinstance(val, pd.Timestamp): t=val.to_pydatetime().time(); return t.hour,t.minute
            from datetime import time as T, datetime as D
            if isinstance(val, T): return val.hour,val.minute
            if isinstance(val, D): t=val.time(); return t.hour,t.minute
            if isinstance(val,(int,float)) and not pd.isna(val):
                mins=int(round((float(val)%1.0)*24*60)); return mins//60, mins%60
            s=str(val).strip()
            if not s: return None
            sL=s.lower()
            ampm=None
            if "am" in sL or "pm" in sL:
                ampm="pm" if "pm" in sL else "am"
                sL=sL.replace("am","").replace("pm","").strip()
            parts=re.split(r'[:h]', sL)
            h=int(parts[0]); m=int(parts[1]) if len(parts)>1 else 0
            if ampm:
                if ampm=="pm" and h!=12: h+=12
                if ampm=="am" and h==12: h=0
            return h,m
        except: return None
    ot=parse_t(row.iloc[0].get("open_time")); ct=parse_t(row.iloc[0].get("close_time"))
    if not ot or not ct: return False, "Closed"
    o_h,o_m=ot; c_h,c_m=ct
    start_min=start_t.hour*60+start_t.minute; end_min=end_t.hour*60+end_t.minute
    open_min=o_h*60+o_m; close_min=c_h*60+c_m
    return (open_min<=start_min) and (end_min<=close_min), f"{o_h:02d}:{o_m:02d}–{c_h:02d}:{c_m:02d}"

def next_booking_id():
    B=sheets["Bookings"]
    return int(pd.to_numeric(B["booking_id"], errors="coerce").fillna(0).max())+1 if not B.empty else 1

with tabs[0]:
    st.subheader("Book a Machine")
    if not me:
        st.info("Sign in to book.")
    else:
        allowed, not_allowed = machine_options_for(int(me["user_id"]))
        opts = [f"{r.machine_id} - {r.machine_name}" for r in allowed.itertuples()]
        if not opts:
            st.error("No active licences found. Ask an admin.")
        else:
            choice = st.selectbox("Machine", opts, key="book_m")
            mid = int(choice.split(" - ")[0])
            day = st.date_input("Day", value=date.today(), key="book_day")
            start_time = st.time_input("Start time", value=time(9,0), key="book_start")
            mrow = sheets["Machines"][sheets["Machines"]["machine_id"]==mid].iloc[0]
            max_mins = int(mrow.get("max_duration_minutes",120) or 120)
            dur = st.slider("Duration (minutes)", 30, max_mins, min(60,max_mins), step=30, key="book_dur")
            st.caption("Availability:")
            st.dataframe(day_bookings(mid, day)[["start","end","purpose","status"]], use_container_width=True, hide_index=True)
            end_dt = datetime.combine(day, start_time) + timedelta(minutes=int(dur))
            ok_hours, msg = is_open(day, start_time, end_dt.time())
            if not ok_hours: st.error(f"Outside operating hours ({msg}).")
            overlap=False
            for r in day_bookings(mid, day).itertuples():
                if not (end_dt <= r.start or datetime.combine(day, start_time) >= r.end):
                    overlap=True; break
            if overlap: st.error("Overlaps an existing booking.")
            if st.button("Confirm booking", key="book_go") and ok_hours and not overlap:
                B=sheets["Bookings"]
                new=pd.DataFrame([[next_booking_id(), int(me["user_id"]), mid, datetime.combine(day,start_time), end_dt, "use","", "confirmed"]],
                                 columns=B.columns)
                sheets["Bookings"]=pd.concat([B,new], ignore_index=True); save_db(sheets); st.success("Booked."); st.rerun()

with tabs[1]:
    st.subheader("Calendar")
    M = sheets["Machines"]; sel = st.selectbox("Machine", [f"{r.machine_id} - {r.machine_name}" for r in M.itertuples()], key="cal_m")
    mid=int(sel.split(" - ")[0]); base_day = st.date_input("Day", value=date.today(), key="cal_day")
    view = st.radio("View", ["Day","Week"], horizontal=True, key="cal_view")
    if view=="Day":
        st.dataframe(day_bookings(mid, base_day)[["start","end","purpose","status"]], use_container_width=True, hide_index=True)
    else:
        start_w = base_day - timedelta(days=base_day.weekday())
        rows=[]
        for d in range(7):
            dd=start_w+timedelta(days=d)
            for r in day_bookings(mid, dd).itertuples():
                rows.append([dd, r.start.time(), r.end.time(), r.purpose])
        st.dataframe(pd.DataFrame(rows, columns=["day","start","end","purpose"]), use_container_width=True, hide_index=True)

with tabs[2]:
    st.subheader("Issues & Maintenance")
    msel = st.selectbox("Machine", [f"{r.machine_id} - {r.machine_name}" for r in sheets["Machines"].itertuples()], key="iss_m")
    mid = int(msel.split(" - ")[0])
    text = st.text_area("Describe an issue")
    if me and st.button("Submit issue"):
        I=sheets["Issues"]; iid=int(pd.to_numeric(I["issue_id"], errors="coerce").fillna(0).max())+1 if not I.empty else 1
        new=pd.DataFrame([[iid, mid, int(me["user_id"]), pd.Timestamp.today(), "open", text]], columns=I.columns)
        sheets["Issues"]=pd.concat([I,new], ignore_index=True); save_db(sheets); st.success("Issue logged.")
    I = sheets["Issues"].copy()
    if I.empty: st.info("No issues yet.")
    else: st.dataframe(I.sort_values("created", ascending=False), use_container_width=True, hide_index=True)

with tabs[3]:
    if not me or me["role"] not in ("admin","superuser"):
        st.info("Admins only.")
    else:
        at = st.tabs(["Users","Licences","Machines","Schedule","Hours & Holidays","Notifications"])

        with at[0]:
            st.markdown("### Users")
            st.dataframe(sheets["Users"][["user_id","name","role","email","phone","birth_date","joined_date","newsletter_opt_in"]], use_container_width=True, hide_index=True)

        with at[1]:
            st.markdown("### Licences")
            st.dataframe(sheets["Licences"], use_container_width=True, hide_index=True)

        with at[2]:
            st.markdown("### Machines")
            st.dataframe(sheets["Machines"], use_container_width=True, hide_index=True)
            st.markdown("#### Inline machines editor (name, serial, next service)")
            M = sheets["Machines"].copy()
            for c in ["machine_id","machine_name","serial_no","next_service_due"]:
                if c not in M.columns: M[c]=None
            try:
                cfg = {
                    "machine_id": st.column_config.NumberColumn("ID", disabled=True),
                    "machine_name": st.column_config.TextColumn("Name"),
                    "serial_no": st.column_config.TextColumn("Serial #"),
                    "next_service_due": st.column_config.DateColumn("Next service"),
                }
            except Exception:
                cfg = None
            edited = st.data_editor(M[["machine_id","machine_name","serial_no","next_service_due"]], num_rows="fixed", hide_index=True, column_config=cfg if cfg else None, key="adm_m_table")
            if st.button("Save machine changes", key="adm_m_save"):
                M2 = M.set_index("machine_id"); E2 = edited.set_index("machine_id")
                for col in ["machine_name","serial_no","next_service_due"]:
                    if col in E2.columns:
                        M2[col]=E2[col]
                M2["next_service_due"]=pd.to_datetime(M2["next_service_due"], errors="coerce")
                sheets["Machines"]=M2.reset_index(); save_db(sheets); st.success("Saved."); st.rerun()

            st.markdown("#### Edit max duration (per machine)")
            em1,em2,em3=st.columns([2,1,1])
            with em1:
                em_sel = st.selectbox("Machine", [f"{r.machine_id} - {r.machine_name}" for r in sheets["Machines"].itertuples()], key="adm_m_sel")
            with em2:
                em_val = st.number_input("Max minutes", 30, 480, int(sheets["Machines"].loc[sheets["Machines"]["machine_id"]==int(em_sel.split(" - ")[0]),"max_duration_minutes"].iloc[0]), step=30, key="adm_m_val")
            with em3:
                if st.button("Save max duration", key="adm_m_set"):
                    M3=sheets["Machines"]; mid=int(em_sel.split(" - ")[0]); M3.loc[M3["machine_id"]==mid,"max_duration_minutes"]=int(em_val); sheets["Machines"]=M3; save_db(sheets); st.success("Saved.")

        with at[3]:
            st.markdown("### Day roster")
            d = st.date_input("Day", value=date.today(), key="adm_roster_day")
            rows = []
            for m in sheets["Machines"].itertuples():
                for r in day_bookings(m.machine_id, d).itertuples():
                    rows.append([m.machine_name, r.start.time(), r.end.time(), r.purpose])
            st.dataframe(pd.DataFrame(rows, columns=["Machine","Start","End","Purpose"]), use_container_width=True, hide_index=True)

        with at[4]:
            st.markdown("### Weekly operating hours")
            day_names=["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
            OH = sheets.get("OperatingHours").copy()
            if OH.empty or set(OH["day_of_week"].tolist()) != set(range(7)):
                rows = []
                for d in range(7):
                    rows.append([d, "09:00" if d in (1,2,3,4,5) else "", "16:00" if d in (1,2,3,4,5) else ""])
                OH = pd.DataFrame(rows, columns=["day_of_week","open_time","close_time"])
            cols = st.columns([1,1,1,1,1,1,1])
            open_vals={}; open_times={}; close_times={}
            for d in range(7):
                with cols[d]:
                    st.write(f"**{day_names[d]}**")
                    ot = str(OH.loc[OH["day_of_week"]==d,"open_time"].iloc[0]) if not OH.empty else ""
                    ct = str(OH.loc[OH["day_of_week"]==d,"close_time"].iloc[0]) if not OH.empty else ""
                    is_open = bool(str(ot).strip()) and bool(str(ct).strip())
                    open_vals[d] = st.checkbox("Open", value=is_open, key=f"hrs_open_{d}")
                    import datetime as _dt
                    def _to_time(s, default_h):
                        try:
                            s = str(s).strip()
                            if not s or s.lower() in ("nan","none","null"): return _dt.time(default_h,0)
                            parts=re.split(r'[:h]', s); h=int(parts[0]); m=int(parts[1]) if len(parts)>1 else 0
                            return _dt.time(h%24, m%60)
                        except: return _dt.time(default_h,0)
                    open_times[d] = st.time_input("Open", value=_to_time(ot,9), key=f"hrs_ot_{d}")
                    close_times[d] = st.time_input("Close", value=_to_time(ct,16), key=f"hrs_ct_{d}")
            if st.button("Copy Tue → Mon–Fri", key="hrs_copy_tue_weekdays"):
                src_open=open_times.get(1); src_close=close_times.get(1); src_flag=open_vals.get(1,False)
                for d2 in range(0,5):
                    open_vals[d2]=bool(src_flag); open_times[d2]=src_open; close_times[d2]=src_close
                st.info("Copied Tuesday hours to Mon–Fri. Click 'Save weekly hours' to persist.")
            if st.button("Copy Tue → Tue–Sat", key="hrs_copy_tue_tuesat"):
                src_open=open_times.get(1); src_close=close_times.get(1); src_flag=open_vals.get(1,False)
                for d2 in range(1,6):
                    open_vals[d2]=bool(src_flag); open_times[d2]=src_open; close_times[d2]=src_close
                st.info("Copied Tuesday hours to Tue–Sat. Click 'Save weekly hours' to persist.")
            cbtn1,cbtn2=st.columns([1,1])
            with cbtn1:
                if st.button("Set weekdays open 09:00–16:00", key="hrs_weekdays_open"):
                    import datetime as _dt
                    for d2 in range(0,5):
                        open_vals[d2]=True; open_times[d2]=_dt.time(9,0); close_times[d2]=_dt.time(16,0)
                    st.info("Weekday hours staged. Click 'Save weekly hours' to persist.")
            with cbtn2:
                if st.button("Close weekdays", key="hrs_weekdays_close"):
                    for d2 in range(0,5): open_vals[d2]=False
                    st.info("Weekdays set to closed (staged). Click 'Save weekly hours' to persist.")
            if st.button("Save weekly hours", key="hrs_save"):
                rows=[]
                for d in range(7):
                    if open_vals[d]:
                        rows.append([d, f"{open_times[d].hour:02d}:{open_times[d].minute:02d}", f"{close_times[d].hour:02d}:{close_times[d].minute:02d}"])
                    else:
                        rows.append([d,"",""])
                sheets["OperatingHours"]=pd.DataFrame(rows, columns=["day_of_week","open_time","close_time"]); save_db(sheets); st.success("Saved.")

            st.divider()
            st.markdown("### Closed dates")
            CD = sheets.get("ClosedDates", pd.DataFrame(columns=["date","reason"])).copy()
            if CD.empty: st.info("No closed dates configured.")
            else: st.dataframe(CD.sort_values("date"), use_container_width=True, hide_index=True)
            nd = st.date_input("Closed date", key="closed_add_date"); nr = st.text_input("Reason", key="closed_add_reason")
            a1,a2=st.columns([1,1])
            with a1:
                if st.button("Add closed date", key="closed_add_btn"):
                    CD = pd.concat([CD, pd.DataFrame([[pd.Timestamp(nd), nr]], columns=["date","reason"])], ignore_index=True); sheets["ClosedDates"]=CD; save_db(sheets); st.success("Added."); st.rerun()
            with a2:
                if not CD.empty:
                    del_ix = st.number_input("Delete row # (see index on left)", min_value=0, max_value=len(CD)-1, value=0, key="closed_del_ix")
                    if st.button("Delete selected", key="closed_del_btn"):
                        CD2 = CD.drop(CD.index[int(del_ix)]).reset_index(drop=True); sheets["ClosedDates"]=CD2; save_db(sheets); st.success("Removed."); st.rerun()

        with at[5]:
            st.markdown("### Notifications")
            today=pd.Timestamp.today().normalize()
            msgs=[]
            # Newsletter due
            try:
                issue_day=int(get_setting(sheets,"newsletter_issue_day",1) or 1)
            except: issue_day=1
            this_issue=pd.Timestamp(year=today.year, month=today.month, day=min(issue_day,28))
            next_issue=this_issue if this_issue>=today else (this_issue + pd.DateOffset(months=1))
            days=(next_issue - today).days
            if 0<=days<=7: msgs.append(f"Newsletter due on {next_issue.date()}")
            # Subscriptions expiring
            S=sheets.get("Subscriptions", pd.DataFrame())
            if not S.empty:
                soon=S[pd.to_datetime(S["end_date"], errors="coerce") <= (today + pd.Timedelta(days=14))]
                if not soon.empty: msgs.append(f"{len(soon)} subscription(s) expiring within 14 days.")
            if msgs: st.write("\\n".join([f"• {m}" for m in msgs]))
            else: st.info("No notifications.")
