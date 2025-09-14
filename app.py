
import streamlit as st
import pandas as pd
from datetime import datetime, date, time, timedelta
from pathlib import Path
import json, re

st.set_page_config(page_title="Woodturners Scheduler", page_icon="🪵", layout="wide")

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "data" / "db.xlsx"
ASSETS = BASE_DIR / "assets"

@st.cache_data(show_spinner=False)
def load_db():
    xls = pd.ExcelFile(DB_PATH, engine="openpyxl")
    sheets = {name: pd.read_excel(DB_PATH, engine="openpyxl", sheet_name=name) for name in xls.sheet_names}
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

def user_licence_ids(sheets, uid):
    UL = sheets.get("UserLicences", pd.DataFrame(columns=["user_id","licence_id","valid_from","valid_to"]))
    today = pd.Timestamp.today().normalize()
    UL["valid_from"] = pd.to_datetime(UL["valid_from"], errors="coerce")
    UL["valid_to"] = pd.to_datetime(UL["valid_to"], errors="coerce")
    L = UL[(UL["user_id"]==uid) & (UL["valid_from"]<=today) & (UL["valid_to"]>=today)]
    return set(L["licence_id"].astype(int).tolist())

def machine_options_for(sheets, uid):
    lids = user_licence_ids(sheets, uid)
    M = sheets["Machines"]
    allowed = M[M["licence_id"].isin(lids)].copy()
    blocked = M[~M["licence_id"].isin(lids)].copy()
    return allowed, blocked

def day_bookings(sheets, mid, day):
    B = sheets["Bookings"].copy()
    B["start"] = pd.to_datetime(B["start"], errors="coerce")
    B["end"] = pd.to_datetime(B["end"], errors="coerce")
    day_start = pd.Timestamp.combine(day, time(0,0))
    day_end   = day_start + timedelta(days=1)
    m = B[(B["machine_id"]==mid) & (B["start"]<day_end) & (B["end"]>day_start)].copy()
    return m.sort_values("start")

def is_open(sheets, day, start_t, end_t):
    CD = sheets.get("ClosedDates", pd.DataFrame(columns=["date","reason"])).copy()
    if not CD.empty:
        CD["date"] = pd.to_datetime(CD["date"], errors="coerce").dt.normalize()
    day_norm = pd.Timestamp(day).normalize()
    if not CD.empty and (CD["date"] == day_norm).any():
        return False, "Closed date"
    OH = sheets.get("OperatingHours", pd.DataFrame(columns=["day_of_week","open_time","close_time"])) 
    dow = pd.Timestamp(day).dayofweek
    row = OH[OH["day_of_week"]==dow]
    if row.empty: return False, "Closed"
    def parse_t(val):
        try:
            if val is None: return None
            if isinstance(val, pd.Timestamp): t=val.to_pydatetime().time(); return t.hour,t.minute
            from datetime import time as T, datetime as D
            if isinstance(val, T): return val.hour,val.minute
            if isinstance(val, D): t=val.time(); return t.hour,t.minute
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

sheets = load_db()

# Header: centred logo (no heading)
c1,c2,c3 = st.columns([1,2,1], vertical_alignment="center")
with c2:
    st.image(str(ASSETS / get_setting(sheets,"active_logo","logo1.png")), use_column_width=True)

# Auth (sidebar)
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

with tabs[0]:
    st.subheader("Book a Machine")
    if not me:
        st.info("Sign in to book.")
    else:
        allowed, blocked = machine_options_for(sheets, int(me["user_id"]))
        opts = [f"{r.machine_id} - {r.machine_name}" for r in allowed.itertuples()]
        st.caption("Machines you can use:")
        st.write(", ".join([r.machine_name for r in allowed.itertuples()]) or "None")
        if len(blocked):
            st.caption("Machines you are not licensed for (greyed out):")
            st.write(", ".join([r.machine_name for r in blocked.itertuples()]))
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
            st.dataframe(day_bookings(sheets, mid, day)[["start","end","purpose","status"]], use_container_width=True, hide_index=True)
            end_dt = datetime.combine(day, start_time) + timedelta(minutes=int(dur))
            ok_hours, msg = is_open(sheets, day, start_time, end_dt.time())
            if not ok_hours: st.error(f"Outside operating hours ({msg}).")
            overlap=False
            for r in day_bookings(sheets, mid, day).itertuples():
                if not (end_dt <= r.start or datetime.combine(day, start_time) >= r.end):
                    overlap=True; break
            if overlap: st.error("Overlaps an existing booking.")
            if st.button("Confirm booking", key="book_go") and ok_hours and not overlap:
                B=sheets["Bookings"]
                new=pd.DataFrame([[int(pd.to_numeric(B["booking_id"], errors="coerce").fillna(0).max())+1 if not B.empty else 1, int(me["user_id"]), mid, datetime.combine(day,start_time), end_dt, "use","", "confirmed"]],
                                 columns=B.columns)
                sheets["Bookings"]=pd.concat([B,new], ignore_index=True); save_db(sheets); st.success("Booked."); st.rerun()

with tabs[1]:
    st.subheader("Calendar")
    M = sheets["Machines"]; sel = st.selectbox("Machine", [f"{r.machine_id} - {r.machine_name}" for r in M.itertuples()], key="cal_m")
    mid=int(sel.split(" - ")[0]); base_day = st.date_input("Day", value=date.today(), key="cal_day")
    view = st.radio("View", ["Day","Week"], horizontal=True, key="cal_view")
    if view=="Day":
        st.dataframe(day_bookings(sheets, mid, base_day)[["start","end","purpose","status"]], use_container_width=True, hide_index=True)
    else:
        start_w = base_day - timedelta(days=base_day.weekday())
        rows=[]
        for d in range(7):
            dd=start_w+timedelta(days=d)
            for r in day_bookings(sheets, mid, dd).itertuples():
                rows.append([dd, r.start.time(), r.end.time(), r.purpose])
        st.dataframe(pd.DataFrame(rows, columns=["day","start","end","purpose"]), use_container_width=True, hide_index=True)

with tabs[2]:
    st.subheader("Issues & Maintenance")
    msel = st.selectbox("Machine", [f"{r.machine_id} - {r.machine_name}" for r in sheets["Machines"].itertuples()], key="iss_m")
    mid = int(msel.split(" - ")[0])
    text = st.text_area("Describe an issue")
    if me and st.button("Submit issue", key="iss_submit"):
        I=sheets["Issues"]; iid=int(pd.to_numeric(I["issue_id"], errors="coerce").fillna(0).max())+1 if not I.empty else 1
        new=pd.DataFrame([[iid, mid, int(me["user_id"]), pd.Timestamp.today(), "open", text]], columns=I.columns)
        sheets["Issues"]=pd.concat([I,new], ignore_index=True); save_db(sheets); st.success("Issue logged."); st.rerun()
    I = sheets["Issues"].copy()
    if I.empty: st.info("No issues yet.")
    else: st.dataframe(I.sort_values("created", ascending=False), use_container_width=True, hide_index=True)

with tabs[3]:
    if not me or me["role"] not in ("admin","superuser"):
        st.info("Admins only.")
    else:
        at = st.tabs(["Users","Licences","User Licences","Machines","Subscriptions","Schedule","Hours & Holidays","Newsletter","Settings","Notifications"])

        with at[0]:
            st.markdown("### Users")
            st.dataframe(sheets["Users"][["user_id","name","role","email","phone","birth_date","joined_date","newsletter_opt_in"]], use_container_width=True, hide_index=True)

        with at[1]:
            st.markdown("### Licences")
            st.dataframe(sheets["Licences"], use_container_width=True, hide_index=True)

        with at[2]:
            st.markdown("### User licencing (assign / revoke)")
            U = sheets["Users"]; L = sheets["Licences"]; UL = sheets["UserLicences"].copy()
            c1,c2,c3 = st.columns([2,2,2])
            with c1:
                ulabel = st.selectbox("Member", [f"{r.user_id} - {r.name}" for r in U.itertuples()], key="ul_user")
            with c2:
                llabel = st.selectbox("Licence", [f"{r.licence_id} - {r.licence_name}" for r in L.itertuples()], key="ul_lic")
            with c3:
                vf = st.date_input("Valid from", key="ul_from"); vt = st.date_input("Valid to", key="ul_to")
            if st.button("Grant licence", key="ul_grant"):
                uid = int(ulabel.split(" - ")[0]); lid = int(llabel.split(" - ")[0])
                new = pd.DataFrame([[uid,lid,pd.Timestamp(vf),pd.Timestamp(vt)]], columns=UL.columns)
                sheets["UserLicences"] = pd.concat([UL,new], ignore_index=True); save_db(sheets); st.success("Licence granted."); st.rerun()
            st.markdown("#### Existing licences")
            if sheets["UserLicences"].empty:
                st.info("No user licences yet.")
            else:
                st.dataframe(sheets["UserLicences"].sort_values(["user_id","licence_id"]), use_container_width=True, hide_index=True)
                del_ix = st.number_input("Delete row #", min_value=0, max_value=len(sheets["UserLicences"])-1, value=0, key="ul_del_ix")
                if st.button("Revoke selected", key="ul_revoke"):
                    UL2 = sheets["UserLicences"].drop(sheets["UserLicences"].index[int(del_ix)]).reset_index(drop=True)
                    sheets["UserLicences"] = UL2; save_db(sheets); st.success("Revoked."); st.rerun()

        with at[3]:
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

        with at[4]:
            st.markdown("### Subscriptions")
            S = sheets["Subscriptions"].copy()
            reasons = sheets.get("DiscountReasons", pd.DataFrame(columns=["reason"]))
            c1,c2,c3,c4 = st.columns([2,1,1,1])
            with c1:
                s_user = st.selectbox("Member", [f"{r.user_id} - {r.name}" for r in sheets["Users"].itertuples()], key="sub_user")
            with c2:
                s_amount = st.number_input("Amount", min_value=0, max_value=1000, value=50, step=5, key="sub_amt")
            with c3:
                s_start = st.date_input("Start", key="sub_start")
            with c4:
                s_end = st.date_input("End", key="sub_end")
            s5,s6 = st.columns([1,1])
            with s5:
                s_paid = st.checkbox("Paid", value=True, key="sub_paid")
            with s6:
                s_reason = st.selectbox("Discount reason", [""] + reasons["reason"].astype(str).tolist(), key="sub_reason")
                s_pct = st.number_input("Discount %", min_value=0, max_value=100, value=0, step=5, key="sub_pct")
            if st.button("Add / update subscription", key="sub_add"):
                uid = int(s_user.split(" - ")[0])
                row = pd.DataFrame([[uid, pd.Timestamp(s_start), pd.Timestamp(s_end), int(s_amount), bool(s_paid), s_reason, int(s_pct)]],
                                   columns=S.columns)
                S = pd.concat([S, row], ignore_index=True)
                sheets["Subscriptions"] = S; save_db(sheets); st.success("Saved subscription."); st.rerun()
            st.markdown("#### Current subscriptions")
            if S.empty: st.info("No subscriptions yet.")
            else: st.dataframe(S.sort_values(["end_date","user_id"], ascending=[True,True]), use_container_width=True, hide_index=True)
            st.markdown("#### Discount reasons")
            dre = sheets.get("DiscountReasons", pd.DataFrame(columns=["reason"])).copy()
            dre_edit = st.data_editor(dre, num_rows="dynamic", hide_index=True, key="dre_edit")
            if st.button("Save discount reasons", key="dre_save"):
                sheets["DiscountReasons"] = dre_edit; save_db(sheets); st.success("Saved reasons."); st.rerun()

        with at[5]:
            st.markdown("### Day roster")
            d = st.date_input("Day", value=date.today(), key="adm_roster_day")
            rows = []
            for m in sheets["Machines"].itertuples():
                for r in day_bookings(sheets, m.machine_id, d).itertuples():
                    rows.append([m.machine_name, r.start.time(), r.end.time(), r.purpose])
            st.dataframe(pd.DataFrame(rows, columns=["Machine","Start","End","Purpose"]), use_container_width=True, hide_index=True)

        with at[6]:
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

        with at[7]:
            st.markdown("### Newsletter")
            T = sheets.get("Templates", pd.DataFrame(columns=["key","text"])).copy()
            row = T[T["key"]=="newsletter_prompt"]
            prompt_text = row.iloc[0]["text"] if not row.empty else ""
            st.markdown("#### Prompt template (editable)")
            new_prompt = st.text_area("Template", value=str(prompt_text), height=400, key="nl_prompt")
            if st.button("Save prompt", key="nl_save_prompt"):
                if row.empty:
                    T = pd.concat([T, pd.DataFrame([["newsletter_prompt", new_prompt]], columns=["key","text"])], ignore_index=True)
                else:
                    T.loc[T["key"]=="newsletter_prompt","text"] = new_prompt
                sheets["Templates"] = T; save_db(sheets); st.success("Prompt saved.")
            st.markdown("#### DATA JSON (auto-compiled)")
            def build_data_json(sheets):
                U = sheets["Users"].copy()
                U["first_name"] = U["name"].str.split().str[0]
                U["last_name"] = U["name"].str.split().str[-1]
                U["suburb"] = U["address"].astype(str).str.split(",").str[0]
                members = []
                for r in U.itertuples():
                    members.append({
                        "first_name": r.first_name, "last_name": r.last_name, "email": r.email,
                        "birth_date": None if pd.isna(r.birth_date) else pd.to_datetime(r.birth_date).date().isoformat(),
                        "joined_date": None if pd.isna(r.joined_date) else pd.to_datetime(r.joined_date).date().isoformat(),
                        "suburb": r.suburb, "opted_in": bool(r.newsletter_opt_in)
                    })
                EV = sheets.get("UserEvents", pd.DataFrame(columns=["event_id","user_id","event_name","event_date","notes"])).copy()
                events = []
                for r in EV.itertuples():
                    email = U.loc[U["user_id"]==r.user_id, "email"]
                    events.append({
                        "member_email": email.iloc[0] if not email.empty else "",
                        "date": None if pd.isna(r.event_date) else pd.to_datetime(r.event_date).date().isoformat(),
                        "type": str(r.event_name),
                        "title": str(r.event_name),
                        "detail": str(getattr(r, "notes", ""))
                    })
                updates = []
                CU = sheets.get("ClubUpdates", pd.DataFrame(columns=["title","text","link"]))
                for r in CU.itertuples():
                    updates.append({"title": str(r.title), "detail": str(getattr(r,"text","") or getattr(r,"detail","")), "link": str(getattr(r,"link",""))})
                notices = []
                NO = sheets.get("Notices", pd.DataFrame(columns=["title","text","link"]))
                for r in NO.itertuples():
                    notices.append({"title": str(r.title), "detail": str(getattr(r,"text","")), "link": str(getattr(r,"link",""))})
                spotlight = []
                SP = sheets.get("SpotlightSubmissions", pd.DataFrame())
                for r in SP.itertuples():
                    nm = U.loc[U["user_id"]==r.user_id, "name"]
                    spotlight.append({"member_name": nm.iloc[0] if not nm.empty else "", "title": str(r.title), "summary": str(r.text), "image_url": str(getattr(r,"image_file","")), "link": ""})
                projects = []
                PJ = sheets.get("ProjectSubmissions", pd.DataFrame())
                for r in PJ.itertuples():
                    email = U.loc[U["user_id"]==r.user_id, "email"]
                    projects.append({"member_email": email.iloc[0] if not email.empty else "", "title": str(r.title), "blurb": str(r.description), "image_urls": [str(getattr(r,"image_file",""))] if getattr(r,"image_file","") else [], "started_date": None, "category": ""})
                mentors = []
                UL = sheets.get("UserLicences", pd.DataFrame())
                L = sheets.get("Licences", pd.DataFrame())
                today = pd.Timestamp.today().normalize()
                active = UL[(pd.to_datetime(UL["valid_from"], errors="coerce")<=today) & (pd.to_datetime(UL["valid_to"], errors="coerce")>=today)]
                for uid in U[U["role"].isin(["admin","superuser"])]["user_id"].tolist():
                    lids = active[active["user_id"]==uid]["licence_id"].astype(int).tolist()
                    skills = L[L["licence_id"].isin(lids)]["licence_name"].astype(str).tolist()
                    mentors.append({"member_email": U.loc[U["user_id"]==uid,"email"].iloc[0], "skills": skills, "availability": "TBC", "location": U.loc[U["user_id"]==uid,"address"].iloc[0], "accepts_beginners": True, "notes": ""})
                MI = sheets.get("MeetingInfo", pd.DataFrame())
                meet = {}
                if not MI.empty:
                    r = MI.iloc[0]
                    meet = {"title": str(r.get("title","")), "date": (None if pd.isna(r.get("date")) else pd.to_datetime(r.get("date")).date().isoformat()), "location": str(r.get("location","")), "agenda_link": str(r.get("agenda_link","")), "rsvp_link": str(r.get("rsvp_link",""))}
                links = {
                    "upload_link": get_setting(sheets,"link_upload","{{upload_link}}"),
                    "mentorship_link": get_setting(sheets,"link_mentorship","{{mentorship_link}}"),
                    "join_link": get_setting(sheets,"link_join","{{join_link}}"),
                    "rsvp_link": get_setting(sheets,"link_rsvp","{{rsvp_link}}"),
                    "unsubscribe_link": (get_setting(sheets,"app_public_url","") + "?unsubscribe=1&uid={{user_id}}") if get_setting(sheets,"app_public_url","") else "{{unsubscribe_link}}",
                }
                last_issue_date = get_setting(sheets,"last_issue_date", "")
                data = {
                    "members": members,
                    "significant_events": events,
                    "club_updates": updates,
                    "notices": notices,
                    "spotlight_submissions": spotlight,
                    "project_submissions": projects,
                    "mentors_offering": mentors,
                    "mentorship_requests": [],
                    "meeting_info": meet,
                    "links": links,
                    "last_issue_date": last_issue_date
                }
                return json.dumps(data, indent=2)
            data_json = build_data_json(sheets)
            st.code(data_json, language="json")
            st.markdown("#### Full prompt (ready to copy to ChatGPT)")
            org_name = get_setting(sheets, "org_name", "Woodturners of the Hunter")
            logo_url = (get_setting(sheets,"app_public_url","").rstrip("/") + "/assets/" + get_setting(sheets,"active_logo","logo1.png")) if get_setting(sheets,"app_public_url","") else "{{logo_url}}"
            compiled = (new_prompt or prompt_text).replace("🔧ORG_NAME", org_name).replace("{DATA_JSON}", data_json).replace("{{logo_url}}", logo_url)
            st.code(compiled, language="markdown")
            st.info("Copy the prompt above into ChatGPT to generate three newsletter drafts.")

        with at[8]:
            st.markdown("### Settings")
            S = sheets.get("Settings").copy()
            def getv(k, default=""):
                m=S[S["key"]==k]
                return default if m.empty else str(m.iloc[0]["value"])
            org = st.text_input("Organisation name", value=getv("org_name","Woodturners of the Hunter"), key="set_org")
            app_url = st.text_input("App public URL (for unsubscribe/logo links)", value=getv("app_public_url",""), key="set_url")
            lock = st.checkbox("Lock booking to signed-in member only", value=(getv("lock_booking_to_member","false").lower() in ("1","true","yes")), key="set_lock")
            logo = st.selectbox("Active logo file", ["logo1.png","logo2.png","logo3.png"], index=["logo1.png","logo2.png","logo3.png"].index(getv("active_logo","logo1.png")), key="set_logo")
            if st.button("Save settings", key="set_save"):
                def upsert(df, k, v):
                    if (df["key"]==k).any():
                        df.loc[df["key"]==k,"value"]=v
                        return df
                    else:
                        return pd.concat([df, pd.DataFrame([[k,v]], columns=["key","value"])], ignore_index=True)
                S = upsert(S, "org_name", org)
                S = upsert(S, "app_public_url", app_url)
                S = upsert(S, "lock_booking_to_member", ("true" if lock else "false"))
                S = upsert(S, "active_logo", logo)
                sheets["Settings"]=S; save_db(sheets); st.success("Settings saved."); st.rerun()

        with at[9]:
            st.markdown("### Notifications")
            today=pd.Timestamp.today().normalize()
            msgs=[]
            try:
                issue_day=int(get_setting(sheets,"newsletter_issue_day",1) or 1)
            except: issue_day=1
            this_issue=pd.Timestamp(year=today.year, month=today.month, day=min(issue_day,28))
            next_issue=this_issue if this_issue>=today else (this_issue + pd.DateOffset(months=1))
            days=(next_issue - today).days
            if 0<=days<=7: msgs.append(f"Newsletter due on {next_issue.date()}")
            S=sheets.get("Subscriptions", pd.DataFrame())
            if not S.empty:
                soon=S[pd.to_datetime(S["end_date"], errors="coerce") <= (today + pd.Timedelta(days=14))]
                if not soon.empty: msgs.append(f"{len(soon)} subscription(s) expiring within 14 days.")
            if msgs: st.write("\\n".join([f"• {m}" for m in msgs]))
            else: st.info("No notifications.")
