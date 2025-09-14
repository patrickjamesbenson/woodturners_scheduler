# app.py — Woodturners Scheduler (clean UI, unique keys, dd/mm/yyyy, mentoring+competency)
import streamlit as st
import pandas as pd
from datetime import datetime, date, time, timedelta
from pathlib import Path
import json, re

# ──────────────────────────────────────────────────────────────────────────────
# App config
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Woodturners Scheduler", page_icon="🪵", layout="wide")
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "data" / "db.xlsx"
ASSETS = BASE_DIR / "assets"
DFMT = "DD/MM/YYYY"

# ──────────────────────────────────────────────────────────────────────────────
# Data helpers
# ──────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_db():
    xls = pd.ExcelFile(DB_PATH, engine="openpyxl")
    return {name: pd.read_excel(DB_PATH, engine="openpyxl", sheet_name=name) for name in xls.sheet_names}

def save_db(sheets: dict):
    with pd.ExcelWriter(DB_PATH, engine="openpyxl", mode="w") as w:
        for name, df in sheets.items():
            if isinstance(df, pd.DataFrame):
                df.to_excel(w, sheet_name=name, index=False)

def get_setting(sheets, key, default=""):
    S = sheets.get("Settings", pd.DataFrame(columns=["key","value"]))
    m = S[S["key"]==key]
    return default if m.empty else ("" if pd.isna(m.iloc[0]["value"]) else str(m.iloc[0]["value"]))

def upsert_setting(sheets, key, value):
    S = sheets.get("Settings", pd.DataFrame(columns=["key","value"])).copy()
    if (S["key"]==key).any():
        S.loc[S["key"]==key, "value"] = value
    else:
        S = pd.concat([S, pd.DataFrame([[key, value]], columns=["key","value"])], ignore_index=True)
    sheets["Settings"] = S

def user_licence_ids(sheets, uid: int):
    UL = sheets.get("UserLicences", pd.DataFrame(columns=["user_id","licence_id","valid_from","valid_to"]))
    if UL.empty: return set()
    UL["valid_from"] = pd.to_datetime(UL["valid_from"], errors="coerce")
    UL["valid_to"]   = pd.to_datetime(UL["valid_to"], errors="coerce")
    today = pd.Timestamp.today().normalize()
    rows = UL[(UL["user_id"]==uid) & (UL["valid_from"]<=today) & (UL["valid_to"]>=today)]
    return set(rows["licence_id"].astype(int).tolist())

def machine_options_for(sheets, uid: int):
    lids = user_licence_ids(sheets, uid)
    M = sheets.get("Machines", pd.DataFrame())
    if M.empty: return M.head(0), M.head(0)
    allowed = M[M["licence_id"].isin(lids)].copy()
    blocked = M[~M["licence_id"].isin(lids)].copy()
    return allowed, blocked

def day_bookings(sheets, mid: int, d: date):
    B = sheets.get("Bookings", pd.DataFrame(columns=["booking_id","user_id","machine_id","start","end","purpose","notes","status"])).copy()
    if B.empty: return B
    B["start"] = pd.to_datetime(B["start"], errors="coerce")
    B["end"]   = pd.to_datetime(B["end"], errors="coerce")
    day_start = pd.Timestamp.combine(d, time(0,0))
    day_end   = day_start + timedelta(days=1)
    m = B[(B["machine_id"]==mid) & (B["start"]<day_end) & (B["end"]>day_start)].copy()
    return m.sort_values("start")

def _parse_time(val):
    """Accept '09:00', '9:00', '9', '9am', '9pm', '09h00'. Return (h,m) or None."""
    try:
        if val is None: return None
        s = str(val).strip()
        if not s or s.lower() in ("nan","none","null"): return None
        sL = s.lower().replace(" ", "")
        ampm = None
        if sL.endswith("am") or sL.endswith("pm"):
            ampm = sL[-2:]
            sL = sL[:-2]
        parts = re.split(r'[:h]', sL)
        h = int(parts[0]); m = int(parts[1]) if len(parts) > 1 else 0
        if ampm == "pm" and h != 12: h += 12
        if ampm == "am" and h == 12: h = 0
        return (h % 24, m % 60)
    except Exception:
        return None

def is_open(sheets, d: date, st_t: time, en_t: time):
    # Closed dates
    CD = sheets.get("ClosedDates", pd.DataFrame(columns=["date","reason"])).copy()
    if not CD.empty:
        CD["date"] = pd.to_datetime(CD["date"], errors="coerce").dt.normalize()
    dn = pd.Timestamp(d).normalize()
    if not CD.empty and (CD["date"] == dn).any():
        return False, "Closed date"

    # Hours
    OH = sheets.get("OperatingHours", pd.DataFrame(columns=["day_of_week","open_time","close_time"]))
    row = OH[OH["day_of_week"] == pd.Timestamp(d).dayofweek]
    if row.empty: return False, "Closed"

    ot = _parse_time(row.iloc[0]["open_time"])
    ct = _parse_time(row.iloc[0]["close_time"])
    if not ot or not ct: return False, "Closed"

    o_h, o_m = ot; c_h, c_m = ct
    st_min = st_t.hour*60 + st_t.minute
    en_min = en_t.hour*60 + en_t.minute
    open_min = o_h*60 + o_m
    close_min = c_h*60 + c_m
    return (open_min <= st_min) and (en_min <= close_min), f"{o_h:02d}:{o_m:02d}–{c_h:02d}:{c_m:02d}"

# ──────────────────────────────────────────────────────────────────────────────
# Load sheets
# ──────────────────────────────────────────────────────────────────────────────
sheets = load_db()

# ──────────────────────────────────────────────────────────────────────────────
# Header — centred logo, no heading text
# ──────────────────────────────────────────────────────────────────────────────
active_logo = get_setting(sheets, "active_logo", "logo1.png")
c1,c2,c3 = st.columns([1,2,1], vertical_alignment="center")
with c2:
    st.image(str(ASSETS / active_logo), use_column_width=True)

# ──────────────────────────────────────────────────────────────────────────────
# Sidebar sign-in (role-aware)
# ──────────────────────────────────────────────────────────────────────────────
U = sheets.get("Users", pd.DataFrame(columns=["user_id","name","role","password"]))
labels = [f"{r.name} ({r.role})" for r in U.itertuples()]
id_by_label = {f"{r.name} ({r.role})": int(r.user_id) for r in U.itertuples()}

st.sidebar.header("Sign in")
label = st.sidebar.selectbox("Your name", [""] + labels, index=0, key="signin_name")
me = None
if label:
    uid = id_by_label[label]
    row = U[U["user_id"]==uid].iloc[0]
    if str(row.get("role","")) in ("admin","superuser") and str(row.get("password","")).strip():
        pwd = st.sidebar.text_input("Password", type="password", key="signin_pwd")
        if st.sidebar.button("Sign in", key="signin_btn"):
            if pwd == str(row["password"]):
                st.session_state["me_id"] = int(uid)
                st.sidebar.success("Signed in")
            else:
                st.sidebar.error("Wrong password")
    else:
        if st.sidebar.button("Continue", key="signin_continue"):
            st.session_state["me_id"] = int(uid)
if "me_id" in st.session_state:
    me = U[U["user_id"]==st.session_state["me_id"]].iloc[0].to_dict()
    st.sidebar.info(f"Signed in as: {me['name']} ({me['role']})")
else:
    st.sidebar.warning("Select your name and sign in to continue.")

# Booking lock toggle (admin setting)
LOCK_BOOK = get_setting(sheets, "lock_booking_to_member", "false").lower() in ("1","true","yes")

# ──────────────────────────────────────────────────────────────────────────────
# Tabs
# ──────────────────────────────────────────────────────────────────────────────
tabs = st.tabs(["Book a Machine","Calendar","Mentoring","Issues & Maintenance","Admin"])

# ──────────────────────────────────────────────────────────────────────────────
# Book a Machine
# ──────────────────────────────────────────────────────────────────────────────
with tabs[0]:
    st.subheader("Book a Machine")
    if not me:
        st.info("Sign in to book.")
    else:
        allowed, blocked = machine_options_for(sheets, int(me["user_id"]))
        if allowed.empty:
            st.error("No active licences found. Ask an admin or submit a mentoring request.")
        else:
            mopts = [f"{r.machine_id} - {r.machine_name}" for r in allowed.itertuples()]
            msel = st.selectbox("Machine", mopts, key="book_m")
            mid = int(msel.split(" - ")[0])

            day = st.date_input("Day", value=date.today(), key="book_day", format=DFMT)
            start_time = st.time_input("Start time", value=time(9,0), key="book_start")

            mrow = sheets["Machines"][sheets["Machines"]["machine_id"]==mid].iloc[0]
            max_mins = int(mrow.get("max_duration_minutes", 120) or 120)
            dur = st.slider("Duration (minutes)", 30, max_mins, min(60, max_mins), step=30, key="book_dur")

            st.caption("Availability:")
            st.dataframe(day_bookings(sheets, mid, day)[["start","end","purpose","status"]],
                         use_container_width=True, hide_index=True)

            start_dt = datetime.combine(day, start_time)
            end_dt   = start_dt + timedelta(minutes=int(dur))
            ok_hours, hours_msg = is_open(sheets, day, start_time, end_dt.time())

            if not ok_hours:
                st.error(f"Outside operating hours ({hours_msg}).")

            overlap = False
            for r in day_bookings(sheets, mid, day).itertuples():
                if not (end_dt <= r.start or start_dt >= r.end):
                    overlap = True
                    break
            if overlap:
                st.error("Overlaps an existing booking.")

            clicked = st.button("Confirm booking", key="book_go", type="primary",
                                disabled=not(ok_hours and (not overlap)))
            if clicked and ok_hours and (not overlap):
                B = sheets.get("Bookings", pd.DataFrame(columns=["booking_id","user_id","machine_id","start","end","purpose","notes","status"]))
                next_id = 1 if B.empty else int(pd.to_numeric(B["booking_id"], errors="coerce").fillna(0).max())+1
                user_id = int(me["user_id"]) if LOCK_BOOK or not B.empty else int(me["user_id"])
                new = pd.DataFrame([[next_id, user_id, mid, start_dt, end_dt, "use", "", "confirmed"]], columns=B.columns)
                sheets["Bookings"] = pd.concat([B, new], ignore_index=True)
                save_db(sheets)
                st.success("Booked.")
                st.rerun()

        # Show machines you cannot use (greyed out conceptually)
        if not blocked.empty:
            st.caption("Machines you are not licensed for:")
            st.write(", ".join(blocked["machine_name"].astype(str).tolist()))

# ──────────────────────────────────────────────────────────────────────────────
# Calendar
# ──────────────────────────────────────────────────────────────────────────────
with tabs[1]:
    st.subheader("Calendar")
    if sheets.get("Machines", pd.DataFrame()).empty:
        st.info("No machines configured.")
    else:
        cal_opts = [f"{r.machine_id} - {r.machine_name}" for r in sheets["Machines"].itertuples()]
        csel = st.selectbox("Machine", cal_opts, key="cal_m")
        cmid = int(csel.split(" - ")[0])
        base_day = st.date_input("Day", value=date.today(), key="cal_day", format=DFMT)
        view = st.radio("View", ["Day","Week"], horizontal=True, key="cal_view")
        if view == "Day":
            st.dataframe(day_bookings(sheets, cmid, base_day)[["start","end","purpose","status"]],
                         use_container_width=True, hide_index=True)
        else:
            start_w = base_day - timedelta(days=base_day.weekday())
            rows = []
            for d in range(7):
                dd = start_w + timedelta(days=d)
                for r in day_bookings(sheets, cmid, dd).itertuples():
                    rows.append([dd, r.start.time(), r.end.time(), r.purpose, r.status])
            dfw = pd.DataFrame(rows, columns=["day","start","end","purpose","status"])
            st.dataframe(dfw, use_container_width=True, hide_index=True)

# ──────────────────────────────────────────────────────────────────────────────
# Mentoring (member request) — plus your past requests
# ──────────────────────────────────────────────────────────────────────────────
with tabs[2]:
    st.subheader("Mentoring & Competency Requests")
    if not me:
        st.info("Sign in to request mentoring.")
    else:
        L = sheets.get("Licences", pd.DataFrame(columns=["licence_id","licence_name"]))
        lic_map = {f"{r.licence_id} - {r.licence_name}": int(r.licence_id) for r in L.itertuples()}
        if not lic_map:
            st.info("No licences defined yet.")
        else:
            sel = st.selectbox("Skill / Machine licence you want help with", list(lic_map.keys()), key="ment_lic")
            msg = st.text_area("What do you need help with? (optional)", key="ment_msg")

            # Suggested mentors (admins/superusers with active licence)
            UL = sheets.get("UserLicences", pd.DataFrame())
            Utab = sheets.get("Users", pd.DataFrame())
            today = pd.Timestamp.today().normalize()
            if not UL.empty:
                UL["valid_from"] = pd.to_datetime(UL["valid_from"], errors="coerce")
                UL["valid_to"]   = pd.to_datetime(UL["valid_to"], errors="coerce")
            lic_id = lic_map[sel]
            mentors_ids = []
            if not UL.empty:
                mentors_ids = UL[(UL["licence_id"]==lic_id) & (UL["valid_from"]<=today) & (UL["valid_to"]>=today)]["user_id"].astype(int).tolist()
            mentors = Utab[(Utab["user_id"].isin(mentors_ids)) & (Utab["role"].isin(["admin","superuser"]))][["name","email","phone"]]
            if mentors.empty:
                st.warning("No listed mentors for this skill yet.")
            else:
                st.markdown("**Suggested mentors for this skill:**")
                st.dataframe(mentors, hide_index=True, use_container_width=True)

            if st.button("Submit mentoring request", type="primary", key="ment_submit"):
                AR = sheets.get("AssistanceRequests", pd.DataFrame(columns=["request_id","requester_user_id","licence_id","message","created","status","handled_by","handled_on","outcome","notes"]))
                req_id = 1 if AR.empty else int(pd.to_numeric(AR["request_id"], errors="coerce").fillna(0).max())+1
                new = pd.DataFrame([[req_id, int(me["user_id"]), lic_id, msg, pd.Timestamp.today(), "open", None, None, None, None]], columns=AR.columns)
                sheets["AssistanceRequests"] = pd.concat([AR, new], ignore_index=True)
                save_db(sheets)
                st.success("Request submitted. A mentor/admin will contact you.")
                st.rerun()

    st.markdown("#### Your past requests")
    if me:
        AR2 = sheets.get("AssistanceRequests", pd.DataFrame())
        mine = AR2[AR2["requester_user_id"]==int(me["user_id"])]
        if mine.empty: st.info("No requests yet.")
        else: st.dataframe(mine.sort_values("created", ascending=False), hide_index=True, use_container_width=True)

# ──────────────────────────────────────────────────────────────────────────────
# Issues & Maintenance (member logging)
# ──────────────────────────────────────────────────────────────────────────────
with tabs[3]:
    st.subheader("Issues & Maintenance")
    if sheets.get("Machines", pd.DataFrame()).empty:
        st.info("No machines configured.")
    else:
        isel = st.selectbox("Machine", [f"{r.machine_id} - {r.machine_name}" for r in sheets["Machines"].itertuples()], key="iss_m")
        imid = int(isel.split(" - ")[0])
        itxt = st.text_area("Describe an issue", key="iss_text")
        if me and st.button("Submit issue", key="iss_submit"):
            I = sheets.get("Issues", pd.DataFrame(columns=["issue_id","machine_id","user_id","created","status","text"]))
            iid = 1 if I.empty else int(pd.to_numeric(I["issue_id"], errors="coerce").fillna(0).max())+1
            new = pd.DataFrame([[iid, imid, int(me["user_id"]), pd.Timestamp.today(), "open", itxt]], columns=I.columns)
            sheets["Issues"] = pd.concat([I, new], ignore_index=True)
            save_db(sheets)
            st.success("Issue logged.")
            st.rerun()

    Ishow = sheets.get("Issues", pd.DataFrame())
    if Ishow.empty: st.info("No issues yet.")
    else: st.dataframe(Ishow.sort_values("created", ascending=False), use_container_width=True, hide_index=True)

# ──────────────────────────────────────────────────────────────────────────────
# Admin area
# ──────────────────────────────────────────────────────────────────────────────
with tabs[4]:
    if not me or str(me.get("role","")) not in ("admin","superuser"):
        st.info("Admins only.")
    else:
        at = st.tabs(["Users","Licences","User Licences","Competency","Machines","Subscriptions","Schedule","Hours & Holidays","Newsletter","Settings","Notifications"])

        # Users
        with at[0]:
            st.markdown("### Users")
            DF = sheets.get("Users", pd.DataFrame())
            st.dataframe(DF[["user_id","name","role","email","phone","birth_date","joined_date","newsletter_opt_in"]] if not DF.empty else DF,
                         use_container_width=True, hide_index=True)

        # Licences
        with at[1]:
            st.markdown("### Licences")
            st.dataframe(sheets.get("Licences", pd.DataFrame()), use_container_width=True, hide_index=True)

        # User Licences
        with at[2]:
            st.markdown("### User licencing (assign / revoke)")
            Utab = sheets.get("Users", pd.DataFrame())
            Ltab = sheets.get("Licences", pd.DataFrame())
            UL = sheets.get("UserLicences", pd.DataFrame(columns=["user_id","licence_id","valid_from","valid_to"])).copy()
            c1,c2,c3 = st.columns([2,2,2])
            with c1:
                ulabel = st.selectbox("Member", [f"{r.user_id} - {r.name}" for r in Utab.itertuples()], key="ul_user")
            with c2:
                llabel = st.selectbox("Licence", [f"{r.licence_id} - {r.licence_name}" for r in Ltab.itertuples()], key="ul_lic")
            with c3:
                vf = st.date_input("Valid from", key="ul_from", format=DFMT)
                vt = st.date_input("Valid to", key="ul_to", format=DFMT)
            if st.button("Grant licence", key="ul_grant"):
                uid = int(ulabel.split(" - ")[0]); lid = int(llabel.split(" - ")[0])
                new = pd.DataFrame([[uid, lid, pd.Timestamp(vf), pd.Timestamp(vt)]], columns=UL.columns)
                sheets["UserLicences"] = pd.concat([UL, new], ignore_index=True)
                save_db(sheets); st.success("Licence granted."); st.rerun()

            st.markdown("#### Existing licences")
            ULshow = sheets.get("UserLicences", pd.DataFrame())
            if ULshow.empty: st.info("No user licences yet.")
            else:
                st.dataframe(ULshow.sort_values(["user_id","licence_id"]),
                             use_container_width=True, hide_index=True)
                del_ix = st.number_input("Delete row #", min_value=0, max_value=len(ULshow)-1, value=0, key="ul_del_ix")
                if st.button("Revoke selected", key="ul_revoke"):
                    sheets["UserLicences"] = ULshow.drop(ULshow.index[int(del_ix)]).reset_index(drop=True)
                    save_db(sheets); st.success("Revoked."); st.rerun()

        # Competency assessments
        with at[3]:
            st.markdown("### Competency Assessments")
            st.caption("Superusers/Admins: record outcomes and (optionally) issue licences.")
            AR = sheets.get("AssistanceRequests", pd.DataFrame(columns=["request_id","requester_user_id","licence_id","message","created","status","handled_by","handled_on","outcome","notes"])).copy()
            for c in ["status","handled_by","handled_on","outcome","notes"]:
                if c not in AR.columns: AR[c] = None
            Utab = sheets.get("Users", pd.DataFrame()); Ltab = sheets.get("Licences", pd.DataFrame())
            open_reqs = AR[AR["status"].fillna("open").isin(["open","in_review"])]
            if open_reqs.empty:
                st.info("No open mentoring/competency requests.")
            else:
                st.dataframe(open_reqs.merge(Utab[["user_id","name","email"]], left_on="requester_user_id", right_on="user_id", how="left"),
                             hide_index=True, use_container_width=True)
                sel_ids = open_reqs["request_id"].tolist()
                sel = st.selectbox("Select request to process", sel_ids, key="comp_sel")
                req = open_reqs[open_reqs["request_id"]==sel].iloc[0]
                st.write(f"**Member:** {Utab.loc[Utab['user_id']==req.requester_user_id,'name'].iloc[0]}  •  **Licence:** {Ltab.loc[Ltab['licence_id']==req.licence_id,'licence_name'].iloc[0]}")
                notes = st.text_area("Assessment notes", key="comp_notes")
                outcome = st.radio("Outcome", ["pass","more_training","fail"], horizontal=True, key="comp_outcome")
                grant = st.checkbox("Issue this licence on pass", value=True, key="comp_grant")
                valid_to = st.date_input("Valid to", key="comp_valid_to", format=DFMT)
                if st.button("Save outcome", type="primary", key="comp_save"):
                    AR.loc[AR["request_id"]==sel, ["status","handled_by","handled_on","outcome","notes"]] = ["closed", int(me["user_id"]), pd.Timestamp.today(), outcome, notes]
                    sheets["AssistanceRequests"] = AR
                    if outcome == "pass" and grant:
                        UL = sheets.get("UserLicences", pd.DataFrame(columns=["user_id","licence_id","valid_from","valid_to"]))
                        new = pd.DataFrame([[int(req.requester_user_id), int(req.licence_id), pd.Timestamp.today().normalize(), pd.Timestamp(valid_to)]], columns=UL.columns)
                        sheets["UserLicences"] = pd.concat([UL, new], ignore_index=True)
                    save_db(sheets); st.success("Saved."); st.rerun()

        # Machines (inline editor + max duration)
        with at[4]:
            st.markdown("### Machines")
            M = sheets.get("Machines", pd.DataFrame(columns=["machine_id","machine_name","licence_id","max_duration_minutes","serial_no","next_service_due","hours_used"])).copy()
            st.dataframe(M, use_container_width=True, hide_index=True)

            st.markdown("#### Inline machines editor (name, serial, next service)")
            for c in ["machine_id","machine_name","serial_no","next_service_due"]:
                if c not in M.columns: M[c] = None
            try:
                cfg = {
                    "machine_id": st.column_config.NumberColumn("ID", disabled=True),
                    "machine_name": st.column_config.TextColumn("Name"),
                    "serial_no": st.column_config.TextColumn("Serial #"),
                    "next_service_due": st.column_config.DateColumn("Next service"),
                }
            except Exception:
                cfg = None
            edited = st.data_editor(M[["machine_id","machine_name","serial_no","next_service_due"]],
                                    num_rows="fixed", hide_index=True,
                                    column_config=cfg if cfg else None,
                                    key="adm_m_table")
            if st.button("Save machine changes", key="adm_m_save"):
                M2 = M.set_index("machine_id"); E2 = edited.set_index("machine_id")
                for col in ["machine_name","serial_no","next_service_due"]:
                    if col in E2.columns:
                        M2[col] = E2[col]
                M2["next_service_due"] = pd.to_datetime(M2["next_service_due"], errors="coerce")
                sheets["Machines"] = M2.reset_index()
                save_db(sheets); st.success("Saved."); st.rerun()

            st.markdown("#### Edit max duration (per machine)")
            em1,em2,em3 = st.columns([2,1,1])
            with em1:
                em_sel = st.selectbox("Machine", [f"{r.machine_id} - {r.machine_name}" for r in sheets["Machines"].itertuples()], key="adm_m_sel")
            with em2:
                current = int(sheets["Machines"].loc[sheets["Machines"]["machine_id"]==int(em_sel.split(" - ")[0]),"max_duration_minutes"].iloc[0])
                em_val = st.number_input("Max minutes", 30, 480, current, step=30, key="adm_m_val")
            with em3:
                if st.button("Save max duration", key="adm_m_set"):
                    mid = int(em_sel.split(" - ")[0])
                    M3 = sheets["Machines"].copy()
                    M3.loc[M3["machine_id"]==mid, "max_duration_minutes"] = int(em_val)
                    sheets["Machines"] = M3
                    save_db(sheets); st.success("Saved.")

        # Subscriptions + discount reasons
        with at[5]:
            st.markdown("### Subscriptions")
            S = sheets.get("Subscriptions", pd.DataFrame(columns=["user_id","start_date","end_date","amount","paid","discount_reason","discount_pct"])).copy()
            reasons = sheets.get("DiscountReasons", pd.DataFrame(columns=["reason"]))
            c1,c2,c3,c4 = st.columns([2,1,1,1])
            with c1:
                s_user = st.selectbox("Member", [f"{r.user_id} - {r.name}" for r in sheets["Users"].itertuples()], key="sub_user")
            with c2:
                s_amount = st.number_input("Amount", min_value=0, max_value=1000, value=50, step=5, key="sub_amt")
            with c3:
                s_start = st.date_input("Start", key="sub_start", format=DFMT)
            with c4:
                s_end = st.date_input("End", key="sub_end", format=DFMT)
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

        # Schedule (day roster)
        with at[6]:
            st.markdown("### Day roster")
            d = st.date_input("Day", value=date.today(), key="adm_roster_day", format=DFMT)
            rows = []
            for m in sheets["Machines"].itertuples():
                for r in day_bookings(sheets, m.machine_id, d).itertuples():
                    rows.append([m.machine_name, r.start.time(), r.end.time(), r.purpose, r.status])
            st.dataframe(pd.DataFrame(rows, columns=["Machine","Start","End","Purpose","Status"]), use_container_width=True, hide_index=True)

        # Hours & Holidays
        with at[7]:
            st.markdown("### Weekly operating hours")
            OH = sheets.get("OperatingHours", pd.DataFrame(columns=["day_of_week","open_time","close_time"])).copy()
            if OH.empty or set(OH["day_of_week"].tolist()) != set(range(7)):
                rows = []
                for d in range(7):
                    rows.append([d, "09:00" if d in (1,2,3,4,5) else "", "16:00" if d in (1,2,3,4,5) else ""])
                OH = pd.DataFrame(rows, columns=["day_of_week","open_time","close_time"])
                sheets["OperatingHours"] = OH; save_db(sheets)

            st.dataframe(OH.sort_values("day_of_week"), use_container_width=True, hide_index=True)

            b1,b2,b3,b4 = st.columns([1,1,2,2])
            with b1:
                if st.button("Set weekdays open 09:00–16:00", key="hrs_weekdays_open"):
                    for d in range(0,5):
                        OH.loc[OH["day_of_week"]==d, ["open_time","close_time"]] = ["09:00","16:00"]
                    sheets["OperatingHours"] = OH; save_db(sheets); st.success("Saved.")
            with b2:
                if st.button("Close weekdays", key="hrs_weekdays_close"):
                    for d in range(0,5):
                        OH.loc[OH["day_of_week"]==d, ["open_time","close_time"]] = ["",""]
                    sheets["OperatingHours"] = OH; save_db(sheets); st.success("Saved.")
            with b3:
                if st.button("Copy Tue → Mon–Fri", key="hrs_copy_tue_weekdays"):
                    src = OH[OH["day_of_week"]==1].iloc[0]
                    for d in range(0,5):
                        OH.loc[OH["day_of_week"]==d, ["open_time","close_time"]] = [src["open_time"], src["close_time"]]
                    sheets["OperatingHours"] = OH; save_db(sheets); st.success("Saved.")
            with b4:
                if st.button("Copy Tue → Tue–Sat", key="hrs_copy_tue_tuesat"):
                    src = OH[OH["day_of_week"]==1].iloc[0]
                    for d in range(1,6):
                        OH.loc[OH["day_of_week"]==d, ["open_time","close_time"]] = [src["open_time"], src["close_time"]]
                    sheets["OperatingHours"] = OH; save_db(sheets); st.success("Saved.")

            st.divider()
            st.markdown("### Closed dates")
            CD = sheets.get("ClosedDates", pd.DataFrame(columns=["date","reason"])).copy()
            if CD.empty: st.info("No closed dates configured.")
            else: st.dataframe(CD.sort_values("date"), use_container_width=True, hide_index=True)
            nd = st.date_input("Closed date", key="closed_add_date", format=DFMT)
            nr = st.text_input("Reason", key="closed_add_reason")
            ca, cb = st.columns([1,1])
            with ca:
                if st.button("Add closed date", key="closed_add_btn"):
                    CD = pd.concat([CD, pd.DataFrame([[pd.Timestamp(nd), nr]], columns=["date","reason"])], ignore_index=True)
                    sheets["ClosedDates"] = CD; save_db(sheets); st.success("Added."); st.rerun()
            with cb:
                if not CD.empty:
                    del_ix = st.number_input("Delete row # (see index on left)", min_value=0, max_value=len(CD)-1, value=0, key="closed_del_ix")
                    if st.button("Delete selected", key="closed_del_btn"):
                        sheets["ClosedDates"] = CD.drop(CD.index[int(del_ix)]).reset_index(drop=True)
                        save_db(sheets); st.success("Removed."); st.rerun()

        # Newsletter prompt + DATA JSON preview (copy/paste to ChatGPT)
        with at[8]:
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
                U = sheets.get("Users", pd.DataFrame()).copy()
                if U.empty: return json.dumps({}, indent=2)
                U["first_name"] = U["name"].astype(str).str.split().str[0]
                U["last_name"]  = U["name"].astype(str).str.split().str[-1]
                U["suburb"]     = U["address"].astype(str).str.split(",").str[0]
                members=[]
                for r in U.itertuples():
                    members.append({
                        "first_name": r.first_name, "last_name": r.last_name, "email": r.email,
                        "birth_date": None if pd.isna(r.birth_date) else pd.to_datetime(r.birth_date).date().isoformat(),
                        "joined_date": None if pd.isna(r.joined_date) else pd.to_datetime(r.joined_date).date().isoformat(),
                        "suburb": r.suburb, "opted_in": bool(getattr(r,"newsletter_opt_in", True))
                    })
                events=[]
                EV = sheets.get("UserEvents", pd.DataFrame(columns=["event_id","user_id","event_name","event_date","notes"])).copy()
                for r in EV.itertuples():
                    email = U.loc[U["user_id"]==r.user_id, "email"]
                    events.append({
                        "member_email": email.iloc[0] if not email.empty else "",
                        "date": None if pd.isna(r.event_date) else pd.to_datetime(r.event_date).date().isoformat(),
                        "type": str(r.event_name), "title": str(r.event_name), "detail": str(getattr(r,"notes",""))
                    })
                updates=[]; CU=sheets.get("ClubUpdates", pd.DataFrame(columns=["title","text","link"]))
                for r in CU.itertuples():
                    updates.append({"title": str(r.title), "detail": str(getattr(r,"text","")), "link": str(getattr(r,"link",""))})
                notices=[]; NO=sheets.get("Notices", pd.DataFrame(columns=["title","text","link"]))
                for r in NO.itertuples():
                    notices.append({"title": str(r.title), "detail": str(getattr(r,"text","")), "link": str(getattr(r,"link",""))})
                spotlight=[]; SP=sheets.get("SpotlightSubmissions", pd.DataFrame())
                for r in SP.itertuples():
                    nm = U.loc[U["user_id"]==r.user_id, "name"]
                    spotlight.append({"member_name": nm.iloc[0] if not nm.empty else "", "title": str(r.title), "summary": str(getattr(r,"text","")), "image_url": str(getattr(r,"image_file","")), "link": ""})
                projects=[]; PJ=sheets.get("ProjectSubmissions", pd.DataFrame())
                for r in PJ.itertuples():
                    email = U.loc[U["user_id"]==r.user_id, "email"]
                    projects.append({"member_email": email.iloc[0] if not email.empty else "", "title": str(r.title), "blurb": str(getattr(r,"description","")), "image_urls": [str(getattr(r,"image_file",""))] if getattr(r,"image_file","") else [], "started_date": None, "category": ""})
                # Mentors = admins/superusers with any active licence
                mentors=[]; UL=sheets.get("UserLicences", pd.DataFrame()); L=sheets.get("Licences", pd.DataFrame())
                today=pd.Timestamp.today().normalize()
                if not UL.empty:
                    UL["valid_from"]=pd.to_datetime(UL["valid_from"], errors="coerce")
                    UL["valid_to"]=pd.to_datetime(UL["valid_to"], errors="coerce")
                    active = UL[(UL["valid_from"]<=today) & (UL["valid_to"]>=today)]
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
                last_issue_date = get_setting(sheets,"last_issue_date","")
                data = {
                    "members": members, "significant_events": events, "club_updates": updates,
                    "notices": notices, "spotlight_submissions": spotlight, "project_submissions": projects,
                    "mentors_offering": mentors, "mentorship_requests": [],
                    "meeting_info": meet, "links": links, "last_issue_date": last_issue_date
                }
                return json.dumps(data, indent=2)
            data_json = build_data_json(sheets)
            st.code(data_json, language="json")

            st.markdown("#### Full prompt (ready to copy to ChatGPT)")
            org_name = get_setting(sheets, "org_name", "Woodturners of the Hunter")
            logo_url = (get_setting(sheets,"app_public_url","").rstrip("/") + "/assets/" + get_setting(sheets,"active_logo","logo1.png")) if get_setting(sheets,"app_public_url","") else "{{logo_url}}"
            compiled = (new_prompt or prompt_text).replace("🔧ORG_NAME", org_name).replace("{DATA_JSON}", data_json).replace("{{logo_url}}", logo_url)
            st.code(compiled, language="markdown")
            st.info("Copy the prompt into ChatGPT to generate three newsletter drafts.")

        # Settings
        with at[9]:
            st.markdown("### Settings")
            org = st.text_input("Organisation name", value=get_setting(sheets,"org_name","Woodturners of the Hunter"), key="set_org")
            app_url = st.text_input("App public URL (for unsubscribe/logo links)", value=get_setting(sheets,"app_public_url",""), key="set_url")
            lock = st.checkbox("Lock booking to signed-in member only", value=(get_setting(sheets,"lock_booking_to_member","false").lower() in ("1","true","yes")), key="set_lock")
            logo = st.selectbox("Active logo file", ["logo1.png","logo2.png","logo3.png"], index=["logo1.png","logo2.png","logo3.png"].index(get_setting(sheets,"active_logo","logo1.png")), key="set_logo")
            if st.button("Save settings", key="set_save"):
                upsert_setting(sheets, "org_name", org)
                upsert_setting(sheets, "app_public_url", app_url)
                upsert_setting(sheets, "lock_booking_to_member", "true" if lock else "false")
                upsert_setting(sheets, "active_logo", logo)
                save_db(sheets); st.success("Settings saved."); st.rerun()

        # Notifications
        with at[10]:
            st.markdown("### Notifications")
            msgs = []
            today = pd.Timestamp.today().normalize()
            try:
                issue_day = int(get_setting(sheets,"newsletter_issue_day","1") or 1)
            except:
                issue_day = 1
            this_issue = pd.Timestamp(year=today.year, month=today.month, day=min(issue_day,28))
            next_issue = this_issue if this_issue>=today else (this_issue + pd.DateOffset(months=1))
            days = (next_issue - today).days
            if 0 <= days <= 7:
                msgs.append(f"Newsletter due on {next_issue.date()}")
            S = sheets.get("Subscriptions", pd.DataFrame())
            if not S.empty:
                soon = S[pd.to_datetime(S["end_date"], errors="coerce") <= (today + pd.Timedelta(days=14))]
                if not soon.empty:
                    msgs.append(f"{len(soon)} subscription(s) expiring within 14 days.")
            if msgs:
                st.write("\n".join([f"• {m}" for m in msgs]))
            else:
                st.info("No notifications.")
