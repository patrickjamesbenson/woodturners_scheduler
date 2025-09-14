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
                        new = pd.DataFrame([[int(req.requester_user_id), int(req.licence]()]()
