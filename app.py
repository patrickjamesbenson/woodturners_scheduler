import streamlit as st
import pandas as pd
from datetime import datetime, date, time, timedelta
from pathlib import Path
import re

# ---------- App config ----------
st.set_page_config(page_title="Woodturners Scheduler", page_icon="🪵", layout="wide")
BASE = Path(__file__).resolve().parent
DB = BASE / "data" / "db.xlsx"
ASSETS = BASE / "assets"
DATE_FMT = "DD/MM/YYYY"  # Streamlit display format

# ---------- Data IO ----------
@st.cache_data
def load_db() -> dict:
    """Read all sheets from the Excel DB into dataframes."""
    try:
        xls = pd.ExcelFile(DB, engine="openpyxl")
    except Exception as e:
        st.error(f"Could not open database at {DB}. Error: {e}")
        raise
    sheets = {}
    for name in xls.sheet_names:
        sheets[name] = pd.read_excel(DB, sheet_name=name, engine="openpyxl")
    return sheets

def save_db(sheets: dict):
    """Write all known sheets back to Excel and clear cache so next run reloads."""
    with pd.ExcelWriter(DB, engine="openpyxl", mode="w") as w:
        for name, df in sheets.items():
            if isinstance(df, pd.DataFrame):
                df.to_excel(w, sheet_name=name, index=False)
    try:
        load_db.clear()  # invalidate cache so st.rerun() sees fresh data
    except Exception:
        pass

def ensure_sheet(sheets: dict, name: str, columns: list) -> pd.DataFrame:
    if name not in sheets or not isinstance(sheets[name], pd.DataFrame):
        sheets[name] = pd.DataFrame(columns=columns)
    # add any missing columns (won't drop extras)
    for c in columns:
        if c not in sheets[name].columns:
            sheets[name][c] = None
    return sheets[name]

# ---------- Helpers ----------
def get_setting(sheets: dict, key: str, default: str = "") -> str:
    S = sheets.get("Settings", pd.DataFrame(columns=["key", "value"]))
    row = S[S["key"] == key]
    return default if row.empty else str(row.iloc[0]["value"])

def parse_hhmm_or_ampm(s: str):
    """Accept '9:00', '09:00', '9am', '5:30pm', '17', '1700' etc → (h, m) or None."""
    if s is None:
        return None
    try:
        s = str(s).strip()
        if not s:
            return None
        s2 = s.lower().replace(" ", "")
        ampm = None
        if s2.endswith("am") or s2.endswith("pm"):
            ampm = s2[-2:]
            s2 = s2[:-2]
        # accept 1700 style
        if re.fullmatch(r"\d{3,4}", s2):
            if len(s2) == 3:
                h = int(s2[0])
                m = int(s2[1:])
            else:
                h = int(s2[:2])
                m = int(s2[2:])
            if ampm == "pm" and h != 12:
                h += 12
            if ampm == "am" and h == 12:
                h = 0
            return h, m
        parts = re.split(r"[:h]", s2)
        h = int(parts[0])
        m = int(parts[1]) if len(parts) > 1 else 0
        if ampm == "pm" and h != 12:
            h += 12
        if ampm == "am" and h == 12:
            h = 0
        return h, m
    except Exception:
        return None

def is_open(sheets: dict, d: date, start_t: time, end_t: time):
    """Check closed dates + operating hours for a given day/time window."""
    CD = sheets.get("ClosedDates", pd.DataFrame(columns=["date", "reason"])).copy()
    if not CD.empty:
        CD["date"] = pd.to_datetime(CD["date"], errors="coerce").dt.normalize()
    dn = pd.Timestamp(d).normalize()
    if not CD.empty and (CD["date"] == dn).any():
        return False, "Closed (holiday/maintenance)"
    OH = sheets.get("OperatingHours", pd.DataFrame(columns=["day_of_week", "open_time", "close_time"]))
    row = OH[OH["day_of_week"] == pd.Timestamp(d).dayofweek]
    if row.empty:
        return False, "Closed"
    ot = parse_hhmm_or_ampm(row.iloc[0]["open_time"])
    ct = parse_hhmm_or_ampm(row.iloc[0]["close_time"])
    if not ot or not ct:
        return False, "Closed"
    o_h, o_m = ot
    c_h, c_m = ct
    st_min = start_t.hour * 60 + start_t.minute
    en_min = end_t.hour * 60 + end_t.minute
    ok = (o_h * 60 + o_m) <= st_min and en_min <= (c_h * 60 + c_m)
    return ok, f"{o_h:02d}:{o_m:02d}–{c_h:02d}:{c_m:02d}"

def user_licence_ids(sheets: dict, uid: int) -> set:
    UL = sheets.get("UserLicences", pd.DataFrame(columns=["user_id", "licence_id", "valid_from", "valid_to"])).copy()
    if UL.empty:
        return set()
    UL["valid_from"] = pd.to_datetime(UL["valid_from"], errors="coerce")
    UL["valid_to"] = pd.to_datetime(UL["valid_to"], errors="coerce")
    today = pd.Timestamp.today().normalize()
    valid = UL[(UL["user_id"].astype("Int64") == uid) & (UL["valid_from"] <= today) & (UL["valid_to"] >= today)]
    return set(pd.to_numeric(valid["licence_id"], errors="coerce").dropna().astype(int))

def machine_lists_for_user(sheets: dict, uid: int):
    lids = user_licence_ids(sheets, uid)
    M = sheets.get("Machines", pd.DataFrame(columns=["machine_id", "machine_name", "licence_id", "max_duration_minutes"])).copy()
    for c in ["machine_id", "licence_id", "max_duration_minutes"]:
        if c in M.columns:
            M[c] = pd.to_numeric(M[c], errors="coerce")
    allowed = M[M["licence_id"].isin(lids)]
    blocked = M[~M["licence_id"].isin(lids)]
    return allowed, blocked

def day_bookings(sheets: dict, machine_id: int, d: date) -> pd.DataFrame:
    B = sheets.get("Bookings", pd.DataFrame(columns=["booking_id", "user_id", "machine_id", "start", "end", "purpose", "notes", "status"])).copy()
    if B.empty:
        return B
    B["start"] = pd.to_datetime(B["start"], errors="coerce")
    B["end"] = pd.to_datetime(B["end"], errors="coerce")
    ds = pd.Timestamp.combine(d, time(0, 0))
    de = ds + timedelta(days=1)
    view = B[(pd.to_numeric(B["machine_id"], errors="coerce") == machine_id) & (B["start"] < de) & (B["end"] > ds)].copy()
    return view.sort_values("start")

def make_human(df: pd.DataFrame, sheets: dict) -> pd.DataFrame:
    """Merge in human-friendly labels for user/machine/licence where possible."""
    if df is None or df.empty:
        return df
    U = sheets.get("Users", pd.DataFrame())
    M = sheets.get("Machines", pd.DataFrame())
    L = sheets.get("Licences", pd.DataFrame())
    if "user_id" in df.columns and "user_id" in U.columns and "name" in U.columns:
        df = df.merge(U[["user_id", "name"]], on="user_id", how="left")
    if "machine_id" in df.columns and {"machine_id", "machine_name"}.issubset(M.columns):
        df = df.merge(M[["machine_id", "machine_name"]], on="machine_id", how="left")
    if "licence_id" in df.columns and {"licence_id", "licence_name"}.issubset(L.columns):
        df = df.merge(L[["licence_id", "licence_name"]], on="licence_id", how="left")
    return df

# ---------- Load data ----------
sheets = load_db()

# ---------- Header (logo only, centred) ----------
logo_file = get_setting(sheets, "active_logo", "logo1.png")
c1, c2, c3 = st.columns([1, 2, 1])
with c2:
    logo_path = ASSETS / logo_file
    if logo_path.exists():
        st.image(str(logo_path), use_column_width=True)
    else:
        st.write("")  # keep spacing

# ---------- Auth (very simple: pick your name; admins require password) ----------
U = ensure_sheet(sheets, "Users", ["user_id", "name", "role", "email", "phone", "birth_date", "joined_date", "newsletter_opt_in", "password"])
user_labels = [f"{r.name} ({r.role})" for r in U.itertuples()]
id_by_label = {f"{r.name} ({r.role})": int(r.user_id) for r in U.itertuples()}

st.sidebar.header("Sign in")
sel_label = st.sidebar.selectbox("Your name", [""] + user_labels, key="auth_name")
me = None
if sel_label:
    uid = id_by_label[sel_label]
    row = U[U["user_id"] == uid].iloc[0]
    # Admin/superuser needs password
    if str(row.get("role", "")).lower() in ("admin", "superuser") and str(row.get("password", "")).strip():
        pwd = st.sidebar.text_input("Password", type="password", key="auth_pwd")
        if st.sidebar.button("Sign in", key="auth_btn"):
            if pwd == str(row["password"]):
                st.session_state["me_id"] = uid
                st.sidebar.success("Signed in")
            else:
                st.sidebar.error("Wrong password")
    else:
        if st.sidebar.button("Continue", key="auth_continue"):
            st.session_state["me_id"] = uid

if "me_id" in st.session_state:
    me_row = U[U["user_id"] == st.session_state["me_id"]]
    if not me_row.empty:
        me = me_row.iloc[0].to_dict()
        st.sidebar.info(f"Signed in as: {me['name']} ({me['role']})")
else:
    st.sidebar.warning("Select your name and sign in to continue.")

# ---------- Tabs ----------
tabs = st.tabs(["Book a Machine", "Calendar", "Mentoring", "Issues & Maintenance", "Admin"])

# --- Book a Machine ---
with tabs[0]:
    if not me:
        st.info("Sign in to book a machine.")
    else:
        st.subheader("Book a Machine")
        allowed, blocked = machine_lists_for_user(sheets, int(me["user_id"]))
        if allowed.empty:
            st.warning("No machines available for your current licences.")
        sel_machine = st.selectbox(
            "Machine (only those you’re licensed for appear)",
            [f"{r.machine_id} - {r.machine_name}" for r in allowed.itertuples()],
            key="book_m_sel",
        )
        if sel_machine:
            mid = int(sel_machine.split(" - ")[0])
            # Inputs
            book_day = st.date_input("Day", value=date.today(), format=DATE_FMT, key="book_day")
            book_start = st.time_input("Start time", value=time(9, 0), key="book_start")
            # Max duration from Machines; default 240 mins
            M = sheets.get("Machines", pd.DataFrame())
            max_mins = 240
            try:
                max_mins = int(M.loc[M["machine_id"] == mid, "max_duration_minutes"].iloc[0])
            except Exception:
                pass
            dur = st.slider("Duration (minutes)", 30, max_mins, min(60, max_mins), step=30, key="book_dur")
            start_dt = datetime.combine(book_day, book_start)
            end_dt = start_dt + timedelta(minutes=dur)
            st.caption(f"{start_dt.strftime('%d/%m/%Y %H:%M')} → {end_dt.strftime('%H:%M')}  ({dur} min)")

            # Availabilities today for this machine
            todays = day_bookings(sheets, mid, book_day)
            show = make_human(todays.copy(), sheets)
            if not show.empty:
                cols = [c for c in ["start", "end", "name", "purpose", "status"] if c in show.columns]
                st.write("Already booked:")
                st.dataframe(show[cols] if cols else show, use_container_width=True, hide_index=True)
            else:
                st.write("No bookings yet today.")

            # Hours + overlap
            ok_hours, hours_msg = is_open(sheets, book_day, start_dt.time(), end_dt.time())
            overlap = False
            for r in todays.itertuples():
                if not (end_dt <= r.start or start_dt >= r.end):
                    overlap = True
                    break
            if not ok_hours:
                st.error(f"Outside operating hours ({hours_msg}).")
            if overlap:
                st.error("Overlaps an existing booking.")

            clicked = st.button("Confirm booking", type="primary", key="book_confirm", disabled=not (ok_hours and (not overlap)))
            if clicked:
                B = ensure_sheet(
                    sheets,
                    "Bookings",
                    ["booking_id", "user_id", "machine_id", "start", "end", "purpose", "notes", "status"],
                )
                next_id = 1 if B.empty else int(pd.to_numeric(B["booking_id"], errors="coerce").fillna(0).max()) + 1
                new = pd.DataFrame(
                    [[next_id, int(me["user_id"]), mid, start_dt, end_dt, "use", "", "confirmed"]],
                    columns=B.columns,
                )
                sheets["Bookings"] = pd.concat([B, new], ignore_index=True)
                save_db(sheets)
                st.success("Booked.")
                st.rerun()

# --- Calendar ---
with tabs[1]:
    st.subheader("Calendar")
    M = sheets.get("Machines", pd.DataFrame())
    sel_cal_m = st.selectbox(
        "Machine",
        [f"{r.machine_id} - {r.machine_name}" for r in M.itertuples()],
        key="cal_m_sel",
    )
    if sel_cal_m:
        mid = int(sel_cal_m.split(" - ")[0])
        base_day = st.date_input("Day", value=date.today(), format=DATE_FMT, key="cal_day")
        view = st.radio("View", ["Day", "Week"], horizontal=True, key="cal_view")
        if view == "Day":
            DB = day_bookings(sheets, mid, base_day).copy()
            DB = make_human(DB, sheets)
            cols = [c for c in ["start", "end", "name", "purpose", "status"] if c in DB.columns]
            st.dataframe(DB[cols] if cols else DB, use_container_width=True, hide_index=True)
        else:
            start_w = base_day - timedelta(days=base_day.weekday())
            rows = []
            for d in range(7):
                dd = start_w + timedelta(days=d)
                for r in day_bookings(sheets, mid, dd).itertuples():
                    rows.append([dd.strftime("%d/%m/%Y"), r.start.strftime("%H:%M"), r.end.strftime("%H:%M"), getattr(r, "purpose", ""), r.user_id])
            W = pd.DataFrame(rows, columns=["day", "start", "end", "purpose", "user_id"])
            W = make_human(W, sheets)
            if "name" in W.columns:
                W = W[["day", "start", "end", "name", "purpose"]]
            st.dataframe(W, use_container_width=True, hide_index=True)

# --- Mentoring ---
with tabs[2]:
    st.subheader("Mentoring & Competency Requests")
    if not me:
        st.info("Sign in to request mentoring.")
    else:
        L = ensure_sheet(sheets, "Licences", ["licence_id", "licence_name", "notes"])
        lic_map = {f"{r.licence_id} - {r.licence_name}": int(r.licence_id) for r in L.itertuples()}
        sel_lic = st.selectbox("Skill / Machine licence", list(lic_map.keys()), key="ment_lic")
        msg = st.text_area("What do you need help with?", key="ment_msg")
        # Suggested mentors (members who already hold this licence; admins/superusers shown)
        UL = ensure_sheet(sheets, "UserLicences", ["user_id", "licence_id", "valid_from", "valid_to"]).copy()
        U = sheets["Users"]
        today = pd.Timestamp.today().normalize()
        UL["valid_from"] = pd.to_datetime(UL["valid_from"], errors="coerce")
        UL["valid_to"] = pd.to_datetime(UL["valid_to"], errors="coerce")
        lic_id = lic_map[sel_lic]
        holder_ids = UL[(pd.to_numeric(UL["licence_id"], errors="coerce") == lic_id) & (UL["valid_from"] <= today) & (UL["valid_to"] >= today)]["user_id"].astype(int).tolist()
        mentors = U[(U["user_id"].isin(holder_ids)) & (U["role"].isin(["admin", "superuser"]))][["name", "email", "phone"]]
        if mentors.empty:
            st.warning("No listed mentors for this licence yet.")
        else:
            st.markdown("**Suggested mentors:**")
            st.dataframe(mentors, hide_index=True, use_container_width=True)

        if st.button("Submit mentoring request", type="primary", key="ment_submit"):
            AR = ensure_sheet(
                sheets,
                "AssistanceRequests",
                ["request_id", "requester_user_id", "licence_id", "message", "created", "status", "handled_by", "handled_on", "outcome", "notes"],
            )
            req_id = 1 if AR.empty else int(pd.to_numeric(AR["request_id"], errors="coerce").fillna(0).max()) + 1
            new = pd.DataFrame([[req_id, int(me["user_id"]), int(lic_id), msg, pd.Timestamp.today(), "open", None, None, None, None]], columns=AR.columns)
            sheets["AssistanceRequests"] = pd.concat([AR, new], ignore_index=True)
            save_db(sheets)
            st.success("Request submitted.")
            st.rerun()

        # My requests
        AR2 = sheets.get("AssistanceRequests", pd.DataFrame())
        mine = AR2[AR2["requester_user_id"].astype("Int64") == int(me["user_id"])]
        if mine.empty:
            st.info("No requests yet.")
        else:
            st.dataframe(mine.sort_values("created", ascending=False), hide_index=True, use_container_width=True)

# --- Issues & Maintenance ---
with tabs[3]:
    st.subheader("Issues & Maintenance")
    M = sheets.get("Machines", pd.DataFrame())
    sel_iss_m = st.selectbox("Machine", [f"{r.machine_id} - {r.machine_name}" for r in M.itertuples()], key="iss_m_sel")
    issue_txt = st.text_area("Describe an issue", key="iss_txt")
    if me and st.button("Submit issue", key="iss_btn"):
        I = ensure_sheet(sheets, "Issues", ["issue_id", "machine_id", "user_id", "created", "status", "notes"])
        iid = 1 if I.empty else int(pd.to_numeric(I["issue_id"], errors="coerce").fillna(0).max()) + 1
        mid = int(sel_iss_m.split(" - ")[0]) if sel_iss_m else None
        new = pd.DataFrame([[iid, mid, int(me["user_id"]), pd.Timestamp.today(), "open", issue_txt]], columns=I.columns)
        sheets["Issues"] = pd.concat([I, new], ignore_index=True)
        save_db(sheets)
        st.success("Issue logged.")
        st.rerun()

    Iv = sheets.get("Issues", pd.DataFrame()).copy()
    Iv = make_human(Iv, sheets)
    cols = [c for c in ["issue_id", "created", "machine_name", "name", "status", "notes"] if c in Iv.columns]
    st.dataframe(Iv[cols] if cols else Iv, use_container_width=True, hide_index=True)

# --- Admin ---
with tabs[4]:
    if not me or str(me.get("role", "")).lower() not in ("admin", "superuser"):
        st.info("Admins only.")
    else:
        at = st.tabs(["Users", "Licences", "User Licences", "Competency", "Machines", "Subscriptions", "Hours & Holidays", "Newsletter", "Settings"])

        with at[0]:
            st.markdown("### Users")
            cols = [c for c in ["user_id", "name", "role", "email", "phone", "birth_date", "joined_date", "newsletter_opt_in"] if c in U.columns]
            st.dataframe(U[cols] if cols else U, use_container_width=True, hide_index=True)

        with at[1]:
            st.markdown("### Licences")
            L = ensure_sheet(sheets, "Licences", ["licence_id", "licence_name", "notes"])
            st.dataframe(L, use_container_width=True, hide_index=True)

        with at[2]:
            st.markdown("### User Licences")
            U = sheets["Users"]; L = sheets["Licences"]
            UL = ensure_sheet(sheets, "UserLicences", ["user_id", "licence_id", "valid_from", "valid_to"])
            ulabel = st.selectbox("Member", [f"{r.user_id} - {r.name}" for r in U.itertuples()], key="ul_user")
            llabel = st.selectbox("Licence", [f"{r.licence_id} - {r.licence_name}" for r in L.itertuples()], key="ul_lic")
            vf = st.date_input("Valid from", value=date.today(), format=DATE_FMT, key="ul_from")
            vt = st.date_input("Valid to", value=date.today() + timedelta(days=365), format=DATE_FMT, key="ul_to")
            if st.button("Grant licence", key="ul_grant"):
                uid = int(ulabel.split(" - ")[0])
                lid = int(llabel.split(" - ")[0])
                new = pd.DataFrame([[uid, lid, pd.Timestamp(vf), pd.Timestamp(vt)]], columns=UL.columns)
                sheets["UserLicences"] = pd.concat([UL, new], ignore_index=True)
                save_db(sheets)
                st.success("Licence granted.")
                st.rerun()
            ULv = make_human(sheets.get("UserLicences", pd.DataFrame()).copy(), sheets)
            # Reorder: show names first
            order = [c for c in ["name", "licence_name", "valid_from", "valid_to", "user_id", "licence_id"] if c in ULv.columns]
            st.dataframe(ULv[order] if order else ULv, use_container_width=True, hide_index=True)

        with at[3]:
            st.markdown("### Competency Assessments")
            AR = ensure_sheet(
                sheets,
                "AssistanceRequests",
                ["request_id", "requester_user_id", "licence_id", "message", "created", "status", "handled_by", "handled_on", "outcome", "notes"],
            ).copy()
            U = sheets["Users"]; L = sheets["Licences"]
            open_reqs = AR[AR["status"].fillna("open").isin(["open", "in_review"])]
            if open_reqs.empty:
                st.info("No open requests.")
            else:
                disp = open_reqs.copy()
                disp = disp.merge(U[["user_id", "name", "email"]], left_on="requester_user_id", right_on="user_id", how="left")
                disp = disp.merge(L[["licence_id", "licence_name"]], on="licence_id", how="left")
                cols = [c for c in ["request_id", "name", "email", "licence_name", "message", "status", "created"] if c in disp.columns]
                st.dataframe(disp[cols] if cols else disp, use_container_width=True, hide_index=True)

                sel_req = st.selectbox("Select request id", open_reqs["request_id"].tolist(), key="comp_sel")
                req = open_reqs[open_reqs["request_id"] == sel_req].iloc[0]
                st.write(f"**Member:** {U.loc[U['user_id']==req.requester_user_id,'name'].iloc[0]}  •  **Licence:** {L.loc[L['licence_id']==req.licence_id,'licence_name'].iloc[0]}")
                notes = st.text_area("Assessment notes", key="comp_notes")
                outcome = st.radio("Outcome", ["pass", "more_training", "fail"], horizontal=True, key="comp_outcome")
                grant = st.checkbox("Issue licence on pass", value=True, key="comp_grant")
                valid_to = st.date_input("Valid to", value=date.today() + timedelta(days=365), format=DATE_FMT, key="comp_validto")
                if st.button("Save outcome", key="comp_save"):
                    AR.loc[AR["request_id"] == sel_req, ["status", "handled_by", "handled_on", "outcome", "notes"]] = [
                        "closed",
                        int(me["user_id"]),
                        pd.Timestamp.today(),
                        outcome,
                        notes,
                    ]
                    sheets["AssistanceRequests"] = AR
                    if outcome == "pass" and grant:
                        UL = ensure_sheet(sheets, "UserLicences", ["user_id", "licence_id", "valid_from", "valid_to"])
                        new = pd.DataFrame([[int(req.requester_user_id), int(req.licence_id), pd.Timestamp.today().normalize(), pd.Timestamp(valid_to)]], columns=UL.columns)
                        sheets["UserLicences"] = pd.concat([UL, new], ignore_index=True)
                    save_db(sheets)
                    st.success("Saved.")
                    st.rerun()

        with at[4]:
            st.markdown("### Machines (inline editor)")
            M = ensure_sheet(sheets, "Machines", ["machine_id", "machine_name", "licence_id", "serial", "next_service", "max_duration_minutes"])
            edited = st.data_editor(M, num_rows="dynamic", use_container_width=True, key="mach_edit")
            if st.button("Save machines", key="mach_save"):
                sheets["Machines"] = edited
                save_db(sheets)
                st.success("Machines saved.")
                st.rerun()

        with at[5]:
            st.markdown("### Subscriptions")
            Sv = ensure_sheet(sheets, "Subscriptions", ["user_id", "type", "start_date", "end_date", "amount", "paid", "discount_percent", "discount_reason"])
            Svv = Sv.copy()
            Svv = make_human(Svv, sheets)
            order = [c for c in ["name", "type", "start_date", "end_date", "amount", "paid", "discount_percent", "discount_reason", "user_id"] if c in Svv.columns]
            st.dataframe(Svv[order] if order else Svv, use_container_width=True, hide_index=True)

        with at[6]:
            st.markdown("### Weekly operating hours & holidays")
            OH = ensure_sheet(sheets, "OperatingHours", ["day_of_week", "open_time", "close_time"])
            CD = ensure_sheet(sheets, "ClosedDates", ["date", "reason"])
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**Operating hours (0=Mon … 6=Sun)**")
                oh_edited = st.data_editor(OH, use_container_width=True, key="oh_edit")
                if st.button("Save hours", key="oh_save"):
                    sheets["OperatingHours"] = oh_edited
                    save_db(sheets)
                    st.success("Operating hours saved.")
                    st.rerun()
                if st.button("Set all weekdays open 09:00–17:00", key="oh_weekdays"):
                    sheets["OperatingHours"] = pd.DataFrame(
                        [{"day_of_week": d, "open_time": "09:00", "close_time": "17:00"} for d in range(5)]
                    )
                    save_db(sheets)
                    st.success("Set Mon–Fri 09:00–17:00.")
                    st.rerun()
            with col2:
                st.markdown("**Closed dates**")
                cd_edited = st.data_editor(CD, use_container_width=True, key="cd_edit")
                if st.button("Save closed dates", key="cd_save"):
                    sheets["ClosedDates"] = cd_edited
                    save_db(sheets)
                    st.success("Closed dates saved.")
                    st.rerun()

        with at[7]:
            st.markdown("### Newsletter")
            T = ensure_sheet(sheets, "Templates", ["key", "text"])
            row = T[T["key"] == "newsletter_prompt"]
            default_txt = row.iloc[0]["text"] if not row.empty else ""
            txt = st.text_area("Newsletter Prompt", default_txt, height=260, key="nl_prompt")
            if st.button("Save prompt", key="nl_save"):
                if row.empty:
                    T = pd.concat([T, pd.DataFrame([["newsletter_prompt", txt]], columns=["key", "text"])], ignore_index=True)
                else:
                    T.loc[T["key"] == "newsletter_prompt", "text"] = txt
                sheets["Templates"] = T
                save_db(sheets)
                st.success("Prompt saved.")
                st.rerun()

        with at[8]:
            st.markdown("### Settings")
            S = ensure_sheet(sheets, "Settings", ["key", "value"])
            S_edit = st.data_editor(S, use_container_width=True, key="settings_edit")
            if st.button("Save settings", key="settings_save"):
                sheets["Settings"] = S_edit
                save_db(sheets)
                st.success("Settings saved.")
                st.rerun()