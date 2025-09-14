
import streamlit as st, pandas as pd
from datetime import datetime, date, time, timedelta
from pathlib import Path
import re, json

st.set_page_config(page_title="Woodturners Scheduler", page_icon="🪵", layout="wide")
BASE = Path(__file__).resolve().parent
DB = BASE/"data"/"db.xlsx"
ASSETS = BASE/"assets"
DFMT = "DD/MM/YYYY"

@st.cache_data
def load_db():
    xls = pd.ExcelFile(DB, engine="openpyxl")
    return {n: pd.read_excel(DB, engine="openpyxl", sheet_name=n) for n in xls.sheet_names}

def save_db(sheets: dict):
    import pandas as _pd
    with _pd.ExcelWriter(DB_PATH, engine="openpyxl", mode="w") as w:
        for name, df in sheets.items():
            if isinstance(df, _pd.DataFrame):
                df.to_excel(w, sheet_name=name, index=False)
    try:
        load_db.clear()  # clears @st.cache_data so next run reloads the file
    except Exception:
        pass

def get_setting(s, key, default=""):
    S=s.get("Settings", pd.DataFrame(columns=["key","value"]))
    m=S[S["key"]==key]
    return default if m.empty else str(m.iloc[0]["value"])

def user_licence_ids(s, uid):
    UL=s.get("UserLicences", pd.DataFrame(columns=["user_id","licence_id","valid_from","valid_to"]))
    UL["valid_from"]=pd.to_datetime(UL["valid_from"], errors="coerce")
    UL["valid_to"]=pd.to_datetime(UL["valid_to"], errors="coerce")
    today=pd.Timestamp.today().normalize()
    return set(UL[(UL["user_id"]==uid)&(UL["valid_from"]<=today)&(UL["valid_to"]>=today)]["licence_id"].astype(int))

def machine_options_for(s, uid):
    lids=user_licence_ids(s, uid)
    M=s["Machines"]
    return M[M["licence_id"].isin(lids)].copy(), M[~M["licence_id"].isin(lids)].copy()

def day_bookings(s, mid, d):
    B=s["Bookings"].copy()
    B["start"]=pd.to_datetime(B["start"], errors="coerce")
    B["end"]=pd.to_datetime(B["end"], errors="coerce")
    day_start=pd.Timestamp.combine(d, time(0,0)); day_end=day_start+timedelta(days=1)
    m=B[(B["machine_id"]==mid)&(B["start"]<day_end)&(B["end"]>day_start)]
    return m.sort_values("start")

def is_open(s, d, st_t, en_t):
    CD=s.get("ClosedDates", pd.DataFrame(columns=["date","reason"])).copy()
    if not CD.empty: CD["date"]=pd.to_datetime(CD["date"], errors="coerce").dt.normalize()
    dn=pd.Timestamp(d).normalize()
    if not CD.empty and (CD["date"]==dn).any(): return False,"Closed date"
    OH=s.get("OperatingHours", pd.DataFrame(columns=["day_of_week","open_time","close_time"]))
    row=OH[OH["day_of_week"]==pd.Timestamp(d).dayofweek]
    if row.empty: return False,"Closed"
    def parse(v):
        try:
            s=str(v).strip()
            if not s: return None
            sL=s.lower().replace(" ", "")
            ampm=None
            if sL.endswith("am") or sL.endswith("pm"):
                ampm=sL[-2:]; sL=sL[:-2]
            parts=re.split(r'[:h]', sL); h=int(parts[0]); m=int(parts[1]) if len(parts)>1 else 0
            if ampm=="pm" and h!=12: h+=12
            if ampm=="am" and h==12: h=0
            return h,m
        except: return None
    ot=parse(row.iloc[0]["open_time"]); ct=parse(row.iloc[0]["close_time"])
    if not ot or not ct: return False,"Closed"
    o_h,o_m=ot; c_h,c_m=ct
    st_min=st_t.hour*60+st_t.minute; en_min=en_t.hour*60+en_t.minute
    return (o_h*60+o_m)<=st_min and en_min<=(c_h*60+c_m), f"{o_h:02d}:{o_m:02d}–{c_h:02d}:{c_m:02d}"

sheets=load_db()

# Header logo centred
c1,c2,c3=st.columns([1,2,1])
with c2:
    st.image(str(ASSETS / get_setting(sheets,"active_logo","logo1.png")), use_column_width=True)

U=sheets["Users"]
labels=[f"{r.name} ({r.role})" for r in U.itertuples()]
id_by_label={f"{r.name} ({r.role})": int(r.user_id) for r in U.itertuples()}

st.sidebar.header("Sign in")
label=st.sidebar.selectbox("Your name", [""]+labels, index=0)
me=None
if label:
    uid=id_by_label[label]
    row=U[U["user_id"]==uid].iloc[0]
    if row["role"] in ("admin","superuser") and str(row.get("password","")).strip():
        pwd=st.sidebar.text_input("Password", type="password")
        if st.sidebar.button("Sign in"):
            if pwd==str(row["password"]): st.session_state["me_id"]=uid; st.sidebar.success("Signed in")
            else: st.sidebar.error("Wrong password")
    else:
        if st.sidebar.button("Continue"): st.session_state["me_id"]=uid
if "me_id" in st.session_state:
    me=U[U["user_id"]==st.session_state["me_id"]].iloc[0].to_dict()
    st.sidebar.info(f"Signed in as: {me['name']} ({me['role']})")
else:
    st.sidebar.warning("Select your name and sign in to continue.")

tabs=st.tabs(["Book a Machine","Calendar","Mentoring","Issues & Maintenance","Admin"])

with tabs[0]:
    st.subheader("Book a Machine")

    if not me:
        st.info("Sign in to book.")
    else:
        # Which machines this user is allowed to book
        allowed, blocked = machine_options_for(sheets, int(me["user_id"]))

        choice = st.selectbox(
            "Machine",
            [f"{r.machine_id} - {r.machine_name}" for r in allowed.itertuples()],
            key="book_machine_sel",
        )
        mid = int(choice.split(" - ")[0])

        # Inputs
        day = st.date_input("Day", value=date.today(), format=DFMT, key="book_day")
        start = st.time_input("Start time", value=time(9, 0), key="book_start")

        # Respect per-machine max duration from Admin
        max_mins = int(
            sheets["Machines"]
            .loc[sheets["Machines"]["machine_id"] == mid, "max_duration_minutes"]
            .iloc[0]
        )
        dur = st.slider(
            "Duration (minutes)",
            30,
            max_mins,
            min(60, max_mins),
            step=30,
            key="book_duration",
        )

        # Timing
        start_dt = datetime.combine(day, start)
        end_dt = start_dt + timedelta(minutes=dur)
        st.caption(
            f"{start_dt.strftime('%d/%m/%Y %H:%M')} → {end_dt.strftime('%H:%M')}  ({dur} min)"
        )

        # Availabilities today for this machine
        today_rows = day_bookings(sheets, mid, day).copy()
        show_cols = [c for c in ["start", "end", "purpose", "status", "user_id"] if c in today_rows.columns]
        if not today_rows.empty and show_cols:
            st.write("Availability (today):")
            st.dataframe(
                today_rows.sort_values("start")[show_cols],
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.write("No bookings yet today.")

        # Hours open + overlap check
        ok_hours, hours_msg = is_open(sheets, day, start_dt.time(), end_dt.time())
        overlap = False
        for r in today_rows.itertuples():
            if not (end_dt <= r.start or start_dt >= r.end):
                overlap = True
                break

        if not ok_hours:
            st.error(f"Outside operating hours ({hours_msg}).")
        if overlap:
            st.error("Overlaps an existing booking.")

        # ONE button only (unique key)
        clicked = st.button(
            "Confirm booking",
            type="primary",
            key="confirm_booking_btn",
            disabled=not (ok_hours and (not overlap)),
        )

        # IMPORTANT: this must stay indented inside the tab block
        if clicked:
            # Ensure Bookings exists
            if "Bookings" not in sheets:
                sheets["Bookings"] = pd.DataFrame(
                    columns=[
                        "booking_id",
                        "user_id",
                        "machine_id",
                        "start",
                        "end",
                        "purpose",
                        "notes",
                        "status",
                    ]
                )

            B = sheets["Bookings"]

            # Next ID
            next_id = (
                1
                if B.empty
                else int(pd.to_numeric(B["booking_id"], errors="coerce").fillna(0).max()) + 1
            )

            # Append booking
            new = pd.DataFrame(
                [
                    [
                        next_id,
                        int(me["user_id"]),
                        mid,
                        start_dt,
                        end_dt,
                        "use",
                        "",
                        "confirmed",
                    ]
                ],
                columns=B.columns,
            )
            sheets["Bookings"] = pd.concat([B, new], ignore_index=True)

            # Save and force reload so the new booking shows immediately
            save_db(sheets)   # make sure your save_db calls load_db.clear() inside
            st.success("Booked.")
            st.rerun()

with tabs[1]:
    st.subheader("Calendar")
    sel = st.selectbox("Machine", [f"{r.machine_id} - {r.machine_name}" for r in sheets["Machines"].itertuples()])
    mid=int(sel.split(" - ")[0])
    base_day = st.date_input("Day", value=date.today(), format=DFMT)
    view = st.radio("View", ["Day","Week"], horizontal=True)
    if view=="Day":
        DB = day_bookings(sheets, mid, base_day).copy()
        U = sheets["Users"]
        if not DB.empty:
            DB = DB.merge(U[["user_id","name"]], on="user_id", how="left")
            cols = [c for c in ["start","end","name","purpose","status"] if c in DB.columns]
            if cols: DB = DB[cols]
        st.dataframe(DB, use_container_width=True, hide_index=True)

    else:
        start_w = base_day - timedelta(days=base_day.weekday())
        rows=[]
        for d in range(7):
            dd=start_w+timedelta(days=d)
            for r in day_bookings(sheets, mid, dd).itertuples():
                rows.append([dd, r.start.time(), r.end.time(), r.purpose])
        st.dataframe(pd.DataFrame(rows, columns=["day","start","end","purpose"]), use_container_width=True, hide_index=True)

with tabs[2]:
    st.subheader("Mentoring & Competency Requests")
    if not me: st.info("Sign in to request mentoring.")
    else:
        L=sheets["Licences"]
        lic_map={f"{r.licence_id} - {r.licence_name}": int(r.licence_id) for r in L.itertuples()}
        sel = st.selectbox("Skill / Machine licence", list(lic_map.keys()))
        msg = st.text_area("What do you need help with? (optional)")
        UL=sheets.get("UserLicences", pd.DataFrame())
        U=sheets["Users"]
        today=pd.Timestamp.today().normalize()
        UL["valid_from"]=pd.to_datetime(UL["valid_from"], errors="coerce")
        UL["valid_to"]=pd.to_datetime(UL["valid_to"], errors="coerce")
        lic_id=lic_map[sel]
        mentors=UL[(UL["licence_id"]==lic_id)&(UL["valid_from"]<=today)&(UL["valid_to"]>=today)]["user_id"].astype(int).tolist()
        mentors=U[(U["user_id"].isin(mentors)) & (U["role"].isin(["admin","superuser"]))][["name","email","phone"]]
        if mentors.empty: st.warning("No listed mentors for this skill yet.")
        else:
            st.markdown("**Suggested mentors:**")
            st.dataframe(mentors, hide_index=True, use_container_width=True)
        if st.button("Submit mentoring request", type="primary"):
            AR = sheets.get("AssistanceRequests", pd.DataFrame(columns=["request_id","requester_user_id","licence_id","message","created","status","handled_by","handled_on","outcome","notes"]))
            req_id=1 if AR.empty else int(pd.to_numeric(AR["request_id"], errors="coerce").fillna(0).max())+1
            new=pd.DataFrame([[req_id, int(me["user_id"]), lic_id, msg, pd.Timestamp.today(), "open", None, None, None, None]], columns=AR.columns)
            sheets["AssistanceRequests"]=pd.concat([AR,new], ignore_index=True); save_db(sheets); st.success("Request submitted."); st.rerun()
    if me:
        AR2=sheets.get("AssistanceRequests", pd.DataFrame())
        mine=AR2[AR2["requester_user_id"]==int(me["user_id"])]
        if mine.empty: st.info("No requests yet.")
        else: st.dataframe(mine.sort_values("created", ascending=False), hide_index=True, use_container_width=True)

with tabs[3]:
    st.subheader("Issues & Maintenance")
    msel = st.selectbox("Machine", [f"{r.machine_id} - {r.machine_name}" for r in sheets["Machines"].itertuples()]
        key="book_machine_sel")
    mid = int(msel.split(" - ")[0])
    text = st.text_area("Describe an issue")
    if me and st.button("Submit issue"):
        I=sheets["Issues"]; iid=int(pd.to_numeric(I["issue_id"], errors="coerce").fillna(0).max())+1 if not I.empty else 1
        new=pd.DataFrame([[iid, mid, int(me["user_id"]), pd.Timestamp.today(), "open", text]], columns=I.columns)
        sheets["Issues"]=pd.concat([I,new], ignore_index=True); save_db(sheets); st.success("Issue logged."); st.rerun()
    I = sheets["Issues"].copy()
    Iv = I.copy()
    U = sheets["Users"]; M = sheets["Machines"]
    if not Iv.empty:
        if "user_id" in Iv.columns:
            Iv = Iv.merge(U[["user_id","name"]], on="user_id", how="left")
        if "machine_id" in Iv.columns:
            Iv = Iv.merge(M[["machine_id","machine_name"]], on="machine_id", how="left")
        cols = [c for c in ["issue_id","created","machine_id","machine_name","user_id","name","category","severity","status","notes"] if c in Iv.columns]
        if cols:
            Iv = Iv[cols]
    st.dataframe(Iv.sort_values("created", ascending=False), use_container_width=True, hide_index=True)


with tabs[4]:
    if not me or me["role"] not in ("admin","superuser"): st.info("Admins only.")
    else:
        at = st.tabs(["Users","Licences","User Licences","Competency","Machines","Subscriptions","Hours & Holidays","Newsletter","Settings"])
        with at[0]:
            st.markdown("### Users")
            st.dataframe(sheets["Users"][["user_id","name","role","email","phone","birth_date","joined_date","newsletter_opt_in"]], use_container_width=True, hide_index=True)
        with at[1]:
            st.markdown("### Licences")
            st.dataframe(sheets["Licences"], use_container_width=True, hide_index=True)
        with at[2]:
            st.markdown("### User licencing")
            U=sheets["Users"]; L=sheets["Licences"]; UL=sheets["UserLicences"].copy()
            ulabel=st.selectbox("Member", [f"{r.user_id} - {r.name}" for r in U.itertuples()])
            llabel=st.selectbox("Licence", [f"{r.licence_id} - {r.licence_name}" for r in L.itertuples()])
            vf=st.date_input("Valid from", format=DFMT); vt=st.date_input("Valid to", format=DFMT)
            if st.button("Grant licence"):
                uid=int(ulabel.split(" - ")[0]); lid=int(llabel.split(" - ")[0])
                new=pd.DataFrame([[uid,lid,pd.Timestamp(vf),pd.Timestamp(vt)]], columns=UL.columns)
                sheets["UserLicences"]=pd.concat([UL,new], ignore_index=True); save_db(sheets); st.success("Licence granted."); st.rerun()
            ULv = sheets["UserLicences"].copy()
            U = sheets["Users"]; L = sheets["Licences"]
            if not ULv.empty:
                ULv = ULv.merge(U[["user_id","name"]], on="user_id", how="left")
                ULv = ULv.merge(L[["licence_id","licence_name"]], on="licence_id", how="left")
                # reorder if present
                cols = [c for c in ["user_id","name","licence_id","licence_name","valid_from","valid_to","notes"] if c in ULv.columns]
                if cols:
                    ULv = ULv[cols]
            st.dataframe(ULv, use_container_width=True, hide_index=True)

        with at[3]:
            st.markdown("### Competency Assessments")
            AR=sheets.get("AssistanceRequests", pd.DataFrame(columns=["request_id","requester_user_id","licence_id","message","created","status","handled_by","handled_on","outcome","notes"])).copy()
            for c in ["status","handled_by","handled_on","outcome","notes"]:
                if c not in AR.columns: AR[c]=None
            U=sheets["Users"]; L=sheets["Licences"]
            open_reqs=AR[AR["status"].fillna("open").isin(["open","in_review"])]
            if open_reqs.empty: st.info("No open requests.")
            else:
                open_display = open_reqs.copy()
                open_display = open_display.merge(U[["user_id","name","email"]], left_on="request_user_id", right_on="user_id", how="left")
                open_display = open_display.merge(L[["licence_id","licence_name"]], on="licence_id", how="left")
                cols = [c for c in ["request_id","request_user_id","name","email","licence_id","licence_name","message","status","created_on"] if c in open_display.columns]
                if cols:
                    open_display = open_display[cols]
                st.dataframe(open_display, use_container_width=True, hide_index=True)

                sel = st.selectbox("Select request id", open_reqs["request_id"].tolist())
                req = open_reqs[open_reqs["request_id"]==sel].iloc[0]
                st.write(f"**Member:** {U.loc[U['user_id']==req.requester_user_id,'name'].iloc[0]}  •  **Licence:** {L.loc[L['licence_id']==req.licence_id,'licence_name'].iloc[0]}")
                notes=st.text_area("Assessment notes")
                outcome=st.radio("Outcome", ["pass","more_training","fail"], horizontal=True)
                grant=st.checkbox("Issue licence on pass", value=True)
                valid_to=st.date_input("Valid to", format=DFMT)
                if st.button("Save outcome"):
                    AR.loc[AR["request_id"]==sel, ["status","handled_by","handled_on","outcome","notes"]] = ["closed", int(me["user_id"]), pd.Timestamp.today(), outcome, notes]
                    sheets["AssistanceRequests"]=AR
                    if outcome=="pass" and             Mv = sheets["Machines"].copy()
            L = sheets["Licences"]
            if "licence_id" in Mv.columns and not Mv.empty:
                           Sv = sheets["Subscriptions"].copy()
            U = sheets["Users"]
            if not Sv.empty and "user_id" in Sv.columns:
                Sv = Sv.merge(U[["user_id","name"]], on="user_id", how="left")
                cols = [c for c in ["user_id","name","type","start_date","end_date","amount","paid","discount_percent","discount_reason"] if c in Sv.columns]
                if cols:
                    Sv = Sv[cols]
            st.dataframe(Sv, use_container_width=True, hide_index=True)
  # position human label near id
                if "licence_name" in Mv.columns:
                    lic = Mv.pop("licence_name")
                    Mv.insert(list(Mv.columns).index("licence_id")+1, "licence", lic)
            st.dataframe(Mv, use_container_width=True, hide_index=True)
ser_id","licence_id","valid_from","valid_to"]))
                        new = pd.DataFrame([[int(req.requester_user_id), int(req.licence_id), pd.Timestamp.today().normalize(), pd.Timestamp(valid_to)]], columns=UL.columns)
                        sheets["UserLicences"]=pd.concat([UL,new], ignore_index=True)
                    save_db(sheets); st.success("Saved."); st.rerun()
        with at[4]:
            st.markdown("### Machines")
            st.dataframe(sheets["Machines"], use_container_width=True, hide_index=True)
        with at[5]:
            st.markdown("### Subscriptions")
            st.dataframe(sheets["Subscriptions"], use_container_width=True, hide_index=True)
        with at[6]:
            st.markdown("### Weekly operating hours")
            st.dataframe(sheets["OperatingHours"], use_container_width=True, hide_index=True)
        with at[7]:
            st.markdown("### Newsletter")
            T=sheets.get("Templates", pd.DataFrame(columns=["key","text"]))
            row=T[T["key"]=="newsletter_prompt"]
            prompt=row.iloc[0]["text"] if not row.empty else ""
            txt=st.text_area("Newsletter Prompt", prompt, height=200)
            if st.button("Save prompt"):
                if row.empty: T=pd.concat([T, pd.DataFrame([["newsletter_prompt", txt]], columns=["key","text"])], ignore_index=True)
                else: T.loc[T["key"]=="newsletter_prompt","text"]=txt
                sheets["Templates"]=T; save_db(sheets); st.success("Prompt saved.")
        with at[8]:
            st.markdown("### Settings")
            S=sheets.get("Settings")
            st.dataframe(S, use_container_width=True, hide_index=True)
