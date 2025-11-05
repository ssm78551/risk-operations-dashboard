import streamlit as st
import win32com.client
import datetime
import holidays
import pythoncom


# ---------------------------------------------------------
# Hjælpefunktion: find forrige bankdag via holidays.DK()
# ---------------------------------------------------------
def previous_business_day(date=None):
    if date is None:
        date = datetime.date.today()
    dk_holidays = holidays.DK(years=[date.year, date.year - 1])
    day = date - datetime.timedelta(days=1)
    while day.weekday() >= 5 or day in dk_holidays:
        day -= datetime.timedelta(days=1)
    return day


# ---------------------------------------------------------
# Hjælpefunktion: søg mails i mappe
# ---------------------------------------------------------
def scan_folder(folder, subjects_to_find, today):
    try:
        items = getattr(folder, "Items", None)
        if items is None:
            return
        items.Sort("[ReceivedTime]", True)
        for msg in items:
            if not hasattr(msg, "Subject") or not hasattr(msg, "ReceivedTime"):
                continue
            if msg.ReceivedTime.date() != today:
                continue
            subj = msg.Subject.strip()
            for s in subjects_to_find:
                if s in subj and not subjects_to_find[s]["found"]:
                    subjects_to_find[s]["found"] = True
                    subjects_to_find[s]["folder"] = folder.Name
                    subjects_to_find[s]["received"] = msg.ReceivedTime.strftime("%d.%m.%Y %H:%M")
    except Exception:
        pass


# ---------------------------------------------------------
# Rekursiv søgning i alle mapper
# ---------------------------------------------------------
def walk_folders(folder, subjects_to_find, today):
    try:
        scan_folder(folder, subjects_to_find, today)
        if all(v["found"] for v in subjects_to_find.values()):
            return True
        for sub in getattr(folder, "Folders", []):
            if walk_folders(sub, subjects_to_find, today):
                return True
    except Exception:
        pass
    return False


# ---------------------------------------------------------
# Tjek hele postkassen
# ---------------------------------------------------------
def check_all_mails():
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    today = datetime.date.today()
    prev_day = previous_business_day(today)

    today_str = today.strftime("%Y%m%d")
    prev_day_str = prev_day.strftime("%Y%m%dT230100")

    root_folder = None
    for account in outlook.Folders:
        if account.Name.lower() == "risikoafd@sparnord.dk":
            root_folder = account
            break
    if not root_folder:
        return None, today, prev_day

    eod_subject = f"FM batch finished loading data for EOD - {prev_day_str}"
    intraday1_subject = f"FM batch finished loading data for IntraDay - {today_str}T094500"
    intraday2_subject = f"FM batch finished loading data for IntraDay - {today_str}T131500"

    subjects_to_find = {
        "wf_Downloading_Files succeeded": {"found": False, "folder": None, "received": None},
        eod_subject: {"found": False, "folder": None, "received": None},
        "DSA_S_Calypso_PnL: EOD loaded": {"found": False, "folder": None, "received": None},
        intraday1_subject: {"found": False, "folder": None, "received": None},
        intraday2_subject: {"found": False, "folder": None, "received": None},
    }

    inbox = root_folder.Folders("Indbakke")
    scan_folder(inbox, subjects_to_find, today)
    if not all(v["found"] for v in subjects_to_find.values()):
        walk_folders(root_folder, subjects_to_find, today)
    return subjects_to_find, today, prev_day


# ---------------------------------------------------------
# Hjælpefunktioner til status
# ---------------------------------------------------------
def status_label(found, total, phase):
    pct = (found / total) if total else 0
    now = datetime.datetime.now().time()

    if phase == "EOD":
        if pct == 1:
            return "🟩 EOD status 3/3 OK", "ok"
        else:
            return f"🟥 EOD status {found}/{total} – mangler filer", "bad"

    if phase == "INTRA":
        if now < datetime.time(10, 30):
            return "🕒 Afventer Intra 1 (efter 10:30)", "wait"
        elif now < datetime.time(14, 0) and found < total:
            return f"🟨 Intra 1 klar – Intra 2 afventer", "warn"
        elif pct == 1:
            return "🟩 Intra status 2/2 OK", "ok"
        else:
            return f"🟥 Intra status {found}/{total}", "bad"


# ---------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------
st.set_page_config(page_title="Risk Operations Dashboard", page_icon="📊", layout="wide")

# --- CSS styling ---
st.markdown("""
<style>
body {background-color:#273142;color:#E3E6EB;font-family:'Segoe UI',sans-serif;}
h1.custom-title {
    text-align:center;
    font-family:'Segoe UI Semibold',sans-serif;
    font-size:28px;
    color:#CFE4FA;
    text-shadow:2px 2px 4px #1B2838;
    margin-bottom:4px;
    letter-spacing:0.5px;
}
h3.custom-subtitle {
    text-align:center;
    color:#8FBDFE;
    font-size:18px;
    margin-top:0;
}
div[data-testid="stVerticalBlock"] {
    background-color:#1E2835;
    border-radius:8px;
    padding:10px 18px 8px 18px;
    margin-bottom:8px;
    border:1px solid #2F3B4D;
}
div[data-testid="stMarkdownContainer"] { line-height:1.05; }

button[kind="primary"] {
    background-color:#4E9FFF!important;
    color:white!important;
    border-radius:4px!important;
    font-weight:500!important;
}
button[kind="primary"]:hover { background-color:#2E8BFF!important; }

/* Statusbjælker */
.success-banner {
    background-color: rgba(40,167,69,0.2);
    border-left: 5px solid #28A745;
    padding: 10px 14px;
    border-radius: 4px;
    font-weight: 500;
    margin-bottom: 12px;
}
.blink-banner {
    animation: blinkRed 1.2s infinite alternate;
    border-left: 5px solid #ff3b3b;
    padding: 10px 14px;
    border-radius: 4px;
    font-weight: 600;
    margin-bottom: 12px;
}
@keyframes blinkRed {
  0%   { background-color: rgba(255,0,0,0.1); color: #FFD1D1; }
  100% { background-color: rgba(255,0,0,0.3); color: white; }
}
</style>
""", unsafe_allow_html=True)

# --- Titel ---
st.markdown("<h1 class='custom-title'>📊 Risk Operations Dashboard</h1>", unsafe_allow_html=True)
st.markdown("<h3 class='custom-subtitle'>📬 Datawarehouse mails</h3>", unsafe_allow_html=True)
st.caption("Status for mails modtaget i 'dagen' på risikoafd@sparnord.dk")

# --- Scanning ---
if "results" not in st.session_state:
    with st.spinner("Scanner Outlook..."):
        results, today, prev_day = check_all_mails()
        st.session_state.results = results
        st.session_state.today = today
        st.session_state.prev_day = prev_day
else:
    results = st.session_state.results
    today = st.session_state.today
    prev_day = st.session_state.prev_day

if st.button("🔄 Opdater status"):
    with st.spinner("Opdaterer..."):
        results, today, prev_day = check_all_mails()
        st.session_state.results = results
        st.session_state.today = today
        st.session_state.prev_day = prev_day

results = st.session_state.get("results")
today = st.session_state.get("today", datetime.date.today())
prev_day = st.session_state.get("prev_day", previous_business_day())

# --- Beregn EOD status for banner ---
eod_keys = [k for k in results if "EOD" in k or "wf_Downloading" in k or "PnL" in k]
eod_found = sum(1 for k in eod_keys if results[k]["found"])

if eod_found == len(eod_keys):
    banner_class = "success-banner"
    banner_text = f"✅ Automatisk scanning gennemført! Alle {len(eod_keys)} EOD-mails er modtaget."
else:
    missing = len(eod_keys) - eod_found
    banner_class = "blink-banner"
    banner_text = (
        f"⚠️ Automatisk scanning færdig – men {missing} EOD-mail(s) mangler! "
        "Kontroller om FM batch er forsinket eller ikke kørt."
    )

# --- Vis banner ---
st.markdown(f"<div class='{banner_class}'>{banner_text}</div>", unsafe_allow_html=True)

# --- Opdaterknap ---
st.button("🔁 Opdater status manuelt")

st.markdown(f"📅 **Dagens dato:** {today.strftime('%A %d.%m.%Y')}")
st.markdown(f"🏦 **Forrige bankdag:** {prev_day.strftime('%A %d.%m.%Y')}")
st.divider()

# ---------------------------------------------------------
# Foldbare sektioner (uden blink)
# ---------------------------------------------------------
if results:
    eod_label, eod_state = status_label(eod_found, len(eod_keys), "EOD")
    intra_keys = [k for k in results if "IntraDay" in k]
    intra_found = sum(1 for k in intra_keys if results[k]["found"])
    intra_label, intra_state = status_label(intra_found, len(intra_keys), "INTRA")

    # --- EOD ---
    with st.expander(f"🌙 {eod_label}", expanded=(eod_state != "ok")):
        for subj in eod_keys:
            info = results[subj]
            if info["found"]:
                st.markdown(f"✅ **{subj}** (mappe: {info['folder']}, modtaget: {info['received']})")
            else:
                st.markdown(f"⚠️ **{subj}** (ikke fundet)")
        if eod_state == "ok":
            st.balloons()

    # --- IntraDay ---
    with st.expander(f"🌤️ {intra_label}", expanded=(intra_state != "ok")):
        for subj in intra_keys:
            info = results[subj]
            if info["found"]:
                st.markdown(f"✅ **{subj}** (mappe: {info['folder']}, modtaget: {info['received']})")
            else:
                st.markdown(f"⚠️ **{subj}** (ikke fundet)")
        if intra_state == "ok":
            st.balloons()
else:
    st.info("Ingen scanning gennemført endnu – klik 'Opdater status'.")
