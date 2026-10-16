from datetime import datetime, timedelta
import io
import os
import platform
import re
import shutil
import subprocess
import tempfile
import unicodedata
from pathlib import Path

import streamlit as st

try:
  from zoneinfo import ZoneInfo

  ATHENS_TZ = ZoneInfo("Europe/Athens")
except Exception:  # π.χ. Windows χωρίς το πακέτο tzdata
  ATHENS_TZ = None

st.set_page_config(
    page_title="Αυτοματοποίηση Αδειών Σχολείου", page_icon="📝", layout="wide"
)

PLACEHOLDER_TEACHER = "-- Επιλέξτε Εκπαιδευτικό --"
PLACEHOLDER_TEMPLATE = "-- Επιλέξτε το Έγγραφο της Άδειας --"
TEACHERS_PATH = "teachers.xlsx"
TEMPLATE_FOLDER = "templates"
LOG_PATH = "log_adeies.xlsx"
INIT_PATH = "arxikes_adeies.xlsx"

NUM_TO_GREEK = {
    1: "μίας ημέρας (1)",
    2: "δύο ημερών (2)",
    3: "τριών ημερών (3)",
    4: "τεσσάρων ημερών (4)",
    5: "πέντε ημερών (5)",
    6: "έξι ημερών (6)",
    7: "επτά ημερών (7)",
    8: "οκτώ ημερών (8)",
    9: "εννέα ημερών (9)",
    10: "δέκα ημερών (10)",
}

meres_evdomadas = {
    0: "ΔΕΥΤΕΡΑ",
    1: "ΤΡΙΤΗ",
    2: "ΤΕΤΑΡΤΗ",
    3: "ΠΕΜΠΤΗ",
    4: "ΠΑΡΑΣΚΕΥΗ",
    5: "ΣΑΒΒΑΤΟ",
    6: "ΚΥΡΙΑΚΗ",
}

LEAVE_LABELS = {
    "normal": "Κανονική",
    "sick": "Αναρρωτική",
    "sick_decl": "Αναρρωτική με υπεύθυνη δήλωση",
    "other": "Άλλη",
}
LABEL_TO_TYPE = {v: k for k, v in LEAVE_LABELS.items()}

# Ο τύπος της άδειας αναγνωρίζεται από το όνομα του αρχείου προτύπου
# (χωρίς τόνους, πεζά/κεφαλαία αδιάφορα). Η σειρά μετράει: πρώτα ελέγχεται
# η αναρρωτική με υπεύθυνη δήλωση. Προσαρμόστε τις λέξεις στα ονόματα των
# δικών σας προτύπων.
TYPE_KEYWORDS = [
    ("sick_decl", ["υπευθυν", "ypeuthyn", "ypefthyn"]),
    ("sick", ["αναρρωτ", "anarrot"]),
    ("normal", ["κανονικ", "kanonik"]),
]

# Όρια ανά ημερολογιακό έτος.
LIMITS = {
    "normal_days": 10,  # κανονικές: μέχρι 10 ημέρες
    "sick_days": 8,  # όλες οι αναρρωτικές μαζί: μέχρι 8 ημέρες
    "decl_count": 2,  # αναρρωτικές με υπεύθυνη δήλωση: μέχρι 2 άδειες
    "decl_max_days": 1,  # και διάρκειας μίας ημέρας η καθεμία
    "decl_forbidden_weekdays": (0, 4),  # όχι Δευτέρα (0) και Παρασκευή (4)
}

# Περίοδος στην οποία δεν δίνονται άδειες: από 24/12 έως 7/1 (και τις δύο ημέρες).
BLACKOUT_START = (12, 24)
BLACKOUT_END = (1, 7)

# Στήλες του αρχείου αρχικοποίησης ↔ μετρητές. Οι αναρρωτικές (ημέρες και
# άδειες) περιλαμβάνουν και αυτές με υπεύθυνη δήλωση.
INIT_COLS = {
    "normal_days": "Κανονικές ημέρες",
    "normal_count": "Κανονικές άδειες",
    "sick_days": "Αναρρωτικές ημέρες",
    "sick_count": "Αναρρωτικές άδειες",
    "decl_count": "Αναρρωτικές με υπ. δήλωση (άδειες)",
}


def in_blackout(d):
  md = (d.month, d.day)
  return md >= BLACKOUT_START or md <= BLACKOUT_END


def normalize_text(text):
  text = unicodedata.normalize("NFD", str(text))
  text = "".join(c for c in text if unicodedata.category(c) != "Mn")
  return text.lower()


def detect_leave_type(template_name):
  name = normalize_text(template_name or "")
  for leave_type, keywords in TYPE_KEYWORDS:
    if any(k in name for k in keywords):
      return leave_type
  return "other"


def days_by_year(start, n):
  """Πόσες από τις n ημέρες πέφτουν σε κάθε ημερολογιακό έτος."""
  result = {}
  for i in range(int(n)):
    year = (start + timedelta(days=i)).year
    result[year] = result.get(year, 0) + 1
  return result


def _clean(value):
  if value is None or value != value:  # None ή NaN
    return ""
  return str(value).strip()


def parse_date(value):
  if value is None or value != value:
    return None
  if hasattr(value, "date") and callable(value.date):
    return value.date()
  text = str(value).strip()
  for fmt in ("%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
    try:
      return datetime.strptime(text, fmt).date()
    except ValueError:
      continue
  return None


def _to_int(value, default=0):
  try:
    return int(float(_clean(value)))
  except ValueError:
    return default


def _same_teacher(row, mitrwo, eponymo, onoma):
  """Ταύτιση με ΑΜ όταν υπάρχει, αλλιώς με επώνυμο και όνομα."""
  row_am = _clean(row.get("ΑΜ"))
  mitrwo = str(mitrwo).strip()
  if row_am and mitrwo:
    return row_am == mitrwo
  return _clean(row.get("Επώνυμο")) == str(eponymo).strip() and _clean(
      row.get("Όνομα")
  ) == str(onoma).strip()


def leave_stats(log_df, mitrwo, eponymo, onoma, year, init_df=None):
  """Μετρά τις άδειες του εκπαιδευτικού για το δοσμένο έτος: οι αρχικές τιμές
  (αν έχουν οριστεί) συν όσες έχουν καταγραφεί στο ιστορικό."""
  stats = {key: 0 for key in INIT_COLS}

  if init_df is not None:
    for row in init_df.to_dict("records"):
      if _to_int(row.get("Έτος"), -1) == year and _same_teacher(
          row, mitrwo, eponymo, onoma
      ):
        for key, column in INIT_COLS.items():
          stats[key] += _to_int(row.get(column))

  if log_df is None:
    return stats

  for row in log_df.to_dict("records"):
    if not _same_teacher(row, mitrwo, eponymo, onoma):
      continue

    leave_type = LABEL_TO_TYPE.get(
        _clean(row.get("Τύπος Άδειας"))
    ) or detect_leave_type(_clean(row.get("Είδος Έντυπου")))
    start = parse_date(row.get("Έναρξη"))
    try:
      n = int(float(_clean(row.get("Ημέρες"))))
    except ValueError:
      continue
    if start is None or n < 1:
      continue

    days_in_year = days_by_year(start, n).get(year, 0)
    if days_in_year == 0:
      continue

    if leave_type == "normal":
      stats["normal_days"] += days_in_year
      stats["normal_count"] += 1
    elif leave_type in ("sick", "sick_decl"):
      stats["sick_days"] += days_in_year
      stats["sick_count"] += 1
      if leave_type == "sick_decl":
        stats["decl_count"] += 1
  return stats


def check_leave_rules(
    leave_type, start, n, log_df, mitrwo, eponymo, onoma, init_df=None
):
  """Επιστρέφει λίστα με τους κανόνες που παραβιάζονται (κενή = εντάξει)."""
  problems = []
  n = int(n)

  blocked = [
      start + timedelta(days=i)
      for i in range(n)
      if in_blackout(start + timedelta(days=i))
  ]
  if blocked:
    problems.append(
        "Δεν δίνονται άδειες από 24/12 έως 8/1 (η άδεια περιλαμβάνει την"
        f" {blocked[0].strftime('%d/%m/%Y')})."
    )

  if leave_type not in ("normal", "sick", "sick_decl"):
    return problems

  for year, days in days_by_year(start, n).items():
    s = leave_stats(log_df, mitrwo, eponymo, onoma, year, init_df)
    if leave_type == "normal":
      remaining = LIMITS["normal_days"] - s["normal_days"]
      if days > remaining:
        problems.append(
            f"Κανονική άδεια: για το {year} απομένουν {max(remaining, 0)} από"
            f" {LIMITS['normal_days']} ημέρες, ενώ ζητούνται {days}."
        )
    else:
      remaining = LIMITS["sick_days"] - s["sick_days"]
      if days > remaining:
        problems.append(
            f"Αναρρωτικές: για το {year} απομένουν {max(remaining, 0)} από"
            f" {LIMITS['sick_days']} ημέρες, ενώ ζητούνται {days}."
        )

  if leave_type == "sick_decl":
    if n > LIMITS["decl_max_days"]:
      problems.append(
          "Αναρρωτική με υπεύθυνη δήλωση: η διάρκεια πρέπει να είναι μία"
          " ημέρα."
      )
    forbidden = [
        meres_evdomadas[(start + timedelta(days=i)).weekday()]
        for i in range(n)
        if (start + timedelta(days=i)).weekday()
        in LIMITS["decl_forbidden_weekdays"]
    ]
    if forbidden:
      problems.append(
          "Αναρρωτική με υπεύθυνη δήλωση: δεν επιτρέπεται να περιλαμβάνει"
          f" {' και '.join(dict.fromkeys(forbidden))}."
      )
    s = leave_stats(log_df, mitrwo, eponymo, onoma, start.year, init_df)
    if s["decl_count"] >= LIMITS["decl_count"]:
      problems.append(
          f"Αναρρωτική με υπεύθυνη δήλωση: έχουν ήδη εκδοθεί {s['decl_count']}"
          f" από {LIMITS['decl_count']} για το {start.year}."
      )
  return problems


def render_balance(log_df, mitrwo, eponymo, onoma, year, init_df=None):
  s = leave_stats(log_df, mitrwo, eponymo, onoma, year, init_df)
  st.subheader(f"Άδειες εκπαιδευτικού για το έτος {year}")
  c1, c2, c3 = st.columns(3)
  with c1:
    st.metric(
        "Κανονικές – υπόλοιπο",
        f"{max(LIMITS['normal_days'] - s['normal_days'], 0)} ημέρες",
    )
    st.caption(
        f"Έχει λάβει {s['normal_count']} άδειες ({s['normal_days']} από"
        f" {LIMITS['normal_days']} ημέρες)"
    )
  with c2:
    st.metric(
        "Αναρρωτικές (σύνολο) – υπόλοιπο",
        f"{max(LIMITS['sick_days'] - s['sick_days'], 0)} ημέρες",
    )
    st.caption(
        f"Έχει λάβει {s['sick_count']} άδειες ({s['sick_days']} από"
        f" {LIMITS['sick_days']} ημέρες)"
    )
  with c3:
    st.metric(
        "Αναρρωτικές με υπ. δήλωση – υπόλοιπο",
        f"{max(LIMITS['decl_count'] - s['decl_count'], 0)} άδειες",
    )
    st.caption(f"Έχει λάβει {s['decl_count']} από {LIMITS['decl_count']}")


def now_athens():
  """Τρέχουσα ώρα Ελλάδας, ανεξάρτητα από τη ζώνη ώρας του server."""
  return datetime.now(ATHENS_TZ) if ATHENS_TZ else datetime.now()


def safe_filename_part(text):
  """Αφήνει μόνο γράμματα (και ελληνικά), ψηφία, _ και -."""
  cleaned = re.sub(r"[^\w\-]+", "_", str(text).strip())
  return cleaned.strip("_") or "xwris_onoma"


@st.cache_data
def _load_teachers_cached(excel_path, mtime):
  """Το mtime είναι μέρος του κλειδιού cache: όταν αλλάξει το αρχείο,
  η cache ανανεώνεται αυτόματα."""
  import pandas as pd

  df = pd.read_excel(excel_path, dtype={"mitrwo": str})
  df = df.fillna("")

  # Διπλότυπα ονοματεπώνυμα: προσθέτουμε το ΑΜ ώστε να μη χάνεται κανείς.
  counts = df["full_name"].value_counts()
  duplicates = set(counts[counts > 1].index)
  if duplicates:
    df["full_name"] = df.apply(
        lambda r: (
            f"{r['full_name']} (ΑΜ {r['mitrwo']})"
            if r["full_name"] in duplicates
            else r["full_name"]
        ),
        axis=1,
    )

  return df.set_index("full_name").to_dict("index")


def load_teachers():
  if not os.path.exists(TEACHERS_PATH):
    st.error("Δεν βρέθηκε το αρχείο teachers.xlsx!")
    return {}
  try:
    return _load_teachers_cached(TEACHERS_PATH, os.path.getmtime(TEACHERS_PATH))
  except Exception as e:
    st.error(f"Σφάλμα κατά τη φόρτωση του αρχείου teachers.xlsx: {e}")
    return {}


@st.cache_data
def _list_templates_cached(folder, folder_mtime):
  return [PLACEHOLDER_TEMPLATE] + [
      f
      for f in sorted(os.listdir(folder))
      if f.endswith(".docx") and not f.startswith("~$")
  ]


def get_available_templates(folder):
  if not os.path.exists(folder):
    os.makedirs(folder)
  return _list_templates_cached(folder, os.path.getmtime(folder))


def convert_docx_to_pdf(doc_path, pdf_path):
  if platform.system() == "Windows":
    import pythoncom
    from docx2pdf import convert

    pythoncom.CoInitialize()
    try:
      convert(doc_path, pdf_path)
    finally:
      pythoncom.CoUninitialize()
  else:
    # Ξεχωριστό προφίλ LibreOffice ανά μετατροπή, ώστε να μη συγκρούονται
    # δύο ταυτόχρονοι χρήστες, και timeout ώστε να μην κρεμάσει.
    profile_dir = tempfile.mkdtemp(prefix="lo_profile_")
    try:
      cmd = [
          "libreoffice",
          f"-env:UserInstallation={Path(profile_dir).as_uri()}",
          "--headless",
          "--convert-to",
          "pdf",
          doc_path,
          "--outdir",
          os.path.dirname(pdf_path) or ".",
      ]
      subprocess.run(cmd, check=True, timeout=120, capture_output=True)
    finally:
      shutil.rmtree(profile_dir, ignore_errors=True)


@st.cache_data
def _read_log_cached(path, mtime):
  import pandas as pd

  return pd.read_excel(path, dtype={"ΑΜ": str})


def read_log():
  if not os.path.exists(LOG_PATH):
    return None
  try:
    return _read_log_cached(LOG_PATH, os.path.getmtime(LOG_PATH))
  except Exception:
    return None


def save_to_log(data):
  import pandas as pd

  new_entry = pd.DataFrame([data])

  if os.path.exists(LOG_PATH):
    df = pd.read_excel(LOG_PATH, dtype={"ΑΜ": str})
    df = pd.concat([df, new_entry], ignore_index=True)
  else:
    df = new_entry

  df.to_excel(LOG_PATH, index=False)


@st.cache_data
def _read_init_cached(path, mtime):
  import pandas as pd

  return pd.read_excel(path, dtype={"ΑΜ": str})


def read_init():
  if not os.path.exists(INIT_PATH):
    return None
  try:
    return _read_init_cached(INIT_PATH, os.path.getmtime(INIT_PATH))
  except Exception:
    return None


def save_init(row):
  """Αποθηκεύει αρχικές τιμές. Αν υπάρχει ήδη εγγραφή για τον ίδιο
  εκπαιδευτικό και έτος, αντικαθίσταται."""
  import pandas as pd

  new_entry = pd.DataFrame([row])
  if os.path.exists(INIT_PATH):
    df = pd.read_excel(INIT_PATH, dtype={"ΑΜ": str})
    keep = [
        not (
            _to_int(r.get("Έτος"), -1) == row["Έτος"]
            and _same_teacher(r, row["ΑΜ"], row["Επώνυμο"], row["Όνομα"])
        )
        for r in df.to_dict("records")
    ]
    df = pd.concat([df[keep], new_entry], ignore_index=True)
  else:
    df = new_entry
  df.to_excel(INIT_PATH, index=False)


EKPAIDEYTIKOI_DATA = load_teachers()

st.title("📝 Έκδοση Άδειας Εκπαιδευτικού")

teacher_options = [PLACEHOLDER_TEACHER] + list(EKPAIDEYTIKOI_DATA.keys())
selected_name = st.selectbox(
    "1. Αναζήτηση Εκπαιδευτικού", options=teacher_options
)

# Τα υπόλοιπα και η αρχικοποίηση εμφανίζονται εδώ, αμέσως μετά την επιλογή
# εκπαιδευτικού. Το πλαίσιο γεμίζει στο τέλος της εκτέλεσης, ώστε να
# περιλαμβάνει και την άδεια που μόλις εκδόθηκε.
summary_box = st.container()

available_templates = get_available_templates(TEMPLATE_FOLDER)
selected_template_name = PLACEHOLDER_TEMPLATE

if len(available_templates) > 1:
  selected_template_name = st.selectbox(
      "2. Επιλέξτε το έντυπο προς έκδοση:", options=available_templates
  )
  template_path = (
      os.path.join(TEMPLATE_FOLDER, selected_template_name)
      if selected_template_name != PLACEHOLDER_TEMPLATE
      else None
  )
else:
  st.warning("⚠️ Δεν βρέθηκαν αρχεία .docx στον φάκελο /templates.")
  template_path = None

st.divider()

if selected_name != PLACEHOLDER_TEACHER:
  teacher_data = EKPAIDEYTIKOI_DATA[selected_name]
else:
  teacher_data = {
      "eponymo": "",
      "onoma": "",
      "eponymo_gen": "",
      "onoma_gen": "",
      "klados": "",
      "mitrwo": "",
      "email": "",
  }

st.subheader("Στοιχεία Άδειας")
col1, col2, col3, col4 = st.columns(4)

with col1:
  eponymo = st.text_input(
      "Επώνυμο", value=teacher_data.get("eponymo", ""), disabled=True
  )
  onoma = st.text_input(
      "Όνομα", value=teacher_data.get("onoma", ""), disabled=True
  )
  mitrwo = st.text_input(
      "ΑΜ", value=teacher_data.get("mitrwo", ""), disabled=True
  )
  klados = st.text_input(
      "Κλάδος", value=teacher_data.get("klados", ""), disabled=True
  )

with col2:
  protocollo_aithshs = st.text_input("Αριθμός Πρωτοκόλλου Αίτησης")
  days_number = st.number_input(
      "Αριθμός Ημερών", min_value=1, max_value=10, step=1
  )
  doctor = st.text_input("Γιατρός (γεν.)")

with col3:
  hmer_protocollou_aithshs = st.date_input(
      "Ημερομηνία Πρωτοκόλλου Αίτησης", format="DD/MM/YYYY"
  )
  arxh = st.date_input("Ημερομηνία Έναρξης Άδειας", format="DD/MM/YYYY")
  hmer_gnomateyshs = st.date_input(
      "Ημερομηνία Ιατρικής Γνωμάτευσης", format="DD/MM/YYYY"
  )

with col4:
  telos = arxh + timedelta(days=int(days_number) - 1)
  st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
  days_lektiko = NUM_TO_GREEK[int(days_number)]
  st.text_input(
      "Ημερομηνία Λήξης Άδειας", value=telos.strftime("%d/%m/%Y"), disabled=True
  )

leave_type = detect_leave_type(selected_template_name) if template_path else None
rule_problems = []
if selected_name != PLACEHOLDER_TEACHER and leave_type:
  rule_problems = check_leave_rules(
      leave_type,
      arxh,
      int(days_number),
      read_log(),
      mitrwo,
      eponymo,
      onoma,
      read_init(),
  )
  for problem in rule_problems:
    st.error(f"⛔ {problem}")
  if leave_type == "other":
    st.info(
        "ℹ️ Το έντυπο δεν αναγνωρίστηκε ως κανονική ή αναρρωτική άδεια, οπότε"
        " δεν εφαρμόζονται όρια ημερών (ισχύει μόνο η απαγόρευση 24/12–8/1)."
        " Ο τύπος αναγνωρίζεται από το όνομα του αρχείου του προτύπου."
    )

st.divider()

with st.form("leave_form"):
  st.subheader("Στοιχεία Πρωτοκόλλου Άδειας")
  protocollo_adeias = st.text_input("Πρωτόκολλο Άδειας")
  hmer_protocollou = st.date_input(
      "Ημερομηνία Πρωτοκόλλου Άδειας", format="DD/MM/YYYY"
  )
  make_pdf = st.checkbox("Δημιουργία και αρχείου PDF ταυτόχρονα", value=True)
  submitted = st.form_submit_button("🚀 Δημιουργία Εγγράφων")

def render_init_panel():
  with st.expander("⚙️ Αρχικοποίηση υπολοίπων εκπαιδευτικού"):
    st.caption(
        "Δηλώστε πόσες άδειες έχει ήδη πάρει ο εκπαιδευτικός μέχρι σήμερα για"
        " το επιλεγμένο έτος. Οι τιμές αυτές προστίθενται σε όσες εκδίδονται"
        " από την εφαρμογή. Αν υπάρχει ήδη αρχικοποίηση για το ίδιο έτος,"
        " αντικαθίσταται."
    )
    flash = st.session_state.pop("init_saved", None)
    if flash:
      st.success(flash)

    year = int(
        st.number_input(
            "Έτος",
            min_value=2020,
            max_value=2100,
            value=arxh.year,
            step=1,
            key=f"ini_year_{mitrwo}_{eponymo}_{onoma}",
        )
    )

    current = {}
    init_df = read_init()
    if init_df is not None:
      for row in init_df.to_dict("records"):
        if _to_int(row.get("Έτος"), -1) == year and _same_teacher(
            row, mitrwo, eponymo, onoma
        ):
          current = row

    k = f"{mitrwo}_{eponymo}_{onoma}_{year}"
    c1, c2, c3 = st.columns(3)
    values = {}
    with c1:
      values["normal_days"] = st.number_input(
          "Κανονικές – ημέρες που έχει πάρει",
          min_value=0,
          max_value=365,
          step=1,
          value=_to_int(current.get(INIT_COLS["normal_days"])),
          key=f"ini_nd_{k}",
      )
      values["normal_count"] = st.number_input(
          "Κανονικές – αριθμός αδειών",
          min_value=0,
          max_value=365,
          step=1,
          value=_to_int(current.get(INIT_COLS["normal_count"])),
          key=f"ini_nc_{k}",
      )
    with c2:
      values["sick_days"] = st.number_input(
          "Αναρρωτικές (όλες) – ημέρες",
          min_value=0,
          max_value=365,
          step=1,
          value=_to_int(current.get(INIT_COLS["sick_days"])),
          key=f"ini_sd_{k}",
      )
      values["sick_count"] = st.number_input(
          "Αναρρωτικές (όλες) – αριθμός αδειών",
          min_value=0,
          max_value=365,
          step=1,
          value=_to_int(current.get(INIT_COLS["sick_count"])),
          key=f"ini_sc_{k}",
      )
    with c3:
      values["decl_count"] = st.number_input(
          "Εκ των οποίων με υπεύθυνη δήλωση (άδειες)",
          min_value=0,
          max_value=365,
          step=1,
          value=_to_int(current.get(INIT_COLS["decl_count"])),
          key=f"ini_dc_{k}",
      )

    if st.button("💾 Αποθήκευση αρχικών τιμών", key=f"ini_save_{k}"):
      if values["decl_count"] > values["sick_count"]:
        st.error(
            "Οι άδειες με υπεύθυνη δήλωση δεν μπορεί να είναι περισσότερες από"
            " το σύνολο των αναρρωτικών αδειών."
        )
      elif values["sick_count"] > values["sick_days"]:
        st.error(
            "Ο αριθμός των αναρρωτικών αδειών δεν μπορεί να υπερβαίνει τις"
            " ημέρες τους."
        )
      elif values["normal_count"] > values["normal_days"]:
        st.error(
            "Ο αριθμός των κανονικών αδειών δεν μπορεί να υπερβαίνει τις"
            " ημέρες τους."
        )
      else:
        row = {
            "ΑΜ": str(mitrwo),
            "Επώνυμο": eponymo,
            "Όνομα": onoma,
            "Έτος": year,
        }
        for key, column in INIT_COLS.items():
          row[column] = int(values[key])
        try:
          save_init(row)
        except PermissionError:
          st.error(
              f"Δεν ήταν δυνατή η αποθήκευση: το αρχείο {INIT_PATH} είναι"
              " ανοιχτό σε άλλο πρόγραμμα. Κλείστε το και ξαναπροσπαθήστε."
          )
        else:
          st.session_state["init_saved"] = (
              f"✅ Αποθηκεύτηκαν οι αρχικές τιμές για το {year}."
          )
          st.rerun()


def render_bottom():
  """Υπόλοιπα και αρχικοποίηση. Καλείται στο τέλος κάθε εκτέλεσης, ώστε να
  περιλαμβάνεται και η άδεια που μόλις εκδόθηκε."""
  if selected_name == PLACEHOLDER_TEACHER:
    return
  with summary_box:
    render_balance(read_log(), mitrwo, eponymo, onoma, arxh.year, read_init())
    render_init_panel()


def stop_page():
  render_bottom()
  st.stop()


if submitted:
  if selected_name == PLACEHOLDER_TEACHER:
    st.warning("⚠️ Παρακαλώ επιλέξτε εκπαιδευτικό.")
    stop_page()

  if not protocollo_adeias.strip():
    st.warning("⚠️ Παρακαλώ συμπληρώστε τον Αριθμό Πρωτοκόλλου για να συνεχίσετε.")
    stop_page()

  if not template_path:
    st.error("⚠️ Παρακαλώ επιλέξτε ένα πρότυπο αρχείο Word.")
    stop_page()

  if rule_problems:
    st.error(
        "⛔ Η άδεια δεν μπορεί να εκδοθεί, γιατί παραβιάζει τα όρια (δείτε τα"
        " μηνύματα παραπάνω)."
    )
    stop_page()

  # Καθαρίζουμε το προηγούμενο αποτέλεσμα, ώστε να μην μείνει παλιό έγγραφο.
  for key in (
      "generated_docx",
      "generated_pdf",
      "generated_for",
      "base_filename",
      "teacher_fullname",
  ):
    st.session_state.pop(key, None)

  try:
    from docxtpl import DocxTemplate

    doc = DocxTemplate(template_path)
    now = now_athens()

    onoma_meras1 = meres_evdomadas[arxh.weekday()]

    if days_number > 1:
      onoma_meras2 = meres_evdomadas[telos.weekday()]
      full_enarksi_text = f"από την {onoma_meras1}, {arxh.strftime('%d/%m/%Y')}"
      full_lhxh_text = f"έως και την {onoma_meras2}, {telos.strftime('%d/%m/%Y')}"
    else:
      full_enarksi_text = f"την {onoma_meras1}, {arxh.strftime('%d/%m/%Y')}"
      full_lhxh_text = ""

    hmer_prot_str = hmer_protocollou.strftime("%d/%m/%Y")
    hmer_ait_str = hmer_protocollou_aithshs.strftime("%d/%m/%Y")
    hmer_gnom_str = hmer_gnomateyshs.strftime("%d/%m/%Y")
    hmer_ekd_str = now.strftime("%d/%m/%Y")

    context = {
        "eponymo": eponymo,
        "onoma": onoma,
        "eponymo_gen": teacher_data.get("eponymo_gen", ""),
        "onoma_gen": teacher_data.get("onoma_gen", ""),
        "hmer_protocollou": hmer_prot_str,
        "protocollo_adeias": protocollo_adeias,
        "hmer_protocollou_aithshs": hmer_ait_str,
        "protocollo_aithshs": protocollo_aithshs,
        "klados": klados,
        "mitrwo": mitrwo,
        "days_number": days_number,
        "days_lektiko": days_lektiko,
        "doctor": doctor,
        "hmer_gnomateyshs": hmer_gnom_str,
        "arxh": full_enarksi_text,
        "telos": full_lhxh_text,
        "hmerominia_ekdosis": hmer_ekd_str,
    }

    doc.render(context)

    target_stream = io.BytesIO()
    doc.save(target_stream)
    docx_bytes = target_stream.getvalue()

    base_filename = (
        f"Adeia_{safe_filename_part(eponymo)}_"
        f"{safe_filename_part(protocollo_adeias)}"
    )

    st.session_state["generated_docx"] = docx_bytes
    st.session_state["generated_pdf"] = None
    st.session_state["generated_for"] = selected_name
    st.session_state["base_filename"] = base_filename
    st.session_state["teacher_fullname"] = f"{eponymo} {onoma}"

    # Καταγραφή στο ιστορικό ΜΟΝΟ αφού δημιουργήθηκε επιτυχώς το έγγραφο.
    log_data = {
        "Ημερομηνία Έκδοσης": now.strftime("%d/%m/%Y %H:%M"),
        "Επώνυμο": eponymo,
        "Όνομα": onoma,
        "ΑΜ": mitrwo,
        "Κλάδος": klados,
        "Είδος Έντυπου": selected_template_name,
        "Τύπος Άδειας": LEAVE_LABELS[leave_type],
        "Ημέρες": days_number,
        "Έναρξη": arxh.strftime("%d/%m/%Y"),
        "Λήξη": telos.strftime("%d/%m/%Y"),
        "Πρωτόκολλο": protocollo_adeias,
        "Ημερομηνία πρωτοκόλλου άδειας": hmer_prot_str,
    }
    try:
      save_to_log(log_data)
      st.toast("✅ Η εγγραφή καταγράφηκε στο ιστορικό!")
    except PermissionError:
      st.error(
          f"Δεν ήταν δυνατή η καταγραφή: το αρχείο {LOG_PATH} είναι ανοιχτό σε"
          " άλλο πρόγραμμα (π.χ. Excel). Κλείστε το και ξαναπροσπαθήστε."
      )
    except Exception as e:
      st.error(f"Σφάλμα κατά την καταγραφή στο ιστορικό: {e}")

    if make_pdf:
      try:
        with st.spinner("Μετατροπή σε PDF..."):
          # Μοναδικός προσωρινός φάκελος, καθαρίζεται αυτόματα.
          with tempfile.TemporaryDirectory() as tmp_dir:
            temp_word = os.path.join(tmp_dir, "adeia.docx")
            temp_pdf = os.path.join(tmp_dir, "adeia.pdf")

            with open(temp_word, "wb") as f:
              f.write(docx_bytes)

            convert_docx_to_pdf(temp_word, temp_pdf)

            if os.path.exists(temp_pdf):
              with open(temp_pdf, "rb") as f:
                st.session_state["generated_pdf"] = f.read()
            else:
              st.warning("⚠️ Η μετατροπή σε PDF δεν παρήγαγε αρχείο.")
      except subprocess.TimeoutExpired:
        st.warning(
            "⚠️ Η μετατροπή σε PDF ξεπέρασε τον χρόνο αναμονής. Το αρχείο Word"
            " είναι διαθέσιμο κανονικά."
        )
      except Exception as e:
        st.warning(
            f"⚠️ Η μετατροπή σε PDF απέτυχε ({e}). Το αρχείο Word είναι"
            " διαθέσιμο κανονικά."
        )

  except Exception as e:
    st.error(f"Παρουσιάστηκε σφάλμα: {e}")

# Εμφάνιση λήψεων μόνο για τον εκπαιδευτικό για τον οποίο δημιουργήθηκε το έγγραφο.
if (
    "generated_docx" in st.session_state
    and st.session_state.get("generated_for") == selected_name
):
  st.success(
      f"Η άδεια για τον/την **{st.session_state['teacher_fullname']}** είναι"
      " έτοιμη!"
  )
  col_dl1, col_dl2 = st.columns(2)

  with col_dl1:
    st.download_button(
        label="📥 Λήψη Word (.docx)",
        data=st.session_state["generated_docx"],
        file_name=f"{st.session_state['base_filename']}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )

  if st.session_state.get("generated_pdf"):
    with col_dl2:
      st.download_button(
          label="📄 Λήψη PDF (.pdf)",
          data=st.session_state["generated_pdf"],
          file_name=f"{st.session_state['base_filename']}.pdf",
          mime="application/pdf",
          use_container_width=True,
      )

render_bottom()