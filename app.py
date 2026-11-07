# =====================================================
# DocuCraft Pro — ملف واحد كامل
# شغّله بـ: streamlit run app.py
# =====================================================
import os, io, re, json, html, base64, sqlite3, hashlib, time
from datetime import datetime, date, timedelta

import streamlit as st

# ---------- imports اختيارية (لو مش مثبتة، التطبيق يشتغل بدونها) ----------
try:
    import google.generativeai as genai
    _HAS_GEMINI = True
except ImportError:
    _HAS_GEMINI = False

try:
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    _HAS_RL = True
except ImportError:
    _HAS_RL = False

try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    _HAS_AR = True
except ImportError:
    _HAS_AR = False

try:
    from docx import Document as DocxDocument
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    _HAS_DOCX = True
except ImportError:
    _HAS_DOCX = False

try:
    from pptx import Presentation
    _HAS_PPTX = True
except ImportError:
    _HAS_PPTX = False

try:
    from gtts import gTTS
    _HAS_TTS = True
except ImportError:
    _HAS_TTS = False

try:
    import speech_recognition as sr
    _HAS_STT = True
except ImportError:
    _HAS_STT = False


# =====================================================
# CONFIG
# =====================================================
DESIGNS = {
    "academic": {"name": "أكاديمي", "primary": "#8B7355", "accent": "#D4AF37",
                 "bg": "#FAF8F3", "text": "#1F1B16", "font_ar": "Amiri"},
    "modern":   {"name": "عصري",   "primary": "#4F46E5", "accent": "#EC4899",
                 "bg": "#FFFFFF", "text": "#0F172A", "font_ar": "Cairo"},
    "minimal":  {"name": "بسيط",   "primary": "#111111", "accent": "#666666",
                 "bg": "#FFFFFF", "text": "#111111", "font_ar": "Tajawal"},
}
ARABIC_FONTS = ["Cairo", "Tajawal", "Amiri", "Noto Naskh Arabic", "Noto Kufi Arabic",
                "Almarai", "Changa", "El Messiri", "Lalezar", "Reem Kufi"]
LATIN_FONTS = ["Georgia", "Inter", "Helvetica", "Times New Roman", "Roboto"]
ALL_FONTS = ARABIC_FONTS + LATIN_FONTS
PAGE_SIZES = ["A4", "A5", "A6", "A3", "Letter", "Legal", "Square"]
PAGE_SIZES_MM = {"A4": (210, 297), "A5": (148, 210), "A6": (105, 148),
                 "A3": (297, 420), "Letter": (216, 279), "Legal": (216, 356),
                 "Square": (200, 200)}
MARGINS = {"ضيقة": 15, "عادية": 25, "واسعة": 40}
TRANSLATION_LANGUAGES = {"ar": "العربية", "en": "الإنجليزية", "fr": "الفرنسية",
                         "tr": "التركية", "ur": "الأردية", "zh": "الصينية",
                         "ja": "اليابانية", "fa": "الفارسية", "es": "الإسبانية",
                         "de": "الألمانية", "ru": "الروسية", "hi": "الهندية",
                         "it": "الإيطالية", "pt": "البرتغالية"}
REVIEW_INTERVALS = [1, 3, 7, 14, 30, 60, 120]
REVIEW_LABELS = ["نسيت", "صعب", "تذكّرت", "سهل"]
PERSONAS = {
    "teacher":    "أنت الأستاذ: اشرح وبسّط وصمم اختبارات بأسلوب تربوي واضح.",
    "researcher": "أنت الباحث الأكاديمي: دقيق، تستشهد، وتقدم مراجع.",
    "writer":     "أنت الكاتب الإبداعي: أسلوبك جذاب وسردي.",
    "lawyer":     "أنت المحامي: صياغة قانونية دقيقة ومحكمة.",
    "analyst":    "أنت المحلل: أرقام، بيانات، وتقارير منظمة.",
    "programmer": "أنت المبرمج: كود نظيف وتوثيق تقني.",
}
VOICES = {"عربي أنثوي": "ar", "عربي مغربي": "ar-MA",
          "إنجليزي أنثوي": "en", "إنجليزي ذكري": "en"}
BADGES = [
    {"id": "first_doc", "name": "أول مستند",      "icon": "📄"},
    {"id": "docs_5",    "name": "٥ مستندات",      "icon": "📚"},
    {"id": "docs_10",   "name": "١٠ مستندات",     "icon": "🏆"},
    {"id": "words_1k",  "name": "ألف كلمة",        "icon": "✍️"},
    {"id": "words_10k", "name": "١٠ آلاف كلمة",   "icon": "📝"},
    {"id": "streak_3",  "name": "٣ أيام متتالية", "icon": "🔥"},
    {"id": "streak_7",  "name": "٧ أيام متتالية", "icon": "⚡"},
    {"id": "quiz_done", "name": "إكمال اختبار",   "icon": "🧪"},
    {"id": "cards_20",  "name": "٢٠ بطاقة",        "icon": "🃏"},
    {"id": "vocab_20",  "name": "٢٠ كلمة مترجمة", "icon": "🌐"},
]
HIGHLIGHT_COLORS = {"أصفر": "#FEF08A", "أخضر": "#BBF7D0", "أزرق": "#BFDBFE",
                    "وردي": "#FBCFE8", "برتقالي": "#FED7AA", "بنفسجي": "#DDD6FE"}
REWRITE_STYLES = ["احترافي", "بسيط", "أكاديمي", "تسويقي", "قصصي", "مختصر"]
CONTENT_FORMATS = ["ثريد X", "كاروسيل LinkedIn", "منشور Instagram", "نشرة بريدية",
                   "سكربت بودكاست", "خطوط TED", "قصة أطفال", "أسئلة شائعة",
                   "سيرة ذاتية", "خطة درس", "سلسلة منشورات", "سكربت ريل"]
DIFFICULTY_LEVELS = ["سهل", "متوسط", "صعب", "متقدم"]
EVAL_LEVELS = ["متساهل", "متوازن", "صارم"]
CHAT_MODES = ["من المستند فقط", "عام", "مدمج"]
DEFAULT_SETTINGS = {
    "theme": "light", "accent_color": "#4F46E5",
    "default_design": "modern", "default_page_size": "A4",
    "default_font": "Cairo", "default_margin": "عادية",
    "author_name": "", "reading_font_size": 18, "watermark": "",
    "daily_writing_goal": 500, "default_translation_lang": "ar",
    "reader_voice": "عربي أنثوي", "eval_level": "متوازن",
    "dyslexia_mode": False, "high_contrast": False, "gemini_api_key": "",
}


# =====================================================
# UTILS
# =====================================================
def rtl_css():
    return """
    <style>
    html, body, .stApp, .main { direction: rtl; text-align: right;
        font-family: 'Cairo','Tajawal','Segoe UI',sans-serif; }
    .stTextInput input, .stTextArea textarea { direction: rtl; text-align: right; }
    code, pre { direction: ltr; text-align: left; }
    .doc-card { background: #fff; border-radius: 12px; padding: 16px;
        border: 1px solid #E5E7EB; margin-bottom: 10px; transition: box-shadow 0.2s; }
    .doc-card:hover { box-shadow: 0 4px 20px rgba(0,0,0,0.08); }
    .stat-card { background: linear-gradient(135deg,#EEF2FF,#FFFFFF);
        border-radius: 12px; padding: 14px; text-align: center;
        border: 1px solid #E0E7FF; }
    .vocab-card { background: #F9FAFB; border-radius: 10px; padding: 12px;
        border-right: 4px solid #4F46E5; margin-bottom: 10px; }
    </style>
    """

def ar_word_count(text):
    return len(re.findall(r"[\w\u0600-\u06FF]+", text or ""))

def ar_reading_minutes(text):
    return max(1, round(ar_word_count(text) / 220))

def detect_language(text):
    if not text: return "ar"
    ar = len(re.findall(r"[\u0600-\u06FF]", text))
    en = len(re.findall(r"[A-Za-z]", text))
    if ar > en * 1.5: return "ar"
    if en > ar * 1.5: return "en"
    return "mixed"

def make_snippet(text, n=180):
    t = (text or "").replace("\n", " ").strip()
    return t[:n] + ("…" if len(t) > n else "")

def safe_html(text):
    return html.escape(text or "")

def keyword_counts(text, min_count=3):
    words = re.findall(r"[\w\u0600-\u06FF]+", text or "")
    stop = set("""في من على إلى عن هذا هذه ذلك التي الذي الذين كان كانت يكون
    هو هي هم أنا أنت نحن أن إن لا ما لم لكن أو و ثم قد كل بعض غير بين عند
    with from that this which the and for are was were has have had""".split())
    counter = {}
    for w in words:
        if len(w) < 3 or w.lower() in stop: continue
        counter[w] = counter.get(w, 0) + 1
    return sorted([(w, c) for w, c in counter.items() if c > min_count],
                  key=lambda x: -x[1])


# =====================================================
# DATABASE
# =====================================================
DB_PATH = "docucraft.db"

def _conn():
    c = sqlite3.connect(DB_PATH, check_same_thread=False)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys = ON")
    return c

def init_db():
    c = _conn()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY AUTOINCREMENT,
      email TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL, name TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS folders(id INTEGER PRIMARY KEY AUTOINCREMENT,
      user_id INTEGER NOT NULL, name TEXT NOT NULL, color TEXT DEFAULT '#4F46E5',
      parent_id INTEGER, created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS documents(id INTEGER PRIMARY KEY AUTOINCREMENT,
      user_id INTEGER NOT NULL, title TEXT, subtitle TEXT, author TEXT,
      raw_text TEXT, formatted_data TEXT, design TEXT DEFAULT 'modern',
      folder_id INTEGER, tags TEXT, is_favorite INTEGER DEFAULT 0,
      is_trashed INTEGER DEFAULT 0, word_count INTEGER DEFAULT 0,
      reading_minutes INTEGER DEFAULT 0, language TEXT DEFAULT 'ar',
      page_size TEXT DEFAULT 'A4', font TEXT DEFAULT 'Cairo',
      watermark TEXT, cover_enabled INTEGER DEFAULT 1,
      summary_enabled INTEGER DEFAULT 1, index_enabled INTEGER DEFAULT 1,
      concepts_enabled INTEGER DEFAULT 1, header_footer_enabled INTEGER DEFAULT 1,
      closing_enabled INTEGER DEFAULT 1,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP,
      updated_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS notes(id INTEGER PRIMARY KEY AUTOINCREMENT,
      document_id INTEGER NOT NULL, content TEXT NOT NULL,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS highlights(id INTEGER PRIMARY KEY AUTOINCREMENT,
      document_id INTEGER NOT NULL, text TEXT, color TEXT, note TEXT,
      section_ref TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS vocabulary(id INTEGER PRIMARY KEY AUTOINCREMENT,
      user_id INTEGER NOT NULL, word TEXT NOT NULL, translation TEXT,
      definition TEXT, part_of_speech TEXT, example TEXT, synonyms TEXT,
      antonyms TEXT, source_lang TEXT, target_lang TEXT,
      next_review TEXT, interval_idx INTEGER DEFAULT 0, last_review TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS review_schedule(id INTEGER PRIMARY KEY AUTOINCREMENT,
      document_id INTEGER NOT NULL, next_review TEXT, interval_idx INTEGER DEFAULT 0,
      last_review TEXT, history TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS flashcards(id INTEGER PRIMARY KEY AUTOINCREMENT,
      document_id INTEGER NOT NULL, question TEXT, answer TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS badges_earned(id INTEGER PRIMARY KEY AUTOINCREMENT,
      user_id INTEGER NOT NULL, badge_id TEXT NOT NULL,
      earned_at TEXT DEFAULT CURRENT_TIMESTAMP, UNIQUE(user_id, badge_id));
    CREATE TABLE IF NOT EXISTS settings(user_id INTEGER PRIMARY KEY, data TEXT);
    CREATE TABLE IF NOT EXISTS chat_history(id INTEGER PRIMARY KEY AUTOINCREMENT,
      document_id INTEGER NOT NULL, role TEXT, content TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS writing_log(id INTEGER PRIMARY KEY AUTOINCREMENT,
      user_id INTEGER, date TEXT, words_written INTEGER DEFAULT 0,
      UNIQUE(user_id, date));
    """)
    c.commit(); c.close()

def _hash(pw): return hashlib.sha256(pw.encode()).hexdigest()

def create_user(email, password, name=""):
    c = _conn()
    try:
        c.execute("INSERT INTO users(email,password_hash,name) VALUES(?,?,?)",
                  (email, _hash(password), name))
        c.commit()
        uid = c.execute("SELECT id FROM users WHERE email=?", (email,)).fetchone()["id"]
        c.execute("INSERT OR IGNORE INTO settings(user_id,data) VALUES(?,?)",
                  (uid, json.dumps(DEFAULT_SETTINGS)))
        c.commit()
        return uid
    except sqlite3.IntegrityError:
        return None
    finally:
        c.close()

def authenticate(email, password):
    c = _conn()
    row = c.execute("SELECT id,name FROM users WHERE email=? AND password_hash=?",
                    (email, _hash(password))).fetchone()
    c.close()
    return dict(row) if row else None

def get_user(uid):
    c = _conn()
    row = c.execute("SELECT id,email,name FROM users WHERE id=?", (uid,)).fetchone()
    c.close()
    return dict(row) if row else None

def get_settings(uid):
    c = _conn()
    row = c.execute("SELECT data FROM settings WHERE user_id=?", (uid,)).fetchone()
    c.close()
    if not row: return DEFAULT_SETTINGS.copy()
    d = DEFAULT_SETTINGS.copy()
    d.update(json.loads(row["data"] or "{}"))
    return d

def save_settings(uid, data):
    c = _conn()
    c.execute("INSERT INTO settings(user_id,data) VALUES(?,?) "
              "ON CONFLICT(user_id) DO UPDATE SET data=excluded.data",
              (uid, json.dumps(data, ensure_ascii=False)))
    c.commit(); c.close()

def list_folders(uid):
    c = _conn()
    rows = c.execute("SELECT * FROM folders WHERE user_id=? ORDER BY name", (uid,)).fetchall()
    c.close()
    return [dict(r) for r in rows]

def create_folder(uid, name, color="#4F46E5", parent_id=None):
    c = _conn()
    c.execute("INSERT INTO folders(user_id,name,color,parent_id) VALUES(?,?,?,?)",
              (uid, name, color, parent_id))
    c.commit(); c.close()

def delete_folder(fid, uid):
    c = _conn()
    c.execute("UPDATE documents SET folder_id=NULL WHERE folder_id=? AND user_id=?", (fid, uid))
    c.execute("DELETE FROM folders WHERE id=? AND user_id=?", (fid, uid))
    c.commit(); c.close()

def create_document(uid, title, raw_text, formatted_data=None, design="modern",
                    author="", subtitle="", language="ar", word_count=0,
                    reading_minutes=0, folder_id=None, tags=""):
    c = _conn()
    c.execute("""INSERT INTO documents
      (user_id,title,subtitle,author,raw_text,formatted_data,design,
       folder_id,tags,language,word_count,reading_minutes)
      VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
      (uid, title, subtitle, author, raw_text,
       json.dumps(formatted_data, ensure_ascii=False) if formatted_data else None,
       design, folder_id, tags, language, word_count, reading_minutes))
    c.commit()
    did = c.execute("SELECT last_insert_rowid() as i").fetchone()["i"]
    c.close()
    _add_writing(uid, word_count)
    return did

def get_document(did, uid):
    c = _conn()
    row = c.execute("SELECT * FROM documents WHERE id=? AND user_id=?",
                    (did, uid)).fetchone()
    c.close()
    if not row: return None
    d = dict(row)
    if d.get("formatted_data"):
        try: d["formatted_data"] = json.loads(d["formatted_data"])
        except: d["formatted_data"] = None
    return d

def update_document(did, uid, **fields):
    if not fields: return
    fields["updated_at"] = datetime.now().isoformat()
    if "formatted_data" in fields and isinstance(fields["formatted_data"], dict):
        fields["formatted_data"] = json.dumps(fields["formatted_data"], ensure_ascii=False)
    set_clause = ", ".join(f"{k}=?" for k in fields)
    vals = list(fields.values()) + [did, uid]
    c = _conn()
    c.execute(f"UPDATE documents SET {set_clause} WHERE id=? AND user_id=?", vals)
    c.commit(); c.close()

def list_documents(uid, folder_id=None, only_trash=False, only_fav=False,
                   no_folder=False, search="", tag=None):
    c = _conn()
    q = "SELECT * FROM documents WHERE user_id=?"; args = [uid]
    q += " AND is_trashed=1" if only_trash else " AND is_trashed=0"
    if only_fav: q += " AND is_favorite=1"
    if no_folder: q += " AND folder_id IS NULL"
    if folder_id is not None: q += " AND folder_id=?"; args.append(folder_id)
    if tag: q += " AND tags LIKE ?"; args.append(f"%{tag}%")
    if search:
        q += " AND (title LIKE ? OR subtitle LIKE ? OR tags LIKE ? OR substr(raw_text,1,4000) LIKE ?)"
        like = f"%{search}%"
        args += [like, like, like, like]
    q += " ORDER BY updated_at DESC"
    rows = c.execute(q, args).fetchall()
    c.close()
    out = []
    for r in rows:
        d = dict(r)
        d["snippet"] = make_snippet(d.get("raw_text", ""))
        out.append(d)
    return out

def trash_document(did, uid, trash=True):
    update_document(did, uid, is_trashed=1 if trash else 0)

def delete_document_permanent(did, uid):
    c = _conn()
    c.execute("DELETE FROM documents WHERE id=? AND user_id=?", (did, uid))
    for tbl in ["notes", "highlights", "flashcards", "review_schedule", "chat_history"]:
        c.execute(f"DELETE FROM {tbl} WHERE document_id=?", (did,))
    c.commit(); c.close()

def empty_trash(uid):
    c = _conn()
    rows = c.execute("SELECT id FROM documents WHERE user_id=? AND is_trashed=1", (uid,)).fetchall()
    for r in rows:
        delete_document_permanent(r["id"], uid)
    c.close()

def toggle_favorite(did, uid):
    d = get_document(did, uid)
    if d: update_document(did, uid, is_favorite=0 if d["is_favorite"] else 1)

def library_stats(uid):
    c = _conn()
    docs = c.execute("SELECT COUNT(*) c FROM documents WHERE user_id=? AND is_trashed=0",
                     (uid,)).fetchone()["c"]
    words = c.execute("SELECT COALESCE(SUM(word_count),0) w FROM documents "
                      "WHERE user_id=? AND is_trashed=0", (uid,)).fetchone()["w"]
    folders = c.execute("SELECT COUNT(*) c FROM folders WHERE user_id=?", (uid,)).fetchone()["c"]
    c.close()
    return {"docs": docs, "words": words, "folders": folders}

def add_note(did, content):
    c = _conn(); c.execute("INSERT INTO notes(document_id,content) VALUES(?,?)",
                           (did, content)); c.commit(); c.close()

def list_notes(did):
    c = _conn()
    rows = c.execute("SELECT * FROM notes WHERE document_id=? ORDER BY id DESC", (did,)).fetchall()
    c.close()
    return [dict(r) for r in rows]

def delete_note(nid):
    c = _conn(); c.execute("DELETE FROM notes WHERE id=?", (nid,)); c.commit(); c.close()

def add_highlight(did, text, color, note="", section_ref=""):
    c = _conn()
    c.execute("INSERT INTO highlights(document_id,text,color,note,section_ref) VALUES(?,?,?,?,?)",
              (did, text, color, note, section_ref))
    c.commit(); c.close()

def list_highlights(did):
    c = _conn()
    rows = c.execute("SELECT * FROM highlights WHERE document_id=? ORDER BY id DESC", (did,)).fetchall()
    c.close()
    return [dict(r) for r in rows]

def delete_highlight(hid):
    c = _conn(); c.execute("DELETE FROM highlights WHERE id=?", (hid,)); c.commit(); c.close()

def add_vocabulary(uid, word, translation, definition="", pos="", example="",
                   synonyms="", antonyms="", source_lang="", target_lang="ar"):
    c = _conn()
    exists = c.execute("SELECT id FROM vocabulary WHERE user_id=? AND word=?",
                       (uid, word)).fetchone()
    if exists:
        c.close(); return exists["id"]
    next_rev = (date.today() + timedelta(days=1)).isoformat()
    c.execute("""INSERT INTO vocabulary
      (user_id,word,translation,definition,part_of_speech,example,
       synonyms,antonyms,source_lang,target_lang,next_review)
      VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
      (uid, word, translation, definition, pos, example, synonyms,
       antonyms, source_lang, target_lang, next_rev))
    c.commit()
    vid = c.execute("SELECT last_insert_rowid() as i").fetchone()["i"]
    c.close()
    check_and_award_badges(uid)
    return vid

def list_vocabulary(uid, due_only=False):
    c = _conn()
    if due_only:
        rows = c.execute("SELECT * FROM vocabulary WHERE user_id=? "
                         "AND (next_review IS NULL OR next_review<=?) ORDER BY next_review",
                         (uid, date.today().isoformat())).fetchall()
    else:
        rows = c.execute("SELECT * FROM vocabulary WHERE user_id=? ORDER BY id DESC",
                         (uid,)).fetchall()
    c.close()
    return [dict(r) for r in rows]

def review_vocab(vid, rating):
    c = _conn()
    row = c.execute("SELECT * FROM vocabulary WHERE id=?", (vid,)).fetchone()
    if not row: c.close(); return
    idx = row["interval_idx"]
    if rating == 0: idx = 0
    elif rating == 1: idx = max(0, idx - 1)
    elif rating == 2: idx = min(len(REVIEW_INTERVALS) - 1, idx + 1)
    else: idx = min(len(REVIEW_INTERVALS) - 1, idx + 2)
    next_rev = (date.today() + timedelta(days=REVIEW_INTERVALS[idx])).isoformat()
    c.execute("UPDATE vocabulary SET interval_idx=?, next_review=?, last_review=? WHERE id=?",
              (idx, next_rev, date.today().isoformat(), vid))
    c.commit(); c.close()

def add_to_review(did):
    c = _conn()
    exists = c.execute("SELECT id FROM review_schedule WHERE document_id=?", (did,)).fetchone()
    if exists: c.close(); return
    next_rev = (date.today() + timedelta(days=1)).isoformat()
    c.execute("INSERT INTO review_schedule(document_id,next_review,history) VALUES(?,?,?)",
              (did, next_rev, "[]"))
    c.commit(); c.close()

def list_review_schedule(uid):
    c = _conn()
    rows = c.execute("""SELECT r.*, d.title, d.word_count FROM review_schedule r
                        JOIN documents d ON d.id=r.document_id
                        WHERE d.user_id=? AND d.is_trashed=0
                        ORDER BY r.next_review""", (uid,)).fetchall()
    c.close()
    return [dict(r) for r in rows]

def review_document(rs_id, rating):
    c = _conn()
    row = c.execute("SELECT * FROM review_schedule WHERE id=?", (rs_id,)).fetchone()
    if not row: c.close(); return
    idx = row["interval_idx"]
    if rating == 0: idx = 0
    elif rating == 1: idx = max(0, idx - 1)
    elif rating == 2: idx = min(len(REVIEW_INTERVALS) - 1, idx + 1)
    else: idx = min(len(REVIEW_INTERVALS) - 1, idx + 2)
    next_rev = (date.today() + timedelta(days=REVIEW_INTERVALS[idx])).isoformat()
    try: hist = json.loads(row["history"] or "[]")
    except: hist = []
    hist.append({"date": date.today().isoformat(), "rating": rating})
    c.execute("UPDATE review_schedule SET interval_idx=?, next_review=?, last_review=?, history=? WHERE id=?",
              (idx, next_rev, date.today().isoformat(), json.dumps(hist), rs_id))
    c.commit(); c.close()

def add_flashcard(did, q, a):
    c = _conn()
    c.execute("INSERT INTO flashcards(document_id,question,answer) VALUES(?,?,?)", (did, q, a))
    c.commit(); c.close()

def list_flashcards(did):
    c = _conn()
    rows = c.execute("SELECT * FROM flashcards WHERE document_id=? ORDER BY id", (did,)).fetchall()
    c.close()
    return [dict(r) for r in rows]

def delete_flashcard(fid):
    c = _conn(); c.execute("DELETE FROM flashcards WHERE id=?", (fid,)); c.commit(); c.close()

def add_chat_msg(did, role, content):
    c = _conn()
    c.execute("INSERT INTO chat_history(document_id,role,content) VALUES(?,?,?)",
              (did, role, content))
    c.commit(); c.close()

def get_chat_history(did):
    c = _conn()
    rows = c.execute("SELECT role,content,created_at FROM chat_history "
                     "WHERE document_id=? ORDER BY id", (did,)).fetchall()
    c.close()
    return [dict(r) for r in rows]

def _add_writing(uid, n):
    if n <= 0: return
    c = _conn()
    today = date.today().isoformat()
    c.execute("""INSERT INTO writing_log(user_id,date,words_written) VALUES(?,?,?)
                 ON CONFLICT(user_id,date) DO UPDATE SET
                 words_written=words_written+excluded.words_written""",
              (uid, today, n))
    c.commit(); c.close()

def get_today_words(uid):
    c = _conn()
    row = c.execute("SELECT words_written FROM writing_log WHERE user_id=? AND date=?",
                    (uid, date.today().isoformat())).fetchone()
    c.close()
    return row["words_written"] if row else 0

def get_streak(uid):
    c = _conn()
    rows = c.execute("SELECT date FROM writing_log WHERE user_id=? AND words_written>0 "
                     "ORDER BY date DESC", (uid,)).fetchall()
    c.close()
    if not rows: return 0
    streak = 0; check = date.today()
    for r in rows:
        if r["date"] == check.isoformat():
            streak += 1; check -= timedelta(days=1)
        elif r["date"] < check.isoformat():
            break
    return streak

def _award(uid, badge_id):
    c = _conn()
    c.execute("INSERT OR IGNORE INTO badges_earned(user_id,badge_id) VALUES(?,?)",
              (uid, badge_id))
    c.commit(); c.close()

def list_badges(uid):
    c = _conn()
    rows = c.execute("SELECT badge_id, earned_at FROM badges_earned WHERE user_id=?",
                     (uid,)).fetchall()
    c.close()
    return {r["badge_id"]: r["earned_at"] for r in rows}

def check_and_award_badges(uid):
    stats = library_stats(uid)
    words = stats["words"]; streak = get_streak(uid)
    if stats["docs"] >= 1:  _award(uid, "first_doc")
    if stats["docs"] >= 5:  _award(uid, "docs_5")
    if stats["docs"] >= 10: _award(uid, "docs_10")
    if words >= 1000:       _award(uid, "words_1k")
    if words >= 10000:      _award(uid, "words_10k")
    if streak >= 3:         _award(uid, "streak_3")
    if streak >= 7:         _award(uid, "streak_7")
    c = _conn()
    vc = c.execute("SELECT COUNT(*) c FROM vocabulary WHERE user_id=?", (uid,)).fetchone()["c"]
    c.close()
    if vc >= 20: _award(uid, "vocab_20")


# =====================================================
# AI ENGINE
# =====================================================
_model = None
_current_key = None

def ai_configure(api_key):
    global _model, _current_key
    if not api_key or not _HAS_GEMINI:
        _model = None; _current_key = None; return False
    if api_key == _current_key and _model: return True
    try:
        genai.configure(api_key=api_key)
        _model = genai.GenerativeModel("gemini-1.5-flash")
        _current_key = api_key
        return True
    except Exception:
        _model = None; _current_key = None
        return False

def ai_ready(): return _model is not None

def _call(prompt, system="", json_mode=False):
    if not _model:
        return "⚠️ محرك الذكاء غير مفعّل. أضف مفتاح Gemini من الإعدادات."
    full = (system + "\n\n" + prompt) if system else prompt
    try:
        if json_mode:
            full += "\n\nأعد الرد بصيغة JSON صحيحة فقط دون أي نص إضافي."
        resp = _model.generate_content(full)
        return (resp.text or "").strip()
    except Exception as e:
        return f"⚠️ خطأ: {e}"

def _parse_json(txt, fallback=None):
    if not txt: return fallback
    try: return json.loads(txt)
    except: pass
    m = re.search(r"```json\s*(.*?)\s*```", txt, re.S)
    if m:
        try: return json.loads(m.group(1))
        except: pass
    m = re.search(r"\{.*\}", txt, re.S)
    if m:
        try: return json.loads(m.group(0))
        except: pass
    m = re.search(r"\[.*\]", txt, re.S)
    if m:
        try: return json.loads(m.group(0))
        except: pass
    return fallback

def ai_summarize(text, points=6):
    return _call(f"لخّص النص التالي في {points} نقاط قصيرة وواضحة:\n\n{text[:15000]}")

def ai_extract_concepts(text):
    raw = _call(f"استخرج من النص أهم ٨ مفاهيم أساسية مع تعريف مختصر لكل منها. "
                f'أعد JSON: [{{"term":"...","definition":"..."}}]\n\n{text[:12000]}',
                json_mode=True)
    return _parse_json(raw, [])

def ai_simplify(text):
    return _call(f"اشرح النص التالي بأسلوب مبسّط كأنك تشرح لمبتدئ:\n\n{text[:12000]}")

def ai_suggest_titles(text):
    raw = _call(f"اقترح ٣ عناوين جذابة للنص. أعد JSON: [\"ع1\",\"ع2\",\"ع3\"]\n\n{text[:8000]}",
                json_mode=True)
    return _parse_json(raw, [])

def ai_extract_info(text):
    raw = _call(f"""استخرج من النص: الأشخاص، التواريخ، الأرقام، الأماكن، المصطلحات.
أعد JSON: {{"persons":[],"dates":[],"numbers":[],"places":[],"terms":[]}}

{text[:12000]}""", json_mode=True)
    return _parse_json(raw, {"persons":[],"dates":[],"numbers":[],"places":[],"terms":[]})

def ai_study_plan(text, days=7):
    return _call(f"اعمل خطة مذاكرة {days} أيام لهذا المحتوى. لكل يوم: الموضوع + المدة + ماذا تراجع. Markdown واضح:\n\n{text[:12000]}")

def ai_mind_map(text):
    raw = _call(f"""خريطة ذهنية JSON:
{{"root":"العنوان","children":[{{"name":"فرع","children":[{{"name":"فرعي"}}]}}]}}

{text[:10000]}""", json_mode=True)
    return _parse_json(raw, {"root": "الموضوع", "children": []})

def ai_flashcards(text, count=10):
    raw = _call(f'ولّد {count} بطاقة (سؤال/جواب). JSON: [{{"q":"...","a":"..."}}]\n\n{text[:12000]}',
                json_mode=True)
    return _parse_json(raw, [])

def ai_chat(text, question, mode="مدمج", persona="teacher", history=None):
    sys = PERSONAS.get(persona, PERSONAS["teacher"])
    if mode == "من المستند فقط":
        sys += "\nأجب من المستند فقط. لو المعلومة غير موجودة قل: 'هذه المعلومة غير موجودة في المستند'."
    elif mode == "عام":
        sys += "\nأجب من معرفتك العامة."
    else:
        sys += "\nأجب من المستند، وأثرِ بمعلومات خارجية عند الحاجة."
    ctx = ""
    if history:
        for h in history[-6:]:
            ctx += f"\n[{h['role']}]: {h['content'][:400]}"
    prompt = f"""{sys}

=== المستند ===
{text[:12000]}

=== سياق المحادثة ==={ctx if ctx else ' (بداية)'}

=== السؤال ===
{question}"""
    return _call(prompt)

def ai_quiz(text, level="متوسط", count=6):
    raw = _call(f"""ولّد {count} أسئلة اختيار من متعدد بمستوى {level}.
لكل سؤال ٤ خيارات + الإجابة الصحيحة + شرح.
JSON: [{{"q":"...","options":["أ","ب","ج","د"],"answer":0,"explanation":"..."}}]

{text[:12000]}""", json_mode=True)
    return _parse_json(raw, [])

def ai_rewrite(text, style):
    return _call(f"أعد صياغة النص بأسلوب {style} مع الحفاظ على المعنى:\n\n{text[:12000]}")

def ai_repurpose(text, fmt):
    return _call(f"حوّل النص إلى {fmt} جاهز للنشر:\n\n{text[:12000]}")

def ai_translate_word(word, context, target_lang="ar"):
    raw = _call(f"""ترجم "{word}" من السياق إلى {target_lang}.
JSON: {{"translation":"...","part_of_speech":"...","example":"...",
"synonyms":[],"antonyms":[],"alternatives":[]}}

السياق: {context[:600]}""", json_mode=True)
    return _parse_json(raw, {"translation":"—","part_of_speech":"","example":"",
                             "synonyms":[],"antonyms":[],"alternatives":[]})

def ai_translate_full(text, target_lang="en"):
    raw = _call(f"""ترجم النص كامل إلى {target_lang} مع الحفاظ على العناوين.
JSON: {{"title":"...","sections":[{{"heading":"...","content":"..."}}]}}

{text[:15000]}""", json_mode=True)
    return _parse_json(raw, {})

def ai_translate_sentence(sentence, target_lang="ar"):
    return _call(f"ترجم الجملة إلى {target_lang}:\n\n{sentence[:2000]}")

def ai_evaluate_audio(question, answer, level="متوازن"):
    raw = _call(f"""قيّم الإجابة بمستوى "{level}".
JSON: {{"score":85,"missed":["..."],"advice":"..."}}

السؤال: {question}
الإجابة: {answer}""", json_mode=True)
    return _parse_json(raw, {"score":0,"missed":[],"advice":""})

def ai_health_check():
    if not _model: return False, "غير مفعّل"
    try:
        r = _model.generate_content("قل: OK")
        return True, (r.text or "").strip()[:50]
    except Exception as e:
        return False, str(e)


# =====================================================
# CONVERTER
# =====================================================
def analyze_text(text):
    wc = ar_word_count(text)
    lang = detect_language(text)
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    headings_count = sum(1 for l in lines
                        if len(l) < 90 and not l.endswith((".", "؟", "!")))
    doc_type = _guess_type(text)
    return {"word_count": wc, "reading_minutes": ar_reading_minutes(text),
            "language": lang, "doc_type": doc_type,
            "headings_count": headings_count, "is_mixed": lang == "mixed"}

def _guess_type(text):
    t = text.lower()
    if any(k in t for k in ["ملخص", "بحث", "دراسة", "منهجية", "abstract", "research"]): return "بحث"
    if any(k in t for k in ["تقرير", "report", "quarterly", "annual"]): return "تقرير"
    if any(k in t for k in ["قصة", "حكاية", "story", "chapter", "فصل"]): return "قصة"
    if any(k in t for k in ["سيرة ذاتية", "cv", "resume", "خبرة"]): return "سيرة ذاتية"
    if any(k in t for k in ["عقد", "بند", "contract", "agreement"]): return "عقد"
    if any(k in t for k in ["خطاب", "letter", "dear", "المحترم"]): return "خطاب"
    return "مقال"

def _split_sections(text):
    lines = text.splitlines()
    sections = []
    current = {"heading": "المقدمة", "content": ""}
    for raw in lines:
        line = raw.strip()
        if not line:
            if current["content"].strip(): current["content"] += "\n"
            continue
        is_heading = (len(line) < 90 and not line.endswith((".", "؟", "!", ":"))
                      and len(line.split()) <= 12
                      and (line.startswith(("# ", "## ", "### ")) or
                           (line and not line[0].isdigit() and
                            len(current["content"].strip()) > 40)))
        if is_heading:
            if current["content"].strip(): sections.append(current)
            heading = line.lstrip("#").strip()
            current = {"heading": heading, "content": ""}
        else:
            current["content"] += line + "\n"
    if current["content"].strip() or not sections: sections.append(current)
    sections = [s for s in sections if s["content"].strip() or s["heading"]]
    return sections

def _auto_title(text):
    for line in text.splitlines():
        s = line.strip()
        if s: return s[:80]
    return "مستند بدون عنوان"

def build_structure(text, design="modern", title="", author="", use_ai=True):
    analysis = analyze_text(text)
    sections = _split_sections(text)
    summary = ""; concepts = []; titles = []
    if use_ai and ai_ready():
        try: summary = ai_summarize(text, 6)
        except: pass
        try: concepts = ai_extract_concepts(text) or []
        except: pass
        if not title:
            try: titles = ai_suggest_titles(text) or []
            except: pass
    final_title = title or (titles[0] if titles else _auto_title(text))
    return {"title": final_title, "subtitle": analysis["doc_type"],
            "author": author, "design": design, "analysis": analysis,
            "sections": sections, "summary": summary, "concepts": concepts,
            "suggested_titles": titles}


# =====================================================
# EXPORTERS
# =====================================================
def _shape(text):
    if not text or not _HAS_AR: return text or ""
    try: return get_display(arabic_reshaper.reshape(text))
    except: return text

def _find_arabic_font():
    for p in ["/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
              "/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf",
              "C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/tahoma.ttf",
              "/System/Library/Fonts/Arial.ttf"]:
        if os.path.exists(p): return p
    return ""

def _wrap_lines(text, font, size, max_width, canvas_obj):
    if not text: return [""]
    words = text.split(); lines, current = [], ""
    for wd in words:
        test = (current + " " + wd) if current else wd
        if canvas_obj.stringWidth(test, font, size) > max_width and current:
            lines.append(current); current = wd
        else: current = test
    if current: lines.append(current)
    return lines or [""]

def export_pdf(doc, formatted):
    if not _HAS_RL: raise RuntimeError("reportlab غير مثبت")
    buf = io.BytesIO()
    pw, ph = PAGE_SIZES_MM.get(doc.get("page_size", "A4"), (210, 297))
    w, h = pw * mm, ph * mm
    c = canvas.Canvas(buf, pagesize=(w, h))
    font_name = "Helvetica"
    fp = _find_arabic_font()
    if fp:
        try: pdfmetrics.registerFont(TTFont("ArFont", fp)); font_name = "ArFont"
        except: pass
    design = DESIGNS.get(doc.get("design", "modern"), DESIGNS["modern"])
    margin = MARGINS.get("عادية", 25) * mm
    text_w = w - 2 * margin

    # Cover
    if doc.get("cover_enabled", 1):
        c.setFillColor(design["primary"]); c.rect(0, 0, w, h, fill=1, stroke=0)
        c.setFillColor("white"); c.setFont(font_name, 30)
        for i, ln in enumerate(_wrap_lines(_shape(doc.get("title", "")), font_name, 30, text_w, c)):
            c.drawCentredString(w/2, h - 150 - i*38, ln)
        c.setFont(font_name, 15)
        if doc.get("subtitle"): c.drawCentredString(w/2, h - 200, _shape(doc["subtitle"]))
        if doc.get("author"): c.drawCentredString(w/2, 100, _shape(doc["author"]))
        c.showPage(); c.setFillColor("black")

    # Index
    if doc.get("index_enabled", 1) and formatted.get("sections"):
        c.setFont(font_name, 20); c.drawRightString(w - margin, h - 100, _shape("الفهرس"))
        c.setFont(font_name, 12); y = h - 140
        for i, sec in enumerate(formatted["sections"], 1):
            if y < 60: c.showPage(); y = h - 80
            c.drawRightString(w - margin, y, _shape(f"{i}. {sec.get('heading','')}")); y -= 22
        c.showPage()

    # Summary
    if doc.get("summary_enabled", 1) and formatted.get("summary"):
        c.setFont(font_name, 20); c.drawRightString(w - margin, h - 100, _shape("خلاصة سريعة"))
        y = h - 140; c.setFont(font_name, 12)
        for line in formatted["summary"].split("\n"):
            if not line.strip(): y -= 10; continue
            for ln in _wrap_lines(_shape(line), font_name, 12, text_w, c):
                if y < 60: c.showPage(); y = h - 80
                c.drawRightString(w - margin, y, ln); y -= 18
        c.showPage()

    # Concepts
    if doc.get("concepts_enabled", 1) and formatted.get("concepts"):
        c.setFont(font_name, 20); c.drawRightString(w - margin, h - 100, _shape("مفاهيم أساسية"))
        y = h - 140
        for concept in formatted["concepts"]:
            if y < 100: c.showPage(); y = h - 80
            c.setFont(font_name, 13)
            c.drawRightString(w - margin, y, _shape("• " + concept.get("term", ""))); y -= 20
            c.setFont(font_name, 11)
            for ln in _wrap_lines(_shape(concept.get("definition", "")), font_name, 11, text_w - 20, c):
                if y < 60: c.showPage(); y = h - 80
                c.drawRightString(w - margin - 15, y, ln); y -= 15
            y -= 10
        c.showPage()

    # Body
    for sec in formatted.get("sections", []):
        c.setFont(font_name, 18)
        for ln in _wrap_lines(_shape(sec.get("heading", "")), font_name, 18, text_w, c):
            c.drawRightString(w - margin, h - 80, ln)
        y = h - 120; c.setFont(font_name, 12)
        for paragraph in (sec.get("content", "") or "").split("\n"):
            if not paragraph.strip(): y -= 8; continue
            for ln in _wrap_lines(_shape(paragraph), font_name, 12, text_w, c):
                if y < 60:
                    c.showPage(); y = h - 80; c.setFont(font_name, 12)
                c.drawRightString(w - margin, y, ln); y -= 17
        c.showPage()

    # Closing
    if doc.get("closing_enabled", 1):
        c.setFont(font_name, 20); c.drawCentredString(w/2, h/2 + 20, _shape("— النهاية —"))
        if doc.get("author"):
            c.setFont(font_name, 12); c.drawCentredString(w/2, h/2 - 20, _shape(doc["author"]))
        c.showPage()

    c.save()
    return buf.getvalue()

def export_docx(doc, formatted):
    if not _HAS_DOCX: raise RuntimeError("python-docx غير مثبت")
    d = DocxDocument()
    try:
        d.styles["Normal"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    except: pass
    h = d.add_heading(doc.get("title", "مستند"), level=0); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if doc.get("subtitle"):
        p = d.add_paragraph(doc["subtitle"]); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if doc.get("author"):
        p = d.add_paragraph(doc["author"]); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if doc.get("summary_enabled", 1) and formatted.get("summary"):
        d.add_heading("خلاصة سريعة", level=1); d.add_paragraph(formatted["summary"])
    if doc.get("concepts_enabled", 1) and formatted.get("concepts"):
        d.add_heading("مفاهيم أساسية", level=1)
        for c in formatted["concepts"]:
            p = d.add_paragraph(); p.add_run(c.get("term", "") + ": ").bold = True
            p.add_run(c.get("definition", ""))
    for sec in formatted.get("sections", []):
        d.add_heading(sec.get("heading", ""), level=1)
        for para in (sec.get("content", "") or "").split("\n"):
            if para.strip(): d.add_paragraph(para)
    buf = io.BytesIO(); d.save(buf); return buf.getvalue()

def export_pptx(doc, formatted):
    if not _HAS_PPTX: raise RuntimeError("python-pptx غير مثبت")
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = doc.get("title", "")
    if len(slide.placeholders) > 1: slide.placeholders[1].text = doc.get("author", "")
    for sec in formatted.get("sections", []):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = sec.get("heading", "")
        if len(slide.placeholders) > 1:
            slide.placeholders[1].text = (sec.get("content", "") or "")[:700]
    buf = io.BytesIO(); prs.save(buf); return buf.getvalue()

def export_html(doc, formatted):
    design = DESIGNS.get(doc.get("design", "modern"), DESIGNS["modern"])
    title = safe_html(doc.get("title", ""))
    parts = [f"""<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="utf-8">
<title>{title}</title>
<style>
body{{font-family:'{design["font_ar"]}','Cairo',sans-serif;background:{design["bg"]};
color:{design["text"]};max-width:780px;margin:0 auto;padding:40px 24px;line-height:1.95;font-size:17px;}}
.cover{{text-align:center;padding:100px 0 80px;border-bottom:4px solid {design["primary"]};margin-bottom:40px;}}
.cover h1{{color:{design["primary"]};font-size:2.4em;}}
h2{{color:{design["primary"]};border-bottom:2px solid {design["accent"]};padding-bottom:8px;margin-top:48px;}}
.summary{{background:{design["accent"]}11;border-right:4px solid {design["accent"]};padding:16px 20px;border-radius:8px;}}
</style></head><body>
"""]
    if doc.get("cover_enabled", 1):
        parts.append(f'<div class="cover"><h1>{title}</h1>')
        if doc.get("subtitle"): parts.append(f'<p>{safe_html(doc["subtitle"])}</p>')
        if doc.get("author"): parts.append(f'<p>{safe_html(doc["author"])}</p>')
        parts.append('</div>')
    if doc.get("summary_enabled", 1) and formatted.get("summary"):
        parts.append('<div class="summary"><h2>خلاصة سريعة</h2>')
        for line in formatted["summary"].split("\n"):
            if line.strip(): parts.append(f"<p>{safe_html(line)}</p>")
        parts.append('</div>')
    if doc.get("concepts_enabled", 1) and formatted.get("concepts"):
        parts.append("<h2>مفاهيم أساسية</h2>")
        for c in formatted["concepts"]:
            parts.append(f"<p><strong>{safe_html(c.get('term',''))}</strong>: {safe_html(c.get('definition',''))}</p>")
    for sec in formatted.get("sections", []):
        parts.append(f"<h2>{safe_html(sec.get('heading',''))}</h2>")
        for para in (sec.get("content", "") or "").split("\n"):
            if para.strip(): parts.append(f"<p>{safe_html(para)}</p>")
    parts.append("</body></html>")
    return "".join(parts).encode("utf-8")

def export_md(doc, formatted):
    lines = [f"# {doc.get('title', 'مستند')}\n"]
    if doc.get("author"): lines.append(f"*{doc['author']}*\n")
    if formatted.get("summary"):
        lines.append("\n## خلاصة سريعة\n"); lines.append(formatted["summary"])
    if formatted.get("concepts"):
        lines.append("\n## مفاهيم أساسية\n")
        for c in formatted["concepts"]:
            lines.append(f"- **{c.get('term','')}**: {c.get('definition','')}")
    for sec in formatted.get("sections", []):
        lines.append(f"\n## {sec.get('heading','')}\n"); lines.append(sec.get("content", ""))
    return "\n".join(lines).encode("utf-8")

def export_txt(doc, formatted):
    title = doc.get("title", "مستند")
    lines = [title, "=" * len(title), ""]
    if doc.get("author"): lines.append(f"المؤلف: {doc['author']}\n")
    if formatted.get("summary"):
        lines.append("--- خلاصة ---\n"); lines.append(formatted["summary"] + "\n")
    for sec in formatted.get("sections", []):
        lines.append(f"\n=== {sec.get('heading','')} ===\n"); lines.append(sec.get("content", ""))
    return "\n".join(lines).encode("utf-8")


# =====================================================
# TRANSLATOR & TTS
# =====================================================
def translate_word(word, context="", target="ar"):
    return ai_translate_word(word, context, TRANSLATION_LANGUAGES.get(target, target))

def translate_sentence(sentence, target="ar"):
    return ai_translate_sentence(sentence, TRANSLATION_LANGUAGES.get(target, target))

def translate_full(text, target="en"):
    return ai_translate_full(text, TRANSLATION_LANGUAGES.get(target, target))

def tts(text, lang="ar"):
    if not _HAS_TTS or not text: return None
    try:
        buf = io.BytesIO()
        gTTS(text=text[:2000], lang=lang.split("-")[0], slow=False).write_to_fp(buf)
        buf.seek(0); return buf.read()
    except: return None

def stt(audio_bytes, lang="ar"):
    if not _HAS_STT or not audio_bytes: return "[التفريغ الصوتي غير متاح]"
    try:
        r = sr.Recognizer()
        with sr.AudioFile(io.BytesIO(audio_bytes)) as src:
            audio = r.record(src)
        return r.recognize_google(audio, language=lang)
    except Exception as e:
        return f"[خطأ: {e}]"

def tts_for_voice(voice_label, text):
    return tts(text, VOICES.get(voice_label, "ar"))
    
# =====================================================
# DOCUMENT PAGE — 11 TABS
# =====================================================
def render_document_page(uid, did, settings):
    doc = get_document(did, uid)
    if not doc:
        st.error("المستند غير موجود"); return
    if not doc.get("formatted_data"):
        struct = build_structure(doc.get("raw_text", ""),
                                 design=doc.get("design", "modern"),
                                 title=doc.get("title", ""),
                                 author=doc.get("author", ""), use_ai=False)
        update_document(did, uid, formatted_data=struct)
        doc["formatted_data"] = struct

    st.markdown(f"## 📖 {doc['title']}")
    st.caption(f"{doc.get('subtitle', '')} • {doc.get('author', '')} • "
               f"{doc['word_count']} كلمة • {doc['reading_minutes']} دقيقة")
    if st.button("← المكتبة"):
        st.session_state.page = "library"
        st.session_state.open_doc_id = None
        st.rerun()

    tabs = st.tabs(["📖 قراءة", "🎨 تنسيق", "✏️ تعديل", "🧠 ذكي", "💬 مساعد",
                    "🧪 اختبار", "🎙️ صوتي", "📝 ملاحظات", "🔀 تحويل",
                    "📤 تصدير", "📊 تحليلات"])
    with tabs[0]:  _tab_read(uid, did, doc, settings)
    with tabs[1]:  _tab_format(uid, did, doc, settings)
    with tabs[2]:  _tab_edit(uid, did, doc, settings)
    with tabs[3]:  _tab_smart(uid, did, doc, settings)
    with tabs[4]:  _tab_chat(uid, did, doc, settings)
    with tabs[5]:  _tab_quiz(uid, did, doc, settings)
    with tabs[6]:  _tab_audio(uid, did, doc, settings)
    with tabs[7]:  _tab_notes(uid, did, doc, settings)
    with tabs[8]:  _tab_convert(uid, did, doc, settings)
    with tabs[9]:  _tab_export(uid, did, doc, settings)
    with tabs[10]: _tab_analytics(uid, did, doc, settings)


# ---------- TAB 1: قراءة ----------
def _tab_read(uid, did, doc, settings):
    st.markdown("### 📖 قراءة المستند")
    c1, c2, c3 = st.columns(3)
    font_size = c1.slider("حجم الخط", 12, 32,
                          value=int(settings.get("reading_font_size", 18)))
    translate_mode = c2.toggle("🌐 الترجمة باللمس", value=False)
    edit_mode = c3.toggle("✏️ التعديل المباشر", value=False)

    formatted = doc.get("formatted_data") or {}
    design = DESIGNS.get(doc.get("design", "modern"), DESIGNS["modern"])
    sections = formatted.get("sections", [])

    if doc.get("cover_enabled", 1):
        st.markdown(f"""
        <div style="background:linear-gradient(135deg,{design['primary']},{design['accent']});
                    color:#fff;padding:60px 40px;border-radius:16px;text-align:center;
                    margin-bottom:24px;box-shadow:0 6px 24px rgba(0,0,0,0.12)">
            <h1 style="color:#fff;margin:0;font-size:2.4em">{safe_html(doc['title'])}</h1>
            <p style="font-size:18px;opacity:0.9;margin:12px 0">{safe_html(doc.get('subtitle',''))}</p>
            <p style="opacity:0.8;margin:0">{safe_html(doc.get('author',''))} • {date.today().isoformat()}</p>
        </div>
        """, unsafe_allow_html=True)

    if doc.get("summary_enabled", 1) and formatted.get("summary"):
        with st.expander("📌 خلاصة سريعة", expanded=True):
            st.markdown(formatted["summary"])

    if doc.get("index_enabled", 1) and sections:
        with st.expander("📑 الفهرس"):
            for i, s in enumerate(sections, 1):
                st.write(f"{i}. {s.get('heading', '')}")

    if doc.get("concepts_enabled", 1) and formatted.get("concepts"):
        with st.expander("💡 مفاهيم أساسية"):
            for c in formatted["concepts"]:
                st.markdown(f"- **{c.get('term','')}**: {c.get('definition','')}")

    st.markdown("---")
    for i, sec in enumerate(sections):
        st.markdown(f"### {i+1}. {sec.get('heading', '')}")
        if sec.get("image_url"):
            try: st.image(sec["image_url"], use_column_width=True)
            except: pass
        content = sec.get("content", "")

        if edit_mode:
            with st.expander("✏️ تعديل هذا القسم"):
                new_h = st.text_input("العنوان", value=sec.get("heading", ""), key=f"eh_{i}")
                new_c = st.text_area("المحتوى", value=content, key=f"ec_{i}", height=220)
                col_a, col_b, col_c = st.columns(3)
                if col_a.button("💾 حفظ", key=f"save_ec_{i}"):
                    sections[i] = {"heading": new_h, "content": new_c,
                                   "image_url": sec.get("image_url")}
                    formatted["sections"] = sections
                    update_document(did, uid, formatted_data=formatted,
                                    word_count=ar_word_count(new_c),
                                    reading_minutes=ar_reading_minutes(new_c))
                    st.success("✅"); st.rerun()
                if col_b.button("🗑️ حذف القسم", key=f"del_ec_{i}"):
                    sections.pop(i); formatted["sections"] = sections
                    update_document(did, uid, formatted_data=formatted); st.rerun()
                if col_c.button("➕ إدراج قسم بعده", key=f"ins_ec_{i}"):
                    sections.insert(i + 1, {"heading": "قسم جديد", "content": ""})
                    formatted["sections"] = sections
                    update_document(did, uid, formatted_data=formatted); st.rerun()

        safe_content = safe_html(content).replace("\n", "<br>")
        st.markdown(
            f'<div style="font-size:{font_size}px;line-height:1.9;padding:16px;'
            f'background:{design["bg"]}22;border-radius:10px;margin-bottom:18px;'
            f'border-right:4px solid {design["accent"]}44">{safe_content}</div>',
            unsafe_allow_html=True)

        if translate_mode:
            _translate_panel(uid, did, content, sec.get("heading", f"قسم{i+1}"))

    if edit_mode:
        with st.expander("➕ إضافة قسم جديد"):
            new_h = st.text_input("عنوان القسم", key="new_h")
            new_c = st.text_area("المحتوى", key="new_c", height=150)
            if st.button("إضافة القسم"):
                sections.append({"heading": new_h or "قسم جديد", "content": new_c})
                formatted["sections"] = sections
                update_document(did, uid, formatted_data=formatted); st.rerun()

    st.markdown("---")
    st.markdown("### 🌐 ترجمة المستند كامل")
    c1, c2 = st.columns([3, 1])
    target = c1.selectbox("اللغة الهدف", list(TRANSLATION_LANGUAGES.keys()),
                          format_func=lambda x: TRANSLATION_LANGUAGES[x], index=1,
                          key="full_tr_target")
    if c2.button("🌐 ترجم"):
        if not ai_ready():
            st.warning("⚠️ مفتاح Gemini غير مفعّل")
        else:
            with st.spinner("جارٍ الترجمة…"):
                tr = translate_full(doc["raw_text"], target)
            if tr:
                st.success("✅ تمت")
                st.markdown(f"## {tr.get('title', '')}")
                for s in tr.get("sections", []):
                    st.markdown(f"### {s.get('heading', '')}")
                    st.write(s.get("content", ""))

    st.markdown("---")
    st.markdown("### 🖍️ التمييزات")
    with st.expander("➕ إضافة تمييز"):
        h_text = st.text_area("النص المُظلَّل", key="h_text", height=80)
        c1, c2 = st.columns(2)
        h_color = c1.selectbox("اللون", list(HIGHLIGHT_COLORS.keys()), key="h_color")
        h_section = c2.text_input("القسم المرجعي", key="h_sec")
        h_note = st.text_input("ملاحظة", key="h_note")
        if st.button("💾 حفظ"):
            if h_text.strip():
                add_highlight(did, h_text, HIGHLIGHT_COLORS[h_color], h_note, h_section)
                st.rerun()

    hl_list = list_highlights(did)
    for h in hl_list:
        st.markdown(f'<div style="background:{h["color"]};padding:10px 14px;'
                    f'border-radius:8px;margin-bottom:10px;color:#111">'
                    f'<strong>{safe_html(h["text"][:400])}</strong>'
                    + (f'<br><small>📝 {safe_html(h["note"])}</small>' if h['note'] else "")
                    + (f'<br><small>📑 {safe_html(h["section_ref"])}</small>' if h['section_ref'] else "")
                    + '</div>', unsafe_allow_html=True)
        if st.button("🗑️ حذف", key=f"dh_{h['id']}"):
            delete_highlight(h['id']); st.rerun()


def _translate_panel(uid, did, context, section_name):
    with st.expander(f"🌐 ترجمة — {section_name}"):
        kws = keyword_counts(context, 0)[:8]
        if kws:
            st.caption("كلمات شائعة — اضغط للترجمة:")
            cols = st.columns(4)
            for i, (w, c) in enumerate(kws):
                if cols[i % 4].button(w, key=f"qkw_{section_name}_{w}",
                                     use_container_width=True):
                    with st.spinner("…"):
                        tr = translate_word(w, context, "ar")
                    st.session_state[f"tr_res_{section_name}"] = (w, tr); st.rerun()

        if st.session_state.get(f"tr_res_{section_name}"):
            word, tr = st.session_state[f"tr_res_{section_name}"]
            _show_translation_card(uid, word, tr, section_name)

        st.caption("أو اكتب كلمة:")
        custom = st.text_input("ترجم", key=f"cw_{section_name}",
                               placeholder="اكتب…", label_visibility="collapsed")
        if st.button("ترجم", key=f"tcb_{section_name}"):
            if custom.strip():
                with st.spinner("…"):
                    tr = translate_word(custom, context, "ar")
                st.session_state[f"tr_res_{section_name}"] = (custom, tr); st.rerun()


def _show_translation_card(uid, word, tr, section_name):
    st.markdown(f"""
    <div style="background:#fff;border:1px solid #E5E7EB;border-radius:10px;
                padding:14px;margin:10px 0">
        <div style="font-size:20px;font-weight:bold;color:#4F46E5">
            {safe_html(word)} → {safe_html(tr.get('translation','—'))}
        </div>
        {f"<div style='color:#6B7280;margin-top:6px'>النوع: {safe_html(tr['part_of_speech'])}</div>" if tr.get('part_of_speech') else ""}
        {f"<div style='color:#374151;margin-top:6px'>مثال: {safe_html(tr['example'])}</div>" if tr.get('example') else ""}
        {f"<div style='color:#059669;margin-top:4px'>مرادفات: {safe_html(', '.join(tr['synonyms']))}</div>" if tr.get('synonyms') else ""}
        {f"<div style='color:#DC2626;margin-top:4px'>أضداد: {safe_html(', '.join(tr['antonyms']))}</div>" if tr.get('antonyms') else ""}
    </div>
    """, unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    if c1.button("➕ أضف لمفرداتي", key=f"av_{section_name}_{word}"):
        add_vocabulary(uid, word, tr.get("translation", ""),
                       example=tr.get("example", ""), pos=tr.get("part_of_speech", ""),
                       synonyms=", ".join(tr.get("synonyms", [])),
                       antonyms=", ".join(tr.get("antonyms", [])), target_lang="ar")
        st.success("✅ أُضيفت")
    if c2.button("🌐 ترجم الجملة كاملة", key=f"ts_{section_name}_{word}"):
        with st.spinner("…"):
            sentence_tr = translate_sentence(word, "ar")
        st.info(sentence_tr)


# ---------- TAB 2: تنسيق ----------
def _tab_format(uid, did, doc, settings):
    st.markdown("### 🎨 التنسيق وإعداد الصفحة")
    formatted = doc.get("formatted_data") or {}

    c1, c2, c3 = st.columns(3)
    page_size = c1.selectbox("مقاس الصفحة", PAGE_SIZES,
                             index=PAGE_SIZES.index(doc.get("page_size", "A4"))
                             if doc.get("page_size") in PAGE_SIZES else 0)
    orientation = c2.radio("الاتجاه", ["عمودي", "أفقي"], horizontal=True, index=0)
    margin = c3.selectbox("الهوامش", list(MARGINS.keys()), index=1)

    c1, c2 = st.columns(2)
    current_font = doc.get("font", "Cairo")
    font_idx = ALL_FONTS.index(current_font) if current_font in ALL_FONTS else 0
    font = c1.selectbox("الخط", ALL_FONTS, index=font_idx)
    text_color = c2.color_picker("لون النص", "#111111")
    c1, c2 = st.columns(2)
    brand_color = c1.color_picker("لون الهوية", DESIGNS[doc.get("design", "modern")]["primary"])
    bg_color = c2.color_picker("خلفية الصفحة", "#FFFFFF")

    st.markdown(f"""
    <div style="border:1px solid #E5E7EB;border-radius:10px;padding:16px;
                background:{bg_color};color:{text_color};margin-top:12px">
        <div style="font-family:'{font}';font-size:24px;font-weight:bold">
            نموذج معاينة الخط: {font}
        </div>
        <div style="font-family:'{font}';font-size:15px;margin-top:8px">
            هذا نص تجريبي ١٢٣٤٥ ABC 123.
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("**مفاتيح صفحات الملف:**")
    c1, c2, c3 = st.columns(3)
    cover = c1.toggle("غلاف", value=bool(doc.get("cover_enabled", 1)))
    summary_t = c2.toggle("خلاصة", value=bool(doc.get("summary_enabled", 1)))
    index_t = c3.toggle("فهرس", value=bool(doc.get("index_enabled", 1)))
    c1, c2, c3 = st.columns(3)
    concepts_t = c1.toggle("مفاهيم", value=bool(doc.get("concepts_enabled", 1)))
    hf_t = c2.toggle("رأس/تذييل", value=bool(doc.get("header_footer_enabled", 1)))
    closing_t = c3.toggle("صفحة ختامية", value=bool(doc.get("closing_enabled", 1)))

    if st.button("💾 حفظ إعدادات الصفحة", type="primary"):
        update_document(did, uid, page_size=page_size, font=font,
                        cover_enabled=int(cover), summary_enabled=int(summary_t),
                        index_enabled=int(index_t), concepts_enabled=int(concepts_t),
                        header_footer_enabled=int(hf_t), closing_enabled=int(closing_t))
        st.success("✅"); st.rerun()

    st.markdown("---")
    st.markdown("### 🖼️ توليد صور الأقسام")
    sections = formatted.get("sections", [])
    scope = st.radio("النطاق",
                     ["كل الأقسام", "نصف الأقسام", "٣ صور فقط", "قسم محدد"],
                     horizontal=True)
    selected_section = None
    if scope == "قسم محدد" and sections:
        selected_section = st.selectbox("اختر", range(len(sections)),
                                        format_func=lambda i: sections[i].get("heading", f"قسم {i+1}"))

    if st.button("🎨 توليد الصور"):
        targets = []
        if scope == "كل الأقسام": targets = list(range(len(sections)))
        elif scope == "نصف الأقسام": targets = list(range(0, len(sections), 2))
        elif scope == "٣ صور فقط": targets = list(range(min(3, len(sections))))
        elif selected_section is not None: targets = [selected_section]
        progress = st.progress(0.0, text="جارٍ التوليد…")
        for idx, i in enumerate(targets):
            keyword = (sections[i].get("heading", "") or "")[:30].replace(" ", ",")
            sections[i]["image_url"] = f"https://loremflickr.com/800/400/{keyword}"
            progress.progress((idx + 1) / len(targets))
        formatted["sections"] = sections
        update_document(did, uid, formatted_data=formatted)
        st.success(f"✅ {len(targets)} صورة أُضيفت"); st.rerun()

    if sections:
        with st.expander("🖼️ الصور الحالية"):
            for i, s in enumerate(sections):
                if s.get("image_url"):
                    c1, c2 = st.columns([4, 1])
                    c1.markdown(f"**{s.get('heading', '')}** — `{s['image_url'][:60]}`")
                    if c2.button("🗑️", key=f"delimg_{i}"):
                        sections[i].pop("image_url", None)
                        formatted["sections"] = sections
                        update_document(did, uid, formatted_data=formatted); st.rerun()

    st.markdown("---")
    if st.button("➕ أضف المستند لخطة المراجعة"):
        add_to_review(did); st.success("✅ أُضيف للمراجعة")


# ---------- TAB 3: تعديل ----------
def _tab_edit(uid, did, doc, settings):
    st.markdown("### ✏️ تعديل بيانات المستند")
    with st.form("edit_meta_form"):
        title = st.text_input("العنوان", value=doc.get("title", ""))
        subtitle = st.text_input("العنوان الفرعي", value=doc.get("subtitle", ""))
        author = st.text_input("المؤلف", value=doc.get("author", ""))
        tags = st.text_input("الوسوم", value=doc.get("tags", ""))
        current_design = doc.get("design", "modern")
        design = st.selectbox("التصميم", list(DESIGNS.keys()),
                              index=list(DESIGNS.keys()).index(current_design)
                              if current_design in DESIGNS else 1,
                              format_func=lambda k: DESIGNS[k]["name"])
        if st.form_submit_button("💾 حفظ", use_container_width=True):
            update_document(did, uid, title=title, subtitle=subtitle,
                           author=author, tags=tags, design=design)
            st.success("✅"); st.rerun()

    st.markdown("---")
    st.markdown("### 🔄 إعادة صياغة")
    raw = st.text_area("النص الخام", value=doc.get("raw_text", ""), height=280, key="rewrite_raw")
    style = st.selectbox("الأسلوب", REWRITE_STYLES, key="rewrite_style")
    c1, c2 = st.columns(2)
    if c1.button("🔄 إعادة صياغة"):
        with st.spinner("…"):
            st.session_state["rewritten_text"] = ai_rewrite(raw[:12000], style)
    if c2.button("↩️ تراجع"):
        st.session_state.pop("rewritten_text", None); st.rerun()

    if st.session_state.get("rewritten_text"):
        st.text_area("النسخة الجديدة", value=st.session_state["rewritten_text"],
                     height=280, key="rewrite_out")
        if st.button("💾 حفظ وإعادة التنسيق", type="primary"):
            new_raw = st.session_state["rewritten_text"]
            struct = build_structure(new_raw, design=doc.get("design", "modern"),
                                     title=doc.get("title", ""), author=doc.get("author", ""),
                                     use_ai=ai_ready())
            update_document(did, uid, raw_text=new_raw, formatted_data=struct,
                            word_count=struct["analysis"]["word_count"],
                            reading_minutes=struct["analysis"]["reading_minutes"])
            st.session_state.pop("rewritten_text", None)
            st.success("✅"); st.rerun()


# ---------- TAB 4: ذكي ----------
def _tab_smart(uid, did, doc, settings):
    st.markdown("### 🧠 الأدوات الذكية")
    text = (doc.get("raw_text") or "")[:12000]
    if not ai_ready():
        st.warning("⚠️ AI غير مفعّل — افتح الإعدادات")

    with st.expander("📌 ملخص في ٦ نقاط"):
        if st.button("توليد ملخص", key="s_sum"):
            with st.spinner("…"): st.session_state["s_summary"] = ai_summarize(text, 6)
        if st.session_state.get("s_summary"): st.markdown(st.session_state["s_summary"])

    with st.expander("💡 مفاهيم أساسية"):
        if st.button("استخراج", key="s_con"):
            with st.spinner("…"): st.session_state["s_concepts"] = ai_extract_concepts(text)
        for c in st.session_state.get("s_concepts", []):
            st.markdown(f"- **{c.get('term','')}**: {c.get('definition','')}")

    with st.expander("🎓 شرح مبسّط"):
        if st.button("تبسيط", key="s_simp"):
            with st.spinner("…"): st.session_state["s_simplify"] = ai_simplify(text)
        if st.session_state.get("s_simplify"): st.markdown(st.session_state["s_simplify"])

    with st.expander("📝 عناوين مقترحة"):
        if st.button("اقترح", key="s_tit"):
            with st.spinner("…"): st.session_state["s_titles"] = ai_suggest_titles(text)
        for t in st.session_state.get("s_titles", []):
            if st.button(t, key=f"apply_tit_{hash(t)}"):
                update_document(did, uid, title=t)
                st.success(f"✅ {t}"); st.session_state.pop("s_titles", None); st.rerun()

    with st.expander("🔍 استخراج المعلومات"):
        if st.button("استخرج", key="s_ex"):
            with st.spinner("…"): st.session_state["s_info"] = ai_extract_info(text)
        info = st.session_state.get("s_info") or {}
        for k, label in [("persons", "أشخاص"), ("dates", "تواريخ"),
                         ("numbers", "أرقام"), ("places", "أماكن"), ("terms", "مصطلحات")]:
            if info.get(k): st.markdown(f"**{label}:** " + ", ".join(info[k]))

    with st.expander("📅 خطة مذاكرة ٧ أيام"):
        if st.button("توليد خطة", key="s_plan"):
            with st.spinner("…"): st.session_state["s_plan"] = ai_study_plan(text, 7)
        if st.session_state.get("s_plan"): st.markdown(st.session_state["s_plan"])

    st.markdown("---")
    st.markdown("### 🗺️ خريطة ذهنية")
    if st.button("توليد خريطة"):
        with st.spinner("…"): st.session_state["s_mindmap"] = ai_mind_map(text)
    mm = st.session_state.get("s_mindmap")
    if mm:
        st.markdown(f"#### 🌳 {mm.get('root', '')}")
        def _rec(nodes, level=1):
            indent = "　" * level
            for n in nodes:
                st.markdown(f"{indent}• {n.get('name', '')}")
                if n.get("children"): _rec(n["children"], level + 1)
        if mm.get("children"): _rec(mm["children"])

    st.markdown("---")
    st.markdown("### 🃏 بطاقات مراجعة")
    if st.button("توليد ١٠ بطاقات"):
        with st.spinner("…"):
            cards = ai_flashcards(text, 10)
        for c in cards:
            add_flashcard(did, c.get("q", ""), c.get("a", ""))
        st.success(f"✅ {len(cards)} بطاقة"); st.rerun()

    cards = list_flashcards(did)
    if cards:
        st.caption(f"لديك {len(cards)} بطاقة")
        for c in cards[:15]:
            with st.expander(f"❓ {c['question'][:70]}"):
                st.markdown(f"**الإجابة:** {c['answer']}")
                if st.button("🗑️ حذف", key=f"dfl_{c['id']}"):
                    delete_flashcard(c["id"]); st.rerun()


# ---------- TAB 5: DocuChat ----------
def _tab_chat(uid, did, doc, settings):
    st.markdown("### 💬 مساعد DocuChat")
    if not ai_ready():
        st.warning("⚠️ AI غير مفعّل"); return

    c1, c2 = st.columns(2)
    mode = c1.selectbox("الوضع", CHAT_MODES, key="chat_mode")
    persona = c2.selectbox("الشخصية", list(PERSONAS.keys()),
                           format_func=lambda k: {
                               "teacher": "👨‍🏫 الأستاذ", "researcher": "🎓 الباحث",
                               "writer": "✍️ الكاتب", "lawyer": "⚖️ المحامي",
                               "analyst": "📊 المحلل", "programmer": "🧑‍💻 المبرمج",
                           }.get(k, k), key="chat_persona")

    st.caption("أسئلة مقترحة:")
    suggestions = ["ايه أهم ٣ نقاط؟", "لخصلي الفصل التاني",
                   "اشرحلي الفقرة الأولى بأسلوب أبسط", "ولدلي ١٠ أسئلة"]
    cols = st.columns(2)
    for i, s in enumerate(suggestions):
        if cols[i % 2].button(s, key=f"sug_{i}", use_container_width=True):
            st.session_state["chat_prefill"] = s

    history = get_chat_history(did)
    for msg in history:
        with st.chat_message(msg["role"]): st.markdown(msg["content"])

    question = st.chat_input("اسأل عن المستند…")
    if not question and st.session_state.get("chat_prefill"):
        question = st.session_state.pop("chat_prefill")

    if question:
        st.chat_message("user").markdown(question)
        add_chat_msg(did, "user", question)
        with st.spinner("…"):
            answer = ai_chat(doc.get("raw_text", "")[:12000], question, mode, persona, history)
        st.chat_message("assistant").markdown(answer)
        add_chat_msg(did, "assistant", answer)
        st.rerun()

    if history and history[-1]["role"] == "assistant":
        if st.button("📥 أدرج آخر رد في المستند"):
            formatted = doc.get("formatted_data") or {}
            formatted.setdefault("sections", []).append({
                "heading": "من المساعد", "content": history[-1]["content"]})
            update_document(did, uid, formatted_data=formatted)
            st.success("✅ أُدرج"); st.rerun()


# ---------- TAB 6: اختبار ----------
def _tab_quiz(uid, did, doc, settings):
    st.markdown("### 🧪 اختبار MCQ")
    if not ai_ready():
        st.warning("⚠️ AI غير مفعّل"); return

    c1, c2 = st.columns(2)
    level = c1.selectbox("مستوى الصعوبة", DIFFICULTY_LEVELS, index=1)
    if c2.button("🎯 توليد الأسئلة"):
        with st.spinner("…"):
            quiz = ai_quiz((doc.get("raw_text") or "")[:8000], level, 6)
        st.session_state[f"quiz_{did}"] = quiz

    quiz = st.session_state.get(f"quiz_{did}", [])
    if not quiz:
        st.info("اضغط «توليد الأسئلة»"); return

    answers = {}
    for i, q in enumerate(quiz):
        st.markdown(f"**{i+1}. {q.get('q', '')}**")
        opts = q.get("options", [])
        opts_fmt = [f"{chr(65+j)}. {o}" for j, o in enumerate(opts)]
        choice = st.radio("اختر", opts_fmt, key=f"q_{did}_{i}",
                          label_visibility="collapsed", index=None)
        answers[i] = opts_fmt.index(choice) if choice else None

    if st.button("✅ صحح", type="primary"):
        correct = 0
        for i, q in enumerate(quiz):
            if answers.get(i) == q.get("answer", 0):
                correct += 1
                st.success(f"✅ س{i+1}: صحيح — {q.get('explanation', '')}")
            else:
                ans_idx = q.get("answer", 0)
                ans_txt = q.get("options", [""])[ans_idx] if q.get("options") else ""
                st.error(f"❌ س{i+1}: خطأ. الصحيح: **{ans_txt}**. {q.get('explanation', '')}")
        st.metric("النتيجة", f"{correct}/{len(quiz)}")
        _award(uid, "quiz_done")


# ---------- TAB 7: صوتي ----------
def _tab_audio(uid, did, doc, settings):
    st.markdown("### 🎙️ اختبار صوتي")
    if not ai_ready():
        st.warning("⚠️ AI غير مفعّل"); return

    c1, c2 = st.columns(2)
    voice = c1.selectbox("الصوت", list(VOICES.keys()),
                         index=list(VOICES.keys()).index(settings.get("reader_voice", "عربي أنثوي"))
                         if settings.get("reader_voice") in VOICES else 0)
    eval_level = c2.selectbox("مستوى التقييم", EVAL_LEVELS,
                              index=EVAL_LEVELS.index(settings.get("eval_level", "متوازن"))
                              if settings.get("eval_level") in EVAL_LEVELS else 1)

    if st.button("🔊 توليد سؤال صوتي"):
        with st.spinner("…"):
            quiz = ai_quiz((doc.get("raw_text") or "")[:3000], "متوسط", 1)
            if quiz:
                q = quiz[0]
                st.session_state["audio_q"] = q
                audio = tts_for_voice(voice, q.get("q", ""))
                if audio: st.session_state["audio_q_bytes"] = audio

    q = st.session_state.get("audio_q")
    if not q:
        st.info("اضغط «توليد سؤال صوتي»"); return

    st.markdown(f"**❓ السؤال:** {q.get('q', '')}")
    if st.session_state.get("audio_q_bytes"):
        st.audio(st.session_state["audio_q_bytes"], format="audio/mp3")

    st.markdown("**🎤 سجل إجابتك:**")
    audio_recording = None
    try:
        audio_recording = st.audio_input("سجل", key="audio_rec")
    except AttributeError:
        uploaded = st.file_uploader("أو ارفع ملف صوتي", type=["wav", "mp3", "m4a"], key="audio_up")
        if uploaded: audio_recording = uploaded

    if audio_recording is not None:
        if st.button("🎯 قيّم إجابتي", type="primary"):
            with st.spinner("…"):
                audio_bytes = audio_recording.read() if hasattr(audio_recording, "read") else audio_recording
                transcript = stt(audio_bytes, "ar")
            st.markdown(f"**📝 إجابتك:** {transcript}")
            with st.spinner("…"):
                result = ai_evaluate_audio(q.get("q", ""), transcript, eval_level)
            st.metric("الدرجة", f"{result.get('score', 0)}/100")
            if result.get("missed"): st.warning("**نقاط فاتت:** " + " • ".join(result["missed"]))
            if result.get("advice"): st.info(f"💡 {result['advice']}")
            if q.get("options") and q.get("answer") is not None:
                ans_idx = q.get("answer", 0)
                if 0 <= ans_idx < len(q["options"]):
                    st.success(f"**الإجابة النموذجية:** {q['options'][ans_idx]}")

    if st.button("🔄 سؤال جديد"):
        st.session_state.pop("audio_q", None); st.session_state.pop("audio_q_bytes", None); st.rerun()


# ---------- TAB 8: ملاحظات ----------
def _tab_notes(uid, did, doc, settings):
    st.markdown("### 📝 ملاحظات المستند")
    with st.form("new_note_form", clear_on_submit=True):
        content = st.text_area("ملاحظة جديدة", height=100)
        if st.form_submit_button("➕ إضافة"):
            if content.strip():
                add_note(did, content.strip()); st.rerun()

    notes = list_notes(did)
    if not notes:
        st.info("لا ملاحظات بعد"); return
    for n in notes:
        with st.expander(f"📌 {n['content'][:80]}"):
            st.write(n["content"])
            st.caption(f"🕒 {n['created_at']}")
            if st.button("🗑️ حذف", key=f"dn_{n['id']}"):
                delete_note(n["id"]); st.rerun()


# ---------- TAB 9: تحويل ----------
def _tab_convert(uid, did, doc, settings):
    st.markdown("### 🔀 تحويل المحتوى")
    if not ai_ready():
        st.warning("⚠️ AI غير مفعّل"); return

    fmt = st.selectbox("اختر الصيغة", CONTENT_FORMATS)
    if st.button("🎬 توليد المحتوى", type="primary"):
        with st.spinner("…"):
            result = ai_repurpose((doc.get("raw_text") or "")[:10000], fmt)
        st.session_state[f"repurpose_{fmt}"] = result

    result = st.session_state.get(f"repurpose_{fmt}")
    if result:
        st.markdown(f"### {fmt}"); st.markdown(result)
        st.download_button("⬇️ تحميل", result.encode("utf-8"),
                           file_name=f"{fmt}.txt", mime="text/plain", key=f"dl_rep_{fmt}")


# ---------- TAB 10: تصدير ----------
def _tab_export(uid, did, doc, settings):
    st.markdown("### 📤 تصدير")
    formatted = doc.get("formatted_data") or {}

    c1, c2 = st.columns([3, 1])
    wm = c1.text_input("العلامة المائية", value=doc.get("watermark", ""), key="wm_input")
    if c2.button("💾 حفظ"):
        update_document(did, uid, watermark=wm); st.rerun()

    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    with col1:
        try:
            pdf_data = export_pdf(doc, formatted)
            st.download_button("📄 PDF", pdf_data, file_name=f"{doc['title']}.pdf",
                               mime="application/pdf", use_container_width=True, key="dl_pdf")
        except Exception as e:
            st.button("📄 PDF (خطأ)", disabled=True, use_container_width=True, key="pdf_err")
            st.caption(f"⚠️ {e}")
        try:
            docx_data = export_docx(doc, formatted)
            st.download_button("📝 Word", docx_data, file_name=f"{doc['title']}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               use_container_width=True, key="dl_docx")
        except Exception as e:
            st.caption(f"⚠️ Word: {e}")
    with col2:
        try:
            pptx_data = export_pptx(doc, formatted)
            st.download_button("📊 PowerPoint", pptx_data, file_name=f"{doc['title']}.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                               use_container_width=True, key="dl_pptx")
        except Exception as e:
            st.caption(f"⚠️ PPTX: {e}")
        html_data = export_html(doc, formatted)
        st.download_button("🌐 HTML", html_data, file_name=f"{doc['title']}.html",
                           mime="text/html", use_container_width=True, key="dl_html")
    with col3:
        md_data = export_md(doc, formatted)
        st.download_button("📄 Markdown", md_data, file_name=f"{doc['title']}.md",
                           mime="text/markdown", use_container_width=True, key="dl_md")
        txt_data = export_txt(doc, formatted)
        st.download_button("📃 TXT", txt_data, file_name=f"{doc['title']}.txt",
                           mime="text/plain", use_container_width=True, key="dl_txt")

    st.markdown("---")
    st.markdown("### 🖨️ معاينة وطباعة")
    b64 = base64.b64encode(export_html(doc, formatted)).decode()
    st.components.v1.html(f"""
    <div style="border:1px solid #E5E7EB;border-radius:8px;overflow:hidden">
        <iframe id="pv" src="data:text/html;base64,{b64}"
                style="width:100%;height:520px;border:none"></iframe>
    </div>
    <button onclick="document.getElementById('pv').contentWindow.print()"
            style="margin-top:12px;padding:10px 20px;background:#4F46E5;
                   color:white;border:none;border-radius:8px;cursor:pointer;
                   font-family:'Cairo',sans-serif;font-weight:600">🖨️ طباعة</button>
    """, height=600)


# ---------- TAB 11: تحليلات ----------
def _tab_analytics(uid, did, doc, settings):
    st.markdown("### 📊 تحليلات")
    formatted = doc.get("formatted_data") or {}
    sections = formatted.get("sections", [])
    hl = list_highlights(did)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("كلمات", doc["word_count"])
    c2.metric("أقسام", len(sections))
    c3.metric("دقائق", doc["reading_minutes"])
    c4.metric("تمييزات", len(hl))

    if sections:
        st.markdown("### 📈 حجم كل قسم")
        chart = {s.get("heading", "؟")[:24]: ar_word_count(s.get("content", ""))
                 for s in sections}
        st.bar_chart(chart)

    st.markdown("---")
    st.markdown("### 🔑 كلمات مفتاحية متكررة")
    min_count = st.selectbox("الحد الأدنى", [3, 5, 10, 20], index=2)
    keywords = keyword_counts(doc.get("raw_text", ""), min_count)

    if not keywords:
        st.info("لا كلمات متكررة بدرجة كافية"); return

    st.caption(f"{len(keywords)} كلمة — اضغط للترجمة")
    cols = st.columns(4)
    for i, (word, cnt) in enumerate(keywords[:40]):
        if cols[i % 4].button(f"{word} ({cnt})", key=f"kw_{did}_{word}",
                              use_container_width=True):
            with st.spinner(f"…"):
                tr = translate_word(word, (doc.get("raw_text") or "")[:2000], "ar")
            st.session_state[f"kw_tr_{did}_{word}"] = tr; st.rerun()

    for word, cnt in keywords[:40]:
        tr = st.session_state.get(f"kw_tr_{did}_{word}")
        if tr:
            st.markdown(f'<div style="background:#EEF2FF;padding:10px 14px;'
                        f'border-radius:8px;margin:8px 0;border-right:4px solid #4F46E5">'
                        f'<strong>{safe_html(word)}</strong> → '
                        f'<strong style="color:#4F46E5">{safe_html(tr.get("translation","—"))}</strong>'
                        f'</div>', unsafe_allow_html=True)
            if st.button(f"➕ أضف «{word}» لمفرداتي", key=f"add_kw_{did}_{word}"):
                add_vocabulary(uid, word, tr.get("translation", ""),
                               example=tr.get("example", ""), pos=tr.get("part_of_speech", ""),
                               target_lang="ar")
                st.success(f"✅ {word}")

    if st.button("➕ أضف أول ١٢ كلمة لمفرداتي"):
        added = 0
        with st.spinner("…"):
            for word, _ in keywords[:12]:
                tr = translate_word(word, (doc.get("raw_text") or "")[:2000], "ar")
                add_vocabulary(uid, word, tr.get("translation", ""),
                               example=tr.get("example", ""), pos=tr.get("part_of_speech", ""),
                               target_lang="ar")
                added += 1
        st.success(f"✅ {added} كلمة")

    st.markdown("---")
    st.markdown("### 📋 تفاصيل")
    details = {
        "اللغة": doc.get("language", "ar"),
        "مقاس الصفحة": doc.get("page_size", "A4"),
        "الخط": doc.get("font", "Cairo"),
        "التصميم": DESIGNS.get(doc.get("design", "modern"), {}).get("name", ""),
        "عدد الحروف": len(doc.get("raw_text", "")),
        "عدد الكلمات": doc.get("word_count", 0),
        "عدد الملاحظات": len(list_notes(did)),
        "عدد البطاقات": len(list_flashcards(did)),
        "عدد التمييزات": len(hl),
        "عدد الأقسام": len(sections),
        "الإنشاء": doc.get("created_at", "")[:16],
        "آخر تعديل": doc.get("updated_at", "")[:16],
    }
    col_a, col_b = st.columns(2)
    items = list(details.items())
    for i, (k, v) in enumerate(items):
        (col_a if i < len(items) // 2 else col_b).text(f"{k}: {v}")
        
# =====================================================
# REVIEW PAGE — المراجعة المتباعدة
# =====================================================
def render_review_page(uid, settings):
    st.markdown("## 🔁 المراجعة")
    streak = get_streak(uid)
    schedule = list_review_schedule(uid)
    vocab = list_vocabulary(uid)
    earned = list_badges(uid)
    due_docs = [r for r in schedule
                if not r.get("next_review") or r["next_review"] <= date.today().isoformat()]
    due_vocab = [v for v in vocab
                 if not v.get("next_review") or v["next_review"] <= date.today().isoformat()]

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f'<div class="stat-card"><h2>🔥 {streak}</h2><p>أيام متتالية</p></div>',
                unsafe_allow_html=True)
    c2.markdown(f'<div class="stat-card"><h2>{len(schedule)}</h2><p>في الخطة</p></div>',
                unsafe_allow_html=True)
    c3.markdown(f'<div class="stat-card"><h2>{len(due_docs) + len(due_vocab)}</h2>'
                f'<p>مستحق اليوم</p></div>', unsafe_allow_html=True)
    c4.markdown(f'<div class="stat-card"><h2>🏅 {len(earned)}</h2><p>شارات</p></div>',
                unsafe_allow_html=True)

    today_words = get_today_words(uid)
    goal = settings.get("daily_writing_goal", 500)
    pct = min(1.0, today_words / max(1, goal))
    st.markdown("### 🎯 هدف الكتابة اليومي")
    st.progress(pct, text=f"{today_words} / {goal} كلمة ({int(pct*100)}%)")
    if today_words >= goal:
        st.success("🎉 أنجزت هدفك!")

    st.divider()
    tabs = st.tabs(["📌 مستحق اليوم", "📅 خطة المراجعة", "🌐 مفرداتي", "🏅 الإنجازات"])
    with tabs[0]: _review_due(uid, due_docs, due_vocab)
    with tabs[1]: _review_plan(schedule)
    with tabs[2]: _review_vocab(vocab)
    with tabs[3]: _review_badges(uid, earned, streak, schedule, vocab)


def _rating_buttons(key_prefix, on_click):
    cols = st.columns(4)
    colors = ["#DC2626", "#F59E0B", "#10B981", "#059669"]
    for i, (label, color) in enumerate(zip(REVIEW_LABELS, colors)):
        with cols[i]:
            if st.button(label, key=f"{key_prefix}_r{i}", use_container_width=True):
                on_click(i)


def _review_due(uid, due_docs, due_vocab):
    if not due_docs and not due_vocab:
        st.info("✨ لا يوجد شيء مستحق اليوم — عمل رائع!")
        return
    if due_docs:
        st.markdown(f"#### 📄 مستندات مستحقة ({len(due_docs)})")
        for r in due_docs:
            with st.container():
                st.markdown('<div class="doc-card">', unsafe_allow_html=True)
                ca, cb = st.columns([3, 2])
                with ca:
                    st.markdown(f"**📖 {r['title']}**")
                    st.caption(f"📝 {r.get('word_count', 0)} كلمة • القادم: {r.get('next_review', 'اليوم')}")
                with cb:
                    _rating_buttons(key_prefix=f"doc_{r['id']}",
                                    on_click=lambda rating, rs_id=r["id"]: (
                                        review_document(rs_id, rating), st.rerun()))
                if st.button("📂 افتح المستند", key=f"open_doc_{r['id']}"):
                    st.session_state.open_doc_id = r["document_id"]
                    st.session_state.page = "doc"; st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)
    if due_vocab:
        st.markdown(f"#### 🌐 كلمات مستحقة ({len(due_vocab)})")
        for v in due_vocab:
            with st.container():
                st.markdown('<div class="vocab-card">', unsafe_allow_html=True)
                ca, cb = st.columns([3, 2])
                with ca:
                    hidden = st.session_state.get(f"v_hidden_{v['id']}", True)
                    if hidden:
                        st.markdown(f"### {v['word']}")
                        if st.button("👁️ إظهار", key=f"show_{v['id']}"):
                            st.session_state[f"v_hidden_{v['id']}"] = False; st.rerun()
                    else:
                        st.markdown(f"### {v['word']} → **{v.get('translation', '—')}**")
                        if v.get("part_of_speech"): st.caption(f"النوع: {v['part_of_speech']}")
                        if v.get("example"): st.caption(f"مثال: {v['example']}")
                with cb:
                    _rating_buttons(key_prefix=f"vocab_{v['id']}",
                                    on_click=lambda rating, vid=v["id"]: (
                                        review_vocab(vid, rating),
                                        st.session_state.update({f"v_hidden_{v['id']}": True}),
                                        st.rerun()))
                st.markdown('</div>', unsafe_allow_html=True)


def _review_plan(schedule):
    st.caption("تقييمك يحدد الموعد القادم: نسيت → من الأول • صعب → خطوة لورا • "
               "تذكّرت → خطوة لقدام • سهل → خطوتين")
    with st.expander("🪜 سلم الفواصل (أيام)"):
        cols = st.columns(len(REVIEW_INTERVALS))
        for i, d in enumerate(REVIEW_INTERVALS):
            cols[i].markdown(f'<div style="text-align:center;background:#EEF2FF;'
                            f'border-radius:8px;padding:8px"><b>{d}</b><br>'
                            f'<small>مرحلة {i+1}</small></div>', unsafe_allow_html=True)
    if not schedule:
        st.info("لا مستندات في الخطة بعد")
        st.caption("من صفحة المستند → «تنسيق» → «أضف لخطة المراجعة»")
        return
    for r in schedule:
        idx = r.get("interval_idx", 0)
        next_rev = r.get("next_review", "—")
        is_due = next_rev <= date.today().isoformat()
        badge = "🔴" if is_due else "🟢"
        with st.expander(f"{badge} **{r['title']}** — القادم: {next_rev} "
                         f"(مرحلة {idx+1}/{len(REVIEW_INTERVALS)})"):
            st.progress((idx + 1) / len(REVIEW_INTERVALS),
                        text=f"التقدم: {idx+1}/{len(REVIEW_INTERVALS)}")
            try: history = json.loads(r.get("history") or "[]")
            except: history = []
            if history:
                st.markdown("**آخر ٣ مراجعات:**")
                for h in history[-3:][::-1]:
                    rating = h.get("rating", 0)
                    emoji = {0: "❌", 1: "⚠️", 2: "✅", 3: "⭐"}.get(rating, "•")
                    st.caption(f"{emoji} {h.get('date', '')} — "
                               f"{REVIEW_LABELS[rating] if 0 <= rating < 4 else '—'}")
            else:
                st.caption("لم تُراجع بعد")
            _rating_buttons(key_prefix=f"plan_{r['id']}",
                            on_click=lambda rating, rs_id=r["id"]: (
                                review_document(rs_id, rating), st.rerun()))


def _review_vocab(vocab):
    st.caption(f"{len(vocab)} كلمة محفوظة")
    if not vocab:
        st.info("لا كلمات بعد")
        st.caption("أضف من صفحة المستند → «قراءة» → الترجمة باللمس")
        return
    c1, c2 = st.columns(2)
    show_tr = c1.toggle("👁️ إظهار كل الترجمات", value=False)
    sort_by = c2.selectbox("ترتيب", ["الأحدث", "المراجعة القادمة", "الأبجدي"], index=0)
    items = vocab
    if sort_by == "المراجعة القادمة":
        items = sorted(vocab, key=lambda v: v.get("next_review") or "9999")
    elif sort_by == "الأبجدي":
        items = sorted(vocab, key=lambda v: (v.get("word") or "").lower())
    for v in items:
        with st.container():
            st.markdown('<div class="vocab-card">', unsafe_allow_html=True)
            col1, col2 = st.columns([3, 2])
            with col1:
                if show_tr:
                    st.markdown(f"### {v['word']} → **{v.get('translation', '—')}**")
                    if v.get("part_of_speech"): st.caption(f"النوع: {v['part_of_speech']}")
                    if v.get("example"): st.caption(f"📝 {v['example']}")
                    if v.get("synonyms"): st.caption(f"✅ مرادفات: {v['synonyms']}")
                    if v.get("antonyms"): st.caption(f"❌ أضداد: {v['antonyms']}")
                else:
                    st.markdown(f"### {v['word']}")
                    if st.button("👁️ إظهار", key=f"vocab_show_{v['id']}"):
                        st.session_state[f"v_vis_{v['id']}"] = True; st.rerun()
                    if st.session_state.get(f"v_vis_{v['id']}"):
                        st.markdown(f"**الترجمة:** {v.get('translation', '—')}")
            with col2:
                idx = v.get("interval_idx", 0)
                st.caption(f"المرحلة {idx+1}/{len(REVIEW_INTERVALS)} • "
                          f"القادم: {v.get('next_review', '—')}")
                _rating_buttons(key_prefix=f"vbtn_{v['id']}",
                                on_click=lambda rating, vid=v["id"]: (
                                    review_vocab(vid, rating), st.rerun()))
            st.markdown('</div>', unsafe_allow_html=True)


def _review_badges(uid, earned, streak, schedule, vocab):
    st.caption(f"حصلت على {len(earned)} من {len(BADGES)} شارة")
    pct = len(earned) / len(BADGES) if BADGES else 0
    st.progress(pct, text=f"{int(pct*100)}%")
    stats = library_stats(uid)
    cards_count = 0
    for d in list_documents(uid):
        cards_count += len(list_flashcards(d["id"]))
    progress_map = {
        "first_doc": (stats["docs"], 1), "docs_5": (stats["docs"], 5),
        "docs_10": (stats["docs"], 10), "words_1k": (stats["words"], 1000),
        "words_10k": (stats["words"], 10000), "streak_3": (streak, 3),
        "streak_7": (streak, 7),
        "quiz_done": (1 if "quiz_done" in earned else 0, 1),
        "cards_20": (cards_count, 20), "vocab_20": (len(vocab), 20),
    }
    st.markdown("---")
    cols = st.columns(5)
    for i, b in enumerate(BADGES):
        is_earned = b["id"] in earned
        cur, tgt = progress_map.get(b["id"], (0, 1))
        col = cols[i % 5]
        if is_earned:
            col.markdown(f'<div style="text-align:center;background:linear-gradient(135deg,'
                        f'#FEF3C7,#FDE68A);border-radius:12px;padding:16px;'
                        f'border:2px solid #F59E0B;height:140px">'
                        f'<div style="font-size:36px">{b["icon"]}</div>'
                        f'<div style="font-weight:bold;margin-top:6px">{b["name"]}</div>'
                        f'<div style="font-size:11px;color:#92400E">✅ مُحققة</div></div>',
                        unsafe_allow_html=True)
        else:
            prog = min(100, int((cur / tgt) * 100)) if tgt > 0 else 0
            col.markdown(f'<div style="text-align:center;background:#F9FAFB;'
                        f'border-radius:12px;padding:16px;border:1px solid #E5E7EB;'
                        f'height:140px;opacity:0.6">'
                        f'<div style="font-size:36px;filter:grayscale(1)">{b["icon"]}</div>'
                        f'<div style="font-weight:bold;margin-top:6px">{b["name"]}</div>'
                        f'<div style="font-size:11px;color:#6B7280">'
                        f'{cur}/{tgt} ({prog}%)</div></div>', unsafe_allow_html=True)
    if st.button("🔄 تحديث الشارات"):
        check_and_award_badges(uid); st.rerun()


# =====================================================
# SETTINGS PAGE
# =====================================================
def render_settings_page(uid, settings):
    st.markdown("## ⚙️ الإعدادات")
    tabs = st.tabs(["🎨 المظهر", "📄 الصفحة الافتراضية", "♿ إمكانية الوصول",
                    "🎙️ الصوت والتقييم", "🤖 محرك AI", "👤 الحساب"])
    with tabs[0]: _set_appearance(uid, settings)
    with tabs[1]: _set_page(uid, settings)
    with tabs[2]: _set_access(uid, settings)
    with tabs[3]: _set_voice(uid, settings)
    with tabs[4]: _set_ai(uid, settings)
    with tabs[5]: _set_account(uid, settings)


def _set_appearance(uid, settings):
    st.markdown("### 🎨 المظهر")
    theme = st.radio("الثيم", ["light", "dark"],
                     format_func=lambda x: "☀️ نهاري" if x == "light" else "🌙 ليلي",
                     horizontal=True,
                     index=0 if settings.get("theme", "light") == "light" else 1)
    accent = st.color_picker("لون الهوية", settings.get("accent_color", "#4F46E5"))
    design_options = list(DESIGNS.keys())
    dd = settings.get("default_design", "modern")
    if dd not in design_options: dd = "modern"
    design = st.selectbox("التصميم الافتراضي", design_options,
                          index=design_options.index(dd),
                          format_func=lambda k: DESIGNS[k]["name"])
    author = st.text_input("اسم المؤلف", value=settings.get("author_name", ""))
    goal = st.number_input("هدف الكتابة اليومي", 50, 10000,
                           value=int(settings.get("daily_writing_goal", 500)), step=50)
    if st.button("💾 حفظ", type="primary", key="sv_app"):
        settings.update({"theme": theme, "accent_color": accent,
                         "default_design": design, "author_name": author,
                         "daily_writing_goal": int(goal)})
        save_settings(uid, settings)
        st.success("✅"); st.rerun()
    st.markdown("---")
    st.markdown(f'<div style="background:{accent};color:white;padding:20px;'
                f'border-radius:10px;text-align:center;font-weight:bold">'
                f'معاينة اللون — {accent}</div>', unsafe_allow_html=True)


def _set_page(uid, settings):
    st.markdown("### 📄 الصفحة الافتراضية")
    ps = st.selectbox("مقاس الصفحة", PAGE_SIZES,
                      index=PAGE_SIZES.index(settings.get("default_page_size", "A4"))
                      if settings.get("default_page_size") in PAGE_SIZES else 0)
    font = st.selectbox("الخط", ALL_FONTS,
                        index=ALL_FONTS.index(settings.get("default_font", "Cairo"))
                        if settings.get("default_font") in ALL_FONTS else 0)
    margin = st.selectbox("الهوامش", list(MARGINS.keys()),
                          index=list(MARGINS.keys()).index(settings.get("default_margin", "عادية"))
                          if settings.get("default_margin") in MARGINS else 1)
    rf = st.slider("حجم خط القراءة", 12, 32, value=int(settings.get("reading_font_size", 18)))
    wm = st.text_input("العلامة المائية الافتراضية", value=settings.get("watermark", ""))
    if st.button("💾 حفظ", type="primary", key="sv_page"):
        settings.update({"default_page_size": ps, "default_font": font,
                         "default_margin": margin, "reading_font_size": rf,
                         "watermark": wm})
        save_settings(uid, settings); st.success("✅")


def _set_access(uid, settings):
    st.markdown("### ♿ إمكانية الوصول")
    dys = st.toggle("وضع الديسليكسيا", value=bool(settings.get("dyslexia_mode", False)),
                    help="خطوط مريحة وتباعد أوسع")
    hc = st.toggle("التباين العالي", value=bool(settings.get("high_contrast", False)))
    if st.button("💾 حفظ", type="primary", key="sv_acc"):
        settings.update({"dyslexia_mode": dys, "high_contrast": hc})
        save_settings(uid, settings); st.success("✅"); st.rerun()


def _set_voice(uid, settings):
    st.markdown("### 🎙️ الصوت والتقييم")
    voices = list(VOICES.keys())
    cv = settings.get("reader_voice", "عربي أنثوي")
    if cv not in voices: cv = voices[0]
    voice = st.selectbox("صوت القارئ", voices, index=voices.index(cv))
    el = st.selectbox("مستوى التقييم", EVAL_LEVELS,
                      index=EVAL_LEVELS.index(settings.get("eval_level", "متوازن"))
                      if settings.get("eval_level") in EVAL_LEVELS else 1)
    if st.button("💾 حفظ", type="primary", key="sv_voice"):
        settings.update({"reader_voice": voice, "eval_level": el})
        save_settings(uid, settings); st.success("✅")


def _set_ai(uid, settings):
    st.markdown("### 🤖 محرك الذكاء الاصطناعي")
    st.markdown(f"**الحالة:** {'✅ مفعّل' if ai_ready() else '⚠️ غير مفعّل'}")
    env_key = os.getenv("GEMINI_API_KEY", "")
    key = st.text_input("Gemini API Key",
                        value=settings.get("gemini_api_key", "") or env_key,
                        type="password",
                        help="مجاني من aistudio.google.com/apikey")
    c1, c2 = st.columns(2)
    if c1.button("💾 حفظ المفتاح", type="primary", key="sv_ai"):
        settings["gemini_api_key"] = key
        save_settings(uid, settings)
        ok = ai_configure(key)
        st.success("✅ المفتاح محفوظ") if ok else st.error("⚠️ فشل التفعيل")
        st.rerun()
    if c2.button("🔍 افحص الاتصال", key="chk_ai"):
        if not ai_ready(): ai_configure(key)
        ok, msg = ai_health_check()
        st.success(f"✅ {msg}") if ok else st.error(f"❌ {msg}")
    st.info("📘 **جيب المفتاح من:** aistudio.google.com/apikey — مجاني وسخي.")


def _set_account(uid, settings):
    st.markdown("### 👤 الحساب")
    user = get_user(uid)
    if not user: st.error("غير موجود"); return
    st.markdown(f"**الاسم:** {user.get('name', '—')}")
    st.markdown(f"**البريد:** {user.get('email', '—')}")
    stats = library_stats(uid)
    c1, c2, c3 = st.columns(3)
    c1.metric("مستندات", stats["docs"])
    c2.metric("كلمات", f"{stats['words']:,}")
    c3.metric("مجلدات", stats["folders"])
    st.markdown("---")
    new_name = st.text_input("تغيير الاسم", value=user.get("name", ""))
    if st.button("💾 تحديث الاسم", key="upd_name"):
        c = _conn()
        c.execute("UPDATE users SET name=? WHERE id=?", (new_name, uid))
        c.commit(); c.close()
        st.session_state.user["name"] = new_name
        st.success("✅"); st.rerun()
    st.markdown("---")
    if st.button("🚪 تسجيل الخروج", type="primary", key="logout_btn"):
        st.session_state.user = None
        st.session_state.open_doc_id = None
        st.session_state.page = "library"
        st.rerun()
    with st.expander("⚠️ منطقة الخطر"):
        if st.button("🗑️ إفراغ المهملات", key="dng1"):
            empty_trash(uid); st.success("✅")
        confirm = st.text_input("اكتب DELETE للتأكيد")
        if st.button("❌ حذف كل المستندات"):
            if confirm == "DELETE":
                c = _conn()
                c.execute("DELETE FROM documents WHERE user_id=?", (uid,))
                c.execute("DELETE FROM folders WHERE user_id=?", (uid,))
                c.execute("DELETE FROM vocabulary WHERE user_id=?", (uid,))
                c.commit(); c.close()
                st.success("✅ تم")
            else:
                st.error("اكتب DELETE")


# =====================================================
# LIBRARY PAGE
# =====================================================
def _relative_time(iso):
    try:
        dt = datetime.fromisoformat(iso)
        delta = (datetime.now() - dt).total_seconds()
        if delta < 60: return "الآن"
        if delta < 3600: return f"قبل {int(delta/60)} دقيقة"
        if delta < 86400: return f"قبل {int(delta/3600)} ساعة"
        return f"قبل {int(delta/86400)} يوم"
    except: return "—"


def render_library_page(uid, settings):
    st.markdown("## 📚 المكتبة")
    stats = library_stats(uid)
    c1, c2, c3 = st.columns(3)
    c1.markdown(f'<div class="stat-card"><h2>{stats["docs"]}</h2><p>مستند</p></div>',
                unsafe_allow_html=True)
    c2.markdown(f'<div class="stat-card"><h2>{stats["words"]:,}</h2><p>كلمة</p></div>',
                unsafe_allow_html=True)
    c3.markdown(f'<div class="stat-card"><h2>{stats["folders"]}</h2><p>مجلد</p></div>',
                unsafe_allow_html=True)
    st.divider()

    with st.expander("📁 المجلدات"):
        folders = list_folders(uid)
        fc1, fc2, fc3 = st.columns([2, 1, 1])
        nf = fc1.text_input("اسم المجلد", key="nf")
        nc = fc2.color_picker("لون", "#4F46E5", key="nfc")
        parent_opts = {f["name"]: f["id"] for f in folders}
        parent_name = fc3.selectbox("داخل", ["— بدون —"] + list(parent_opts.keys()))
        if st.button("➕ إنشاء"):
            if nf.strip():
                pid = None if parent_name == "— بدون —" else parent_opts[parent_name]
                create_folder(uid, nf.strip(), nc, pid); st.rerun()
        if folders:
            st.write("المجلدات:")
            for f in folders:
                cc1, cc2, cc3 = st.columns([4, 1, 1])
                cc1.markdown(f'<span style="color:{f["color"]}">●</span> **{f["name"]}**',
                            unsafe_allow_html=True)
                if cc2.button("📂", key=f"of{f['id']}"):
                    st.session_state.library_folder = f["id"]; st.rerun()
                if cc3.button("🗑️", key=f"df{f['id']}"):
                    delete_folder(f["id"], uid); st.rerun()

    st.markdown("---")
    s1, s2 = st.columns([3, 1])
    search = s1.text_input("🔍 بحث", placeholder="عنوان، وسم، أو نص…")
    filt = s2.selectbox("تصفية", ["الكل", "المفضلة", "بدون مجلد", "المهملات"])
    cur_folder = st.session_state.get("library_folder", None)
    if cur_folder:
        fname = next((f["name"] for f in list_folders(uid) if f["id"] == cur_folder), "")
        st.info(f"📂 داخل: **{fname}**")
        if st.button("إظهار الكل"):
            st.session_state.library_folder = None; st.rerun()

    docs = list_documents(uid, folder_id=cur_folder,
                          only_trash=(filt == "المهملات"),
                          only_fav=(filt == "المفضلة"),
                          no_folder=(filt == "بدون مجلد"), search=search)
    if filt == "المهملات":
        if st.button("🗑️ تفريغ المهملات"):
            empty_trash(uid); st.rerun()

    if not docs:
        st.info("لا يوجد مستندات — ابدأ من «مستند جديد» ✍️")
        return

    for d in docs:
        with st.container():
            st.markdown('<div class="doc-card">', unsafe_allow_html=True)
            ca, cb, cc = st.columns([5, 1, 1])
            with ca:
                star = "⭐ " if d["is_favorite"] else ""
                st.markdown(f"#### {star}{d['title'] or 'بدون عنوان'}")
                st.caption(d["snippet"])
                m1, m2, m3, m4 = st.columns(4)
                m1.caption(f"📝 {d['word_count']}")
                m2.caption(f"⏱️ {d['reading_minutes']}د")
                m3.caption(f"🎨 {d['design']}")
                m4.caption(f"🕒 {_relative_time(d['updated_at'])}")
                if d["tags"]:
                    st.caption("🏷️ " + ", ".join(t.strip() for t in d["tags"].split(",") if t.strip()))
            with cb:
                if st.button("📖", key=f"open_{d['id']}"):
                    st.session_state.open_doc_id = d["id"]
                    st.session_state.page = "doc"; st.rerun()
                if st.button("⭐" if not d["is_favorite"] else "☆", key=f"fav_{d['id']}"):
                    toggle_favorite(d["id"], uid); st.rerun()
            with cc:
                folders = list_folders(uid)
                mv = {"— بدون —": None}
                for f in folders: mv[f["name"]] = f["id"]
                sel = st.selectbox("📂", list(mv.keys()), key=f"mv_{d['id']}",
                                   label_visibility="collapsed")
                if st.button("نقل", key=f"mvb_{d['id']}"):
                    update_document(d["id"], uid, folder_id=mv[sel]); st.rerun()
                if not d["is_trashed"]:
                    if st.button("🗑️", key=f"tr_{d['id']}"):
                        trash_document(d["id"], uid, True); st.rerun()
                else:
                    if st.button("↩️", key=f"rs_{d['id']}"):
                        trash_document(d["id"], uid, False); st.rerun()
                    if st.button("❌", key=f"dp_{d['id']}"):
                        delete_document_permanent(d["id"], uid); st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)


# =====================================================
# NEW DOCUMENT PAGE (المحوّل)
# =====================================================
def render_new_document_page(uid, settings):
    st.markdown("## ✨ مستند جديد — من نص خام لملف احترافي")
    up = st.file_uploader("📎 استيراد ملف نصي (.txt / .md)", type=["txt", "md"])
    if up is not None:
        try:
            st.session_state["raw_text_input"] = up.read().decode("utf-8", errors="ignore")
            if not st.session_state.get("doc_title_input"):
                st.session_state["doc_title_input"] = up.name.rsplit(".", 1)[0]
        except Exception as e:
            st.error(f"تعذّر: {e}")

    raw = st.text_area("📝 النص الخام",
                       value=st.session_state.get("raw_text_input", ""),
                       height=320,
                       placeholder="الصق أي نص… مقال، بحث، محاضرة، قصة…")
    st.session_state["raw_text_input"] = raw

    if raw.strip():
        analysis = analyze_text(raw)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("كلمات", analysis["word_count"])
        c2.metric("دقائق", analysis["reading_minutes"])
        c3.metric("أقسام", analysis["headings_count"])
        c4.metric("النوع", analysis["doc_type"])
        if analysis["is_mixed"]:
            st.info("ℹ️ النص مختلط عربي/إنجليزي")
        if len(raw) > 120000:
            st.warning("⚠️ النص كبير — قد يستغرق وقتًا (الحد الموصى به ١٢٠ ألف حرف)")

    st.markdown("### 🎨 اختر التصميم")
    dcols = st.columns(3)
    design_info = [
        ("academic", {"name": "أكاديمي", "desc": "كلاسيكي، هوامش واسعة", "icon": "📜"}),
        ("modern", {"name": "عصري", "desc": "ألوان حيوية", "icon": "🎨"}),
        ("minimal", {"name": "بسيط", "desc": "أبيض وأسود", "icon": "⬜"}),
    ]
    selected = st.session_state.get("selected_design", "modern")
    for i, (key, cfg) in enumerate(design_info):
        with dcols[i]:
            is_sel = selected == key
            border = "3px solid #4F46E5" if is_sel else "1px solid #E5E7EB"
            st.markdown(f'<div style="border:{border};border-radius:12px;padding:16px;'
                        f'text-align:center"><div style="font-size:40px">{cfg["icon"]}</div>'
                        f'<h4>{cfg["name"]}</h4>'
                        f'<p style="color:#6B7280;font-size:12px">{cfg["desc"]}</p></div>',
                        unsafe_allow_html=True)
            if st.button(f"اختيار", key=f"d_{key}"):
                st.session_state["selected_design"] = key; st.rerun()
    design = st.session_state.get("selected_design", "modern")

    st.markdown("### 📋 بيانات المستند")
    m1, m2 = st.columns(2)
    title = m1.text_input("العنوان (اختياري)", key="doc_title_input",
                          value=st.session_state.get("doc_title_input", ""))
    author = m2.text_input("المؤلف", value=settings.get("author_name", ""))
    tags = st.text_input("الوسوم (مفصولة بفواصل)")

    st.markdown("---")
    c1, c2, c3 = st.columns([1, 1, 1])
    if c1.button("🪄 حوّل لملف احترافي", type="primary", use_container_width=True):
        if not raw.strip():
            st.error("الصق نصًا أولًا"); return
        with st.spinner("جارٍ التحليل والبناء…"):
            struct = build_structure(raw, design=design, title=title,
                                     author=author, use_ai=ai_ready())
        wc = struct["analysis"]["word_count"]
        rm = struct["analysis"]["reading_minutes"]
        did = create_document(uid, title=struct["title"], raw_text=raw,
                              formatted_data=struct, design=design, author=author,
                              subtitle=struct["subtitle"],
                              language=struct["analysis"]["language"],
                              word_count=wc, reading_minutes=rm, tags=tags)
        check_and_award_badges(uid)
        st.session_state.open_doc_id = did
        st.session_state.page = "doc"
        st.success("✅"); st.rerun()

    if c2.button("✨ الوضع السحري (3 تصاميم)", use_container_width=True):
        if not raw.strip():
            st.error("الصق نصًا أولًا"); return
        with st.spinner("جاري توليد ٣ تصاميم…"):
            for d in ["academic", "modern", "minimal"]:
                struct = build_structure(raw, design=d, title=title,
                                         author=author, use_ai=ai_ready())
                create_document(uid, title=f"{struct['title']} — {d}",
                                raw_text=raw, formatted_data=struct, design=d,
                                author=author, subtitle=struct["subtitle"],
                                language=struct["analysis"]["language"],
                                word_count=struct["analysis"]["word_count"],
                                reading_minutes=struct["analysis"]["reading_minutes"],
                                tags=tags)
        check_and_award_badges(uid)
        st.success("✅ 3 مستندات"); st.session_state.page = "library"; st.rerun()

    if c3.button("🧪 جرّب مثالًا", use_container_width=True):
        st.session_state["raw_text_input"] = (
            "الذكاء الاصطناعي هو محاكاة للذكاء البشري في الآلات.\n\n"
            "تاريخ النشأة\n\nظهر في الخمسينيات على يد علماء مثل آلان تورينج.\n\n"
            "التطبيقات\n\nالتعلم الآلي ومعالجة اللغة الطبيعية والرؤية الحاسوبية.\n\n"
            "الفوائد\n\nالأتمتة والدقة والسرعة في المهام المتكررة.\n\n"
            "التحديات\n\nالخصوصية والبطالة وتحيّز الخوارزميات.\n\n"
            "المستقبل\n\nيتجه العالم نحو تكامل أكبر بين الإنسان والذكاء الاصطناعي.")
        st.session_state["doc_title_input"] = "الذكاء الاصطناعي: الفرص والتحديات"
        st.rerun()


# =====================================================
# AUTH SCREEN
# =====================================================
def auth_screen():
    st.markdown("""
    <div style="text-align:center;padding:40px 0">
        <div style="font-size:64px">📘</div>
        <h1 style="margin:10px 0">DocuCraft Pro</h1>
        <p style="color:#6B7280;font-size:18px">من نص خام… لملف احترافي</p>
    </div>
    """, unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["تسجيل الدخول", "حساب جديد"])
    with tab1:
        with st.form("login"):
            email = st.text_input("البريد الإلكتروني")
            pw = st.text_input("كلمة المرور", type="password")
            if st.form_submit_button("دخول", use_container_width=True):
                user = authenticate(email, pw)
                if user:
                    st.session_state.user = user; st.rerun()
                else:
                    st.error("بيانات غير صحيحة")
    with tab2:
        with st.form("signup"):
            email2 = st.text_input("البريد", key="e2")
            name2 = st.text_input("الاسم", key="n2")
            pw2 = st.text_input("كلمة المرور", type="password", key="p2")
            if st.form_submit_button("إنشاء حساب", use_container_width=True):
                if len(pw2) < 6:
                    st.error("كلمة المرور قصيرة (٦ حروف على الأقل)")
                else:
                    uid = create_user(email2, pw2, name2)
                    if uid:
                        st.session_state.user = {"id": uid, "name": name2, "email": email2}
                        st.rerun()
                    else:
                        st.error("البريد مستخدم")


# =====================================================
# THEME & NAV
# =====================================================
def _apply_theme(settings):
    theme = settings.get("theme", "light")
    accent = settings.get("accent_color", "#4F46E5")
    hc = settings.get("high_contrast", False)
    dys = settings.get("dyslexia_mode", False)
    css = ""
    if theme == "dark":
        css += """<style>
        .stApp { background: #0F172A; color: #E5E7EB; }
        h1, h2, h3, h4 { color: #F1F5F9; }
        .doc-card { background: #1E293B; border-color: #334155; color: #E5E7EB; }
        .stat-card { background: linear-gradient(135deg,#1E293B,#0F172A);
            border-color: #334155; color: #E5E7EB; }
        .vocab-card { background: #1E293B; border-color: #334155; color: #E5E7EB; }
        </style>"""
    if hc:
        css += "<style>* { font-weight: 500 !important; }</style>"
    if dys:
        css += """<style>
        html, body, .stApp, p, div, span {
            font-family: 'Comic Sans MS', 'Verdana', sans-serif !important;
            letter-spacing: 0.5px !important; line-height: 1.9 !important;
        }
        </style>"""
    css += f"""<style>
    .stButton>button[kind="primary"] {{ background: {accent} !important;
        border-color: {accent} !important; }}
    a {{ color: {accent} !important; }}
    </style>"""
    st.markdown(css, unsafe_allow_html=True)


def _bottom_nav():
    st.markdown("---")
    cols = st.columns(4)
    current = st.session_state.page
    items = [("library", "📚 المكتبة"), ("new", "➕ مستند جديد"),
             ("review", "🔁 مراجعة"), ("settings", "⚙️ الإعدادات")]
    for i, (key, label) in enumerate(items):
        with cols[i]:
            if st.button(label, key=f"bn_{key}", use_container_width=True,
                        type="primary" if current == key else "secondary"):
                st.session_state.page = key
                if key != "doc":
                    st.session_state.open_doc_id = None
                st.rerun()


# =====================================================
# MAIN ENTRY POINT
# =====================================================
st.set_page_config(page_title="DocuCraft Pro", page_icon="📘", layout="wide")
init_db()
st.markdown(rtl_css(), unsafe_allow_html=True)

if "user" not in st.session_state: st.session_state.user = None
if "page" not in st.session_state: st.session_state.page = "library"
if "open_doc_id" not in st.session_state: st.session_state.open_doc_id = None

if not st.session_state.user:
    auth_screen()
    st.stop()

user = st.session_state.user
uid = user["id"]
settings = get_settings(uid)
ai_configure(settings.get("gemini_api_key", "") or os.getenv("GEMINI_API_KEY", ""))
_apply_theme(settings)

with st.sidebar:
    st.markdown(
        f'<div style="text-align:center;padding:12px 0">'
        f'<div style="font-size:32px">📘</div>'
        f'<h3 style="margin:6px 0;color:{settings.get("accent_color","#4F46E5")}">'
        f'DocuCraft Pro</h3>'
        f'<p style="color:#6B7280;font-size:12px;margin:0">من نص خام… لملف احترافي</p></div>',
        unsafe_allow_html=True)
    st.divider()
    st.caption(f"👤 {user.get('name') or user['email']}")

    current = st.session_state.page
    if st.button("📚 المكتبة", use_container_width=True,
                 type="primary" if current == "library" else "secondary"):
        st.session_state.page = "library"; st.session_state.open_doc_id = None; st.rerun()
    if st.button("➕ مستند جديد", use_container_width=True,
                 type="primary" if current == "new" else "secondary"):
        st.session_state.page = "new"; st.rerun()
    if st.button("🔁 المراجعة", use_container_width=True,
                 type="primary" if current == "review" else "secondary"):
        st.session_state.page = "review"; st.rerun()
    if st.button("⚙️ الإعدادات", use_container_width=True,
                 type="primary" if current == "settings" else "secondary"):
        st.session_state.page = "settings"; st.rerun()

    st.divider()
    _stats = library_stats(uid)
    sc1, sc2 = st.columns(2)
    sc1.metric("مستندات", _stats["docs"])
    sc2.metric("كلمات", f"{_stats['words']:,}")
    tw = get_today_words(uid)
    g = settings.get("daily_writing_goal", 500)
    st.progress(min(1.0, tw / max(1, g)), text=f"اليوم: {tw}/{g}")
    st.caption(f"🔥 {get_streak(uid)} يوم متتالي")
    if not ai_ready():
        st.warning("⚠️ AI غير مفعّل")
    st.divider()
    if st.button("🚪 خروج", use_container_width=True):
        st.session_state.user = None
        st.session_state.open_doc_id = None
        st.session_state.page = "library"; st.rerun()

page = st.session_state.page
if page == "library": render_library_page(uid, settings)
elif page == "new": render_new_document_page(uid, settings)
elif page == "doc" and st.session_state.open_doc_id:
    render_document_page(uid, st.session_state.open_doc_id, settings)
elif page == "review": render_review_page(uid, settings)
elif page == "settings": render_settings_page(uid, settings)
else:
    st.session_state.page = "library"; st.rerun()

_bottom_nav()