import os
import io
import json
import re
import streamlit as st
from google import genai
from google.genai import types
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import gdown

st.set_page_config(page_title="ממשק כתיבת תל\"א", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Alef:wght@400;700&display=swap');

    /* החלת גופן אלף והגדלה מתונה וקריאה של כל הטקסטים */
    html, body, .stMarkdown, p, h1, h2, h3, h4, label, input, textarea, button, select, [class*="css"], details, summary {
        font-family: 'Alef', sans-serif !important;
        direction: rtl !important;
        text-align: right !important;
        font-size: 1.05rem !important;
    }

    h1, .main-title {
        font-size: 2.1rem !important;
        font-weight: 700 !important;
    }

    h2, h3, .stSubheader {
        font-size: 1.4rem !important;
        font-weight: 700 !important;
    }

    /* כותרות חלונות המטרות - גודל פונט זהה לתוכן החלונות ובולד */
    [data-testid="stExpander"] details summary p {
        font-size: 1.05rem !important;
        font-weight: 700 !important;
        color: #1a1a1a !important;
    }

    /* דגל יחיד בלבד מימין לכותרת המטרה */
    [data-testid="stExpander"] details summary p::before {
        content: "";
        display: inline-block;
        width: 17px;
        height: 17px;
        margin-left: 8px;
        vertical-align: -2px;
        background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='%23111111'%3E%3Cpath d='M6 3c-.55 0-1 .45-1 1v16c0 .55.45 1 1 1s1-.45 1-1v-6h11.5c.4 0 .75-.24.9-.6.15-.37.07-.79-.2-1.07L16.5 10l2.7-3.33c.27-.28.35-.7.2-1.07-.15-.36-.5-.6-.9-.6H7V4c0-.55-.45-1-1-1z'/%3E%3C/svg%3E");
        background-size: contain;
        background-repeat: no-repeat;
    }

    /* הסתרת הודעת 'Press Enter to apply' */
    [data-testid="InputInstructions"],
    .stTextInput small,
    div[data-testid="InputInstructions"] {
        display: none !important;
    }

    /* חיוויי הצלחה בירוק - טקסט רגיל ללא בולד */
    div[data-testid="stAlert"] {
        background-color: #e8f5e9 !important;
        border: 1px solid #81c784 !important;
        color: #1b5e20 !important;
        border-radius: 8px !important;
    }
    div[data-testid="stAlert"] * {
        color: #1b5e20 !important;
        font-weight: 400 !important;
        font-size: 1.05rem !important;
    }

    /* תיקון האייקונים של המערכת למניעת עיוותים */
    [data-testid="stIcon"],
    [data-testid="stExpanderToggleIcon"],
    [class*="material-symbols"],
    [class*="material-icons"],
    span[class*="Icon"] {
        font-family: 'Material Symbols Rounded', 'Material Icons', sans-serif !important;
        direction: ltr !important;
        text-align: left !important;
    }

    /* הסתרת סרגל ההגדרות בצד כברירת מחדל */
    [data-testid="stSidebar"] {
        display: none;
    }

    /* כפתור סטטוס דרייב צף וקומפקטי בפינה השמאלית העליונה */
    div[data-testid="stPopover"] {
        position: fixed !important;
        top: 12px !important;
        left: 20px !important;
        right: auto !important;
        width: auto !important;
        z-index: 999999 !important;
    }
    div[data-testid="stPopover"] > button {
        width: auto !important;
        min-width: unset !important;
        background-color: #ffffff !important;
        color: #2E7D32 !important;
        border: 1px solid #c8e6c9 !important;
        border-radius: 20px !important;
        padding: 5px 14px !important;
        font-size: 0.9rem !important;
        font-weight: bold !important;
        box-shadow: 0 2px 5px rgba(0,0,0,0.06) !important;
    }
    div[data-testid="stPopover"] > button:hover {
        background-color: #f1f8e9 !important;
        border-color: #2E7D32 !important;
    }
    div[data-testid="stPopoverBody"] {
        direction: rtl !important;
        text-align: right !important;
    }

    /* עיצוב כפתורים כללי */
    .stButton>button, [data-testid="stFormSubmitButton"]>button {
        width: 100%;
        border-radius: 8px;
        font-weight: bold;
        font-family: 'Alef', sans-serif !important;
        direction: rtl !important;
        text-align: center !important;
        font-size: 1.02rem !important;
        white-space: nowrap !important;
    }
    .main-btn>button {
        background-color: #2E7D32 !important;
        color: white !important;
        font-size: 1.15rem !important;
        padding: 10px !important;
    }

    /* מניעת רווחים מיותרים מתגיות המרקרים */
    div[data-testid="element-container"]:has(.marker-regen),
    div[data-testid="element-container"]:has(.marker-edit),
    div[data-testid="element-container"]:has(.marker-del),
    div[data-testid="element-container"]:has(.marker-plus),
    div[data-testid="element-container"]:has(.marker-download),
    div.stElementContainer:has(.marker-regen),
    div.stElementContainer:has(.marker-edit),
    div.stElementContainer:has(.marker-del),
    div.stElementContainer:has(.marker-plus),
    div.stElementContainer:has(.marker-download) {
        margin: 0 !important;
        padding: 0 !important;
        height: 0 !important;
        min-height: 0 !important;
    }

    /* חץ עגול - נסח מחדש */
    div[data-testid="element-container"]:has(.marker-regen) + div[data-testid="element-container"] button p::before,
    div.stElementContainer:has(.marker-regen) + div.stElementContainer button p::before {
        content: "";
        display: inline-block;
        width: 16px;
        height: 16px;
        margin-left: 6px;
        vertical-align: -2px;
        background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='%23111111'%3E%3Cpath d='M12 4V1L8 5l4 4V6c3.31 0 6 2.69 6 6 0 1.01-.25 1.97-.7 2.8l1.46 1.46C19.54 15.03 20 13.57 20 12c0-4.42-3.58-8-8-8zm0 14c-3.31 0-6-2.69-6-6 0-1.01.25-1.97.7-2.8L5.24 7.74C4.46 8.97 4 10.43 4 12c0 4.42 3.58 8 8 8v3l4-4-4-4v3z'/%3E%3C/svg%3E");
        background-size: contain;
        background-repeat: no-repeat;
    }

    /* עפרון - ערוך לפי תיאור */
    form:has(.marker-edit) button p::before,
    div[data-testid="element-container"]:has(.marker-edit) + div[data-testid="element-container"] button p::before,
    div.stElementContainer:has(.marker-edit) + div.stElementContainer button p::before {
        content: "";
        display: inline-block;
        width: 16px;
        height: 16px;
        margin-left: 6px;
        vertical-align: -2px;
        background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='%23111111'%3E%3Cpath d='M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z'/%3E%3C/svg%3E");
        background-size: contain;
        background-repeat: no-repeat;
    }

    /* פח - מחק מטרה / מחק יעד */
    div[data-testid="element-container"]:has(.marker-del) + div[data-testid="element-container"] button p::before,
    div.stElementContainer:has(.marker-del) + div.stElementContainer button p::before {
        content: "";
        display: inline-block;
        width: 16px;
        height: 16px;
        margin-left: 6px;
        vertical-align: -2px;
        background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='%23111111'%3E%3Cpath d='M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z'/%3E%3C/svg%3E");
        background-size: contain;
        background-repeat: no-repeat;
    }

    /* פלוס - הוסף יעד / הוסף מטרה */
    form:has(.marker-plus) button p::before,
    div[data-testid="element-container"]:has(.marker-plus) + div[data-testid="element-container"] button p::before,
    div.stElementContainer:has(.marker-plus) + div.stElementContainer button p::before {
        content: "";
        display: inline-block;
        width: 16px;
        height: 16px;
        margin-left: 6px;
        vertical-align: -2px;
        background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='%23111111'%3E%3Cpath d='M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z'/%3E%3C/svg%3E");
        background-size: contain;
        background-repeat: no-repeat;
    }

    /* חץ הורדה - כפתור הורדת Word */
    .download-btn>button {
        background-color: #2E7D32 !important;
        color: #ffffff !important;
        font-size: 1.15rem !important;
        font-weight: 700 !important;
        padding: 10px !important;
    }
    .download-btn>button:hover {
        background-color: #1B5E20 !important;
        color: #ffffff !important;
    }
    div[data-testid="element-container"]:has(.marker-download) + div[data-testid="element-container"] button p::before,
    div.stElementContainer:has(.marker-download) + div.stElementContainer button p::before {
        content: "";
        display: inline-block;
        width: 18px;
        height: 18px;
        margin-left: 8px;
        vertical-align: -3px;
        background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='%23ffffff'%3E%3Cpath d='M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z'/%3E%3C/svg%3E");
        background-size: contain;
        background-repeat: no-repeat;
    }

    /* תיקון אזור העלאת קבצים */
    [data-testid="stFileUploader"] section {
        direction: ltr !important;
        text-align: left !important;
    }
    [data-testid="stFileUploader"] section button {
        direction: ltr !important;
    }

    /* מתיחה מדויקת של חלון דרכי ההוראה ומרכוז */
    div.teach-box div[data-baseweb="textarea"],
    div.teach-box textarea,
    textarea[aria-label*="דרכי הוראה"] {
        height: 970px !important;
        min-height: 970px !important;
        max-height: 970px !important;
        padding-top: 340px !important;
        padding-bottom: 40px !important;
        line-height: 2 !important;
        box-sizing: border-box !important;
        font-size: 1.02rem !important;
    }
</style>
""", unsafe_allow_html=True)

DRIVE_FOLDER_URL = "https://drive.google.com/drive/folders/1IHat-atuDDzFIfsmKq24aF7WkDjhbx0S?usp=drive_link"
LOCAL_DRIVE_FOLDER = "drive_examples"

@st.cache_data(show_spinner=False)
def sync_and_load_drive_examples():
    if not os.path.exists(LOCAL_DRIVE_FOLDER):
        os.makedirs(LOCAL_DRIVE_FOLDER)
        try:
            gdown.download_folder(DRIVE_FOLDER_URL, output=LOCAL_DRIVE_FOLDER, quiet=True, use_cookies=False)
        except Exception:
            pass

    combined_examples = []
    if os.path.exists(LOCAL_DRIVE_FOLDER):
        for root, _, files in os.walk(LOCAL_DRIVE_FOLDER):
            for f in files:
                if f.endswith('.docx') and not f.startswith('~$'):
                    try:
                        doc = Document(os.path.join(root, f))
                        text_parts = [p.text for p in doc.paragraphs if p.text.strip()]
                        for table in doc.tables:
                            for row in table.rows:
                                text_parts.append(" | ".join([c.text.replace("\n", " ").strip() for c in row.cells]))
                        combined_examples.append(f"=== דוגמת תל\"א מקצועית מתוך ({f}) ===\n" + "\n".join(text_parts))
                    except Exception:
                        pass
    return "\n\n".join(combined_examples)

def safe_parse_json(text_content):
    t = text_content.strip()
    if t.startswith("```"):
        lines = t.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].startswith("```"):
            lines = lines[:-1]
        t = "\n".join(lines).strip()
    
    try:
        data = json.loads(t)
    except Exception:
        # חילוץ מבוסס ביטוי רגולרי אם יש טקסט מסביב ל-JSON
        match_arr = re.search(r'\[\s*\{.*\}\s*\]', t, re.DOTALL)
        if match_arr:
            data = json.loads(match_arr.group(0))
        else:
            match_obj = re.search(r'\{.*\}', t, re.DOTALL)
            if match_obj:
                data = json.loads(match_obj.group(0))
            else:
                raise ValueError("לא אותר מבנה JSON תקין בתשובה")

    # חילוץ רשימת המטרות גם אם הוחזר מילון עטוף
    if isinstance(data, dict):
        for k in ["goals", "matorot", "tala", "items", "data", "goals_list"]:
            if k in data and isinstance(data[k], list):
                return data[k]
        for val in data.values():
            if isinstance(val, list) and len(val) > 0 and isinstance(val[0], dict):
                return val
        if "goal_title" in data:
            return [data]
        return []
    elif isinstance(data, list):
        return data
    return []

with st.spinner("מסנכרן דוגמאות מתיקיית הדרייב..."):
    examples_context = sync_and_load_drive_examples()

with st.popover("מאגר דרייב"):
    st.markdown("**סטטוס חיבור ל-Google Drive:**")
    if examples_context:
        num_docs = examples_context.count("=== דוגמת תל\"א מקצועית")
        st.success(f"מאגר הדוגמאות מחובר ומסונכרן!\n\nנטענו {num_docs} קובצי תל\"א ללמידת המודל.")
    else:
        st.warning("לא אותרו קבצים בתיקייה (וודאי שהשיתוף פתוח לצפייה לכולם).")

    if st.button("רענן מאגר דרייב", key="refresh_drive_btn"):
        try:
            gdown.download_folder(DRIVE_FOLDER_URL, output=LOCAL_DRIVE_FOLDER, quiet=True, use_cookies=False)
        except Exception:
            pass
        st.cache_data.clear()
        st.rerun()

st.title("ממשק חכם לניסוח תל\"א")
st.caption("מערכת לגזירת מטרות על ויעדים אופרטיביים על בסיס תיאור הילד/ה")

with st.sidebar:
    st.header("הגדרות מערכת")
    api_key = st.text_input(
        "מפתח API:", 
        value="AQ.Ab8RN6KO-aLO8c_IGTJZnUl0ju67TTDDblmIPxEQr4LeH0KGAA", 
        type="password"
    )

if 'goals_list' not in st.session_state:
    st.session_state['goals_list'] = []

# טופס ראשי
col1, col2 = st.columns([1, 3])
with col1:
    gender = st.radio("התאמה מגדרית:", ["ילדה (נקבה)", "ילד (זכר)"])
with col2:
    student_name = st.text_input("שם הילד/ה:", value="", placeholder="הזיני את שם הילד/ה")

st.markdown("---")
st.subheader("1. תיאור רמת התפקוד בתחומים הרלוונטיים")
st.markdown("**הזנת מוקדי כוח ומוקדים לחיזוק**")

uploaded_file = st.file_uploader(
    "העלי קובץ Word של מוקדי כוח ומוקדים לחיזוק (יש לוודא שהקובץ סגור במחשב):", 
    type=["docx"]
)

input_text = ""
if uploaded_file is not None:
    try:
        doc_uploaded = Document(uploaded_file)
        extracted = [p.text for p in doc_uploaded.paragraphs if p.text.strip()]
        for table in doc_uploaded.tables:
            for row in table.rows:
                extracted.append(" | ".join([c.text.replace("\n", " ").strip() for c in row.cells]))
        input_text = "\n".join(extracted)
        st.success("הקובץ נטען בהצלחה.")
    except Exception as err:
        st.error("לא ניתן לקרוא את הקובץ. אנא ודאי שמסמך ה-Word סגור במחשבך ולא נעול, ונסי להעלות שוב.")
else:
    input_text = st.text_area("או הדביקי כאן את הטקסט:", height=120, placeholder="הדביקי כאן מוקדי כוח וחיזוק...")

st.markdown('<div class="main-btn">', unsafe_allow_html=True)
if st.button("הפק מטרות ויעדים"):
    active_name = student_name.strip() if student_name.strip() else ("הילדה" if "נקבה" in gender else "הילד")
    if not input_text.strip():
        st.warning("אנא הזיני נתוני תפקוד.")
    else:
        with st.spinner("מנתח דפוסים ומנסח מטרות ויעדים תפקודיים ממוקדי השתתפות..."):
            try:
                env_pronoun = "לסביבתה" if "נקבה" in gender else "לסביבתו"
                
                system_prompt = (
                    "אתה מומחה פדגוגי וקלינאי תקשורת בכיר לניסוח תכניות לימודים אישיות (תל\"א) בגני חינוך מיוחד.\n"
                    + f"עליך לנסח תל\"א מקצועית עבור '{active_name}' ({gender}) מתוך הישענות עמוקה על מאגר הדוגמאות מתיקיית הדרייב.\n\n"
                    + "### מאגר הדוגמאות המקצועיות ללמידה וחיקוי:\n"
                    + str(examples_context) + "\n\n"
                    + "---\n### עקרונות מחייבים לניסוח מטרות ויעדים:\n\n"
                    + "1. **מטרת-על – רחבה, כוללת ותמציתית:**\n"
                    + f"   - שאיפה תפקודית רחבה (למשל: \"{active_name} תביע כוונות תקשורתיות באמצעות משפטים...\", \"{active_name} תשתתף באופן מילולי במשימות חשיבה...\", \"{active_name} תנהל שיחה באופן הדדי...\").\n"
                    + "   - קצרה, ללא פירוט תנאים ספציפיים וללא חיבור שני תחומים ב-ו' החיבור.\n\n"
                    + "2. **היעדים האופרטיביים – ספציפיים לתפקוד יחיד, חדים וללא סרבול (כלל קריטי):**\n"
                    + "   - **כל יעד עוסק בתפקוד אחד ויחיד בלבד!**\n"
                    + f"   - **איסור מוחלט על העמסה וסרבול:** אין לדחוס מספר פעולות, תנאים ורמות תיווך באותו משפט (למשל, אין לכתוב: \"{active_name} תשתמש במשפטים פשוטים להבעת צרכים, רצונות ורעיונות באופן המובן לסביבתה במהלך פעילויות השגרה והמשחק בגן, בתיווך מבוגר הולך ופוחת\").\n"
                    + "   - **יש לפרק לתפקודים בודדים ומדויקים, לדוגמה:**\n"
                    + f"     * יעד 1: \"{active_name} תביע צרכים ורצונות באמצעות משפטים פשוטים\"\n"
                    + f"     * יעד 2: \"{active_name} תביע רעיונות במשחק עם מבוגר באמצעות שימוש במשפטים פשוטים\"\n"
                    + f"     * יעד 3: \"{active_name} תשתף בחוויה אישית קצרה באמצעות רצף משפטים פשוטים\"\n"
                    + f"   - **בנושא מובנות דיבור והיגוי:** אם קיים קושי במובנות/היגוי, יש להקדיש לו **יעד נפרד וספציפי** (למשל: \"{active_name} תהגה מילים דו-הברתיות באופן מובן בתוך שטף הדיבור\"), ולא להעמיס את הביטוי \"באופן המובן {env_pronoun}\" על שאר יעדי השפה.\n\n"
                    + "3. **התאמה מלאה למוקדי החיזוק של הילד/ה:**\n"
                    + "   - גזור את המטרות והיעדים אך ורק מתוך תחומי הקושי שצוינו בטבלת התפקוד שהוזנה.\n\n"
                    + "4. **דרכי הוראה, שיטות ואמצעים:** פירוט מעשי של אסטרטגיות מתוך שגרת הגן והטיפולים (שיחה בפת שחרית, משחקי קופסה, משחקי דמיון, מדרשי תמונה, מחברת שפה, טיפול פרטני/בזוגות, הדרכת הורים).\n\n"
                    + "5. **כמות מחייבת:** בדיוק 3 מטרות-על. לכל מטרת-על בדיוק 3 יעדים ספציפיים.\n\n"
                    + "---\n### מבנה הפלט הנדרש (JSON בלבד של רשימת 3 מטרות):\n"
                    + "[\n  {\n"
                    + f"    \"goal_title\": \"{active_name} תשתתף / תפעל / תביע...\",\n"
                    + "    \"domains\": \"תחום תפקוד\",\n"
                    + "    \"objectives_list\": [\n"
                    + f"      {{\"text\": \"{active_name} [תפקוד יחיד ממוקד 1]...\", \"timeframe\": \"עד סוף השנה\"}},\n"
                    + f"      {{\"text\": \"{active_name} [תפקוד יחיד ממוקד 2]...\", \"timeframe\": \"עד סוף השנה\"}},\n"
                    + f"      {{\"text\": \"{active_name} [תפקוד יחיד ממוקד 3]...\", \"timeframe\": \"עד סוף השנה\"}}\n"
                    + "    ],\n"
                    + "    \"teaching_methods\": \"• אסטרטגיה 1\\n• אסטרטגיה 2\\n• אסטרטגיה 3\"\n"
                    + "  }\n]"
                )

                client = genai.Client(api_key=api_key)
                response = client.models.generate_content(
                    model='gemini-3.6-flash',
                    contents=f"נתוני התפקוד של {active_name} ({gender}):\n{input_text}",
                    config=types.GenerateContentConfig(
                        system_instruction=system_prompt,
                        response_mime_type="application/json",
                        temperature=0.2
                    )
                )

                parsed_data = safe_parse_json(response.text)
                parsed_data = parsed_data[:3]
                for item in parsed_data:
                    item['ver'] = 0
                    if "objectives" in item and "objectives_list" not in item:
                        lines = [l.strip(" •\n\r") for l in item["objectives"].split("\n") if l.strip()]
                        item["objectives_list"] = [{"text": l, "timeframe": item.get("timeframe", "עד סוף השנה"), "ver": 0} for l in lines]

                    item_objs = item.get("objectives_list", [])[:3]
                    while len(item_objs) < 3:
                        item_objs.append({"text": f"{active_name} תשתתף בפעילות תפקודית מותאמת בהנחיית הצוות", "timeframe": "עד סוף השנה", "ver": 0})
                    for obj in item_objs:
                        obj['ver'] = 0
                    item['objectives_list'] = item_objs

                st.session_state['goals_list'] = parsed_data
                st.session_state['current_input_text'] = input_text
                st.session_state['just_generated'] = True
                st.rerun()
            except Exception as e:
                st.error(f"שגיאה בהפקה: {e}")
st.markdown('</div>', unsafe_allow_html=True)

# חיווי הצלחה ירוק
if st.session_state.get('just_generated', False) and st.session_state.get('goals_list'):
    st.success("המטרות והיעדים הופקו בהצלחה בדגש תפקודי ומוכוון השתתפות!")

# ממשק עריכה אינטראקטיבי
if st.session_state['goals_list']:
    active_name = student_name.strip() if student_name.strip() else ("הילדה" if "נקבה" in gender else "הילד")
    env_pronoun = "לסביבתה" if "נקבה" in gender else "לסביבתו"

    st.markdown("---")
    st.subheader("2. עריכה, דיוק והתאמת המטרות")

    for idx, goal in enumerate(st.session_state['goals_list']):
        g_ver = goal.get('ver', 0)
        current_title = goal.get('goal_title', '')

        with st.expander(f"מטרה {idx+1}: {current_title}", expanded=True, key=f"goal_expander_{idx}"):
            col_a, col_b = st.columns([3, 1])
            with col_a:
                goal['goal_title'] = st.text_input(
                    f"כותרת מטרה {idx+1}:", 
                    value=current_title, 
                    key=f"title_{idx}_{g_ver}"
                )

                goal['domains'] = st.text_input(
                    f"תחומי תפקוד:", 
                    value=goal.get('domains', ''), 
                    key=f"dom_{idx}_{g_ver}"
                )

            with col_b:
                st.write("**פעולות למטרה זו:**")
                
                # כפתור נסח מחדש למטרה
                st.markdown('<span class="marker-regen"></span>', unsafe_allow_html=True)
                if st.button("נסח מחדש", key=f"btn_regen_goal_{idx}"):
                    with st.spinner("מנסח חלופה תפקודית כוללת למטרה..."):
                        client = genai.Client(api_key=api_key)
                        cur_g_title = str(goal.get('goal_title', ''))
                        cur_inp = str(st.session_state.get('current_input_text', ''))
                        
                        regen_prompt = (
                            "אתה מומחה לניסוח תל\"א בגני חינוך מיוחד.\n"
                            + f"הצע ניסוח חלופי, כללי ותפקודי (מוכוון השתתפות פעילה בשגרת הגן) למטרת-העל עבור {active_name} ({gender}).\n"
                            + "התבסס באופן הדוק על הסגנון והשפה במאגר הדוגמאות:\n"
                            + str(examples_context) + "\n\n"
                            + f"הניסוח הנוכחי: '{cur_g_title}'\n"
                            + "רקע נתוני תפקוד:\n" + cur_inp + "\n\n"
                            + "דגשים קריטיים:\n"
                            + f"1. ניסוח תפקודי ברוח הדוגמאות (למשל: '{active_name} תשתתף...', '{active_name} תביע...', '{active_name} תיקח חלק...'). אסור לנסח 'תרכוש מיומנות'.\n"
                            + "2. מטרת-על כללית וכוללת, קצרה ותמציתית, ללא תנאים ספציפיים בכותרת.\n"
                            + f"3. השתמש בשם המפורש '{active_name}' ואל תכתוב 'הילדה' או 'הילד' (אלא אם זהו השם שנבחר).\n"
                            + "4. החזר אך ורק מחרוזת טקסט פשוטה של המטרה ללא מרכאות."
                        )
                        
                        res = client.models.generate_content(
                            model='gemini-3.6-flash',
                            contents=regen_prompt,
                            config=types.GenerateContentConfig(temperature=0.2)
                        )
                        goal['goal_title'] = res.text.strip().replace('"', '').replace("'", "")
                        goal['ver'] = g_ver + 1
                        st.rerun()

                # חלון עריכת מטרה לפי תיאור
                with st.form(key=f"form_edit_g_{idx}", clear_on_submit=False, border=False):
                    prompt_g_val = st.text_input(
                        "תיאור לעריכת מטרה:", 
                        placeholder="למשל: התייחס למובנות הדיבור", 
                        key=f"edit_g_p_{idx}", 
                        label_visibility="collapsed"
                    )
                    st.markdown('<span class="marker-edit"></span>', unsafe_allow_html=True)
                    submit_edit_g = st.form_submit_button("ערוך לפי תיאור")
                    
                    if submit_edit_g and prompt_g_val.strip():
                        with st.spinner("מעדכן ניסוח מטרה..."):
                            client = genai.Client(api_key=api_key)
                            cur_g_title = str(goal.get('goal_title', ''))
                            edit_g_prompt = (
                                f"ערוך את מטרת-העל עבור {active_name} ({gender}): '{cur_g_title}' לפי ההנחיה: '{prompt_g_val}'. "
                                + "שמור על סגנון הדוגמאות מתיקיית הדרייב ועל ניסוח תפקודי כולל, קצר ובהיר. "
                                + f"השתמש בשם המפורש '{active_name}'. החזר טקסט בלבד."
                            )
                            res = client.models.generate_content(
                                model='gemini-3.6-flash',
                                contents=edit_g_prompt,
                                config=types.GenerateContentConfig(temperature=0.2)
                            )
                            goal['goal_title'] = res.text.strip().replace('"', '').replace("'", "")
                            goal['ver'] = g_ver + 1
                            st.session_state[f"edit_g_p_{idx}"] = ""
                            st.rerun()

                # כפתור מחיקת מטרה
                st.markdown('<span class="marker-del"></span>', unsafe_allow_html=True)
                if st.button("מחק מטרה", key=f"del_{idx}"):
                    st.session_state['goals_list'].pop(idx)
                    st.rerun()

            st.markdown("---")

            # מבנה היעדים ודרכי ההוראה
            t_col_left, t_col_right = st.columns([6, 4])
            with t_col_left:
                for o_idx, obj_item in enumerate(goal.get('objectives_list', [])):
                    o_ver = obj_item.get('ver', 0)
                    col_obj_text, col_obj_tf = st.columns([4, 2])
                    with col_obj_text:
                        st.markdown(f"**יעד {o_idx+1}:**")
                        obj_item['text'] = st.text_area(
                            f"טקסט יעד {o_idx+1}", 
                            value=obj_item.get('text', ''), 
                            height=70, 
                            label_visibility="collapsed", 
                            key=f"obj_txt_{idx}_{o_idx}_{o_ver}"
                        )

                        c_b1, c_b2 = st.columns([1, 1])
                        with c_b1:
                            st.markdown('<span class="marker-regen"></span>', unsafe_allow_html=True)
                            if st.button("נסח מחדש", key=f"btn_reg_obj_{idx}_{o_idx}"):
                                with st.spinner("מנסח יעד ספציפי וממוקד..."):
                                    client = genai.Client(api_key=api_key)
                                    cur_g_title = str(goal.get('goal_title', ''))
                                    cur_obj_txt = str(obj_item.get('text', ''))
                                    cur_inp = str(st.session_state.get('current_input_text', ''))
                                    
                                    regen_obj_prompt = (
                                        "אתה מומחה לניסוח יעדים אופרטיביים בתל\"א לגני חינוך מיוחד.\n"
                                        + f"הצע ניסוח חלופי ליעד אופרטיבי זה בלבד עבור {active_name} ({gender}) הנגזר ממטרת-העל '{cur_g_title}'.\n"
                                        + "התבסס באופן מלא על שפת הדוגמאות במאגר:\n"
                                        + str(examples_context) + "\n\n"
                                        + f"היעד הנוכחי: '{cur_obj_txt}'\n"
                                        + "רקע נתוני תפקוד:\n" + cur_inp + "\n\n"
                                        + "דגשים מחייבים:\n"
                                        + "1. **תפקוד יחיד וספציפי בלבד:** על היעד להתמקד בפעולה מדויקת אחת (ללא העמסה וללא סרבול).\n"
                                        + f"2. ניסוח בהיר, קצר וישיר שפותח בשם המפורש '{active_name}'.\n"
                                        + "3. איסור על ניסוח 'תרכוש מיומנות' - השתמש בפועל של עשייה והשתתפות.\n"
                                        + "4. החזר משפט יחיד בלבד ללא מרכאות או תוספות."
                                    )
                                    res = client.models.generate_content(
                                        model='gemini-3.6-flash',
                                        contents=regen_obj_prompt,
                                        config=types.GenerateContentConfig(temperature=0.2)
                                    )
                                    obj_item['text'] = res.text.strip().replace('"', '').replace("'", "")
                                    obj_item['ver'] = o_ver + 1
                                    st.rerun()

                        with c_b2:
                            st.markdown('<span class="marker-del"></span>', unsafe_allow_html=True)
                            if st.button("מחק יעד", key=f"del_obj_{idx}_{o_idx}"):
                                goal['objectives_list'].pop(o_idx)
                                st.rerun()

                        # חלון עריכת יעד לפי תיאור
                        with st.form(key=f"form_pr_obj_{idx}_{o_idx}", clear_on_submit=False, border=False):
                            prompt_obj_val = st.text_input(
                                "ערוך יעד לפי תיאור:", 
                                placeholder="למשל: נסח בצורה פשוטה יותר", 
                                key=f"pr_obj_{idx}_{o_idx}", 
                                label_visibility="collapsed"
                            )
                            st.markdown('<span class="marker-edit"></span>', unsafe_allow_html=True)
                            submit_pr_obj = st.form_submit_button("ערוך לפי תיאור")

                            if submit_pr_obj and prompt_obj_val.strip():
                                with st.spinner("מעדכן יעד..."):
                                    client = genai.Client(api_key=api_key)
                                    cur_g_title = str(goal.get('goal_title', ''))
                                    cur_obj_txt = str(obj_item.get('text', ''))
                                    edit_obj_prompt = (
                                        f"ערוך את היעד של {active_name} ({gender}): '{cur_obj_txt}' לפי ההנחיה: '{prompt_obj_val}'. "
                                        + f"ודא שהיעד נגזר ממטרת-העל: '{cur_g_title}', עוסק בתפקוד יחיד ומוגדר בלבד, ללא סרבול, ופותח בשם '{active_name}'. "
                                        + "החזר משפט יחיד בלבד."
                                    )
                                    res = client.models.generate_content(
                                        model='gemini-3.6-flash',
                                        contents=edit_obj_prompt,
                                        config=types.GenerateContentConfig(temperature=0.2)
                                    )
                                    obj_item['text'] = res.text.strip().replace('"', '').replace("'", "")
                                    obj_item['ver'] = o_ver + 1
                                    st.session_state[f"pr_obj_{idx}_{o_idx}"] = ""
                                    st.rerun()

                    with col_obj_tf:
                        st.markdown("**פרק זמן להשגה**")
                        obj_item['timeframe'] = st.text_input(
                            f"פרק זמן ליעד {o_idx+1}", 
                            value=obj_item.get('timeframe', 'עד סוף השנה'), 
                            key=f"obj_T_{idx}_{o_idx}_{o_ver}", 
                            label_visibility="collapsed"
                        )

                    st.markdown("---")

                # הוספת יעד חדש
                st.markdown("**הוספת יעד חדש:**")
                with st.form(key=f"form_add_obj_{idx}", clear_on_submit=False, border=False):
                    add_obj_val = st.text_input(
                        "תיאור ליעד החדש (אופציונלי):", 
                        placeholder="למשל: הוסף יעד בסיסי יותר", 
                        key=f"add_obj_p_{idx}", 
                        label_visibility="collapsed"
                    )
                    st.markdown('<span class="marker-plus"></span>', unsafe_allow_html=True)
                    submit_add_obj = st.form_submit_button("הוסף יעד למטרה זו")

                    if submit_add_obj:
                        with st.spinner("מנסח יעד תפקודי חדש וממוקד..."):
                            client = genai.Client(api_key=api_key)
                            cur_g_title = str(goal.get('goal_title', ''))
                            prompt_add = (
                                "אתה מומחה לניסוח תל\"א.\n"
                                + f"הוסף יעד אופרטיבי נוסף, ספציפי וממוקד בתפקוד יחיד, שנגזר ישירות ממטרת-העל: '{cur_g_title}' עבור {active_name} ({gender}).\n"
                                + "התבסס על הסגנון והטרמינולוגיה בדוגמאות:\n"
                                + str(examples_context) + "\n"
                            )
                            if add_obj_val.strip():
                                prompt_add += f"\nדגש מיוחד ליעד: {add_obj_val}."
                            prompt_add += f"\nהקפד לפתוח בשם המפורש '{active_name}', לשמור על ניסוח קצר וממוקד בתפקוד יחיד ללא סרבול. החזר משפט יחיד בלבד ללא מרכאות."
                            
                            res = client.models.generate_content(
                                model='gemini-3.6-flash',
                                contents=prompt_add,
                                config=types.GenerateContentConfig(temperature=0.2)
                            )
                            goal.setdefault('objectives_list', []).append({
                                "text": res.text.strip().replace('"', '').replace("'", ""), 
                                "timeframe": "עד סוף השנה",
                                "ver": 0
                            })
                            st.session_state[f"add_obj_p_{idx}"] = ""
                            st.rerun()

            with t_col_right:
                st.markdown("**דרכי הוראה, השיטות והאמצעים**")
                st.markdown('<div class="teach-box">', unsafe_allow_html=True)
                goal['teaching_methods'] = st.text_area(
                    "דרכי הוראה ואמצעים:", 
                    value=goal.get('teaching_methods', ''), 
                    height=970, 
                    key=f"teach_{idx}_{g_ver}", 
                    label_visibility="collapsed"
                )
                st.markdown('</div>', unsafe_allow_html=True)

    # הוספת מטרה נוספת
    st.markdown("#### הוספת מטרה נוספת")
    with st.form(key="form_add_goal", clear_on_submit=False, border=False):
        col_new1, col_new2 = st.columns([3, 1])
        with col_new1:
            custom_prompt = st.text_input(
                "איזו מטרה תרצי להוסיף?", 
                placeholder="למשל: הוסף מטרה בתחום המשחק", 
                key="new_goal_inp", 
                label_visibility="collapsed"
            )
        with col_new2:
            st.markdown('<span class="marker-plus"></span>', unsafe_allow_html=True)
            submit_add_goal = st.form_submit_button("הוסף מטרה זו")

        if submit_add_goal and custom_prompt.strip():
            with st.spinner("מנסח מטרה ויעדים תפקודיים על בסיס הדוגמאות..."):
                client = genai.Client(api_key=api_key)
                add_prompt = (
                    "אתה מומחה לניסוח תל\"א בגני חינוך מיוחד.\n"
                    + f"נסח מטרה חדשה בתחום '{custom_prompt}' עבור {active_name} ({gender}) על בסיס מאגר הדוגמאות:\n"
                    + str(examples_context) + "\n\n"
                    + "כללים:\n"
                    + f"1. מטרת-על כללית וכוללת בדגש על השתתפות ותפקוד יומיומי ברוח הדוגמאות עבור '{active_name}' (ללא 'תרכוש מיומנות').\n"
                    + f"2. בדיוק 3 יעדים אופרטיביים ממוקדים (כל יעד עוסק בתפקוד יחיד וברור, ללא סרבול והעמסה) שנגזרים ממנה ופותחים בשם '{active_name}'.\n"
                    + "3. דרכי הוראה טיפוליות מעשיות.\n\n"
                    + "החזר JSON יחיד בלבד במבנה:\n"
                    + "{\n"
                    + f"  \"goal_title\": \"{active_name} תשתתף / תפעל...\",\n"
                    + "  \"domains\": \"תחום תפקוד\",\n"
                    + "  \"objectives_list\": [\n"
                    + f"    {{\"text\": \"{active_name} [תפקוד יחיד ממוקד 1]...\", \"timeframe\": \"עד סוף השנה\"}},\n"
                    + f"    {{\"text\": \"{active_name} [תפקוד יחיד ממוקד 2]...\", \"timeframe\": \"עד סוף השנה\"}},\n"
                    + f"    {{\"text\": \"{active_name} [תפקוד יחיד ממוקד 3]...\", \"timeframe\": \"עד סוף השנה\"}}\n"
                    + "  ],\n"
                    + "  \"teaching_methods\": \"• דרך הוראה 1\\n• דרך הוראה 2\"\n"
                    + "}"
                )
                
                res = client.models.generate_content(
                    model='gemini-3.6-flash',
                    contents=add_prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        temperature=0.2
                    )
                )
                new_item = safe_parse_json(res.text)
                if isinstance(new_item, list) and len(new_item) > 0:
                    new_item = new_item[0]
                new_item['ver'] = 0
                for obj in new_item.get('objectives_list', []):
                    obj['ver'] = 0
                st.session_state['goals_list'].append(new_item)
                st.rerun()

    # ייצוא קובץ Word
    st.markdown("---")
    st.subheader("3. ייצוא המסמך הסופי")

    def build_docx():
        doc = Document()
        title = doc.add_heading("תכנית לימודים אישית לתלמיד בחינוך המיוחד – תל\"א", level=1)
        title.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        p_name = doc.add_paragraph()
        p_name.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_name.add_run(f"שם התלמיד/ה: {active_name}\n").bold = True

        for idx, g in enumerate(st.session_state['goals_list']):
            p_goal = doc.add_paragraph()
            p_goal.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p_goal.add_run(f"\nמטרה {idx+1}: {g.get('goal_title', '')}\n").bold = True
            p_goal.add_run(f"תחומי התפקוד אליהם מתייחסת : {g.get('domains', '')}")

            table = doc.add_table(rows=1, cols=5)
            table.alignment = WD_TABLE_ALIGNMENT.RIGHT
            table.style = 'Table Grid'

            headers = ["יעדים", "פרק זמן להשגתם", "דרכי הוראה, השיטות והאמצעים", "סיכום הערכת ביניים", "סיכום הערכה מסכמת"]
            for col_i, header_text in enumerate(headers):
                cell = table.cell(0, col_i)
                cell.text = header_text
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                cell.paragraphs[0].runs[0].bold = True

            objs = g.get('objectives_list', [])
            if objs:
                objs_formatted = "\n\n".join([f"• {o.get('text', '')}" for o in objs])
                timeframes_formatted = "\n\n".join([f"{o.get('timeframe', 'עד סוף השנה')}" for o in objs])
            else:
                objs_formatted = g.get('objectives', '')
                timeframes_formatted = g.get('timeframe', 'עד סוף השנה')

            row = table.add_row()
            row_cells = row.cells
            row_cells[0].text = objs_formatted
            row_cells[1].text = timeframes_formatted
            row_cells[2].text = g.get('teaching_methods', '')
            row_cells[3].text = ""
            row_cells[4].text = ""

            for cell in row_cells:
                if cell.paragraphs:
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

        p_sign = doc.add_paragraph()
        p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_sign.add_run("\n\nחתימת הצוות: ____________________________")

        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()

    st.markdown('<span class="marker-download"></span>', unsafe_allow_html=True)
    st.markdown('<div class="download-btn">', unsafe_allow_html=True)
    st.download_button(
        label="הורד קובץ Word מעוצב עם טבלאות מלאות",
        data=build_docx(),
        file_name=f"תלא_{active_name}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    st.markdown('</div>', unsafe_allow_html=True)
