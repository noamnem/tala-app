import os
import io
import json
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

    /* החלת גופן אלף ויישור לימין על רכיבי טקסט בלבד */
    html, body, .stMarkdown, p, h1, h2, h3, h4, label, input, textarea {
        font-family: 'Alef', sans-serif !important;
        direction: rtl !important;
        text-align: right !important;
    }

    /* תיקון האייקונים של המערכת למניעת עיוותים ועליית טקסט */
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
        font-size: 0.85rem !important;
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

    /* כפתורים כלליים */
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        font-weight: bold;
        font-family: 'Alef', sans-serif !important;
        direction: rtl !important;
        text-align: center !important;
    }
    .main-btn>button {
        background-color: #2E7D32 !important;
        color: white !important;
        font-size: 1.1rem !important;
        padding: 10px !important;
    }

    /* תיקון אזור העלאת קבצים */
    [data-testid="stFileUploader"] section {
        direction: ltr !important;
        text-align: left !important;
    }
    [data-testid="stFileUploader"] section button {
        direction: ltr !important;
    }

    /* מתיחה מדויקת של חלון דרכי ההוראה עד לקו המפריד של יעד 3 ומרכוז אנכי */
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
        font-size: 0.95rem !important;
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

with st.spinner("מסנכרן דוגמאות מתיקיית הדרייב..."):
    examples_context = sync_and_load_drive_examples()

# כפתור קומפקטי בפינה השמאלית העליונה
with st.popover("📁 מאגר דרייב"):
    st.markdown("**סטטוס חיבור ל-Google Drive:**")
    if examples_context:
        num_docs = examples_context.count("=== דוגמת תל\"א מקצועית")
        st.success(f"🟢 מחובר ומסונכרן!\n\nנטענו **{num_docs}** קובצי תל\"א ללמידת המודל.")
    else:
        st.warning("⚠️ לא אותרו קבצים בתיקייה (וודאי שהשיתוף פתוח לצפייה לכולם).")
    
    if st.button("🔄 רענן מאגר דרייב", key="refresh_drive_btn"):
        st.cache_data.clear()
        st.rerun()

st.title("🌱 ממשק חכם לניסוח תכנית לימודים אישית לתלמיד בחינוך המיוחד – תל\"א")
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
st.subheader("תיאור רמת התפקוד בתחומים הרלוונטיים")
st.markdown("**1. הזנת מוקדי כוח ומוקדים לחיזוק**")

uploaded_file = st.file_uploader("העלי קובץ Word של מוקדי כוח ומוקדים לחיזוק:", type=["docx"])

input_text = ""
if uploaded_file is not None:
    doc_uploaded = Document(uploaded_file)
    extracted = [p.text for p in doc_uploaded.paragraphs if p.text.strip()]
    for table in doc_uploaded.tables:
        for row in table.rows:
            extracted.append(" | ".join([c.text.replace("\n", " ").strip() for c in row.cells]))
    input_text = "\n".join(extracted)
    st.info("הקובץ נטען בהצלחה.")
else:
    input_text = st.text_area("או הדביקי כאן את הטקסט:", height=120, placeholder="הדביקי כאן מוקדי כוח וחיזוק...")

st.markdown('<div class="main-btn">', unsafe_allow_html=True)
if st.button("🚀 הפק מטרות ויעדים"):
    active_name = student_name.strip() if student_name.strip() else ("הילדה" if "נקבה" in gender else "הילד")
    if not input_text.strip():
        st.warning("אנא הזיני נתוני תפקוד.")
    else:
        with st.spinner("מנתח דפוסים ומנסח מטרות ויעדים תפקודיים ממוקדי השתתפות..."):
            try:
                env_pronoun = "לסביבתה" if "נקבה" in gender else "לסביבתו"
                system_prompt = f"""
אתה מומחה פדגוגי וקליני בכיר לניסוח תכניות לימודים אישיות (תל"א) בגני חינוך מיוחד ושיקומיים.
תפקידך לנסח תל"א מקצועית עבור '{active_name}' ({gender}) על בסיס נתוני התפקוד.

חובה עליך להתבסס באופן מלא על שפת הניסוח, הטרמינולוגיה המקצועית ואופן גזירת המדרג בין מטרות ליעדים במאגר הדוגמאות:
---
### מאגר הדוגמאות המקצועיות:
{examples_context}

---
### עקרונות פדגוגיים וקליניים מחייבים:

1. **דגש תפקודי ומוכוון השתתפות יומיומית (Functioning & Active Participation):**
   - **איסור מוחלט על ניסוחים מופשטים של רכישת ידע:** (אין לנסח: "{active_name} תרכוש מיומנות...", "{active_name} תלמד מושגים...", "{active_name} תפתח יכולת...").
   - **חובה לנסח מטרות ויעדים של השתתפות פעילה, תפקוד ועשייה בשגרת הגן:** (יש לנסח: "{active_name} תשתתף במשחק הדדי...", "{active_name} תביע בחירה ורצון בזמן ארוחה/משחק...", "{active_name} תיקח חלק פעיל במפגש...", "{active_name} תפעל באופן עצמאי בהתארגנות...").

2. **טרמינולוגיה קלינית ומקצועית (קלינאות תקשורת וחינוך מיוחד):**
   - בהתייחסות למובנות דיבור או בהירות קולית, **חובה להשתמש במינוח המקצועי המדויק מתוך הדוגמאות**: למשל, "באופן המובן {env_pronoun} (הקרובה/הרחוקה)", "במובנות דיבור תואמת הקשר", ולא בביטויים פשטניים כגון "דיבור ברור".

3. **כמות מחייבת:** בדיוק 3 מטרות-על. לכל מטרת-על בדיוק 3 יעדים אופרטיביים.

4. **מטרת-העל (Goal Title) – רחבה, כוללת, תפקודית וקצרה:**
   - שאיפה תפקודית רחבה בתחום משמעותי בשגרת היומיום.
   - קצרה ותמציתית, ללא תנאי ביצוע ספציפיים וללא רמות תיווך בכותרת.
   - איסור מוחלט על חיבור שתי פעולות שונות באמצעות ו"ו החיבור.

5. **היעדים האופרטיביים (Objectives) – שלבי השתתפות מדורגים ומדידים:**
   - כל 3 היעדים נגזרים ישירות ובלעדית מאותה מטרת-על ומהווים סולם שלבים התפתחותי / תיווכי / תפקודי להשגתה.
   - כל יעד ייפתח בשם המפורש '{active_name}' (אין לכתוב 'הילדה'/'הילד') + פועל תפקודי נצפה של השתתפות/ביצוע + רמת תיווך מפורשת + הקשר והזדמנות בשגרת הגן.

6. **דרכי הוראה ואמצעים:** פירוט רחב של אסטרטגיות פדגוגיות וטיפוליות בהתאם לדוגמאות.

---
### מבנה הפלט הנדרש (JSON בלבד של בדיוק 3 מטרות):
[
  {{
    "goal_title": "{active_name} תשתתף / תפעל / תביע (שאיפה תפקודית כוללת)...",
    "domains": "תחום תפקוד",
    "objectives_list": [
      {{"text": "{active_name} תשתתף/תבצע (שלב 1 מדורג) בתיווך... בהקשר...", "timeframe": "עד סוף השנה"}},
      {{"text": "{active_name} תשתתף/תבצע (שלב 2 מדורג) בתיווך... בהקשר...", "timeframe": "עד סוף השנה"}},
      {{"text": "{active_name} תשתתף/תבצע (שלב 3 מתקדם/עצמאי) בהקשר...", "timeframe": "עד סוף השנה"}}
    ],
    "teaching_methods": "• אסטרטגיה 1\\n• אסטרטגיה 2\\n• אסטרטגיה 3"
  }}
]
"""
                client = genai.Client(api_key=api_key)
                response = client.models.generate_content(
                    model='gemini-3.6-flash',
                    contents=f"נתוני התפקוד של {active_name} ({gender}):\n{input_text}",
                    config=types.GenerateContentConfig(
                        system_instruction=system_prompt,
                        response_mime_type="application/json"
                    )
                )
                parsed_data = json.loads(response.text)
                
                # הבטחת בדיוק 3 מטרות ו-3 יעדים
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
                st.success("המטרות והיעדים הופקו בהצלחה בדגש תפקודי ומוכוון השתתפות!")
            except Exception as e:
                st.error(f"שגיאה בהפקה: {e}")
st.markdown('</div>', unsafe_allow_html=True)

# ממשק עריכה אינטראקטיבי
if st.session_state['goals_list']:
    active_name = student_name.strip() if student_name.strip() else ("הילדה" if "נקבה" in gender else "הילד")
    env_pronoun = "לסביבתה" if "נקבה" in gender else "לסביבתו"
    
    st.markdown("---")
    st.subheader("2. עריכה, דיוק והתאמת המטרות")

    for idx, goal in enumerate(st.session_state['goals_list']):
        g_ver = goal.get('ver', 0)
        current_title = goal.get('goal_title', '')
        
        with st.expander(f"🎯 מטרה {idx+1}: {current_title}", expanded=True, key=f"goal_expander_{idx}"):
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
                if st.button(f"🔄 נסח מחדש", key=f"btn_regen_goal_{idx}"):
                    with st.spinner("מנסח חלופה תפקודית כוללת למטרה..."):
                        client = genai.Client(api_key=api_key)
                        regen_prompt = f"""אתה מומחה לניסוח תל"א בגני חינוך מיוחד.
הצע ניסוח חלופי, כללי ותפקודי (מוכוון השתתפות פעילה בשגרת הגן) למטרת-העל עבור {active_name} ({gender}).
התבסס על הסגנון והשפה במאגר הדוגמאות:
{examples_context}

הניסוח הנוכחי: '{goal['goal_title']}'
רקע נתוני תפקוד:
{st.session_state.get('current_input_text', '')}

דגשים קריטיים:
1. ניסוח תפקודי של השתתפות ועשייה יומיומית (למשל: '{active_name} תשתתף...', '{active_name} תביע...', '{active_name} תיקח חלק...'). אסור לנסח 'תרכוש מיומנות'.
2. מטרת-על כללית וכוללת, קצרה ותמציתית, ללא תנאים ספציפיים בכותרת.
3. בהתייחסות למובנות דיבור השתמש בביטוי המקצועי: 'באופן המובן {env_pronoun}' (ולא 'דיבור ברור').
4. השתמש בשם המפורש '{active_name}' ואל תכתוב 'הילדה' או 'הילד' (אלא אם זהו השם שנבחר).
5. החזר אך ורק מחרוזת טקסט פשוטה של המטרה ללא מרכאות."""
                        res = client.models.generate_content(
                            model='gemini-3.6-flash',
                            contents=regen_prompt
                        )
                        goal['goal_title'] = res.text.strip().replace('"', '').replace("'", "")
                        goal['ver'] = g_ver + 1
                        st.rerun()
                
                prompt_g_val = st.session_state.get(f"edit_g_p_{idx}", "")
                if st.button("✨ ערוך לפי תיאור", key=f"btn_edit_g_{idx}"):
                    if prompt_g_val.strip():
                        with st.spinner("מעדכן ניסוח מטרה..."):
                            client = genai.Client(api_key=api_key)
                            res = client.models.generate_content(
                                model='gemini-3.6-flash',
                                contents=f"ערוך את מטרת-העל עבור {active_name} ({gender}): '{goal['goal_title']}' לפי ההנחיה: '{prompt_g_val}'. שמור על ניסוח תפקודי של השתתפות פעילה בשגרת הגן, בהתייחסות למובנות דיבור השתמש ב'באופן המובן {env_pronoun}' (ולא 'דיבור ברור'), והשתמש בשם המפורש '{active_name}'. החזר טקסט בלבד."
                            )
                            goal['goal_title'] = res.text.strip().replace('"', '').replace("'", "")
                            goal['ver'] = g_ver + 1
                            st.session_state[f"edit_g_p_{idx}"] = ""
                            st.rerun()
                
                st.text_input("תיאור לעריכת מטרה:", placeholder="למשל: התייחס למובנות הדיבור", key=f"edit_g_p_{idx}", label_visibility="collapsed")
                
                if st.button(f"🗑️ מחק מטרה", key=f"del_{idx}"):
                    st.session_state['goals_list'].pop(idx)
                    st.rerun()

            st.markdown("---")
            
            # מבנה טבלאי של היעדים ודרכי ההוראה
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
                            if st.button("🔄 נסח מחדש", key=f"btn_reg_obj_{idx}_{o_idx}"):
                                with st.spinner("מנסח יעד תפקודי הנגזר מהמטרה..."):
                                    client = genai.Client(api_key=api_key)
                                    regen_obj_prompt = f"""אתה מומחה לניסוח יעדים אופרטיביים בתל"א לגני חינוך מיוחד.
הצע ניסוח חלופי, תפקודי ומדורג (בדגש על השתתפות יומיומית) ליעד אופרטיבי זה בלבד עבור {active_name} ({gender}).
התבסס על הדוגמאות במאגר:
{examples_context}

מטרת-העל אליה היעד משתייך: '{goal['goal_title']}'
היעד הנוכחי: '{obj_item.get('text', '')}'
רקע נתוני תפקוד:
{st.session_state.get('current_input_text', '')}

דגשים מחייבים:
1. היעד חייב להיגזר ישירות ממטרת-העל '{goal['goal_title']}' ולהוות שלב של השתתפות/תפקוד פעיל ומדיד בשגרת הגן.
2. איסור על ניסוח 'תרכוש מיומנות' - השתמש בפועל של עשייה והשתתפות בפועל.
3. בהתייחסות למובנות דיבור השתמש בביטוי: 'באופן המובן {env_pronoun}' (ולא 'דיבור ברור').
4. חובה לפתוח בשם המפורש '{active_name}'.
5. כלול רמת תיווך מפורשת והקשר בשגרת הגן.
6. החזר משפט יחיד בלבד ללא מרכאות או תוספות."""
                                    res = client.models.generate_content(
                                        model='gemini-3.6-flash',
                                        contents=regen_obj_prompt
                                    )
                                    obj_item['text'] = res.text.strip().replace('"', '').replace("'", "")
                                    obj_item['ver'] = o_ver + 1
                                    st.rerun()
                        with c_b2:
                            if st.button("🗑️ מחק יעד", key=f"del_obj_{idx}_{o_idx}"):
                                goal['objectives_list'].pop(o_idx)
                                st.rerun()
                        
                        prompt_obj_val = st.session_state.get(f"pr_obj_{idx}_{o_idx}", "")
                        if st.button("✨ ערוך לפי תיאור", key=f"btn_pr_obj_{idx}_{o_idx}"):
                            if prompt_obj_val.strip():
                                with st.spinner("מעדכן יעד..."):
                                    client = genai.Client(api_key=api_key)
                                    res = client.models.generate_content(
                                        model='gemini-3.6-flash',
                                        contents=f"ערוך את היעד של {active_name} ({gender}): '{obj_item.get('text', '')}' לפי ההנחיה: '{prompt_obj_val}'. ודא שהיעד נגזר ממטרת-העל: '{goal['goal_title']}', מדגיש השתתפות ותפקוד, בהתייחסות למובנות דיבור משתמש ב'באופן המובן {env_pronoun}', פותח בשם '{active_name}' וכולל תיווך והקשר. החזר משפט יחיד בלבד."
                                    )
                                    obj_item['text'] = res.text.strip().replace('"', '').replace("'", "")
                                    obj_item['ver'] = o_ver + 1
                                    st.session_state[f"pr_obj_{idx}_{o_idx}"] = ""
                                    st.rerun()
                                    
                        st.text_input("ערוך יעד לפי תיאור:", placeholder="למשל: יעד בסיסי יותר", key=f"pr_obj_{idx}_{o_idx}", label_visibility="collapsed")

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
                st.markdown("**➕ הוספת יעד חדש:**")
                add_obj_val = st.session_state.get(f"add_obj_p_{idx}", "")
                if st.button("הוסף יעד למטרה זו", key=f"do_add_obj_{idx}"):
                    with st.spinner("מנסח יעד תפקודי חדש הנגזר ישירות ממטרת-העל..."):
                        client = genai.Client(api_key=api_key)
                        prompt_add = f"""אתה מומחה לניסוח תל"א.
הוסף יעד אופרטיבי נוסף של השתתפות ותפקוד יומיומי שנגזר ישירות ובלעדית ממטרת-העל: '{goal['goal_title']}' עבור {active_name} ({gender}).
התבסס על הסגנון והטרמינולוגיה בדוגמאות:
{examples_context}
"""
                        if add_obj_val.strip():
                            prompt_add += f"\nדגש מיוחד ליעד: {add_obj_val}."
                        prompt_add += f"\nהקפד לפתוח בשם המפורש '{active_name}', דגש על השתתפות פעילה, בהתייחסות למובנות להשתמש ב'באופן המובן {env_pronoun}', לציין רמת תיווך והקשר בשגרת הגן. החזר משפט יחיד בלבד ללא מרכאות."
                        res = client.models.generate_content(model='gemini-3.6-flash', contents=prompt_add)
                        goal.setdefault('objectives_list', []).append({
                            "text": res.text.strip().replace('"', '').replace("'", ""), 
                            "timeframe": "עד סוף השנה",
                            "ver": 0
                        })
                        st.session_state[f"add_obj_p_{idx}"] = ""
                        st.rerun()
                st.text_input("תיאור ליעד החדש (אופציונלי):", placeholder="למשל: שלב השתתפות מקדים", key=f"add_obj_p_{idx}", label_visibility="collapsed")

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

    # הוספת מטרה חדשה
    st.markdown("#### ➕ הוספת מטרה נוספת")
    col_new1, col_new2 = st.columns([3, 1])
    with col_new1:
        custom_prompt = st.text_input("איזו מטרה תרצי להוסיף?", placeholder="למשל: נסח מטרה תפקודית בתחום משחק חברתי בחצר")
    with col_new2:
        st.write("")
        st.write("")
        if st.button("הוסף מטרה זו"):
            if custom_prompt.strip():
                with st.spinner("מנסח מטרה ויעדים תפקודיים על בסיס הדוגמאות..."):
                    client = genai.Client(api_key=api_key)
                    add_prompt = f"""
אתה מומחה לניסוח תל"א בגני חינוך מיוחד.
נסח מטרה חדשה בתחום '{custom_prompt}' עבור {active_name} ({gender}) על בסיס מאגר הדוגמאות:
{examples_context}

כללים:
1. מטרת-על כללית וכוללת בדגש על השתתפות ותפקוד יומיומי עבור '{active_name}' (ללא 'תרכוש מיומנות'). בהתייחסות למובנות להשתמש ב'באופן המובן {env_pronoun}'.
2. בדיוק 3 יעדים אופרטיביים המהווים שלבי השתתפות מדורגים (התפתחותיים / תיווך) שנגזרים ממנה ופותחים בשם '{active_name}'.
3. דרכי הוראה טיפוליות מעשיות.

החזר JSON יחיד בלבד במבנה:
{{
  "goal_title": "{active_name} תשתתף / תפעל...",
  "domains": "תחום תפקוד",
  "objectives_list": [
    {{"text": "{active_name} תשתתף/תבצע (שלב 1)...", "timeframe": "עד סוף השנה"}},
    {{"text": "{active_name} תשתתף/תבצע (שלב 2)...", "timeframe": "עד סוף השנה"}},
    {{"text": "{active_name} תשתתף/תבצע (שלב 3)...", "timeframe": "עד סוף השנה"}}
  ],
  "teaching_methods": "• דרך הוראה 1\\n• דרך הוראה 2"
}}
"""
                    res = client.models.generate_content(
                        model='gemini-3.6-flash',
                        contents=add_prompt,
                        config=types.GenerateContentConfig(response_mime_type="application/json")
                    )
                    new_item = json.loads(res.text)
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
        title = doc.add_heading(f"תכנית לימודים אישית לתלמיד בחינוך המיוחד – תל\"א", level=1)
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

    st.download_button(
        label="📥 הורד קובץ Word מעוצב עם טבלאות מלאות",
        data=build_docx(),
        file_name=f"תלא_{active_name}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
