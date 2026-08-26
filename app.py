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

st.title("🌱 ממשק חכם לניסוח תכנית לימודים אישית לתלמיד בחינוך המיוחד – תל\"א")
st.caption("מערכת לגזירת מטרות על ויעדים אופרטיביים על בסיס תיאור הילד/ה")

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

with st.sidebar:
    st.header("הגדרות מערכת")
    api_key = st.text_input(
        "מפתח API:", 
        value="AQ.Ab8RN6KO-aLO8c_IGTJZnUl0ju67TTDDblmIPxEQr4LeH0KGAA", 
        type="password"
    )
    if examples_context:
        st.success("מאגר הדוגמאות מחובר ומנותח!")
    else:
        st.warning("לא אותרו קבצים בתיקיית הדרייב.")
    
    if st.button("🔄 רענן מאגר דרייב"):
        st.cache_data.clear()
        st.rerun()

if 'goals_list' not in st.session_state:
    st.session_state['goals_list'] = []

# טופס ראשי
col1, col2 = st.columns([1, 3])
with col1:
    gender = st.radio("התאמה מגדרית:", ["ילדה (נקבה)", "ילד (זכר)"])
with col2:
    student_name = st.text_input("שם הילד/ה:", value="דנה", placeholder="הזיני את שם הילד/ה")

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
    if not input_text.strip():
        st.warning("אנא הזיני נתוני תפקוד.")
    else:
        with st.spinner("מנתח דפוסים ומנסח מטרות תמציתיות וממוקדות..."):
            try:
                system_prompt = f"""
אתה עוזר פדגוגי ופרא-רפואי מומחה לניסוח תכניות לימודים אישיות (תל"א) בגני חינוך מיוחד ושיקומיים.
תפקידך: לגזור בדיוק 3 מטרות-על תמציתיות וממוקדות, יעדים אופרטיביים ודרכי הוראה עבור הילד/ה '{student_name}' ({gender}) על בסיס מוקדי הכוח והחיזוק.

---
### מאגר הדוגמאות ללמידה וחיקוי:
{examples_context}

---
### כללי ניסוח קריטיים למטרת-העל (תמציתיות ופשטות):
1. **מבנה משפט פשוט וקצר (ללא פיצול):**
   - מטרת-העל חייבת להיות משפט קצר ותמציתי המתאר פעולה תפקודית מרכזית **אחת**.
   - **אסור לחבר שתי פעולות שונות באמצעות ו"ו החיבור!**
2. **אינטגרציה אורגנית (אם נדרשת):** מותר לשלב רק רכיבים שמגדירים את אותה הפעולה עצמה.
3. **מבנה היעדים האופרטיביים:**
   - כל יעד ייפתח בשם הילד/ה ('{student_name}') + פועל תפקודי נצפה.
   - חובה לציין רמת תיווך מדורגת והקשר בשגרת הגן.
4. **דרכי הוראה עשירות:** פירוט אסטרטגיות טיפוליות מתוך הדוגמאות.

---
### מבנה הפלט הנדרש (JSON בלבד):
[
  {{
    "goal_title": "ניסוח מטרת-על קצרה, תמציתית וחד-פעולתית עבור {student_name}",
    "domains": "תחום תפקוד יחיד או משולב טבעי",
    "objectives_list": [
      {{"text": "יעד אופרטיבי 1", "timeframe": "עד סוף השנה"}},
      {{"text": "יעד אופרטיבי 2", "timeframe": "עד סוף השנה"}},
      {{"text": "יעד אופרטיבי 3", "timeframe": "עד סוף השנה"}}
    ],
    "teaching_methods": "• דרך הוראה 1\\n• דרך הוראה 2\\n• דרך הוראה 3"
  }}
]
"""
                client = genai.Client(api_key=api_key)
                response = client.models.generate_content(
                    model='gemini-3.6-flash',
                    contents=f"נתוני התפקוד של {student_name}:\n{input_text}",
                    config=types.GenerateContentConfig(
                        system_instruction=system_prompt,
                        response_mime_type="application/json"
                    )
                )
                parsed_data = json.loads(response.text)
                for item in parsed_data:
                    if "objectives" in item and "objectives_list" not in item:
                        lines = [l.strip(" •\n\r") for l in item["objectives"].split("\n") if l.strip()]
                        item["objectives_list"] = [{"text": l, "timeframe": item.get("timeframe", "עד סוף השנה")} for l in lines]
                st.session_state['goals_list'] = parsed_data
                st.session_state['current_input_text'] = input_text
                st.success("המטרות והיעדים הופקו בהצלחה!")
            except Exception as e:
                st.error(f"שגיאה בהפקה: {e}")
st.markdown('</div>', unsafe_allow_html=True)

# ממשק עריכה אינטראקטיבי
if st.session_state['goals_list']:
    st.markdown("---")
    st.subheader("2. עריכה, דיוק והתאמת המטרות")

    for idx, goal in enumerate(st.session_state['goals_list']):
        current_title = goal.get('goal_title', '')
        if "objectives_list" not in goal and "objectives" in goal:
            lines = [l.strip(" •\n\r") for l in goal["objectives"].split("\n") if l.strip()]
            goal["objectives_list"] = [{"text": l, "timeframe": goal.get('timeframe', 'עד סוף השנה')} for l in lines]

        with st.expander(f"🎯 מטרה {idx+1}: {current_title}", expanded=True):
            col_a, col_b = st.columns([3, 1])
            with col_a:
                new_title = st.text_input(f"כותרת מטרה {idx+1} (עדכון בלחיצת Enter):", value=current_title, key=f"title_{idx}")
                if new_title != goal['goal_title']:
                    goal['goal_title'] = new_title
                    st.rerun()

                goal['domains'] = st.text_input(f"תחומי תפקוד:", value=goal.get('domains', ''), key=f"dom_{idx}")
            
            with col_b:
                st.write("**פעולות למטרה זו:**")
                if st.button(f"🔄 נסח מחדש", key=f"regen_{idx}"):
                    with st.spinner("מנסח חלופה תמציתית..."):
                        client = genai.Client(api_key=api_key)
                        regen_prompt = f"הצע ניסוח חלופי תמציתי וממוקד בפעולה אחת עבור {student_name} ({gender}) בהתבסס על מוקדי הכוח והחיזוק:\n{goal['goal_title']}\nהחזר מחרוזת טקסט פשוטה בלבד."
                        res = client.models.generate_content(
                            model='gemini-3.6-flash',
                            contents=regen_prompt
                        )
                        goal['goal_title'] = res.text.strip().replace('"', '')
                        st.rerun()
                
                prompt_g_val = st.session_state.get(f"edit_g_p_{idx}", "")
                if st.button("✨ ערוך לפי תיאור", key=f"btn_edit_g_{idx}"):
                    if prompt_g_val.strip():
                        with st.spinner("מעדכן ניסוח מטרה..."):
                            client = genai.Client(api_key=api_key)
                            res = client.models.generate_content(
                                model='gemini-3.6-flash',
                                contents=f"ערוך את מטרת-העל עבור {student_name} ({gender}): '{goal['goal_title']}' לפי ההנחיה: '{prompt_g_val}'. החזר טקסט בלבד."
                            )
                            goal['goal_title'] = res.text.strip().replace('"', '')
                            st.rerun()
                
                st.text_input("תיאור לעריכת מטרה:", placeholder="למשל: הוסף התייחסות למובנות דיבור", key=f"edit_g_p_{idx}", label_visibility="collapsed")
                
                if st.button(f"🗑️ מחק מטרה", key=f"del_{idx}"):
                    st.session_state['goals_list'].pop(idx)
                    st.rerun()

            st.markdown("---")
            
            # מבנה טבלאי
            t_col_left, t_col_right = st.columns([6, 4])
            with t_col_left:
                for o_idx, obj_item in enumerate(goal.get('objectives_list', [])):
                    col_obj_text, col_obj_tf = st.columns([4, 2])
                    with col_obj_text:
                        st.markdown(f"**יעד {o_idx+1}:**")
                        goal['objectives_list'][o_idx]['text'] = st.text_area(f"טקסט יעד {o_idx+1}", value=obj_item.get('text', ''), height=70, label_visibility="collapsed", key=f"obj_txt_{idx}_{o_idx}")
                        
                        c_b1, c_b2 = st.columns([1, 1])
                        with c_b1:
                            if st.button("🔄 נסח מחדש", key=f"reg_obj_{idx}_{o_idx}"):
                                with st.spinner("מנסח יעד..."):
                                    client = genai.Client(api_key=api_key)
                                    res = client.models.generate_content(
                                        model='gemini-3.6-flash',
                                        contents=f"נסח מחדש יעד אופרטיבי זה עבור {student_name} ({gender}): '{obj_item.get('text', '')}'. מטרת-העל היא: '{goal['goal_title']}'. החזר משפט יחיד בלבד."
                                    )
                                    goal['objectives_list'][o_idx]['text'] = res.text.strip().replace('"', '')
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
                                        contents=f"ערוך את היעד עבור {student_name} ({gender}): '{obj_item.get('text', '')}' לפי ההנחיה: '{prompt_obj_val}'. החזר משפט יחיד בלבד."
                                    )
                                    goal['objectives_list'][o_idx]['text'] = res.text.strip().replace('"', '')
                                    st.rerun()
                                    
                        st.text_input("ערוך יעד לפי תיאור:", placeholder="למשל: יעד בסיסי יותר", key=f"pr_obj_{idx}_{o_idx}", label_visibility="collapsed")

                    with col_obj_tf:
                        st.markdown("**פרק זמן להשגה**")
                        goal['objectives_list'][o_idx]['timeframe'] = st.text_input(f"פרק זמן ליעד {o_idx+1}", value=obj_item.get('timeframe', 'עד סוף השנה'), key=f"obj_T_{idx}_{o_idx}", label_visibility="collapsed")

                    st.markdown("---")

                # הוספת יעד חדש
                st.markdown("**➕ הוספת יעד חדש:**")
                add_obj_val = st.session_state.get(f"add_obj_p_{idx}", "")
                if st.button("הוסף יעד למטרה זו", key=f"do_add_obj_{idx}"):
                    with st.spinner("מוסיף יעד חדש..."):
                        client = genai.Client(api_key=api_key)
                        prompt_add = f"נסח יעד אופרטיבי נוסף למטרת-העל: '{goal['goal_title']}' עבור {student_name} ({gender})."
                        if add_obj_val.strip():
                            prompt_add += f" דגש: {add_obj_val}."
                        prompt_add += " החזר משפט יחיד בלבד."
                        res = client.models.generate_content(model='gemini-3.6-flash', contents=prompt_add)
                        goal.setdefault('objectives_list', []).append({"text": res.text.strip().replace('"', ''), "timeframe": "עד סוף השנה"})
                        st.rerun()
                st.text_input("תיאור ליעד החדש (אופציונלי):", placeholder="למשל: יעד הדרגתי לשלב המוקדם", key=f"add_obj_p_{idx}", label_visibility="collapsed")

            with t_col_right:
                st.markdown("**דרכי הוראה, השיטות והאמצעים**")
                st.markdown('<div class="teach-box">', unsafe_allow_html=True)
                goal['teaching_methods'] = st.text_area(
                    "דרכי הוראה ואמצעים:", 
                    value=goal.get('teaching_methods', ''), 
                    height=970, 
                    key=f"teach_{idx}", 
                    label_visibility="collapsed"
                )
                st.markdown('</div>', unsafe_allow_html=True)

    # הוספת מטרה חדשה
    st.markdown("#### ➕ הוספת מטרה נוספת")
    col_new1, col_new2 = st.columns([3, 1])
    with col_new1:
        custom_prompt = st.text_input("איזו מטרה תרצי להוסיף?", placeholder="למשל: נסח מטרה תמציתית בתחום משחק חברתי בחצר")
    with col_new2:
        st.write("")
        st.write("")
        if st.button("הוסף מטרה זו"):
            if custom_prompt.strip():
                with st.spinner("מנסח מטרה חדשה..."):
                    client = genai.Client(api_key=api_key)
                    add_prompt = f"""
נסח מטרה חדשה בתחום '{custom_prompt}' עבור {student_name} ({gender}) ללא חיבור שתי פעולות.
החזר JSON יחיד במבנה:
{{
  "goal_title": "ניסוח מטרת-על קצרה",
  "domains": "תחום תפקוד",
  "objectives_list": [
    {{"text": "יעד 1", "timeframe": "עד סוף השנה"}},
    {{"text": "יעד 2", "timeframe": "עד סוף השנה"}}
  ],
  "teaching_methods": "• דרך הוראה 1\\n• דרך הוראה 2"
}}
"""
                    res = client.models.generate_content(
                        model='gemini-3.6-flash',
                        contents=add_prompt,
                        config=types.GenerateContentConfig(response_mime_type="application/json")
                    )
                    st.session_state['goals_list'].append(json.loads(res.text))
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
        p_name.add_run(f"שם התלמיד/ה: {student_name}\n").bold = True

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
        file_name=f"תלא_{student_name}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
