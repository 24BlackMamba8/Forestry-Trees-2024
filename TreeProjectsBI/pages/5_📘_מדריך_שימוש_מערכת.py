# -*- coding: utf-8 -*-
from __future__ import annotations
import streamlit as st
from textwrap import dedent
from streamlit.components.v1 import html as html_comp

try:
    from style_pack import inject_base_css, apply_plotly_theme, hero_header
except Exception:
    inject_base_css = apply_plotly_theme = hero_header = None

# ========= הגדרות עמוד =========
st.set_page_config(page_title="📘 מדריך שימוש — Forestry 2024", layout="wide")

if inject_base_css:
    inject_base_css(bg_main="assets/bg_main.jpg", bg_sidebar="assets/bg_sidebar.jpg")
if apply_plotly_theme:
    apply_plotly_theme()

if hero_header:
    hero_header(
        "📘 מדריך שימוש — Forestry & Trees 2024",
        "הסבר מלא: אילו קבצים צריך, איך להריץ את המיזוג, ואיך להשתמש בדפי ה-BI.",
        logo_path=None,
    )
else:
    st.markdown(
        "<h2 style='direction:rtl;text-align:right'>📘 מדריך שימוש — Forestry & Trees 2024</h2>",
        unsafe_allow_html=True,
    )

# ========= CSS + HTML – כרטיס לבן אטום =========
css = dedent("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Heebo:wght@400;600;800&family=Assistant:wght@400;600;700&display=swap');

  html, body{
    margin:0; padding:0;
    background: transparent;
    font-family: 'Heebo','Assistant',system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;
    direction: rtl;
  }

  .guide-wrap{
    display:flex;
    justify-content:center;
    align-items:flex-start;
    padding: 10px 12px 28px;
  }

  .guide-card{
    width: min(1100px, 100%);
    background:#ffffff !important;           /* לבן מלא, בלי שקיפות */
    border-radius: 20px;
    padding: 20px 26px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.18);
    border: 1px solid rgba(0,0,0,0.08);
    color:#111;
    text-align:right;
  }

  .guide-card h3{
    margin-top: 14px;
    margin-bottom: 6px;
    font-size: 20px;
    font-weight: 800;
  }

  .guide-card h4{
    margin-top: 10px;
    margin-bottom: 4px;
    font-size: 17px;
    font-weight: 700;
  }

  .guide-card p,
  .guide-card li{
    font-size: 15px;
    line-height: 1.9;
  }

  .step-badge{
    display:inline-flex;
    align-items:center;
    justify-content:center;
    border-radius:999px;
    background:#106b21;
    color:#fff;
    font-weight:700;
    font-size:12px;
    width:22px;height:22px;
    margin-left:6px;
  }

  .file-tag{
    display:inline-block;
    background:#eef7ee;
    border-radius:999px;
    padding:2px 9px;
    margin:0 0 3px 4px;
    font-size:13px;
    border:1px solid #cde2ce;
  }

  .tip-box{
    border-radius:12px;
    border:1px solid rgba(0,0,0,0.08);
    padding:10px 12px;
    background:#f4fbf4;
    margin-top:8px;
    margin-bottom:8px;
    font-size:14px;
  }

  @media print{
    .guide-card{
      background:#ffffff !important;
      box-shadow:none !important;
      border:1px solid #ccc !important;
    }
  }
</style>
""")

html = """
<div class="guide-wrap">
  <div class="guide-card">

    <h3>1. מה המערכת יודעת לעשות?</h3>
    <p>
      מערכת <strong>Forestry & Trees 2024</strong> מיועדת לשלושה שימושים עיקריים:
    </p>
    <ul>
      <li>
        <strong>מיזוג וניקוי נתונים</strong> – איחוד כל דוחות הכריתה לקובץ אחד,
        פיענוח קודים (יישוב, מין עץ, פעולה, סיבה), תיקון תאריכים וניקוי רשומות בעייתיות.
      </li>
      <li>
        <strong>יצירת קובץ BI מאוחד לכריתה</strong> – הפקת קובץ
        <code>merged_forest_reports_FINAL_dates_fixed.xlsx</code> – בסיס אחד מסודר
        שאפשר להטעין ל-Streamlit, Power BI, Excel Pivot ועוד.
      </li>
      <li>
        <strong>ניתוח BI לדוחות כריתה ולדוחות ערעורים</strong> – דף BI לכריתה (TOP-יישובים, TOP-עצים,
        פילוחים לפי זמן/אזור/סיבה) ודף BI לערעורים (אחוזי הצלחה, עצים שניצלו לשימור, יישובים מובילים ועוד).
      </li>
    </ul>

    <h3>2. אילו קבצים נדרשים לעבודה?</h3>
    <p>לצורך עבודה מלאה עם כל הכלים במערכת נדרשים:</p>

    <h4>2.1 דוחות כריתה מאוחדים</h4>
    <p>
      קובץ מהגוף הממשלתי (למשל <em>"forestry_and_trees_report2024 דוחות כריתה אחוד.xlsx"</em>), שמכיל:
    </p>
    <ul>
      <li>מספר גליונות כריתה לפי אזורים/מחוזות.</li>
      <li>גיליון <code>רשימת ערים לפי קודים</code>.</li>
      <li>גיליון <code>רשימת עצים לפי קודים</code>.</li>
    </ul>

    <h4>2.2 קובצי רשימות ייעודיים</h4>
    <p>מהמחשב שלך נשתמש בשלושה קבצים נפרדים:</p>
    <ul>
      <li>
        <span class="file-tag">קובץ 1 – דוחות כריתה מאוחדים (ללא גליונות רשימות)</span><br/>
        קובץ שבו נשארים רק גליונות הכריתה עצמם.
      </li>
      <li>
        <span class="file-tag">קובץ 2 – רשימת ערים לפי קודים.xlsx</span><br/>
        כולל עמודות כמו <code>ישוב</code> ו-<code>סמל ישוב</code>.
      </li>
      <li>
        <span class="file-tag">קובץ 3 – רשימת עצים לפי קודים.xlsx</span><br/>
        כולל עמודות כמו <code>Tree</code> ו-<code>שם עץ</code>.
      </li>
    </ul>

    <h4>2.3 קובץ דוחות ערעורים</h4>
    <p>
      קובץ נפרד (למשל <em>"forestry_and_trees_ararim2024- דוח ערערים.xlsx"</em>)
      המשמש רק לדף BI של הערעורים.
    </p>

    <div class="tip-box">
      💡 <strong>טיפ:</strong> מומלץ ליצור תיקייה מסודרת במחשב
      (למשל: <code>TreeProjectsBI/inputs</code>)
      ושם לשמור את 4 הקבצים: כריתה מאוחד, רשימת ערים, רשימת עצים, ודוח ערעורים.
    </div>

    <h3>3. איך להכין את שלושת קובצי הכריתה (ידנית באקסל)?</h3>

    <h4><span class="step-badge">1</span> פתיחת הקובץ המקורי</h4>
    <ul>
      <li>פתח/י באקסל את הקובץ שקיבלת (למשל: <em>"forestry_and_trees_report2024 דוחות כריתה אחוד.xlsx"</em>).</li>
      <li>ודא/י שיש בו גליונות עם שמות דומים ל: <code>רשימת ערים לפי קודים</code> ו-<code>רשימת עצים לפי קודים</code>.</li>
    </ul>

    <h4><span class="step-badge">2</span> שמירת רשימת ערים לקובץ נפרד</h4>
    <ul>
      <li>עמוד/י על הגיליון <code>רשימת ערים לפי קודים</code>.</li>
      <li>לחץ/י <em>File → Save As</em> ושמור/י כקובץ חדש בשם: <code>רשימת ערים לפי קודים.xlsx</code>.</li>
      <li>ודא/י שהעמודות הראשיות בו הן: <code>ישוב</code> ו-<code>סמל ישוב</code>.</li>
    </ul>

    <h4><span class="step-badge">3</span> שמירת רשימת עצים לקובץ נפרד</h4>
    <ul>
      <li>חזור/י לקובץ המקורי, עבור/י לגיליון <code>רשימת עצים לפי קודים</code>.</li>
      <li>שמור/י אותו כקובץ חדש בשם: <code>רשימת עצים לפי קודים.xlsx</code>.</li>
      <li>ודא/י שיש בו עמודות: <code>Tree</code> (קוד העץ) ו-<code>שם עץ</code> (שם העץ בעברית).</li>
    </ul>

    <h4><span class="step-badge">4</span> יצירת קובץ דוחות כריתה "ללא רשימות"</h4>
    <ul>
      <li>בקובץ המקורי – מחק/י (או הסתר/י ושמור/י כעותק חדש) את הגיליונות:
          <code>רשימת ערים לפי קודים</code> ו-<code>רשימת עצים לפי קודים</code>.</li>
      <li>שמור/י את הקובץ בשם, למשל:
          <em>"forestry_and_trees_report2024 דוחות כריתה אחוד ללא הרשימות.xlsx"</em>.</li>
      <li>זה יהיה <strong>קובץ הדוחות הראשי</strong> שנעלה בדף <em>Merge &amp; Export</em>.</li>
    </ul>

    <h3>4. הרצת המיזוג — דף Merge &amp; Export</h3>
    <p>בדף הצד של האפליקציה בחר/י את <strong>🧩 Merge &amp; Export</strong>.</p>

    <h4><span class="step-badge">1</span> העלאת הקבצים</h4>
    <p>במסך המיזוג תופיע טופס העלאה עם שלושה שדות:</p>
    <ul>
      <li>קובץ דוחות כריתה מאוחד <strong>ללא</strong> גליונות רשימות  
          <span class="file-tag">forestry_and_trees_report2024 דוחות כריתה אחוד ללא הרשימות.xlsx</span></li>
      <li><span class="file-tag">רשימת ערים לפי קודים.xlsx</span></li>
      <li><span class="file-tag">רשימת עצים לפי קודים.xlsx</span></li>
    </ul>
    <p>לאחר בחירת שלושת הקבצים לחץ/י על כפתור <strong>"🚀 הרץ מיזוג והמרות"</strong>.</p>

    <h4><span class="step-badge">2</span> מה עושה המנוע בזמן הריצה?</h4>
    <ul>
      <li>קריאה של כל גליונות הכריתה בקובץ הראשי.</li>
      <li>איתור אוטומטי של שורת הכותרת גם אם:
        <ul>
          <li>שורה 1 היא כותרת כללית,</li>
          <li>שורה 2 ריקה,</li>
          <li>הכותרות האמיתיות נמצאות בשורה 3/4.</li>
        </ul>
      </li>
      <li>מיפוי עמודות לשמות אחידים (<code>אזור</code>, <code>יישוב</code>, <code>שם   מין עץ</code>,
          <code>מ-תאריך</code>, <code>עד-תאריך</code> וכו').</li>
      <li>המרת קודי יישוב → שם יישוב (בעזרת <strong>רשימת ערים לפי קודים</strong>).</li>
      <li>המרת קודי עץ → שם מין עץ (בעזרת <strong>רשימת עצים לפי קודים</strong>).</li>
      <li>פיענוח:
        <ul>
          <li><code>פעולה</code> → <code>פעולה_מפוענחת</code> (כריתה / העתקה).</li>
          <li><code>סיבה</code> → <code>סיבה_מפוענחת</code> (בטיחות, בנייה, עץ מת, וכו').</li>
        </ul>
      </li>
      <li>יצירת דגלים:
        <ul>
          <li><code>__is_cut__</code> – האם מדובר בשורת כריתה.</li>
          <li><code>__is_move__</code> – האם מדובר בשורת העתקה/שימור.</li>
        </ul>
      </li>
    </ul>

    <h4><span class="step-badge">3</span> תוצאה – קובץ BI מוכן</h4>
    <p>
      בסוף התהליך יופיע כפתור הורדה לקובץ:
      <code>merged_forest_reports_FINAL_dates_fixed.xlsx</code><br/>
      זהו הקובץ המרכזי שאיתו עובדים בדף BI של דוחות הכריתה,
      ואפשר להשתמש בו גם ב-Power BI / Pivot / כל כלי BI נוסף.
    </p>

    <h3>5. שימוש בקובץ הממוזג — דף BI של דוחות כריתה</h3>
    <p>בדף הצד בחר/י את <strong>📊 BI – דוחות כריתה</strong>.</p>
    <p>שם מופיע:</p>
    <ul>
      <li>רכיב העלאת קובץ — טען/י את  
          <span class="file-tag">merged_forest_reports_FINAL_dates_fixed.xlsx</span></li>
      <li>מסננים (שנים, יישובים, סוג פעולה, סיבה, ועוד).</li>
      <li>אוסף גרפים ו-TOP-10:
        <ul>
          <li><strong>TOP-10 יישובים בכמות כריתות</strong> – לפי ספירת שורות כריתה (<code>__is_cut__ = True</code>).</li>
          <li><strong>TOP-10 מיני עץ שנכרתו</strong> – לפי <code>שם   מין עץ</code> + <code>__is_cut__</code>.</li>
          <li>פילוח לפי סיבות, אזורים, טווחי תאריכים ועוד.</li>
        </ul>
      </li>
    </ul>
    <p>
      אחרי הטעינה ניתן לסנן לפי שנה/יישוב/סיבה, להוריד גרפים כ-PNG או להעתיק למצגת,
      ולהוריד קבצי CSV מסוננים להמשך ניתוח.
    </p>

    <h3>6. ניתוח דוחות ערעורים — דף BI ערעורים</h3>
    <p>בדף הצד בחר/י את <strong>📝 BI – דוחות ערעורים</strong>.</p>
    <p>שם תמצא/י:</p>
    <ul>
      <li>רכיב העלאת קובץ ערעורים
          (למשל: <span class="file-tag">forestry_and_trees_ararim2024- דוח ערערים.xlsx</span>).</li>
      <li>מנגנון אוטומטי שמזהה:
        <ul>
          <li>תאריך הערעור.</li>
          <li>יישוב (מתוך כתובת/טקסט).</li>
          <li>סיבת הערעור (בנייה / בטיחות / מטרד / אחר).</li>
          <li>סטטוס (התקבל / התקבל חלקית / נדחה / לא נדון).</li>
          <li>מספר עצים לשימור שניצלו.</li>
        </ul>
      </li>
    </ul>
    <p>לאחר הטעינה ניתן לראות:</p>
    <ul>
      <li>TOP-10 יישובים בכמות ערעורים.</li>
      <li>TOP-יישובים בהצלחת ערעורים.</li>
      <li>אחוזי הצלחה לפי סיבת ערעור.</li>
      <li>רשימת הערעורים "הגדולים" שהצילו הכי הרבה עצים.</li>
    </ul>

    <h3>7. מסלול עבודה מומלץ (סיכום קצר)</h3>
    <ul>
      <li><strong>איסוף קבצים</strong> – ודא/י שכל ארבעת הקבצים קיימים.</li>
      <li><strong>הפרדת רשימות</strong> – יצירת שלושת קבצי הכריתה:
          דוחות כריתה ללא רשימות + רשימת ערים + רשימת עצים.</li>
      <li><strong>Merge &amp; Export</strong> – הרצת המיזוג והפקת הקובץ
          <code>merged_forest_reports_FINAL_dates_fixed.xlsx</code>.</li>
      <li><strong>BI כריתה</strong> – טעינת קובץ המיזוג וניתוח דוחות הכריתה.</li>
      <li><strong>BI ערעורים</strong> – טעינת דוח הערעורים וניתוח ההחלטות/הצלחות.</li>
      <li><strong>יצוא והגשה</strong> – הורדת גרפים/CSV ושילובם בדוחות/מצגות/עבודות.</li>
    </ul>

    <h3>8. הערות נוספות ושיפור עתידי</h3>
    <ul>
      <li>אם עדיין מופיעים יישובים כקוד (למשל <code>1457</code>) – ניתן להוסיף אותם ידנית
          לקובץ <strong>רשימת ערים לפי קודים</strong> ולהריץ שוב את המיזוג.</li>
      <li>באופן דומה, אם מין עץ לא מזוהה, אפשר לעדכן את קובץ
          <strong>רשימת עצים לפי קודים</strong>.</li>
      <li>מומלץ לשמור גיבוי לקובצי המקור לפני כל שינוי.</li>
    </ul>
    <p>
      המדריך הזה נבנה כדי שתוכל/י גם להדפיס אותו כ-PDF ולצרף כנספח
      מקצועי לפרויקט ה-BI שלך.
    </p>

  </div>
</div>
"""

html_comp(css + html, height=1200, scrolling=True)
