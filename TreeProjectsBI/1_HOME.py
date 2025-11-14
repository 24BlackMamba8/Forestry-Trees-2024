# -*- coding: utf-8 -*-
from __future__ import annotations
import streamlit as st
from streamlit.components.v1 import html as html_comp
from textwrap import dedent

# ייבוא חבילת העיצוב (אם קיימת)
try:
    from style_pack import inject_base_css, apply_plotly_theme, hero_header
except Exception:
    inject_base_css = apply_plotly_theme = hero_header = None

# ========= הגדרות עמוד =========
st.set_page_config(page_title="Forestry — 2024", page_icon="🌳", layout="wide")

# עיצוב בסיס (אם style_pack קיים)
if inject_base_css:
    inject_base_css(bg_main="assets/Capture2.jpg", bg_sidebar="assets/bg_sidebar.jpg")
if apply_plotly_theme:
    apply_plotly_theme()

# ========= HERO =========
if hero_header:
    hero_header(
        title="🌳 Forestry & Trees — 2024",
        subtitle="מערכת BI חברתית לניתוח דוחות כריתה וערעורים בשיתוף 'רחובות של עצים'.",
        logo_path="assets/logo.avif"   # ← הלוגו של העמותה / הפרויקט
    )
else:
    st.markdown(dedent("""
    <div style="direction:rtl;text-align:center;margin-top:8px">
      <div style="font-size:clamp(26px,5vw,40px);font-weight:800">🌳 Forestry & Trees — 2024</div>
      <div style="font-size:clamp(16px,2.4vw,20px);opacity:.9;margin-top:4px">
        מערכת BI חברתית לניתוח דוחות כריתה וערעורים בשיתוף "רחובות של עצים".
      </div>
    </div>
    """), unsafe_allow_html=True)

# ========= CSS + HTML (דף בית – Welcome + מפת ניווט קצרה + צוות) =========
css = dedent("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Heebo:wght@400;600;800&family=Assistant:wght@400;600;700&display=swap');

  :root{
    --card-bg: rgba(255,255,255,0.95);
    --card-border: rgba(0,0,0,0.08);
    --card-shadow: 0 10px 30px rgba(0,0,0,0.16);
    --text-color: #111;
    --muted: #444;
    --accent: #106b21;
  }

  html, body{
    margin:0; padding:0;
    background: transparent;
    font-family: 'Heebo','Assistant',system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;
    color: var(--text-color);
    direction: rtl;
  }

  .wrap{
    display:flex;
    align-items:flex-start;
    justify-content:center;
    padding: 10px 12px 28px;
  }

  .info-layout{
    width: min(1120px, 98vw);
    display:grid;
    grid-template-columns: 1.1fr 0.9fr;
    gap: 18px;
    align-items: stretch;
  }
  @media (max-width: 900px){
    .info-layout{
      grid-template-columns: 1fr;
    }
  }

  .info-card, .project-card{
    background: var(--card-bg);
    border-radius: 20px;
    padding: clamp(16px, 3vw, 26px) clamp(16px, 3vw, 26px);
    border: 1px solid var(--card-border);
    box-shadow: var(--card-shadow);
    backdrop-filter: blur(10px);
    -webkit-backdrop-filter: blur(10px);
  }

  .info-card h3,
  .project-card h3{
    margin: 0 0 10px 0;
    font-size: clamp(18px,2.4vw,24px);
    font-weight: 800;
    letter-spacing: 0.2px;
    text-align: right;
  }

  .info-card p,
  .project-card p{
    margin: 8px 0 10px 0;
    line-height: 1.85;
    font-size: clamp(14px,2.1vw,17px);
    color: var(--muted);
    text-align: right;
  }

  .tagline{
    font-size: clamp(14px,2vw,16px);
    margin-top: 6px;
    margin-bottom: 4px;
  }

  .tagline span{
    color: var(--accent);
    font-weight: 700;
  }

  .tool-list{
    list-style: none;
    padding: 0;
    margin: 10px 0 0 0;
  }

  .tool{
    display: grid;
    grid-template-columns: 40px 1fr;
    gap: 8px 12px;
    align-items: start;
    background: rgba(0,0,0,0.03);
    border-radius: 16px;
    padding: 10px 12px;
    margin: 8px 0;
    box-shadow: 0 1px 0 rgba(0,0,0,0.04);
    border: 1px solid rgba(0,0,0,0.04);
  }

  .tool .icon{
    font-size: 24px;
    line-height: 1;
    display:flex; align-items:center; justify-content:center;
  }

  .tool-title{
    font-weight: 800;
    font-size: clamp(14px,2vw,17px);
    margin: 0 0 3px 0;
    color: #0c3f12;
  }

  .tool-desc{
    margin: 0;
    font-size: clamp(13px,1.9vw,16px);
    line-height: 1.7;
    color: #222;
  }

  .nav-steps{
    counter-reset: step;
    list-style:none;
    padding:0;
    margin: 6px 0 0 0;
  }

  .nav-steps li{
    position:relative;
    margin: 8px 0;
    padding-right: 36px;
    font-size: clamp(13px,1.9vw,15px);
    line-height: 1.7;
    color:#222;
  }

  .nav-steps li::before{
    counter-increment: step;
    content: counter(step);
    position:absolute;
    right:0;
    top:4px;
    width:24px;
    height:24px;
    border-radius:999px;
    border:2px solid var(--accent);
    color:var(--accent);
    font-weight:700;
    font-size:13px;
    display:flex; align-items:center; justify-content:center;
    background:#fff;
  }

  .nav-steps strong{
    color: var(--accent);
  }

  .project-label{
    display:inline-block;
    background:#0f766e10;
    color:#0f766e;
    border-radius: 999px;
    padding: 4px 10px;
    font-size: 12px;
    font-weight: 700;
    margin-bottom: 6px;
    border: 1px solid #0f766e40;
  }

  .ngo-name{
    font-weight: 800;
    color:#14532d;
  }

  .team-box{
    margin-top: 10px;
    padding: 10px 12px;
    border-radius: 14px;
    background: #f4faf4;
    border: 1px solid rgba(16,107,33,0.16);
    font-size: 13px;
  }

  .team-box h4{
    margin: 0 0 6px 0;
    font-size: 14px;
    font-weight: 800;
    color:#166534;
  }

  .team-list{
    list-style:none;
    padding-right: 0;
    margin: 0;
  }
  .team-list li{
    margin: 2px 0;
  }
  .team-en{
    font-size: 12px;
    color:#555;
  }

  .mini-note{
    margin-top: 10px;
    font-size: 12px;
    color:#555;
  }

  @media (max-width: 560px){
    .tool{
      grid-template-columns: 32px 1fr;
      padding: 8px 10px;
    }
    .tool .icon{ font-size: 20px; }
  }

  @media print{
    .info-card, .project-card{
      background:#fff !important;
      border:1px solid #ddd !important;
      box-shadow:none !important;
      backdrop-filter:none !important;
      -webkit-backdrop-filter:none !important;
    }
    .tool{
      background:#fff !important;
      box-shadow:none !important;
      border:1px solid #eee !important;
    }
  }
</style>
""")

html = """
<div class="wrap">
  <div class="info-layout">

    <!-- כרטיס ימין: Welcome + מה יש במערכת -->
    <section class="info-card">
      <h3>ברוכים הבאים למערכת Forestry & Trees 2024</h3>
      <p>
        זהו ממשק BI אינטראקטיבי שנועד לעזור לנתח את
        <strong>דוחות הכריתה</strong> ו<strong>דוחות הערעורים</strong>
        של "רחובות של עצים" – בצורה נוחה, ויזואלית ומוכנה למצגות.
      </p>

      <p class="tagline">
        🔍 מסלול עבודה מומלץ:
        <span> Merge&amp Export → BI  כריתה →  ערעורים </span>
      </p>

      <ul class="tool-list">
        <li class="tool">
          <div class="icon">🧩</div>
          <div class="content">
            <div class="tool-title">Merge &amp; Export – מיזוג ופיענוח</div>
            <p class="tool-desc">
              מיזוג כל גליונות הכריתה לקובץ אחד, המרת קודים (יישוב / עץ),
              פיענוח פעולה/סיבה ותיקון תאריכים. בסוף התהליך מתקבל קובץ
              <strong>merged_forest_reports_FINAL_dates_fixed.xlsx</strong>
              המוכן ל-BI.
            </p>
          </div>
        </li>

        <li class="tool">
          <div class="icon">📊</div>
          <div class="content">
            <div class="tool-title">BI – דוחות כריתה</div>
            <p class="tool-desc">
              ניתוח ויזואלי של כריתות:
              TOP-יישובים, TOP-מיני עץ, פילוח לפי שנים, אזורים וסיבות כריתה,
              כולל גרפים אינטראקטיביים וייצוא ל-PNG / CSV.
            </p>
          </div>
        </li>

        <li class="tool">
          <div class="icon">📝</div>
          <div class="content">
            <div class="tool-title">BI – דוחות ערעורים</div>
            <p class="tool-desc">
              ניתוח ערעורים: אחוזי הצלחה, עצים שניצלו לשימור,
              יישובים מובילים, סיבות ערעור וסטטוס החלטה.
            </p>
          </div>
        </li>

        <li class="tool">
          <div class="icon">📘</div>
          <div class="content">
            <div class="tool-title">מדריך שימוש</div>
            <p class="tool-desc">
              דף ייעודי עם הסבר מפורט צעד־אחר־צעד:
              אילו קבצים צריך, איך להכין אותם באקסל ואיך לעבוד בתוך כל דף במערכת.
            </p>
          </div>
        </li>
      </ul>
    </section>

    <!-- כרטיס שמאל: פרויקט חברתי + צוות -->
    <section class="project-card">
      <span class="project-label">פרויקט חברתי־תכנותי</span>
      <h3>על הפרויקט והשותפים</h3>
      <p>
        המערכת פותחה כפרויקט BI חברתי־תכנותי בשיתוף
        <span class="ngo-name">עמותת "רחובות של עצים"</span>,
        במטרה להנגיש נתונים על כריתת עצים וערעורים,
        ולאפשר קבלת החלטות מושכלת ושקופה יותר.
      </p>

      <p>
        הפרויקט נבנה על־ידי סטודנטים למדעי המחשב מהקריה האקדמית אונו,
        בשילוב טכנולוגיות Python, Streamlit ו־BI.
      </p>

      <div class="team-box">
        <h4>צוות הפיתוח (הקריה האקדמית אונו)</h4>
        <ul class="team-list">
          <li>תמיר סנבטו – <span class="team-en">Tamir Sanbato</span></li>
          <li>יובל מדרי – <span class="team-en">Yuval Madari</span></li>
          <li>חן סמרה – <span class="team-en">Hen Samara</span></li>
        </ul>
      </div>

      <p class="mini-note">
        💡 לפרטים טכניים מלאים (הכנת קבצים, שלבי מיזוג ושימוש ב-BI) –
        ניתן לעבור לדף <strong>"📘 מדריך שימוש"</strong> מתוך סרגל הצד.
      </p>
    </section>

  </div>
</div>
"""

html_comp(css + html, height=1000, scrolling=False)
