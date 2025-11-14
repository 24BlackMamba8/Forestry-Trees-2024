# -*- coding: utf-8 -*-
from __future__ import annotations
import base64, os
import streamlit as st
import plotly.io as pio

# ---------- Plotly theme ----------
PLOTLY_LAYOUT = dict(
    margin=dict(l=40, r=30, t=60, b=40),
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(255,255,255,0.65)",
    font=dict(family="Heebo, Assistant, Rubik, Arial, sans-serif", size=14, color="#0f172a"),
    xaxis=dict(showgrid=True, gridcolor="rgba(15,23,42,0.08)", zeroline=False),
    yaxis=dict(showgrid=True, gridcolor="rgba(15,23,42,0.08)", zeroline=False),
)

def apply_plotly_theme():
    # בסיס: plotly_white + התאמות
    pio.templates["he_theme"] = pio.templates["plotly_white"]
    pio.templates["he_theme"].layout.update(PLOTLY_LAYOUT)
    pio.templates.default = "he_theme"

# ---------- helpers ----------
def _read_bytes(path: str) -> bytes:
    if not path:
        return b""
    try:
        if not os.path.isfile(path):
            return b""
        with open(path, "rb") as f:
            return f.read()
    except Exception:
        return b""

def _b64_or_empty(path: str) -> str:
    data = _read_bytes(path)
    return base64.b64encode(data).decode("utf-8") if data else ""

# ---------- CSS injection ----------
def inject_base_css(
    bg_main: str | None = None,
    bg_sidebar: str | None = None,
    blur_glass: int = 14,
    fallback_gradient: bool = True,
    overlay_alpha: float = 0.45,  # ← שכבת “הלבנה/הכהיה” מעל הרקע כדי שהטקסט יבלוט
):
    """
    מזריק CSS כולל RTL, טיפוגרפיה, וזכוכית.
    אם תמונות רקע לא קיימות/לא תקינות — משתמש בגרדיאנט עדין (אם fallback_gradient=True).
    overlay_alpha קובע כמה לרכך את הרקע (0=ללא, 0.45 מומלץ).
    """

    main_b64 = _b64_or_empty(bg_main) if bg_main else ""
    side_b64 = _b64_or_empty(bg_sidebar) if bg_sidebar else main_b64

    # רקע ראשי: תמונה + שכבת overlay לבנה שקופה, או גרדיאנט
    if main_b64:
        bg_main_css = (
            f"background-image: linear-gradient(rgba(255,255,255,{overlay_alpha}), rgba(255,255,255,{overlay_alpha})), "
            f"url('data:image/jpeg;base64,{main_b64}');"
            "background-size: cover; background-position: center; background-attachment: fixed;"
        )
    elif fallback_gradient:
        bg_main_css = "background: radial-gradient(1200px 600px at 100% 0%, #ebffef 0%, #ffffff 55%, #ffffff 100%);"
    else:
        bg_main_css = ""

    # רקע סרגל צד (אם אין תמונה — נשאר צבעוני)
    bg_side_css = f"background: url('data:image/jpeg;base64,{side_b64}') center/cover fixed no-repeat;" if side_b64 else ""

    st.markdown(f"""
    <style>
      /* גופנים (Google) */
      @import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;600;700&family=Assistant:wght@300;400;600;700&display=swap');

      html, body, [data-testid="stAppViewContainer"] * {{
        font-family: Heebo, Assistant, Rubik, Arial, sans-serif;
        direction: rtl;      /* RTL מלא */
        text-align: right;
      }}

      /* רקע ראשי */
      [data-testid="stAppViewContainer"] {{
        {bg_main_css}
        background-color: #f6faf7;
      }}

      /* רקע סרגל צד */
      [data-testid="stSidebar"] > div:first-child {{
        {bg_side_css}
        background-color: rgba(255,255,255,0.9);
      }}

      /* שקיפות קלה לאיזון */
      [data-testid="stSidebar"] {{
        backdrop-filter: blur({blur_glass}px);
      }}

      /* כותרת ראשית */
      h1, .stMarkdown h1 {{
        font-weight: 800;
        letter-spacing: 0.2px;
      }}

      /* כרטיסים מזכוכית */
      .glass {{
        background: rgba(255,255,255,0.68);
        border-radius: 18px;
        box-shadow: 0 8px 28px rgba(2, 6, 23, 0.12);
        backdrop-filter: blur({blur_glass}px);
        -webkit-backdrop-filter: blur({blur_glass}px);
        border: 1px solid rgba(15,23,42,0.08);
        padding: 18px 20px;
        margin-bottom: 16px;
      }}

      /* גרסה כהה לכרטיסי זכוכית (לרקעים בהירים מאוד) */
      .glass-dark {{
        background: rgba(12,16,24,0.55);
        color: #fff;
        border-radius: 18px;
        box-shadow: 0 8px 28px rgba(2, 6, 23, 0.12);
        backdrop-filter: blur({blur_glass}px);
        -webkit-backdrop-filter: blur({blur_glass}px);
        border: 1px solid rgba(15,23,42,0.14);
        padding: 18px 20px;
        margin-bottom: 16px;
      }}

      /* כפתורים מודרניים */
      .stButton>button {{
        border-radius: 12px;
        padding: 0.6rem 1rem;
        font-weight: 600;
        box-shadow: 0 6px 18px rgba(22,163,74,0.25);
      }}

      /* טבלאות */
      .stDataFrame tbody tr:hover {{
        background: rgba(22,163,74,0.06);
      }}

      /* תיבות אזהרה/טיפים */
      .stAlert {{ border-radius: 14px !important; }}

      /* Expander עגול/נקי */
      details[data-testid="stExpander"] > summary {{ font-weight: 700; }}
      details[data-testid="stExpander"] {{
        background: rgba(255,255,255,0.76);
        border-radius: 16px;
        padding: 4px 8px;
        border: 1px solid rgba(15,23,42,0.06);
      }}

      /* הילה לטקסט כדי שיבלוט על רקע עמוס */
      .text-halo {{
        text-shadow:
          0 1px 1px rgba(0,0,0,.06),
          0 3px 10px rgba(0,0,0,.14);
      }}
      .text-halo-strong {{
        text-shadow:
          0 1px 1px rgba(0,0,0,.18),
          0 4px 14px rgba(0,0,0,.25);
      }}
    </style>
    """, unsafe_allow_html=True)

def hero_header(
    title: str,
    subtitle: str | None = None,
    logo_path: str | None = None,
    bg_alpha: float = 0.80,   # ← שקיפות תיבת הכותרת (0=שקוף, 1=אטום)
    dark: bool = False,       # ← מצב כהה לתיבה (לרקעים בהירים)
):
    """כותרת Hero; מתעלמת בשקט אם הלוגו לא קיים/לא תקין."""
    col_logo, col_text = st.columns([1, 5], gap="large")
    with col_logo:
        if logo_path and os.path.isfile(logo_path):
            try:
                st.image(logo_path, use_column_width=True)
            except Exception:
                pass
    klass = "glass-dark" if dark else "glass"
    rgba_base = "12,16,24" if dark else "255,255,255"
    st.markdown(f"""
    <div class="{klass}" style="padding:18px 22px; background: rgba({rgba_base},{bg_alpha});">
      <h1 class="text-halo" style="margin:0 0 6px 0;">{title}</h1>
      {"<p class='text-halo' style='font-size:1.05rem; margin:0; opacity:.95;'>"+subtitle+"</p>" if subtitle else ""}
    </div>
    """, unsafe_allow_html=True)

def glass_container(opacity: float = 0.72, dark: bool = False):
    """
    מחזיר context manager לבלוק 'קלף זכוכית' עם שליטה בשקיפות/מצב כהה.
    שימוש:
        with glass_container(opacity=0.8): ...
        with glass_container(opacity=0.6, dark=True): ...
    """
    from contextlib import contextmanager
    @contextmanager
    def _ctx():
        rgba_base = "12,16,24" if dark else "255,255,255"
        text_color = "color:#fff;" if dark else ""
        st.markdown(
            f'<div class="{ "glass-dark" if dark else "glass" }" '
            f'style="background: rgba({rgba_base},{opacity}); {text_color}">', 
            unsafe_allow_html=True
        )
        yield
        st.markdown('</div>', unsafe_allow_html=True)
    return _ctx()
