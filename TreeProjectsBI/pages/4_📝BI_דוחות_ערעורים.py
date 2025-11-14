# -*- coding: utf-8 -*-
# =========================
# 📝 ערעורים — פילוחים וגרפים (דף עצמאי)
# =========================
# pages/📝BI_דוחות_ערעורים.py

import re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.io as pio

from style_pack import inject_base_css, apply_plotly_theme, hero_header, glass_container

# ---------- עיצוב עמוד ----------
st.set_page_config(page_title="BI – ערעורים", layout="wide")
apply_plotly_theme()
inject_base_css(bg_main="assets/bg_main.jpg", bg_sidebar="assets/bg_sidebar.jpg")
hero_header("📊 BI – ערעורים", "ניתוח ערעורים: פילוחים, הצלחות וטרנדים")

# ---------- יצוא PNG (אופציונלי) ----------
try:
    import kaleido  # noqa
    HAVE_KALEIDO = True
    pio.kaleido.scope.plotlyjs = None
    pio.kaleido.scope.default_scale = 2
except Exception:
    HAVE_KALEIDO = False

PLOTLY_CONFIG = {
    "displaylogo": False,
    "modeBarButtonsToAdd": ["toImage"],
    "toImageButtonOptions": {"format": "png", "scale": 2, "filename": "chart"},
}
def fig_download_png(fig, name: str):
    if not HAVE_KALEIDO:
        return
    try:
        st.download_button(
            f"⬇️ הורד {name} כ־PNG",
            data=fig.to_image(format="png", scale=2),
            file_name=f"{name}.png",
            mime="image/png",
            use_container_width=True,
        )
    except Exception:
        pass

# ---------- עוזרים ----------
def _norm(s: str) -> str:
    if s is None:
        return ""
    return str(s).replace("\u200f","").replace("\u200e","").strip()

def _clean_cat_value(v):
    if pd.isna(v):
        return None
    s = _norm(v)
    if re.fullmatch(r"-?\d+(\.\d+)?", s):
        try:
            f = float(s)
            if f.is_integer():
                s = str(int(f))
        except Exception:
            pass
    return s or None

def normalize_cat_col(series: pd.Series) -> pd.Series:
    return series.map(_clean_cat_value) if series is not None else pd.Series([None])

def parse_any_date(v):
    """ממיר טקסט/מספר (Excel serial)/datetime ל־Timestamp; אחרת NaT."""
    if pd.isna(v):
        return pd.NaT
    if isinstance(v, (int, float, np.integer, np.floating)):
        # Excel serial (מקור 1899-12-30)
        return pd.to_datetime(v, unit="D", origin="1899-12-30", errors="coerce")
    s = _norm(v)
    # קודם dayfirst (כמו 02.01.2024), אם נכשל — נסה הפוך
    dt = pd.to_datetime(s, dayfirst=True, errors="coerce", infer_datetime_format=True)
    if pd.isna(dt):
        dt = pd.to_datetime(s, dayfirst=False, errors="coerce")
    return dt

CITY_STOPWORDS = {
    "רחוב","רח","שדרות","שד","דרך","כיכר","ככר","סמטה","שכונה","שכ",
    "מס","בית","בניין","בנין","דירה","ד","מס'", "מס’"
}
def extract_city(addr):
    """מנסה לחלץ את העיר מהשדה 'ישוב/כתובת' (למשל 'הדף היומי 1 ירושלים' → 'ירושלים')."""
    if pd.isna(addr):
        return None
    s = _norm(addr)
    # הסר מספרים ותווי מפריד בסיסיים, ואחד רווחים
    s = re.sub(r"[0-9\-_,./]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    if not s:
        return None
    toks = [t for t in s.split(" ") if t and t not in CITY_STOPWORDS]
    if not toks:
        return None
    city = toks[-1]
    # נירמולים נפוצים
    city = (city.replace("ת\"א", "תל אביב")
                 .replace("תל-אביב", "תל אביב")
                 .replace("ירושלם", "ירושלים"))
    return city

def detect_header_row(df_headless: pd.DataFrame, scan_rows: int = 12) -> int:
    """
    בוחר את שורת הכותרת הסבירה ביותר מתוך השורות הראשונות.
    נותן ניקוד למילים כמו תאריך/מועד/ישוב/יישוב/עיר/סיבה/החלטה/הערות/מס'.
    """
    kws = {
        "תאריך","מועד","ישוב","יישוב","עיר","כתובת","סיבת","סיבה","החלטת פקיד",
        "פקיד יערות אזורי","אזורי","פקיד יערות ממשלתי","ממשלתי","הערות","מס'","מספר"
    }
    best_row, best_score = 0, -1
    for r in range(min(scan_rows, len(df_headless))):
        row_vals = [_norm(v) for v in df_headless.iloc[r].tolist()]
        non_empty = sum(1 for v in row_vals if v)
        hits = sum(1 for v in row_vals if any(k in v for k in kws))
        score = hits * 5 + non_empty
        if score > best_score:
            best_score, best_row = score, r
    return best_row

def build_col_map(df_cols) -> dict:
    """מיפוי עמודות גמיש → מפתחות פנימיים: date, city, reason, regional, gov, notes, idx."""
    col_map_local = {}
    for c in df_cols:
        n = _norm(c)
        if any(k in n for k in ["מס'", "מספר"]):                 col_map_local.setdefault("idx", c)
        if ("תאריך" in n) or ("מועד" in n):                      col_map_local.setdefault("date", c)
        if any(k in n for k in ["ישוב","יישוב","עיר","כתובת"]): col_map_local.setdefault("city", c)
        if ("סיבת הגשת" in n) or ("סיבה" in n):                  col_map_local.setdefault("reason", c)
        if ("פקיד יערות אזורי" in n) or ("אזורי" in n):          col_map_local.setdefault("regional", c)
        if ("פקיד יערות ממשלתי" in n) or ("ממשלתי" in n):        col_map_local.setdefault("gov", c)
        if "הערות" in n:                                          col_map_local.setdefault("notes", c)
    return col_map_local

# ---------- קליטת קובץ ----------
with glass_container():
    st.markdown("### 📥 העלאת קובץ ערעורים")
    f_appeals = st.file_uploader(
        "forestry_and_trees_ararim2024- דוח ערערים.xlsx",
        type=["xlsx"],
        key="appeals_file",
    )

if f_appeals is None:
    st.info("יש להעלות קובץ XLSX של דוח ערעורים.")
    st.stop()

# ---------- קריאה ומיפוי עמודות (כולל איתור כותרת אם אינה בשורה הראשונה) ----------
try:
    appeals_raw = pd.read_excel(f_appeals, sheet_name=0)
except Exception as e:
    st.error(f"שגיאה בקריאת הקובץ: {e}")
    st.stop()

col_map = build_col_map(appeals_raw.columns)

# אם חסר date/city — נסה איתור שורת כותרת אוטומטי
if not all(k in col_map for k in ("date", "city")):
    appeals_h = pd.read_excel(f_appeals, sheet_name=0, header=None)
    hdr = detect_header_row(appeals_h, scan_rows=12)
    # הגדר כותרות מתוך השורה שנמצאה, וקח את הטבלה מתחת
    new_cols = [_norm(x) if x is not None else "" for x in appeals_h.iloc[hdr].tolist()]
    appeals_h.columns = new_cols
    appeals_h = appeals_h.iloc[hdr + 1:].reset_index(drop=True)
    # הסר כפילויות שמות עמודות
    appeals_h = appeals_h.loc[:, ~appeals_h.columns.duplicated()]
    appeals_h = appeals_h.dropna(how="all")
    appeals_raw = appeals_h
    col_map = build_col_map(appeals_raw.columns)

# בדיקה סופית
missing = [k for k in ("date", "city") if k not in col_map]
if missing:
    st.error(
        "חסרות עמודות חיוניות בקובץ: " + ", ".join(missing)
        + "\nעמודות שנקראו בפועל:\n" + ", ".join(map(str, appeals_raw.columns))
    )
    st.stop()

# ---------- הכנה וטיוב נתונים ----------
ap = appeals_raw.copy()

# תאריך/שנה
ap["תאריך"] = ap[col_map["date"]].map(parse_any_date)
ap["שנה"]   = ap["תאריך"].dt.year.astype("Int64")

# יישוב (מנרמל מתוך ישוב/כתובת)
ap["יישוב_raw"] = ap[col_map["city"]]
ap["יישוב_cat"] = ap["יישוב_raw"].map(extract_city)

# שדות תוכן
ap["סיבת ערעור גולמית"] = ap.get(col_map.get("reason")).astype(str) if "reason" in col_map else ""
ap["החלטה אזורי"]  = ap.get(col_map.get("regional")).astype(str) if "regional" in col_map else ""
ap["החלטה ממשלתי"] = ap.get(col_map.get("gov")).astype(str) if "gov" in col_map else ""
ap["הערות"] = ap.get(col_map.get("notes")).astype(str) if "notes" in col_map else ""

# נירמול סיבת ערעור
def map_reason(s):
    s = str(s)
    if re.search(r"בטיחות", s):     return "בטיחות"
    if re.search(r"בריאות", s):     return "בריאות"
    if re.search(r"מטרד", s):       return "מטרד"
    if re.search(r"בניה|בנייה", s): return "בנייה"
    return "אחר"
ap["סיבת ערעור"] = ap["סיבת ערעור גולמית"].map(map_reason)

# סטטוס ערעור
def map_status(row):
    t = f"{row['החלטה ממשלתי']} {row['הערות']}"
    t = t.replace("\u200f","").replace("\u200e","")
    if re.search(r"נכרתו.*טרם|נכרתו.*דיון|נכרת.*טרם", t): return "לא נדון (כבר נכרת)"
    if re.search(r"ערר\s*התקבל\s*חלקית", t): return "התקבל חלקית"
    if re.search(r"ערר\s*התקבל", t):         return "התקבל"
    if re.search(r"ערר\s*נדחה", t):           return "נדחה"
    # fallback
    if re.search(r"התקבל\s*חלקית", t): return "התקבל חלקית"
    if re.search(r"התקבל", t):         return "התקבל"
    if re.search(r"נדחה", t):           return "נדחה"
    return "לא ידוע"
ap["סטטוס ערעור"] = ap.apply(map_status, axis=1)

# עצים לשימור/שניצלו מתוך טקסט
def extract_saved(t):
    t = str(t)
    m = re.search(r"(\d+)\s*עצ(?:ים)?\s*לשימור", t)
    if m: return int(m.group(1))
    return 1 if re.search(r"\bעץ\s*לשימור\b", t) else 0
ap["עצים לשימור"] = ap["החלטה ממשלתי"].map(extract_saved) + ap["הערות"].map(extract_saved)

# סוג המקור (מה החלטת האזורי)
def src_kind(s):
    s = str(s)
    if re.search(r" דחה .*בקשה |דחה בקשה", s): return "דחיית בקשה"
    if re.search(r" אישר |אישר כרית", s):      return "רישיון שאושר"
    return "אחר"
ap["סוג מקור"] = ap["החלטה אזורי"].map(src_kind)

# ---------- מסננים ----------
with glass_container():
    st.markdown("### 🔎 מסננים")
    c1, c2, c3, c4 = st.columns(4)
    years = sorted([int(y) for y in ap["שנה"].dropna().unique()])
    with c1: f_years  = st.multiselect("שנים", years, default=years or [])
    with c2: f_cities = st.multiselect("יישובים", sorted(ap["יישוב_cat"].dropna().unique()), default=[])
    with c3: f_reason = st.multiselect("סיבת ערעור", ["בנייה","בטיחות","מטרד","בריאות","אחר"], default=[])
    with c4: f_stat   = st.multiselect("סטטוס החלטה", ["התקבל","התקבל חלקית","נדחה","לא נדון (כבר נכרת)","לא ידוע"], default=[])

mask = pd.Series(True, index=ap.index)
if f_years:  mask &= ap["שנה"].isin(f_years)
if f_cities: mask &= ap["יישוב_cat"].isin(f_cities)
if f_reason: mask &= ap["סיבת ערעור"].isin(f_reason)
if f_stat:   mask &= ap["סטטוס ערעור"].isin(f_stat)
apv = ap[mask].copy()

with st.expander("⚙️ Top-N יישובים", expanded=True):
    cN, cEx = st.columns([1,3])
    with cN: topN = st.number_input("N יישובים מוצגים", 1, 50, 10, 1)
    with cEx:
        excl = st.multiselect("החרג יישובים", sorted(apv["יישוב_cat"].dropna().unique()), default=[])
if excl:
    apv = apv[~apv["יישוב_cat"].isin(excl)]

st.markdown("---")

# ---------- KPI ----------
k1,k2,k3,k4 = st.columns(4)
k1.metric("סה\"כ ערעורים (מסונן)", f"{len(apv):,}")
k2.metric("% הצלחה", f"{(apv['סטטוס ערעור'].isin(['התקבל','התקבל חלקית']).mean()*100):.1f}%")
k3.metric("עצים לשימור (סה\"כ)", f"{int(apv['עצים לשימור'].sum()):,}")
k4.metric("# יישובים ייחודיים", f"{apv['יישוב_cat'].nunique(dropna=True):,}")

st.markdown("---")

# ---------- גרפים / שאילתות ----------
# 1) TOP-10 יישובים בכמות ערעורים
g1 = (apv.groupby("יישוב_cat").size()
        .sort_values(ascending=False).head(topN)
        .reset_index(name="ערעורים").rename(columns={"יישוב_cat":"יישוב"}))
fig1 = px.bar(g1, x="יישוב", y="ערעורים", title="TOP-10 יישובים — כמות ערעורים")
st.plotly_chart(fig1, use_container_width=True, config=PLOTLY_CONFIG); fig_download_png(fig1, "appeals_top_cities")

# 2) TOP-10 יישובים — ערעורים שהתקבלו (מלא/חלקית)
acc = apv[apv["סטטוס ערעור"].isin(["התקבל","התקבל חלקית"])]
g2 = (acc.groupby("יישוב_cat").size()
        .sort_values(ascending=False).head(topN)
        .reset_index(name="ערעורים שהתקבלו").rename(columns={"יישוב_cat":"יישוב"}))
fig2 = px.bar(g2, x="יישוב", y="ערעורים שהתקבלו", title="TOP-10 יישובים — ערעורים שהתקבלו (מלא/חלקית)")
st.plotly_chart(fig2, use_container_width=True, config=PLOTLY_CONFIG); fig_download_png(fig2, "appeals_top_cities_accepted")

# 3) TOP-10 יישובים — עצים לשימור/שניצלו
g3 = (apv.groupby("יישוב_cat")["עצים לשימור"].sum()
        .sort_values(ascending=False).head(topN)
        .reset_index().rename(columns={"יישוב_cat":"יישוב"}))
fig3 = px.bar(g3, x="יישוב", y="עצים לשימור", title="TOP-10 יישובים — עצים שניצלו/לשימור")
st.plotly_chart(fig3, use_container_width=True, config=PLOTLY_CONFIG); fig_download_png(fig3, "trees_saved_by_city")

# 4) הערעורים הגדולים שהתקבלו (Top 15 לפי עצים לשימור)
big = (apv[apv["סטטוס ערעור"].isin(["התקבל","התקבל חלקית"])]
       .sort_values("עצים לשימור", ascending=False).head(15)
       [["תאריך","יישוב_cat","סיבת ערעור","סטטוס ערעור","עצים לשימור"]]
       .rename(columns={"יישוב_cat":"יישוב"}))
st.markdown("#### הערעורים הגדולים שהתקבלו (Top 15 לפי עצים לשימור)")
st.dataframe(big, use_container_width=True)

# 5) ערעורים לפי סוג מקור (+ אחוזי הצלחה)
g5c = apv.groupby("סוג מקור").size().reset_index(name="מספר ערעורים")
g5s = (apv.assign(הצלחה=apv["סטטוס ערעור"].isin(["התקבל","התקבל חלקית"]).astype(int))
          .groupby("סוג מקור")["הצלחה"].mean().mul(100).reset_index(name="אחוזי הצלחה"))
g5 = g5c.merge(g5s, on="סוג מקור", how="left")
fig5 = px.bar(g5, x="סוג מקור", y="מספר ערעורים", text="אחוזי הצלחה",
              title="ערעורים לפי סוג מקור", labels={"מספר ערעורים":"כמות"})
st.plotly_chart(fig5, use_container_width=True, config=PLOTLY_CONFIG); fig_download_png(fig5, "appeals_by_source")

# 6) אחוזי הצלחה לפי סיבת ערעור
g6 = (apv.assign(הצלחה=apv["סטטוס ערעור"].isin(["התקבל","התקבל חלקית"]).astype(int))
          .groupby("סיבת ערעור")["הצלחה"].mean().mul(100).reset_index())
fig6 = px.bar(g6.sort_values("הצלחה", ascending=False),
              x="סיבת ערעור", y="הצלחה", title="אחוזי הצלחה לפי סיבת ערעור",
              labels={"הצלחה":"% הצלחה"})
st.plotly_chart(fig6, use_container_width=True, config=PLOTLY_CONFIG); fig_download_png(fig6, "appeals_success_by_reason")

# 7) ערעורים שלא נדונו (כבר נכרתו)
nd = apv[apv["סטטוס ערעור"] == "לא נדון (כבר נכרת)"]
g7 = (nd.groupby("יישוב_cat").size()
        .sort_values(ascending=False).head(topN)
        .reset_index(name="ערעורים שלא נדונו").rename(columns={"יישוב_cat":"יישוב"}))
if not g7.empty:
    fig7 = px.bar(g7, x="יישוב", y="ערעורים שלא נדונו",
                  title="יישובים — ערעורים שלא נדונו (העצים כבר נכרתו)")
    st.plotly_chart(fig7, use_container_width=True, config=PLOTLY_CONFIG); fig_download_png(fig7, "appeals_not_discussed")
else:
    st.info("לא נמצאו ערעורים שלא נדונו (העצים כבר נכרתו) במסננים הנוכחיים.")

# 8) פילוח סטטוס כללי
pie = apv.groupby("סטטוס ערעור").size().reset_index(name="מספר")
fig8 = px.pie(pie, names="סטטוס ערעור", values="מספר", title="פילוח סטטוס ערעורים")
st.plotly_chart(fig8, use_container_width=True, config=PLOTLY_CONFIG); fig_download_png(fig8, "appeals_status_pie")

st.markdown("---")
st.download_button(
    "⬇️ הורד CSV — ערעורים (אחרי מסננים)",
    data=apv.to_csv(index=False).encode("utf-8-sig"),
    file_name="appeals_filtered.csv",
    mime="text/csv",
    use_container_width=True,
)

if not HAVE_KALEIDO:
    st.info("להורדת PNG השתמש/י בכפתור המצלמה המובנה בגרף או התקן/י kaleido.")
