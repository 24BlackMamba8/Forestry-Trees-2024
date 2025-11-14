# -*- coding: utf-8 -*-
# =========================
# 🌳 BI – דוחות כריתה והעתקה
# =========================
# pages/🌳BI_דוחות_כריתה.py

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.io as pio

from style_pack import inject_base_css, apply_plotly_theme, hero_header, glass_container

# ---------- הגדרות עמוד ----------
st.set_page_config(page_title="BI – דוחות כריתה", layout="wide")
apply_plotly_theme()
inject_base_css(bg_main="assets/bg_main.jpg", bg_sidebar="assets/bg_sidebar.jpg")
hero_header("🌳 BI – דוחות כריתה והעתקה",
            "ניתוח דוחות הכריתה המאוחדים: כריתות, העתקות, יישובים ומיני עצים")

# ---------- תמיכה בייצוא PNG (אופציונלי) ----------
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
        # אם kaleido לא מותקן/לא עובד – פשוט לא נציג את כפתור ההורדה
        pass


# ---------- עוזרים כלליים ----------

def _norm(s):
    if s is None or pd.isna(s):
        return ""
    return str(s).replace("\u200f", "").replace("\u200e", "").strip()

def ensure_bool(series: pd.Series) -> pd.Series:
    """ממיר עמודות בוליאניות ('TRUE'/'FALSE'/0/1 וכו') ל־bool אמיתי."""
    if series.dtype == bool:
        return series.fillna(False)
    s = series.astype(str).str.strip().str.upper()
    return s.map({"TRUE": True, "FALSE": False, "1": True, "0": False}).fillna(False)

def pick_first_existing(df: pd.DataFrame, candidates) -> str | None:
    """מחזיר את שם העמודה הראשון שקיים מתוך רשימה."""
    for c in candidates:
        if c in df.columns:
            return c
    return None

def safe_to_datetime(s: pd.Series) -> pd.Series:
    """המרת תאריכים (כולל serial של אקסל) בצורה עמידה."""
    dt = pd.to_datetime(s, errors="coerce", infer_datetime_format=True, dayfirst=True)
    need = dt.isna()
    if need.any():
        numeric = pd.to_numeric(s[need], errors="coerce")
        good = numeric[numeric.notna()]
        if not good.empty:
            dt.loc[good.index] = pd.to_datetime(
                good,
                unit="D",
                origin="1899-12-30",
                errors="coerce",
            )
    return dt

# ---------- טעינת קובץ הכריתות ----------

with glass_container():
    st.markdown("### 📥 העלאת קובץ כריתות מאוחד")
    st.caption("קובץ לדוגמה: **merged_forest_reports_FINAL_dates_fixed.xlsx** (גליון 'Merged').")
    f_main = st.file_uploader(
        "בחר/י קובץ Excel של דוחות הכריתה המאוחדים",
        type=["xlsx"],
        key="cuts_file",
    )

if f_main is None:
    st.info("יש להעלות קובץ XLSX של דוחות כריתה מאוחדים (הקובץ שיצרנו במיזוג).")
    st.stop()

# ננסה קודם את הגיליון בשם 'Merged', ואם אין – את הגיליון הראשון
try:
    try:
        df_raw = pd.read_excel(f_main, sheet_name="Merged")
    except Exception:
        df_raw = pd.read_excel(f_main, sheet_name=0)
except Exception as e:
    st.error(f"שגיאה בקריאת הקובץ: {e}")
    st.stop()

df = df_raw.copy()

# ---------- עיבוד בסיסי של הדאטה ----------

# שם יישוב
city_col = pick_first_existing(df, ["יישוב", "ישוב", "עיר"])
if city_col is None:
    st.error("לא נמצאה עמודת יישוב ('יישוב' / 'ישוב' / 'עיר') בקובץ.")
    st.write("עמודות שנקראו:", list(df.columns))
    st.stop()
df["יישוב_cat"] = df[city_col].map(lambda x: _norm(x) or "לא ידוע")

# שם מין עץ
tree_col = pick_first_existing(df, ["שם   מין עץ", "שם מין עץ", "מין עץ"])
if tree_col is None:
    # אם אין – ניצור עמודה ריקה כדי שהקוד ישאר אחיד
    df["שם מין עץ (BI)"] = ""
    tree_col_bi = "שם מין עץ (BI)"
else:
    tree_col_bi = "שם מין עץ (BI)"
    df[tree_col_bi] = df[tree_col].map(_norm)

# מספר עצים – אם אין עמודה מתאימה, נניח 1 לכל רשומה
count_col = pick_first_existing(df, ["מספר עצים", "מספר   עצים", "כמות עצים"])
if count_col is None:
    df["מספר עצים (BI)"] = 1
else:
    df["מספר עצים (BI)"] = (
        pd.to_numeric(df[count_col], errors="coerce")
          .fillna(1)
          .clip(lower=1)
    )

# תאריכים: נשתמש ב"מ-תאריך" כבסיס; אם אין – "עד-תאריך"
date_col = pick_first_existing(df, ["מ-תאריך", "מתאריך", "תאריך", "עד-תאריך"])
if date_col:
    df["תאריך"] = safe_to_datetime(df[date_col])
    df["שנה"]   = df["תאריך"].dt.year.astype("Int64")
else:
    df["תאריך"] = pd.NaT
    df["שנה"]   = pd.NA

# פעולה כריתה / העתקה – מתוך הדגלים שהוספנו במיזוג
cut_col  = pick_first_existing(df, ["__is_cut__", "is_cut"])
move_col = pick_first_existing(df, ["__is_move__", "is_move"])

if cut_col:
    df["__is_cut__"] = ensure_bool(df[cut_col])
else:
    df["__is_cut__"] = False

if move_col:
    df["__is_move__"] = ensure_bool(df[move_col])
else:
    df["__is_move__"] = False

# פעולה/סיבה מפוענחות – אם קיימות
action_text_col = pick_first_existing(df, ["פעולה_מפוענחת", "פעולה"])
reason_text_col = pick_first_existing(df, ["סיבה_מפוענחת", "סיבה", "סיבה  מילולית"])

df["פעולה BI"] = np.select(
    [df["__is_cut__"], df["__is_move__"]],
    ["כריתה", "העתקה/שימור"],
    default=df[action_text_col].astype(str) if action_text_col else "לא ידוע",
)

if reason_text_col:
    df["סיבה BI"] = df[reason_text_col].astype(str).apply(_norm)
else:
    df["סיבה BI"] = ""

# ---------- מסננים ----------

with glass_container():
    st.markdown("### 🔎 מסננים")
    c1, c2, c3, c4 = st.columns(4)

    years = sorted([int(y) for y in df["שנה"].dropna().unique()])
    cities_all = sorted(df["יישוב_cat"].dropna().unique())
    trees_all  = sorted(df[tree_col_bi].dropna().unique())
    actions_all = ["כריתה", "העתקה/שימור"]

    with c1:
        f_years = st.multiselect("שנים", years, default=years or [])
    with c2:
        f_cities = st.multiselect("יישובים", cities_all, default=[])
    with c3:
        f_trees = st.multiselect("מיני עצים", trees_all, default=[])
    with c4:
        f_actions = st.multiselect("סוג פעולה", actions_all, default=actions_all)

    st.markdown("---")
    cN, cEx = st.columns([1, 3])
    with cN:
        topN = st.number_input("N ל־TOP", 1, 50, 10, 1)
    with cEx:
        excl_cities = st.multiselect("החרג יישובים מ־TOP", cities_all, default=[])

mask = pd.Series(True, index=df.index)
if f_years:
    mask &= df["שנה"].isin(f_years)
if f_cities:
    mask &= df["יישוב_cat"].isin(f_cities)
if f_trees:
    mask &= df[tree_col_bi].isin(f_trees)
if f_actions:
    mask &= df["פעולה BI"].isin(f_actions)

dfv = df[mask].copy()
if excl_cities:
    dfv = dfv[~dfv["יישוב_cat"].isin(excl_cities)]

st.markdown("---")

# ---------- KPI מרכזיים ----------

total_trees     = int(dfv["מספר עצים (BI)"].sum())
total_cuts      = int(dfv.loc[dfv["__is_cut__"], "מספר עצים (BI)"].sum())
total_moves     = int(dfv.loc[dfv["__is_move__"], "מספר עצים (BI)"].sum())
total_records   = len(dfv)
cut_ratio       = (total_cuts / total_trees * 100) if total_trees else 0
move_ratio      = (total_moves / total_trees * 100) if total_trees else 0
unique_cities   = dfv["יישוב_cat"].nunique(dropna=True)
unique_trees    = dfv[tree_col_bi].nunique(dropna=True)

k1, k2, k3, k4 = st.columns(4)
k1.metric("סה\"כ עצים בדוחות (מסונן)", f"{total_trees:,}")
k2.metric("עצים שנכרתו", f"{total_cuts:,}", f"{cut_ratio:.1f}%")
k3.metric("עצים שהועתקו/לשימור", f"{total_moves:,}", f"{move_ratio:.1f}%")
k4.metric("יישובים / מיני עצים", f"{unique_cities} / {unique_trees}")

st.markdown("---")

# ---------- גרפים עיקריים ----------

# 1) TOP-N יישובים – עצים שנכרתו
cut_df = dfv[dfv["__is_cut__"]].copy()
g1 = (
    cut_df.groupby("יישוב_cat")["מספר עצים (BI)"]
          .sum()
          .sort_values(ascending=False)
          .head(topN)
          .reset_index()
          .rename(columns={"יישוב_cat": "יישוב", "מספר עצים (BI)": "עצים שנכרתו"})
)
fig1 = px.bar(
    g1,
    x="יישוב",
    y="עצים שנכרתו",
    title=f"TOP-{len(g1)} יישובים – עצים שנכרתו",
)
st.plotly_chart(fig1, use_container_width=True, config=PLOTLY_CONFIG)
fig_download_png(fig1, "top_cities_cuts")

# 2) TOP-N מיני עצים שנכרתו
g2 = (
    cut_df.groupby(tree_col_bi)["מספר עצים (BI)"]
          .sum()
          .sort_values(ascending=False)
          .head(topN)
          .reset_index()
          .rename(columns={tree_col_bi: "מין עץ", "מספר עצים (BI)": "עצים שנכרתו"})
)
fig2 = px.bar(
    g2,
    x="מין עץ",
    y="עצים שנכרתו",
    title=f"TOP-{len(g2)} מיני עצים שנכרתו",
)
st.plotly_chart(fig2, use_container_width=True, config=PLOTLY_CONFIG)
fig_download_png(fig2, "top_tree_species_cuts")

# 3) TOP-N יישובים – עצים שהועתקו
move_df = dfv[dfv["__is_move__"]].copy()
g3 = (
    move_df.groupby("יישוב_cat")["מספר עצים (BI)"]
           .sum()
           .sort_values(ascending=False)
           .head(topN)
           .reset_index()
           .rename(columns={"יישוב_cat": "יישוב", "מספר עצים (BI)": "עצים שהועתקו"})
)
fig3 = px.bar(
    g3,
    x="יישוב",
    y="עצים שהועתקו",
    title=f"TOP-{len(g3)} יישובים – עצים שהועתקו/שומרו",
)
st.plotly_chart(fig3, use_container_width=True, config=PLOTLY_CONFIG)
fig_download_png(fig3, "top_cities_moves")

# 4) פילוח סיבות כריתה (רק שורות כריתה)
if not cut_df.empty:
    g4 = (
        cut_df.groupby("סיבה BI")["מספר עצים (BI)"]
              .sum()
              .reset_index()
              .rename(columns={"מספר עצים (BI)": "עצים"})
    )
    fig4 = px.bar(
        g4.sort_values("עצים", ascending=False),
        x="סיבה BI",
        y="עצים",
        title="עצים שנכרתו לפי סיבה",
        labels={"סיבה BI": "סיבה"},
    )
    st.plotly_chart(fig4, use_container_width=True, config=PLOTLY_CONFIG)
    fig_download_png(fig4, "cut_reasons")
else:
    st.info("אין נתוני כריתה במסננים הנוכחיים להצגת פילוח סיבות.")

# 5) מגמת כריתות/העתקות לפי שנה
if dfv["שנה"].notna().any():
    trend = (
        dfv.assign(שנה=dfv["שנה"].astype("Int64"))
           .groupby(["שנה", "פעולה BI"])["מספר עצים (BI)"]
           .sum()
           .reset_index()
    )
    fig5 = px.line(
        trend.sort_values("שנה"),
        x="שנה",
        y="מספר עצים (BI)",
        color="פעולה BI",
        markers=True,
        title="מגמת עצים שנכרתו/הועתקו לפי שנים",
        labels={"מספר עצים (BI)": "מספר עצים"},
    )
    st.plotly_chart(fig5, use_container_width=True, config=PLOTLY_CONFIG)
    fig_download_png(fig5, "trend_by_year")
else:
    st.info("לא נמצאו תאריכים תקינים ליצירת מגמת שנים.")

# 6) פילוח כריתה מול העתקה (Pie)
sum_by_action = (
    dfv.groupby("פעולה BI")["מספר עצים (BI)"]
       .sum()
       .reset_index()
       .rename(columns={"מספר עצים (BI)": "עצים"})
)
fig6 = px.pie(
    sum_by_action,
    names="פעולה BI",
    values="עצים",
    title="פילוח עצים – כריתה מול העתקה/שימור",
)
st.plotly_chart(fig6, use_container_width=True, config=PLOTLY_CONFIG)
fig_download_png(fig6, "cut_vs_move_pie")

st.markdown("---")

# 7) הרישיונות הגדולים (Top 20 לפי מספר עצים)
top_licenses = (
    dfv.sort_values("מספר עצים (BI)", ascending=False)
       .head(20)
       .copy()
)
# נשאיר רק עמודות שימושיות להצגה
cols_for_table = []
for c in ["אזור", "מספר רישיון", city_col, tree_col_bi,
          "מספר עצים (BI)", "פעולה BI", "סיבה BI", "תאריך", "__source_sheet__"]:
    if c in top_licenses.columns:
        cols_for_table.append(c)

st.markdown("#### הרישיונות הגדולים (Top 20 לפי מספר עצים)")
if cols_for_table:
    st.dataframe(top_licenses[cols_for_table], use_container_width=True)
else:
    st.dataframe(top_licenses, use_container_width=True)

st.markdown("---")

# 8) הורדת הדאטה המסונן
st.download_button(
    "⬇️ הורד CSV — דוחות כריתה (אחרי מסננים)",
    data=dfv.to_csv(index=False).encode("utf-8-sig"),
    file_name="forest_cuts_filtered.csv",
    mime="text/csv",
    use_container_width=True,
)

if not HAVE_KALEIDO:
    st.info("להורדת PNG אפשר להשתמש בכפתור המצלמה המובנה בגרף, או להתקין את חבילת kaleido.")
