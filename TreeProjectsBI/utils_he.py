# -*- coding: utf-8 -*-
from __future__ import annotations
import re
import pandas as pd

# ---------- ניקוי ונירמול בסיסי ----------
def clean_text(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    # הסרת סימני כיוון RTL/LTR
    s = s.replace("\u200f", "").replace("\u200e", "")
    return s

def norm(s: str) -> str:
    s = clean_text(s).lower()
    s = (s.replace("״", '"').replace("”", '"').replace("„", '"')
           .replace("’", "'").replace("׳", "'").replace("“", '"'))
    s = s.replace('קק"ל', "קקל").replace('קק\"ל', "קקל")
    s = re.sub(r"[^א-תa-z0-9\/\-\s\._]", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def norm_key(x):
    """מפתח אחיד לקודים (48.0 -> '48')."""
    s = clean_text(x)
    try:
        if s and re.fullmatch(r"-?\d+(\.\d+)?", s):
            f = float(s)
            if f.is_integer():
                return str(int(f))
    except Exception:
        pass
    return s

# ---------- תבנית יעד ומיפוי כותרות ----------
TARGET_COLS = [
    "אזור","מספר רישיון","פעולה","שם בעל הרישיו","סיבה","סיבה  מילולית",
    "יישוב","רחוב","מס'","גוש","חלקה","מ-תאריך","עד-תאריך",
    "שם   מאשר הרישיון","שם   מין עץ","מספר עצים","פעולה (2)","הערות","__source_sheet__"
]

ALIASES = {
    # אזור
    "ezor":"אזור","data1":"אזור",
    # מספר רישיון
    "מספר רישיון":"מספר רישיון","מספר רשיון":"מספר רישיון","license":"מספר רישיון",
    # פעולה
    "פעולה":"פעולה","peula":"פעולה","action":"פעולה",
    # שם בעל הרישיון / מבקש
    "data1name":"שם בעל הרישיו","data1_printname":"שם בעל הרישיו","מבקש":"שם בעל הרישיו",
    # סיבות
    "סיבה":"סיבה","siba":"סיבה","reason":"סיבה",
    "סיבה  מילולית":"סיבה  מילולית","sibatext":"סיבה  מילולית",
    # יישוב/עיר
    "יישוב":"יישוב","city":"יישוב","cityname":"יישוב",
    # רחוב
    "רחוב":"רחוב","street":"רחוב","רחוב ומספר בית":"רחוב",
    # מספר בית
    "מס":"מס'","מספר בית":"מס'","homenumber":"מס'","homenun":"מס'","homenun ":"מס'",
    # גוש/חלקה
    "גוש":"גוש","gush":"גוש","חלקה":"חלקה","helka":"חלקה",
    # תאריכים
    "מ-תאריך":"מ-תאריך","fromdate":"מ-תאריך","exedate":"מ-תאריך","data1_exedate":"מ-תאריך","date":"מ-תאריך",
    "עד-תאריך":"עד-תאריך","todate":"עד-תאריך",
    # מאשר רישיון
    "שם   מאשר הרישיון":"שם   מאשר הרישיון","rname":"שם   מאשר הרישיון","data2_okname":"שם   מאשר הרישיון",
    "תפקיד ושם המאשר":"שם   מאשר הרישיון",
    # מין עץ וכמות
    "שם   מין עץ":"שם   מין עץ","שם מין עץ":"שם   מין עץ","מין העץ":"שם   מין עץ","tree":"שם   מין עץ",
    "מספר עצים":"מספר עצים","quant":"מספר עצים","quantity":"מספר עצים","כמות עצים":"מספר עצים","כמות":"מספר עצים",
    # הערות
    "הערות":"הערות"
}

def map_col(col: str):
    return ALIASES.get(norm(col))

def detect_header_row(df_headless: pd.DataFrame, scan_rows: int = 8):
    """בחר את שורת הכותרת הסבירה ביותר מתוך השורות הראשונות."""
    kws = {
        "date","fromdate","todate","city","cityname","יישוב","רחוב","מס","homenumber","homenun","gush","helka",
        "data1","data1name","data1_exedate","rname","data2_okname","data2_oknamejob",
        "siba","sibatext","tree","quant","quantity","כמות עצים","מין העץ","סיבה","פעולה","שם בעל הרישיו","מספר רישיון",
        "ezor","אזור","גוש","חלקה","שם מאשר הרישיון","שם   מאשר הרישיון","שם מין עץ","שם   מין עץ"
    }
    kws = {norm(k) for k in kws}
    best_row, best_score = 0, -1
    for r in range(min(scan_rows, len(df_headless))):
        row_vals = [clean_text(v) for v in df_headless.iloc[r].tolist()]
        non_empty = sum(1 for v in row_vals if v != "")
        keyword_hits = sum(1 for v in row_vals if norm(v) in kws)
        score = keyword_hits*5 + non_empty
        if score > best_score:
            best_score, best_row = score, r
    return best_row

# ---------- קריאת לוקאפ עצים/יישובים ----------
def build_city_tree_luts(xls: pd.ExcelFile):
    """מוציא מילוני קוד→שם מהגיליונות 'רשימת ערים לפי קודים' ו'רשימת עצים לפי קודים' גם אם יש 'כותרות עמוקות'."""
    # ערים
    cities_raw = pd.read_excel(xls, sheet_name="רשימת ערים לפי קודים", header=None)
    cities = cities_raw.iloc[3:].reset_index(drop=True)
    if cities.shape[1] >= 3:
        cities.columns = ["rownum","שם יישוב","קוד"] + [f"extra_{i}" for i in range(cities.shape[1]-3)]
    else:
        cities.columns = [f"c{i}" for i in range(cities.shape[1])]
    city_lut = {}
    if "קוד" in cities.columns and "שם יישוב" in cities.columns:
        for _, r in cities.dropna(how="all").iterrows():
            code = norm_key(r["קוד"])
            name = clean_text(r["שם יישוב"])
            if code and name:
                city_lut[code] = name

    # עצים
    trees_raw = pd.read_excel(xls, sheet_name="רשימת עצים לפי קודים", header=None)
    trees = trees_raw.iloc[3:].reset_index(drop=True)
    if trees.shape[1] >= 5:
        trees.columns = ["קוד","סוג עץ","שם עץ","שם באנגלית","growth_form"] + [f"extra_{i}" for i in range(trees.shape[1]-5)]
    else:
        trees.columns = [f"t{i}" for i in range(trees.shape[1])]
    tree_lut = {}
    if "קוד" in trees.columns and "שם עץ" in trees.columns:
        for _, r in trees.dropna(how="all").iterrows():
            code = norm_key(r["קוד"])
            name = clean_text(r["שם עץ"])
            if code and name:
                tree_lut[code] = name

    return city_lut, tree_lut

# ---------- המרת קודים לשמות ----------
def apply_lookups(merged_df: pd.DataFrame, city_lut: dict, tree_lut: dict):
    COL_CITY = "יישוב"; COL_TREE = "שם   מין עץ"
    def replace_with_lut(val, lut):
        key = norm_key(val)
        return lut.get(key, val)
    if COL_CITY in merged_df.columns:
        merged_df[COL_CITY] = merged_df[COL_CITY].apply(lambda v: replace_with_lut(v, city_lut))
    if COL_TREE in merged_df.columns:
        merged_df[COL_TREE] = merged_df[COL_TREE].apply(lambda v: replace_with_lut(v, tree_lut))
    return merged_df

# ---------- תאריכים/מספרים ----------
def parse_dates_inplace(df: pd.DataFrame):
    for c in ["מ-תאריך", "עד-תאריך"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=False, infer_datetime_format=True)

def ensure_numeric(df: pd.DataFrame, cols):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df
def action_flags(df: pd.DataFrame):
    """
    מזהה דגלים לכריתה/העתקה מתוך העמודות 'פעולה' ו'פעולה (2)'.
    מחזיר שני Series בוליאניים: (cut_flag, move_flag) באורך df.
    """
    def is_cut(x):
        s = clean_text(x)
        # "כריתה", "כרות", וגם גיבוי באנגלית אם מופיע
        return ("כרית" in s) or ("כרות" in s) or ("cut" in s.lower())

    def is_move(x):
        s = clean_text(x)
        # "העתקה", "שימור", וגם גיבוי באנגלית אם מופיע
        return ("העתק" in s) or ("שימור" in s) or ("move" in s.lower()) or ("transplant" in s.lower()) or ("preserv" in s.lower())

    idx = df.index
    a1 = df.get("פעולה", pd.Series([None]*len(idx), index=idx)).apply(is_cut)
    a2 = df.get("פעולה (2)", pd.Series([None]*len(idx), index=idx)).apply(is_cut)
    cut_flag = (a1 | a2).fillna(False)

    b1 = df.get("פעולה", pd.Series([None]*len(idx), index=idx)).apply(is_move)
    b2 = df.get("פעולה (2)", pd.Series([None]*len(idx), index=idx)).apply(is_move)
    move_flag = (b1 | b2).fillna(False)

    return cut_flag, move_flag
