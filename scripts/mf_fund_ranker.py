"""
MF Fund Ranker — DUAL ENGINE Model
====================================
Reads your dashboard_data.xlsx, applies a dual-layer scoring system,
and outputs a ranked Excel file: mf_ranked_screener.xlsx

DUAL ENGINE ARCHITECTURE:
  ENGINE 1 - LONG-TERM QUALITY (Conservative consistency)
    • Filters: 2Y CAGR > 10%, 3Y CAGR > 12%
    • Scores: 1Y (25%), 2Y CAGR (30%), 3Y CAGR (45%)
    • Focus: Structural soundness & proven compounding
    
  ENGINE 2 - SHORT-TERM MOMENTUM (Tactical trend-following)
    • Scores: 6M (30%), 3M (20%), 1Y (25%), 1M (25%)
    • Focus: Current performance & recovery trajectory
    • Trend confirmation: +5 bonus if 6M > 3M > 1M
    
  COMPOSITE SCORE (Blended)
    • 55% Momentum (Engine 2) + 45% Quality (Engine 1)
    • Only qualifies funds passing quality filters
    • Ranked within category via percentile scoring

Usage:
  pip install pandas openpyxl
  python mf_fund_ranker_dual_engine.py

Place dashboard_data.xlsx in the same folder as this script.
Output: mf_ranked_screener.xlsx (with Engine 1 & Engine 2 scores visible)
"""

import pandas as pd
import os
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter

# ── CONFIG ──────────────────────────────────────────────────────────────────
INPUT_FILE  = "dashboard_data.xlsx"
OUTPUT_FILE = "mf_ranked_screener.xlsx"

# ── QUALITY FILTER THRESHOLDS (NEW - Dual Engine Model) ──────────────────
# NOTE: 2Y filter disabled if 2Y return column not available
QUALITY_FILTERS = {
    "cagr_2y_min": 0.10,      # 2Y CAGR must be > 10% (if column exists)
    "cagr_3y_min": 0.12,      # 3Y CAGR must be > 12%
}

# ── ENGINE WEIGHTS (NEW - Dual Engine Model) ──────────────────────────────
# ENGINE 1: Long-Term Quality Scoring (within qualified funds)
ENGINE1_WEIGHTS = {
    "return_1y":  0.25,   # 1-year return
    "return_2y":  0.30,   # 2-year CAGR (converted from 2Y return)
    "return_3y":  0.45,   # 3-year CAGR (long-term emphasis)
}

# ENGINE 2: Short-Term Momentum Scoring
ENGINE2_WEIGHTS = {
    "return_6m":  0.30,   # 6-month momentum
    "return_3m":  0.20,   # 3-month momentum
    "return_1y":  0.25,   # 1-year bridge
    "return_1m":  0.25,   # 1-month latest (if available)
}

# ── COMPOSITE BLEND (NEW - Dual Engine Model) ──────────────────────────
COMPOSITE_BLEND = {
    "engine1_quality": 0.45,   # Long-term quality weight
    "engine2_momentum": 0.55,  # Short-term momentum weight (higher = trend-focused)
}

# ── TREND CONFIRMATION BONUS (NEW - Dual Engine Model) ──────────────────
TREND_BONUS = 5.0  # +5 points if 6M > 3M > 1M (confirmed uptrend)

# ── FILTERS ─────────────────────────────────────────────────────────────────
# Set a value to None to skip that filter entirely.
# Column names must match exactly what is in your Excel file (case-insensitive).
FILTERS = {
    "cat_level_1": "Open Ended Schemes",
    "cat_level_2": "Equity Scheme",
    "plan_type":   "Regular",
    "option_type": "Growth",
}

# ── COLUMN AUTO-DETECTION ───────────────────────────────────────────────────
COLUMN_ALIASES = {
    "scheme_name": [
        "scheme_name", "scheme name", "fund name", "name", "schemename",
        "fund", "scheme"
    ],
    "category": [
        "cat_level_3", "cat_level_2", "cat_level_1",
        "category", "scheme_category", "scheme category", "cat",
        "fund_category", "fund category", "sub_category", "sub category",
        "category_name", "sub_type"
    ],
    "amc": [
        "amc", "amc_name", "amc name", "fund_house", "fund house",
        "amcname", "asset_management_company", "house"
    ],
    "return_1m": [
        "return_30d", "return_1m", "1m_return", "1month_return", "ret_1m",
        "trailing_1m", "1m return", "1 month return"
    ],
    "return_3m": [
        "return_90d", "return_3m", "3m_return", "3month_return", "ret_3m",
        "trailing_3m", "3m return", "3 month return"
    ],
    "return_6m": [
        "return_180d", "return_6m", "6m_return", "6month_return", "ret_6m",
        "trailing_6m", "6m return", "6 month return"
    ],
    "return_1y": [
        "return_365d", "return_1y", "1y_return", "1yr_return", "returns_1y",
        "1y return", "1 year return", "1yr", "ret_1y", "cagr_1y", "1y_cagr",
        "trailing_1y", "1year_return", "ann_return_1y"
    ],
    "return_2y": [
        "return_2y", "2y_return", "2yr_return", "returns_2y",
        "2y return", "2 year return", "2yr", "ret_2y", "cagr_2y", "2y_cagr",
        "trailing_2y", "2year_return", "ann_return_2y"
    ],
    "return_3y": [
        "return_1095d", "return_3y", "3y_return", "3yr_return", "returns_3y",
        "3y return", "3 year return", "3yr", "ret_3y", "cagr_3y", "3y_cagr",
        "trailing_3y", "3year_return", "ann_return_3y"
    ],
    "nav": [
        "nav", "net_asset_value", "current_nav", "latest_nav"
    ],
    "plan": [
        "plan", "plan_type", "scheme_plan", "direct_regular"
    ],
}

COLORS = {
    "header_bg":   "1A3A5C",   # dark navy
    "header_fg":   "FFFFFF",
    "rank1_bg":    "FFD700",   # gold
    "rank2_bg":    "E8E8E8",   # silver
    "rank3_bg":    "D4956A",   # bronze
    "positive":    "1E7A4B",   # dark green
    "negative":    "C0392B",   # red
    "cat_header":  "2D6A8A",   # teal
    "cat_fg":      "FFFFFF",
    "alt_row":     "F5F0E8",   # cream
    "score_bg":    "EBF5FB",
    "border":      "CCCCCC",
    "engine1_bg":  "E3F2FD",   # Light blue (Quality)
    "engine2_bg":  "FFF3E0",   # Light amber (Momentum)
}


def detect_column(df_cols, key):
    """Find the actual column name from aliases."""
    cols_lower = {c.lower().strip(): c for c in df_cols}
    for alias in COLUMN_ALIASES.get(key, []):
        if alias.lower() in cols_lower:
            return cols_lower[alias.lower()]
    return None


def apply_filters(df):
    """
    Apply FILTERS to the DataFrame.
    Each key in FILTERS must match a column name in the data (case-insensitive).
    Values are matched case-insensitively; set a filter value to None to skip it.
    """
    cols_lower = {c.lower().strip(): c for c in df.columns}
    active_filters = {k: v for k, v in FILTERS.items() if v is not None}

    if not active_filters:
        print("   ℹ️  No filters configured — using all rows.")
        return df

    print(f"\n🔎 Applying filters ({len(active_filters)} active):")
    for col_key, value in active_filters.items():
        actual_col = cols_lower.get(col_key.lower().strip())
        if actual_col is None:
            print(f"   ⚠️  Filter column '{col_key}' not found in data — skipping this filter.")
            continue
        before = len(df)
        df = df[df[actual_col].astype(str).str.strip().str.lower() == str(value).strip().lower()]
        print(f"   ✅ {col_key} = '{value}'  →  {before} → {len(df)} rows")

    print(f"   📋 Rows after all filters: {len(df)}")
    return df


def load_data(filepath):
    """Load all sheets, combine, apply filters, and return a unified DataFrame."""
    print(f"\n📂 Reading: {filepath}")
    sheets = pd.read_excel(filepath, sheet_name=None)
    print(f"   Sheets found: {list(sheets.keys())}")

    frames = []
    for name, df in sheets.items():
        df["_source_sheet"] = name
        frames.append(df)

    combined = pd.concat(frames, ignore_index=True)
    print(f"   Total rows across all sheets: {len(combined)}")
    print(f"   Columns: {list(combined.columns)}")

    combined = apply_filters(combined)
    return combined


def map_columns(df):
    """Detect and map standard field names."""
    mapping = {}
    for key in COLUMN_ALIASES:
        col = detect_column(df.columns, key)
        mapping[key] = col
        status = f"✅ '{col}'" if col else "❌ not found"
        print(f"   {key:<15} → {status}")
    return mapping


def to_numeric(series):
    """Clean and convert a series to numeric, coercing errors."""
    if series is None:
        return pd.Series(dtype=float)
    s = series.astype(str).str.replace('%', '').str.replace(',', '').str.strip()
    return pd.to_numeric(s, errors='coerce')


def percentile_score(series):
    """Rank-based percentile score 0–100 within the group."""
    ranks = series.rank(method='min', na_option='bottom')
    return (ranks - 1) / max(len(series) - 1, 1) * 100


# ════════════════════════════════════════════════════════════════════════════
# NEW FUNCTIONS - DUAL ENGINE MODEL
# ════════════════════════════════════════════════════════════════════════════

def convert_2y_to_cagr(return_2y):
    """
    Convert 2-year total return to annualized CAGR.
    Formula: CAGR = (1 + return)^(1/years) - 1
    For 2 years: CAGR = sqrt(1 + return) - 1
    
    Args:
        return_2y: Total 2-year return as decimal (e.g., 0.25 for +25%)
    
    Returns:
        Annualized CAGR, or NaN if input is NaN
    """
    if pd.isna(return_2y):
        return float('nan')
    try:
        # Ensure return_2y is numeric
        r = float(return_2y)
        # CAGR = (1 + r)^(1/2) - 1
        cagr = (1 + r) ** (1/2) - 1
        return cagr
    except (ValueError, TypeError):
        return float('nan')


def score_engine1_quality(df, mapping):
    """
    ENGINE 1: LONG-TERM QUALITY SCORING
    
    Purpose: Identify structurally sound funds with proven 2Y & 3Y performance
    
    Filters (hard gates):
      • 2Y CAGR > 10%
      • 3Y CAGR > 12%
    
    Scoring (within qualified funds):
      • 1Y Return (25%)
      • 2Y CAGR (30%)
      • 3Y CAGR (45%)
    
    Returns: DataFrame with _engine1_score column (0-100 per category)
    """
    print("\n⚙️  ENGINE 1: Long-Term Quality Scoring...")
    
    df = df.copy()
    
    # Extract and convert to numeric
    r1y = to_numeric(df[mapping["return_1y"]] if mapping["return_1y"] else None)
    r2y_raw = to_numeric(df[mapping["return_2y"]] if mapping["return_2y"] else None)
    r3y = to_numeric(df[mapping["return_3y"]] if mapping["return_3y"] else None)
    
    # Convert 2Y return to CAGR
    r2y_cagr = r2y_raw.apply(convert_2y_to_cagr)
    
    df["_r1y"] = r1y
    df["_r2y_cagr"] = r2y_cagr
    df["_r3y"] = r3y
    
    # ── QUALITY FILTERS (Hard Gates) ──────────────────────────────────────
    print(f"   Applying quality filters...")
    before_filter = len(df)
    
    # Filter 1: 2Y CAGR > 10% (only if 2Y column exists)
    if mapping["return_2y"] is not None:
        mask_2y = df["_r2y_cagr"] > QUALITY_FILTERS["cagr_2y_min"]
        print(f"     2Y CAGR > {QUALITY_FILTERS['cagr_2y_min']*100}%: {mask_2y.sum()}/{before_filter} funds qualify")
    else:
        print(f"     2Y CAGR > {QUALITY_FILTERS['cagr_2y_min']*100}%: ⚠️  Column not found (skipping this filter)")
        mask_2y = pd.Series(True, index=df.index)  # Don't filter on missing column
    
    # Filter 2: 3Y CAGR > 12%
    mask_3y = df["_r3y"] > QUALITY_FILTERS["cagr_3y_min"]
    print(f"     3Y CAGR > {QUALITY_FILTERS['cagr_3y_min']*100}%: {mask_3y.sum()}/{before_filter} funds qualify")
    
    # Combined mask
    quality_mask = mask_2y & mask_3y
    df["_qualifies_quality"] = quality_mask
    qualified_count = quality_mask.sum()
    print(f"   ✅ Both filters: {qualified_count}/{before_filter} funds qualify for Engine 1")
    
    # ── SCORING (Percentile within category for qualified funds) ───────────
    cat_col = mapping["category"]
    if not cat_col:
        df["_category_clean"] = "All Funds"
    else:
        df["_category_clean"] = df[cat_col].astype(str).str.strip().str.title()
    
    df["_engine1_score"] = 0.0
    
    # Determine which columns are available
    cols_available = {
        "return_1y":  mapping["return_1y"] is not None,
        "return_2y":  mapping["return_2y"] is not None,
        "return_3y":  mapping["return_3y"] is not None,
    }
    available_count = sum(cols_available.values())
    
    if available_count == 0:
        print(f"   ⚠️  No return columns available for Engine 1 — setting score to 0")
        return df
    
    # Recalculate weights if columns are missing
    weights = ENGINE1_WEIGHTS.copy()
    if not cols_available["return_1y"]:
        weights.pop("return_1y", None)
    if not cols_available["return_2y"]:
        weights.pop("return_2y", None)
    if not cols_available["return_3y"]:
        weights.pop("return_3y", None)
    
    # Normalize weights
    total_weight = sum(weights.values())
    weights = {k: v / total_weight for k, v in weights.items()}
    
    print(f"   Scoring metrics (Engine 1):")
    for metric, w in weights.items():
        print(f"     {metric}: {w*100:.0f}%")
    
    # Score per category
    categories = df["_category_clean"].unique()
    for cat in categories:
        cat_mask = df["_category_clean"] == cat
        cat_qualified = cat_mask & quality_mask
        
        if cat_qualified.sum() == 0:
            continue  # No qualified funds in this category
        
        # Initialize composite score (for qualified funds only)
        composite_scores = pd.Series(0.0, index=df[cat_qualified].index)
        
        # Score each available metric with percentile ranking
        if "return_1y" in weights:
            r1y_scores = percentile_score(df.loc[cat_qualified, "_r1y"])
            composite_scores += r1y_scores * weights["return_1y"]
        
        if "return_2y" in weights:
            r2y_scores = percentile_score(df.loc[cat_qualified, "_r2y_cagr"])
            composite_scores += r2y_scores * weights["return_2y"]
        
        if "return_3y" in weights:
            r3y_scores = percentile_score(df.loc[cat_qualified, "_r3y"])
            composite_scores += r3y_scores * weights["return_3y"]
        
        df.loc[cat_qualified, "_engine1_score"] = composite_scores
    
    return df


def score_engine2_momentum(df, mapping):
    """
    ENGINE 2: SHORT-TERM MOMENTUM SCORING
    
    Purpose: Capture current trend strength and tactical opportunities
    
    Scoring:
      • 6M Return (30%)
      • 3M Return (20%)
      • 1Y Return (25%)
      • 1M Return (25%, optional)
    
    Trend Confirmation Bonus:
      • +5 points if 6M > 3M > 1M (confirmed uptrend)
    
    Returns: DataFrame with _engine2_score and _trend_confirmed columns
    """
    print("\n⚙️  ENGINE 2: Short-Term Momentum Scoring...")
    
    df = df.copy()
    
    # Extract and convert to numeric
    r1m = to_numeric(df[mapping["return_1m"]] if mapping["return_1m"] else None)
    r3m = to_numeric(df[mapping["return_3m"]] if mapping["return_3m"] else None)
    r6m = to_numeric(df[mapping["return_6m"]] if mapping["return_6m"] else None)
    r1y = to_numeric(df[mapping["return_1y"]] if mapping["return_1y"] else None)
    
    df["_r1m"] = r1m
    df["_r3m"] = r3m
    df["_r6m"] = r6m
    df["_r1y_for_momentum"] = r1y
    
    # ── MOMENTUM SCORING ─────────────────────────────────────────────────
    cat_col = mapping["category"]
    if not cat_col:
        df["_category_clean"] = "All Funds"
    else:
        df["_category_clean"] = df[cat_col].astype(str).str.strip().str.title()
    
    df["_engine2_score"] = 0.0
    df["_trend_confirmed"] = False
    df["_trend_bonus"] = 0.0
    
    # Determine which columns are available
    cols_available = {
        "return_1m":  mapping["return_1m"] is not None,
        "return_3m":  mapping["return_3m"] is not None,
        "return_6m":  mapping["return_6m"] is not None,
        "return_1y":  mapping["return_1y"] is not None,
    }
    available_count = sum(cols_available.values())
    
    if available_count == 0:
        print(f"   ⚠️  No return columns available for Engine 2 — setting score to 0")
        return df
    
    # Recalculate weights if columns are missing
    weights = ENGINE2_WEIGHTS.copy()
    if not cols_available["return_1m"]:
        weights.pop("return_1m", None)
    if not cols_available["return_3m"]:
        weights.pop("return_3m", None)
    if not cols_available["return_6m"]:
        weights.pop("return_6m", None)
    if not cols_available["return_1y"]:
        weights.pop("return_1y", None)
    
    # Normalize weights
    total_weight = sum(weights.values())
    weights = {k: v / total_weight for k, v in weights.items()}
    
    print(f"   Scoring metrics (Engine 2 - Momentum):")
    for metric, w in weights.items():
        print(f"     {metric}: {w*100:.0f}%")
    
    # Score per category
    categories = df["_category_clean"].unique()
    for cat in categories:
        cat_mask = df["_category_clean"] == cat
        
        # Initialize composite score
        composite_scores = pd.Series(0.0, index=df[cat_mask].index)
        
        # Score each available metric with percentile ranking
        if "return_6m" in weights:
            r6m_scores = percentile_score(df.loc[cat_mask, "_r6m"])
            composite_scores += r6m_scores * weights["return_6m"]
        
        if "return_3m" in weights:
            r3m_scores = percentile_score(df.loc[cat_mask, "_r3m"])
            composite_scores += r3m_scores * weights["return_3m"]
        
        if "return_1y" in weights:
            r1y_scores = percentile_score(df.loc[cat_mask, "_r1y_for_momentum"])
            composite_scores += r1y_scores * weights["return_1y"]
        
        if "return_1m" in weights:
            r1m_scores = percentile_score(df.loc[cat_mask, "_r1m"])
            composite_scores += r1m_scores * weights["return_1m"]
        
        df.loc[cat_mask, "_engine2_score"] = composite_scores
    
    # ── TREND CONFIRMATION BONUS ──────────────────────────────────────────
    # Check if 6M > 3M > 1M (uptrend confirmation)
    print(f"\n   Calculating trend confirmation bonus (+{TREND_BONUS} points for 6M > 3M > 1M)...")
    
    trend_check = (df["_r6m"] > df["_r3m"]) & (df["_r3m"] > df["_r1m"])
    df["_trend_confirmed"] = trend_check
    df.loc[trend_check, "_trend_bonus"] = TREND_BONUS
    
    confirmed_count = trend_check.sum()
    print(f"   ✅ {confirmed_count} funds confirm uptrend (6M > 3M > 1M)")
    
    # Apply bonus to momentum score
    df["_engine2_score"] += df["_trend_bonus"]
    
    # Cap at 100 (percentile + bonus could exceed 100)
    df["_engine2_score"] = df["_engine2_score"].clip(0, 100)
    
    return df


def score_composite(df):
    """
    COMPOSITE SCORE: Blend both engines with momentum weighting
    
    Formula: Composite = (45% × Engine1_QualityScore) + (55% × Engine2_MomentumScore)
    
    Only applies composite to funds that pass quality filters.
    Non-qualified funds get 0 composite score.
    
    Returns: DataFrame with _composite_score and _rank columns
    """
    print("\n⚙️  COMPOSITE SCORING (Dual Engine Blend)...")
    
    df = df.copy()
    
    # Initialize composite score
    df["_composite_score"] = 0.0
    
    # For qualified funds: blend both engines
    qualified = df["_qualifies_quality"]
    blend_e1 = COMPOSITE_BLEND["engine1_quality"]
    blend_e2 = COMPOSITE_BLEND["engine2_momentum"]
    
    print(f"   Blend ratio: {blend_e1*100:.0f}% Quality (Engine 1) + {blend_e2*100:.0f}% Momentum (Engine 2)")
    print(f"   Quality filter status: {qualified.sum()}/{len(df)} funds qualified")
    
    # Composite = blend of both engines
    composite_formula = (
        df["_engine1_score"] * blend_e1 +
        df["_engine2_score"] * blend_e2
    )
    
    df["_composite_score"] = composite_formula
    
    # Non-qualified funds: score = 0 (or could be penalized)
    df.loc[~qualified, "_composite_score"] = 0
    
    # ── RANKING BY CATEGORY ──────────────────────────────────────────────
    df["_rank"] = 0
    cat_col = "_category_clean"  # Already set in prior functions
    
    for cat in df[cat_col].unique():
        cat_mask = df[cat_col] == cat
        cat_df = df[cat_mask].copy()
        
        # Rank by composite score (descending)
        ranks = cat_df["_composite_score"].rank(method='min', ascending=False)
        df.loc[cat_mask, "_rank"] = ranks
    
    return df


# ════════════════════════════════════════════════════════════════════════════
# END DUAL ENGINE FUNCTIONS
# ════════════════════════════════════════════════════════════════════════════


def score_funds(df, mapping):
    """
    Main scoring orchestrator - calls dual engine model.
    
    Flow:
      1. Engine 1 (Long-Term Quality) - filters + scores
      2. Engine 2 (Short-Term Momentum) - scores + trend bonus
      3. Composite Blend - combines both with momentum weighting
      4. Ranking - within each category
    """
    
    # ENGINE 1: LONG-TERM QUALITY
    df = score_engine1_quality(df, mapping)
    
    # ENGINE 2: SHORT-TERM MOMENTUM (includes trend bonus)
    df = score_engine2_momentum(df, mapping)
    
    # COMPOSITE: Blend both engines
    df = score_composite(df)
    
    return df


def get_col_val(row, col_name):
    """Safely retrieve column value."""
    if not col_name:
        return None
    return row.get(col_name)


def sanitize_sheet_title(title, max_length=31):
    """
    Sanitize sheet title for Excel compatibility.
    Excel sheet names can't contain: [ ] : * ? / \
    Max length: 31 characters
    """
    # Remove invalid characters
    invalid_chars = ['[', ']', ':', '*', '?', '/', '\\']
    for char in invalid_chars:
        title = title.replace(char, '-')
    
    # Limit to 31 characters (Excel constraint)
    if len(title) > max_length:
        title = title[:max_length]
    
    return title.strip()


def pct(val):
    """Format numeric value as percentage string."""
    if pd.isna(val) or val is None:
        return "—"
    return f"{float(val):.2f}%"


def fill(color):
    """Create PatternFill with given hex color."""
    return PatternFill(start_color=color, end_color=color, fill_type="solid")


def score_color(score):
    """Return font color based on score."""
    if score >= 75:    return "006633"  # Dark green
    elif score >= 55:  return "1E7A4B"  # Green
    elif score >= 40:  return "CC7722"  # Orange
    else:              return "C0392B"  # Red


def hdr_font():
    """Header font styling."""
    return Font(name="Arial", bold=True, size=10, color=COLORS["header_fg"])


def cell_font(bold=False, size=9):
    """Cell font styling."""
    return Font(name="Arial", bold=bold, size=size)


def build_excel(df, mapping, output_path):
    """
    Build Excel workbook with dual-engine scoring visible.
    
    NEW: Engine 1 and Engine 2 scores now displayed in separate columns
    """
    from openpyxl import Workbook
    
    wb = Workbook()
    wb.remove(wb.active)
    
    border = Border(
        left=Side(style='thin', color=COLORS["border"]),
        right=Side(style='thin', color=COLORS["border"]),
        top=Side(style='thin', color=COLORS["border"]),
        bottom=Side(style='thin', color=COLORS["border"]),
    )
    
    categories = sorted(df["_category_clean"].unique())
    RETURN_COLS = {4, 5, 6, 7, 8}  # 1M, 3M, 6M, 1Y, 2Y/3Y columns
    
    # ── PER-CATEGORY SHEETS ──────────────────────────────────────────────
    for cat in categories:
        cat_df = df[df["_category_clean"] == cat].sort_values("_rank")
        
        # FIXED: Sanitize sheet title to remove invalid Excel characters
        safe_title = sanitize_sheet_title(cat)
        ws = wb.create_sheet(title=safe_title)
        
        # Row 1 — Category header
        ws.merge_cells("A1:K1")
        ws["A1"] = f"🏆 {cat.upper()}"
        ws["A1"].font = Font(name="Arial", bold=True, size=12, color=COLORS["cat_fg"])
        ws["A1"].fill = fill(COLORS["cat_header"])
        ws["A1"].alignment = Alignment(horizontal="left", indent=1, vertical="center")
        ws.row_dimensions[1].height = 24
        
        # Row 2 — Subtitle
        ws.merge_cells("A2:K2")
        ws["A2"] = (
            f"Dual Engine Model: 45% Quality (LT) + 55% Momentum (ST)  "
            f"|  Total funds: {len(cat_df)}"
        )
        ws["A2"].font = Font(name="Arial", italic=True, size=8, color="555555")
        ws["A2"].fill = fill("F0F4F8")
        ws["A2"].alignment = Alignment(horizontal="left", indent=1)
        ws.row_dimensions[2].height = 16
        
        # Row 3 — Column group labels (NEW - showing both engines)
        for blank_col in [1, 2, 3, 11]:
            ws.cell(row=3, column=blank_col).fill = fill(COLORS["header_bg"])
        
        ws.merge_cells("D3:F3")
        ws["D3"] = "◀  Momentum  ▶"
        ws["D3"].font = Font(name="Arial", bold=True, size=8, color="FFFFFF")
        ws["D3"].fill = fill("E67E22")
        ws["D3"].alignment = Alignment(horizontal="center", vertical="center")
        
        ws.merge_cells("G3:H3")
        ws["G3"] = "◀  Long-Term  ▶"
        ws["G3"].font = Font(name="Arial", bold=True, size=8, color="FFFFFF")
        ws["G3"].fill = fill(COLORS["header_bg"])
        ws["G3"].alignment = Alignment(horizontal="center", vertical="center")
        
        ws.merge_cells("I3:K3")
        ws["I3"] = "◀  Engine Scores  ▶"
        ws["I3"].font = Font(name="Arial", bold=True, size=8, color="FFFFFF")
        ws["I3"].fill = fill("1C1C1C")
        ws["I3"].alignment = Alignment(horizontal="center", vertical="center")
        
        ws.row_dimensions[3].height = 14
        
        # Row 4 — Column headers (NEW - added Engine 1, Engine 2, Composite)
        headers = [
            "Rank", "Scheme Name", "AMC",
            "1M Return", "3M Return", "6M Return",
            "1Y Return", "3Y CAGR",
            "Engine 1 (Quality)", "Engine 2 (Momentum)", "Composite Score"
        ]
        for col_idx, hdr in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col_idx, value=hdr)
            cell.font = hdr_font()
            cell.fill = fill(COLORS["header_bg"])
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
        ws.row_dimensions[4].height = 18
        
        # Data rows — start at row 5
        for i, (_, row) in enumerate(cat_df.iterrows(), 5):
            rank  = row.get("_rank", i - 4)
            name  = get_col_val(row, mapping["scheme_name"]) or "—"
            amc   = get_col_val(row, mapping["amc"])         or "—"
            r1m   = row.get("_r1m")
            r3m   = row.get("_r3m")
            r6m   = row.get("_r6m")
            r1y   = row.get("_r1y")
            r3y   = row.get("_r3y")
            
            # NEW: Engine scores (separate columns now visible)
            engine1 = row.get("_engine1_score", 0)
            engine2 = row.get("_engine2_score", 0)
            composite = row.get("_composite_score", 0)
            
            # Signal based on composite score (simpler threshold as per requirements)
            if composite >= 75:   signal = "⭐ Strong Buy"
            elif composite >= 55: signal = "✅ Buy"
            elif composite >= 40: signal = "⚠️ Watch"
            else:                 signal = "❌ Avoid"
            
            # NOTE: Signal stored for reference but not displayed in this layout
            # (columns are: Rank, Scheme, AMC, 1M, 3M, 6M, 1Y, 3Y, E1, E2, Composite)
            
            values = [
                rank, name, amc,
                pct(r1m), pct(r3m), pct(r6m),
                pct(r1y), pct(r3y),
                round(engine1, 1), round(engine2, 1), round(composite, 1)
            ]
            
            # Determine row background
            if rank == 1:    row_bg = COLORS["rank1_bg"]
            elif rank == 2:  row_bg = COLORS["rank2_bg"]
            elif rank == 3:  row_bg = COLORS["rank3_bg"]
            elif i % 2 == 0: row_bg = COLORS["alt_row"]
            else:             row_bg = "FFFFFF"
            
            for col_idx, val in enumerate(values, 1):
                cell = ws.cell(row=i, column=col_idx, value=val)
                cell.font = cell_font(bold=(rank <= 3), size=9)
                cell.fill = fill(row_bg)
                cell.border = border
                cell.alignment = Alignment(
                    horizontal="center", vertical="center",
                    wrap_text=(col_idx == 2)
                )
                
                # Color returns (columns 4-8)
                if col_idx in RETURN_COLS and isinstance(val, str) and val != "—":
                    num = float(val.replace('%', ''))
                    cell.font = Font(
                        name="Arial", size=9, bold=(rank <= 3),
                        color=COLORS["positive"] if num > 0 else COLORS["negative"]
                    )
                
                # Color engine scores (columns 9-11)
                if col_idx in {9, 10, 11}:  # Engine 1, Engine 2, Composite
                    cell.font = Font(name="Arial", bold=True, size=9, color=score_color(val))
                    # Light background for engine columns
                    if col_idx == 9:
                        cell.fill = fill(COLORS["engine1_bg"])
                    elif col_idx == 10:
                        cell.fill = fill(COLORS["engine2_bg"])
            
            ws.row_dimensions[i].height = 16
        
        # Column widths (adjusted for new columns)
        widths = [6, 52, 20, 11, 11, 11, 11, 11, 16, 16, 14]
        for ci, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(ci)].width = w
        
        ws.freeze_panes = "A5"
        print(f"   ✅ {cat} — {len(cat_df)} funds ranked")
    
    # ── Summary Sheet ────────────────────────────────────────────────────
    ws_sum = wb.create_sheet(title="🏆 SUMMARY", index=0)
    
    ws_sum.merge_cells("A1:K1")
    ws_sum["A1"] = "MF INTELLIGENCE — DUAL ENGINE TOP FUND PER CATEGORY"
    ws_sum["A1"].font = Font(name="Arial", bold=True, size=14, color="FFFFFF")
    ws_sum["A1"].fill = fill("0D1117")
    ws_sum["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 30
    
    ws_sum.merge_cells("A2:K2")
    ws_sum["A2"] = (
        "Dual Engine: Long-Term Quality (45% weight, filters 2Y CAGR>10% & 3Y CAGR>12%) "
        "+ Short-Term Momentum (55% weight, 6M/3M/1M returns)  |  Category: cat_level_3"
    )
    ws_sum["A2"].font = Font(name="Arial", italic=True, size=8, color="555555")
    ws_sum["A2"].fill = fill("F5F5F5")
    ws_sum["A2"].alignment = Alignment(horizontal="center")
    ws_sum.row_dimensions[2].height = 16
    
    sum_headers = [
        "Category", "#1 Fund", "AMC",
        "6M Return", "3M Return", "1M Return",
        "1Y Return", "3Y CAGR",
        "Engine 1 (Quality)", "Engine 2 (Momentum)", "Signal"
    ]
    for ci, hdr in enumerate(sum_headers, 1):
        c = ws_sum.cell(row=3, column=ci, value=hdr)
        c.font = hdr_font()
        c.fill = fill(COLORS["header_bg"])
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
    ws_sum.row_dimensions[3].height = 18
    
    SUM_RETURN_COLS = {4, 5, 6, 7, 8}
    SUM_ENGINE_COLS  = {9, 10}
    
    for i, cat in enumerate(categories, 4):
        cat_df = df[df["_category_clean"] == cat].sort_values("_rank")
        if cat_df.empty: 
            continue
        
        top   = cat_df.iloc[0]
        name  = get_col_val(top, mapping["scheme_name"]) or "—"
        amc   = get_col_val(top, mapping["amc"])         or "—"
        r1m   = top.get("_r1m")
        r3m   = top.get("_r3m")
        r6m   = top.get("_r6m")
        r1y   = top.get("_r1y")
        r3y   = top.get("_r3y")
        
        # NEW: Engine scores
        engine1 = top.get("_engine1_score", 0)
        engine2 = top.get("_engine2_score", 0)
        composite = top.get("_composite_score", 0)
        
        signal = "⭐ Strong Buy" if composite >= 75 else "✅ Buy" if composite >= 55 else "⚠️ Watch"
        
        row_bg = COLORS["alt_row"] if i % 2 == 0 else "FFFFFF"
        vals = [cat, name, amc, pct(r6m), pct(r3m), pct(r1m), pct(r1y), pct(r3y), 
                round(engine1, 1), round(engine2, 1), signal]
        
        for ci, v in enumerate(vals, 1):
            c = ws_sum.cell(row=i, column=ci, value=v)
            c.font = cell_font(size=9)
            c.fill = fill(row_bg)
            c.border = border
            c.alignment = Alignment(
                horizontal="center" if ci != 2 else "left",
                vertical="center", wrap_text=(ci == 2)
            )
            
            # Color returns
            if ci in SUM_RETURN_COLS and isinstance(v, str) and v != "—":
                num = float(v.replace('%', ''))
                c.font = Font(
                    name="Arial", size=9,
                    color=COLORS["positive"] if num > 0 else COLORS["negative"]
                )
            
            # Color engine scores
            if ci in SUM_ENGINE_COLS:
                c.font = Font(name="Arial", bold=True, size=9, color=score_color(v))
                if ci == 9:
                    c.fill = fill(COLORS["engine1_bg"])
                elif ci == 10:
                    c.fill = fill(COLORS["engine2_bg"])
        
        ws_sum.row_dimensions[i].height = 18
    
    # Summary column widths
    for ci, w in enumerate([28, 52, 20, 11, 11, 11, 11, 11, 16, 16, 14], 1):
        ws_sum.column_dimensions[get_column_letter(ci)].width = w
    
    ws_sum.freeze_panes = "A4"
    
    wb.save(output_path)
    print(f"\n✅ Output saved → {output_path}")


def main():
    print("\n" + "="*80)
    print("  MF FUND RANKER — DUAL ENGINE MODEL")
    print("  Long-Term Quality (45%) + Short-Term Momentum (55%)")
    print("="*80)
    
    if not os.path.exists(INPUT_FILE):
        print(f"\n❌ File not found: {INPUT_FILE}")
        print("   Place dashboard_data.xlsx in the same folder as this script and re-run.")
        return
    
    df = load_data(INPUT_FILE)
    
    if df.empty:
        print("\n❌ No rows remain after applying filters. Check your FILTERS config.")
        return
    
    print("\n🔍 Auto-detecting columns...")
    mapping = map_columns(df)
    
    missing_critical = [k for k in ("scheme_name", "return_1y", "return_3y") if not mapping[k]]
    if missing_critical:
        print(f"\n⚠️  Critical columns not found: {missing_critical}")
        print("   Available columns:", list(df.columns))
        print("   → Update COLUMN_ALIASES at the top of this script to match your column names.")
        return
    
    print("\n⚙️  Scoring funds with dual engine model...")
    df_scored = score_funds(df, mapping)
    
    print("\n📝 Building Excel output...")
    build_excel(df_scored, mapping, OUTPUT_FILE)
    
    print("\n" + "="*80)
    print(f"  📊 MF Ranked Screener → {OUTPUT_FILE}")
    print(f"\n  📁 What's inside:")
    print(f"     • 🏆 SUMMARY tab — Top fund per category at a glance")
    print(f"     • One tab per fund category — All schemes ranked by composite score")
    print(f"\n  🎯 Scoring Model:")
    print(f"     • Engine 1 (Quality): 2Y CAGR > 10%, 3Y CAGR > 12% [1Y/2Y/3Y returns]")
    print(f"     • Engine 2 (Momentum): 6M/3M/1M trend strength [+5 bonus if 6M>3M>1M]")
    print(f"     • Composite: 45% Quality + 55% Momentum (category-wise percentile)")
    print(f"\n  📈 Signals:")
    print(f"     • ⭐ Strong Buy:  Composite Score ≥ 75")
    print(f"     • ✅ Buy:         Composite Score ≥ 55")
    print(f"     • ⚠️  Watch:      Composite Score ≥ 40")
    print(f"     • ❌ Avoid:       Composite Score < 40 (or fails quality filters)")
    print("="*80 + "\n")


if __name__ == "__main__":
    main()
