import pandas as pd
import re

# Keywords that flag a column as an ID / phone number column.
_ID_PHONE_KEYWORDS = re.compile(
    r'(\bid\b|\bids\b|phone|mobile|msisdn|number|account|reference|ref)',
    re.IGNORECASE
)


def coerce_id_phone_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Convert ID / phone-number-like columns to clean strings.

    Prevents pandas from storing them as float64, which causes scientific-
    notation display (e.g. 2.56789e+11) in both Streamlit dataframes and
    the exported Excel workbooks.

    Rules applied per matching column:
    - NaN values become an empty string.
    - Values that look like whole numbers (e.g. 254712345678.0) have the
      trailing '.0' stripped so they read as '254712345678'.
    """
    df = df.copy()
    for col in df.columns:
        if _ID_PHONE_KEYWORDS.search(str(col)):
            def _clean(v):
                if pd.isna(v):
                    return ''
                s = str(v)
                # strip trailing .0 that pandas adds when int-stored-as-float
                if s.endswith('.0') and s[:-2].lstrip('-').isdigit():
                    s = s[:-2]
                return s
            df[col] = df[col].apply(_clean)
    return df


def process_recon_data(raw_df: pd.DataFrame):
    """Cleans data, applies filters, and computes summary metrics and report subsets."""
    df = coerce_id_phone_columns(raw_df).copy()

    # Standardize string flags
    df['In Xchange'] = df['In Xchange'].astype(str).str.upper().str.strip()
    df['In External'] = df['In External'].astype(str).str.upper().str.strip()

    # Subsets
    ex_df = df[df['In Xchange'] == 'Y']
    ext_df = df[df['In External'] == 'Y']
    match_df = df[(df['In Xchange'] == 'Y') & (df['In External'] == 'Y')]
    xchange_only_df = df[(df['In Xchange'] == 'Y') & (df['In External'] == 'N')]
    external_only_df = df[(df['In Xchange'] == 'N') & (df['In External'] == 'Y')]

    # 1 & 2. Dual-side Summaries
    s1_total = get_dual_summary(ex_df, ext_df)
    s2_match = get_dual_summary(match_df, match_df)

    # 3 & 4. Single-side Summaries
    s3_xchange_only = get_single_summary(xchange_only_df, 'Xchange Status', 'Xchange Amount')
    s4_ext_only = get_single_summary(external_only_df, 'External Status', 'External Amount')

    summaries = {
        "s1_total": s1_total,
        "s2_match": s2_match,
        "s3_xchange_only": s3_xchange_only,
        "s4_ext_only": s4_ext_only
    }

    # Detailed Transaction Reports
    target_cols = ['Date Created', 'Transaction ID', 'Phone', 'Xchange Status', 'External Status', 'Xchange Amount', 'External Amount']
    col_rename = {
        'Date Created': 'Date',
        'Transaction ID': 'Id',
        'Phone': 'phone',
        'Xchange Status': 'exchange status',
        'External Status': 'external status',
        'Xchange Amount': 'xchange amt',
        'External Amount': 'external amt'
    }

    reports = {
        "match_txns": match_df[target_cols].rename(columns=col_rename),
        "xchange_txns": xchange_only_df[target_cols].rename(columns=col_rename),
        "external_txns": external_only_df[target_cols].rename(columns=col_rename)
    }

    # Dynamic filename generator
    recon_filename = generate_recon_filename(df)

    # MUST RETURN 3 VALUES
    return summaries, reports, recon_filename


def get_dual_summary(df_ex: pd.DataFrame, df_ext: pd.DataFrame) -> pd.DataFrame:
    statuses = ['success', 'failed', 'pending']
    records = []

    # Total Row
    ex_tot_vol = len(df_ex)
    ex_tot_val = df_ex['Xchange Amount'].sum() if ex_tot_vol > 0 else 0.0
    ext_tot_vol = len(df_ext)
    ext_tot_val = df_ext['External Amount'].sum() if ext_tot_vol > 0 else 0.0

    records.append({
        'status': 'Total',
        'exchange Volume': ex_tot_vol,
        'exchange Value': ex_tot_val,
        'external Volume': ext_tot_vol,
        'external Value': ext_tot_val
    })

    # Status Breakdown Rows
    for st in statuses:
        sub_ex = df_ex[df_ex['Xchange Status'].astype(str).str.lower() == st]
        sub_ext = df_ext[df_ext['External Status'].astype(str).str.lower() == st]

        records.append({
            'status': st.capitalize(),
            'exchange Volume': len(sub_ex),
            'exchange Value': sub_ex['Xchange Amount'].sum() if len(sub_ex) > 0 else 0.0,
            'external Volume': len(sub_ext),
            'external Value': sub_ext['External Amount'].sum() if len(sub_ext) > 0 else 0.0
        })

    return pd.DataFrame(records)


def get_single_summary(df: pd.DataFrame, status_col: str, amount_col: str) -> pd.DataFrame:
    total_trans = len(df)
    total_amt = df[amount_col].sum() if total_trans > 0 else 0.0

    def stats_for(status_val):
        sub = df[df[status_col].astype(str).str.lower() == status_val]
        return len(sub), sub[amount_col].sum() if len(sub) > 0 else 0.0

    s_trans, s_amt = stats_for('success')
    f_trans, f_amt = stats_for('failed')
    p_trans, p_amt = stats_for('pending')

    return pd.DataFrame([
        {"Category": "Total", "Total Trans": total_trans, "Total Amount": total_amt},
        {"Category": "Success", "Total Trans": s_trans, "Total Amount": s_amt},
        {"Category": "Failed", "Total Trans": f_trans, "Total Amount": f_amt},
        {"Category": "Pending", "Total Trans": p_trans, "Total Amount": p_amt}
    ])


def generate_recon_filename(df: pd.DataFrame) -> str:
    unit_str = "RECON"
    if 'Unit' in df.columns:
        units = df['Unit'].dropna().unique()
        if len(units) > 0:
            unit_str = str(units[0]).replace(" ", "_")

    if 'Date Created' in df.columns:
        parsed_dates = pd.to_datetime(df['Date Created'], errors='coerce')
        min_date = parsed_dates.min()
        max_date = parsed_dates.max()

        if pd.notnull(min_date) and pd.notnull(max_date):
            d_start = min_date.strftime('%d-%m-%Y')
            d_end = max_date.strftime('%d-%m-%Y')
            return f"{unit_str}_{d_start}_to_{d_end}_recon.xlsx"

    return f"{unit_str}_recon.xlsx"