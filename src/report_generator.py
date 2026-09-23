import io
import re
import pandas as pd

# Same pattern as processor.py — any column name that looks like an ID or phone.
_ID_PHONE_KEYWORDS = re.compile(
    r'(\bid\b|\bids\b|phone|mobile|msisdn|number|account|reference|ref)',
    re.IGNORECASE
)


def _apply_text_format_to_id_cols(df: pd.DataFrame, worksheet, workbook) -> None:
    """Apply xlsxwriter '@' (text) cell format to every ID / phone column.

    Must be called AFTER df.to_excel() so the worksheet and column indices
    are already populated. xlsxwriter set_column with a text format tells
    Excel to treat the column as text, preventing scientific-notation display.
    """
    text_fmt = workbook.add_format({'num_format': '@'})
    for col_idx, col_name in enumerate(df.columns):
        if _ID_PHONE_KEYWORDS.search(str(col_name)):
            # +1 offset because to_excel writes the index in column 0 when
            # index=True; we always use index=False so col_idx maps directly.
            worksheet.set_column(col_idx, col_idx, None, text_fmt)


def build_excel_workbook(summaries: dict, reports: dict) -> bytes:
    """Generates an Excel workbook containing all summaries and transaction detail reports."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book

        def write_sheet(df: pd.DataFrame, sheet_name: str) -> None:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            _apply_text_format_to_id_cols(df, ws, wb)

        # Summaries
        write_sheet(summaries["s1_total"],       'Total Summary')
        write_sheet(summaries["s2_match"],        'Match Summary')
        write_sheet(summaries["s3_xchange_only"], 'Xchange Match Summary')
        write_sheet(summaries["s4_ext_only"],     'External Summary')

        # Detailed Reports
        write_sheet(reports["match_txns"],    'Match Transactions')
        write_sheet(reports["xchange_txns"],  'Xchange Match Txns')
        write_sheet(reports["external_txns"], 'External Match Txns')

    return output.getvalue()


def build_raw_workbook(all_sheets: dict) -> bytes:
    """Re-packages every sheet from the uploaded workbook into a single downloadable Excel file.

    Args:
        all_sheets: Mapping of sheet_name -> DataFrame as returned by
                    pd.read_excel(..., sheet_name=None).

    Returns:
        Raw bytes of the Excel workbook.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book
        for sheet_name, df in all_sheets.items():
            safe_name = sheet_name[:31]          # Excel's 31-char sheet name limit
            df.to_excel(writer, sheet_name=safe_name, index=False)
            ws = writer.sheets[safe_name]
            _apply_text_format_to_id_cols(df, ws, wb)

    return output.getvalue()