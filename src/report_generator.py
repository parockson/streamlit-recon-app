import io
import pandas as pd

def build_excel_workbook(summaries: dict, reports: dict) -> bytes:
    """Generates an Excel workbook containing all summaries and transaction detail reports."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Summaries
        summaries["s1_total"].to_excel(writer, sheet_name='Total Summary', index=False)
        summaries["s2_match"].to_excel(writer, sheet_name='Match Summary', index=False)
        summaries["s3_xchange_only"].to_excel(writer, sheet_name='Xchange Match Summary', index=False)
        summaries["s4_ext_only"].to_excel(writer, sheet_name='External Summary', index=False)
        
        # Detailed Reports
        reports["match_txns"].to_excel(writer, sheet_name='Match Transactions', index=False)
        reports["xchange_txns"].to_excel(writer, sheet_name='Xchange Match Txns', index=False)
        reports["external_txns"].to_excel(writer, sheet_name='External Match Txns', index=False)

    return output.getvalue()