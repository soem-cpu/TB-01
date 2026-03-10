from utils.helpers import *
import pandas as pd

def run_ssr11(excel_file):

    xls = pd.ExcelFile(excel_file)

    df_screen = xls.parse("TB screening") if "TB screening" in xls.sheet_names else pd.DataFrame()
    df_register = xls.parse("TB register") if "TB register" in xls.sheet_names else pd.DataFrame()
    df_outcome = xls.parse("TB outcome follow up") if "TB outcome follow up" in xls.sheet_names else pd.DataFrame()
    df_dropdown = xls.parse("Dropdown") if "Dropdown" in xls.sheet_names else pd.DataFrame()
    df_service = xls.parse("Service point") if "Service point" in xls.sheet_names else pd.DataFrame()
    df_tpt = xls.parse("TPT register") if "TPT register" in xls.sheet_names else pd.DataFrame()

    # your existing rule engine
    results = check_rules(excel_file)

    return results
