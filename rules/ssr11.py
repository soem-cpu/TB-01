import pandas as pd
from utils.helpers import *
from utils.dropdown_loader import *
from utils.validation import combine_errors


# =================================================
# TB Screening Rules
# =================================================

def add_channel_category(df):
    """
    Add 'Channel_category' column based on Channel value.
    """
    if df is None or df.empty or "Channel" not in df.columns:
        df["Channel_category"] = ""
        return df

    def categorize(v):
        s = norm_lower(v)

        if s in {"inpatient", "opd"}:
            return "Fixed Clinic"
        if s == "mobile":
            return "Mobile Team"
        if s in {"ichv", "volunteer"}:
            return "Volunteer"
        return ""

    df["Channel_category"] = df["Channel"].apply(categorize)
    return df


def process_tb_screening(df, df_dropdown, df_service, df_register, df_tpt):
    df = df.copy()

    dropdowns = load_dropdowns(df_dropdown)
    service_pairs = load_service_pairs(df_service)
    level_map = load_basecode_level(df_service)

    # Dropdown checks
    df = apply_dropdown_checks_tb_screening(df, df_dropdown)

    # Screening date
    sd_col = find_column(df, "Screening date")
    if sd_col:
        df["Screening_date_Error"] = df[sd_col].apply(
            lambda v: "" if is_valid_date(v) else "Invalid screening date"
        )

    # Age
    age_col = find_column(df, "Age")
    if age_col:
        df["Age_Error"] = df[age_col].apply(
            lambda v: "" if is_numeric(v) else "Invalid age"
        )
    else:
        df["Age_Error"] = "Age column missing"


    # Gender
    def gender_check(row):
        sex = norm_lower(row.get("Gender"))
        name = row.get("Name/Nickname", "")
        prefix = extract_prefix(name)

        if sex not in {s.lower() for s in dropdowns.get("Gender", set())}:
            return "Invalid gender"

        if prefix in {"Ma", "Daw"} and sex not in {"female", "f"}:
            return "Gender mismatch with name"
        if prefix in {"Mg", "U", "Ko"} and sex not in {"male", "m"}:
            return "Gender mismatch with name"
        return ""

    df["Gender_Error"] = df.apply(gender_check, axis=1)

    # Base Code
    df["Base_Code_Error"] = df.apply(
        lambda r: ""
        if (normalize(r.get("Township")), normalize(r.get("Base Code"))) in service_pairs
        else "Invalid Base Code / Township",
        axis=1
    )

    # ICHV Code (conditional)
    df["ICHV_Code_Error"] = df.apply(
        lambda r: (
            "" if normalize(r.get("ICHV Code")) == ""
            else "ICHV Code should be blank for this channel"
        ) if norm_lower(r.get("Channel")) not in {"volunteer", "ichv"} else (
            "" if (normalize(r.get("Township")), normalize(r.get("ICHV Code"))) in service_pairs
            else "Invalid ICHV Code / Township"
        ),
        axis=1
    )

    # Level
    df["Level"] = df["Base Code"].apply(lambda x: level_map.get(normalize(x), ""))

    # Registration No - check if exists in TB register or TPT register
    if "Registration No" in df.columns:
        valid_regs = set(normalize(x) for x in df_register["Township TB No"]) if not df_register.empty and "Township TB No" in df_register.columns else set()
        valid_tpt = set(normalize(x) for x in df_tpt["Code"]) if not df_tpt.empty and "Code" in df_tpt.columns else set()
        df["Registration_No_Error"] = df["Registration No"].apply(
            lambda v: "" if normalize(v) in valid_regs or normalize(v) in valid_tpt else "Not found in TB register or TPT register"
        )

    # Presumptive TB
    exam_cols = ["Microscopy Result", "TrueNat Result", "Gene Xpert Result", "CXR Result"]

    def is_presumptive(r):
        # any lab result present
        if any(normalize(r.get(c)) for c in exam_cols):
            return 1

        s = norm_lower(r.get("TB Screening Results"))

        CLINICAL_TOKENS = ["clinically diagnosed", "clinically dx", "clinically dx tb", "clinically diagnosed tb"]
        BACT_TOKENS = [
            "bact confirmed tb",
            "bact: confirmed tb",
            "bact: confirmed",
            "bact confirmed",
            "bacteriological confirmed",
        ]

        if any(tok in s for tok in CLINICAL_TOKENS):
            return 1
        if any(tok in s for tok in BACT_TOKENS):
            return 1

        return 0

    df["Presumptive_TB_referred"] = df.apply(is_presumptive, axis=1)

    # ---------------- CXR recheck (show in Comment only) ----------------
    cxr_col = find_column(df, "CXR Result")
    if cxr_col:
        BAD_CXR = {"-", "not done", "nd", "pending"}
        err_col = "CXR_Result_Error"
        df[err_col] = df[cxr_col].apply(
            lambda v: "to recheck" if normalize(v) != "" and norm_lower(v) in BAD_CXR else ""
        )

    # TB detected
    TB_POS = {
        "bact confirmed tb", "bact confirmed", "bact: confirmed tb", "bact: confirmed", "bacteriological confirmed", "bc", "2",
        "clinically diagnosed tb", "clinically diagnosed", "clinically dx tb", "1",
        "tb", "positive"
    }

    df["TB_detected"] = df.apply(
        lambda r: 1 if r["Presumptive_TB_referred"] == 1
        and norm_lower(r.get("TB Screening Results")) in TB_POS
        else 0,
        axis=1
    )

    # Bact confirmed
    df["Bact_confirmed_TB"] = df.apply(
        lambda r: 1 if r["Presumptive_TB_referred"] == 1
        and norm_lower(r.get("TB Screening Results")) in {"bact confirmed tb", "bact confirmed", "bact: confirmed tb", "bact: confirmed", "bacteriological confirmed", "bc", "2", "positive"}
        else 0,
        axis=1
    )

    # Result consistency
    def result_check(row):
        sputum = norm_lower(row.get("Microscopy Result"))
        gene = norm_lower(row.get("Gene Xpert Result"))
        truenat = norm_lower(row.get("TrueNat Result"))
        result = norm_text(row.get("TB Screening Results"))

        # ---- sputum positive ----
        sputum_positive = any(
            k in sputum
            for k in ["positive", "scanty", "1+", "2+", "3+"]
        )

        # ---- GeneXpert positive ----
        gene_positive = any(
            k in gene
            for k in [
                "t", "tt", "ti", "rr",
                "mtb detected","mtb detected rr not detected","mtb detected trace","mtb/detected,rif/not detected",
                "rif detected","mtb detected rr detected","mtb detected, rif detected",
                "rif indeterminate","mtb detected rr indeterminate","mtb/detected low, rif/indeterminate"
            ]
        )

        # ---- Truenat positive ----
        truenat_positive = any(
            k in truenat
            for k in ["vt", "rr", "ti", "mtb detected","mtb detected rr not detected","mtb detected trace","mtb/detected,rif/not detected",
                "rif detected","mtb detected rr detected","mtb detected, rif detected",
                "rif indeterminate","mtb detected rr indeterminate","mtb/detected low, rif/indeterminate"]
        )

        positive_exam = sputum_positive or gene_positive or truenat_positive

        # ✅ ACCEPT ALL BACT CONFIRMED VARIANTS
        if positive_exam and result not in {
            "bact confirmed tb",
            "bact confirmed",
            "bacteriological confirmed",
            "bc"
        }:
            return "F"

        return "T"




    df["TB_Screening_Results_check"] = df.apply(result_check, axis=1)

    # Duplicate
    def ordinal(n):
        if 10 <= n % 100 <= 20:
            suffix = 'th'
        else:
            suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
        return str(n) + suffix

    df["Duplicate_check"] = ""
    subset_cols = ["Base Code", "Name/Nickname", "Age", "Gender", "Screening date"]
    for name, group in df.groupby(subset_cols):
        if len(group) > 1:
            # Sort group by index to maintain original order
            group = group.sort_index()
            for i, idx in enumerate(group.index):
                df.loc[idx, "Duplicate_check"] = f"duplicated({ordinal(i+1)})"


    # Ongoing TB
    if not df_register.empty and "Registration No" in df.columns:
        enroll_map = dict(zip(
            df_register["Township TB No"],
            pd.to_datetime(df_register["Enrolled date"], errors="coerce", dayfirst=True)
        ))
        df["Ongoing_TB_case_check"] = df.apply(
            lambda r: "Ongoing TB case"
            if normalize(r.get("Registration No")) in enroll_map
            and pd.to_datetime(r.get("Screening date"), errors="coerce", dayfirst=True)
            > enroll_map.get(normalize(r.get("Registration No")))
            else "",
            axis=1
        )
    else:
        df["Ongoing_TB_case_check"] = ""

    # ---------------- Channel category ----------------
    df = add_channel_category(df)

    # -------- Township TB No_TB register (match by Township, Patient Name, Age, Gender) --------
    df["Township TB No_TB register"] = ""

    if not df_register.empty:
        # Build a lookup map from TB register keyed by (Township, Patient Name, Age, Gender)
        reg_lookup = {}
        for _, reg_row in df_register.iterrows():
            reg_township = normalize(reg_row.get("Township", ""))
            reg_name = normalize(reg_row.get("Patient Name", ""))
            reg_age = normalize(str(reg_row.get("Age", "")))
            reg_gender = normalize(reg_row.get("Gender", ""))
            reg_tbno = normalize(reg_row.get("Township TB No", ""))
            
            key = (reg_township, reg_name, reg_age, reg_gender)
            reg_lookup[key] = reg_tbno

        def match_tb_register(row):
            # Check if TB Screening Results contains clinically dx tb or bact confirmed tb
            screening_result = norm_lower(row.get("TB Screening Results", ""))
            tb_indicators = {"clinically dx tb", "bact: confirmed tb"}
            
            if not any(indicator in screening_result for indicator in tb_indicators):
                return ""
            
            # Try to match against TB register
            scr_township = normalize(row.get("Township", ""))
            scr_name = normalize(row.get("Name/Nickname", ""))
            scr_age = normalize(str(row.get("Age", "")))
            scr_gender = normalize(row.get("Gender", ""))
            
            key = (scr_township, scr_name, scr_age, scr_gender)
            return reg_lookup.get(key, "")

        df["Township TB No_TB register"] = df.apply(match_tb_register, axis=1)

    return df


# =================================================
# TB REGISTER RULES
# =================================================

def process_tb_register(df, df_dropdown, df_service, df_screen):
    df = df.copy()

    dropdowns = load_dropdowns(df_dropdown)
    service_pairs = load_service_pairs(df_service)
    level_map = load_basecode_level(df_service)

    # ---------------- Dropdown checks ----------------
    dropdown_rules = [
        ("Channel", "Channel"),
        ("TYPE of patient", "TYPE of patient"),
        ("Type of disease", "Type of disease"),
        ("Treatment regimen", "Treatment regimen"),
        ("DM status", "DM status"),
        ("HIV status", "HIV status"),

        # ✅ LAB RESULTS (FIXED)
        ("Microscopy Result", "Microscopy Result"),
        ("TrueNat Result", "TrueNat Result"),
        ("Gene Xpert Result", "Gene Xpert Result"),  # optional

        ("Treatment outcome", "Treatment outcome"),
    ]


    for label, var in dropdown_rules:
        col = find_column(df, label)
        if not col:
            continue
        allowed = dropdowns.get(var, set())
        df[f"{label.replace(' ', '_')}_Error"] = df[col].apply(
            lambda v: "" if normalize(v) in allowed else "Invalid value"
        )

    # ---------------- Enrolled date ----------------
    enroll_col = find_column(df, "Enrolled date")
    if enroll_col:
        df["Enrolled_date_Error"] = df[enroll_col].apply(
            lambda v: "" if is_valid_date(v) else "Invalid enrolled date"
        )

    # ---------------- Started date ----------------
    start_col = find_column(df, "Started date")
    refer_ext_col = find_column(df, "Refer/ Transfer out (External)")

    if start_col:
        def started_date_check(row):
            # 🔹 If externally referred out → skip check
            if refer_ext_col and normalize(row.get(refer_ext_col)) != "":
                return ""

            # 🔹 Otherwise validate started date
            return "" if is_valid_date(row.get(start_col)) else "Invalid started date"

        df["Started_date_Error"] = df.apply(started_date_check, axis=1)


    # ---------------- Outcome date ----------------
    outcome_col = find_column(df, "Treatment outcome")
    outcomedate_col = find_column(df, "Outcome date")

    df["Outcome_Date_Error"] = ""

    if outcome_col and outcomedate_col:
        df["Outcome_Date_Error"] = df.apply(
            lambda r: ""
            if normalize(r.get(outcome_col)) == ""
            else (
                ""
                if is_valid_date(r.get(outcomedate_col))
                else "Invalid or missing Outcome date"
            ),
            axis=1
        )


    # ---------------- Age ----------------
    age_col = find_column(df, "Age")
    if age_col:
        df["Age_Error"] = df[age_col].apply(
            lambda v: "" if is_numeric(v) else "Invalid age"
        )
    else:
        df["Age_Error"] = "Age column missing"


    # ---------------- Gender ----------------
    def gender_check(row):
        sex = norm_lower(row.get("Gender"))
        name = row.get("Patient Name", "")
        prefix = extract_prefix(name)

        if sex not in {s.lower() for s in dropdowns.get("Gender", set())}:
            return "Invalid gender"

        if prefix in {"Ma", "Daw"} and sex not in {"female", "f"}:
            return "Gender mismatch with name"
        if prefix in {"Mg", "U", "Ko"} and sex not in {"male", "m"}:
            return "Gender mismatch with name"
        return ""

    df["Gender_Error"] = df.apply(gender_check, axis=1)

    # ---------------- Regimen_check ----------------
    df["Regimen_check"] = ""

    regimen_col = find_column(df, "Treatment regimen")
    type_col = find_column(df, "TYPE of patient")
    age_col = find_column(df, "Age")

    def regimen_check(row):
        regimen = norm_lower(row.get(regimen_col))
        patient_type = norm_lower(row.get(type_col))
        age = row.get(age_col)

        # IR → must be New
        if regimen == "ir" and patient_type != "new":
            return "Invalid"

        # RR → must NOT be New
        if regimen == "rr" and patient_type == "new":
            return "Invalid"

        # CR → Age < 15
        if regimen == "cr":
            try:
                if age == "" or float(age) >= 15:
                    return "Invalid"
            except Exception:
                return "Invalid"

        return ""

    df["Regimen_check"] = df.apply(regimen_check, axis=1)

   
    # ---------------- Base Code ----------------
    df["Base_Code_Error"] = df.apply(
        lambda r: ""
        if (normalize(r.get("Township")), normalize(r.get("Base Code"))) in service_pairs
        else "Invalid Base Code / Township",
        axis=1
    )

    # ---------------- Level ----------------
    df["Level"] = df["Base Code"].apply(lambda x: level_map.get(normalize(x), ""))

    # ---------------- Township TB No+ blank) ----------------
    tbno_col = find_column(df, "Township TB No")
    if tbno_col:
        df["Township_TB_No_Error"] = df[tbno_col].apply(
            lambda v: "Blank Township TB No" if normalize(v) == "" else ""
        )

        def ordinal(n):
            if 10 <= n % 100 <= 20:
                suffix = 'th'
            else:
                suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
            return str(n) + suffix

        screening_date_col = find_column(df, "Screening date")
        if screening_date_col:
            df[screening_date_col] = pd.to_datetime(df[screening_date_col], errors="coerce", dayfirst=True)
            for tbno, group in df[df.duplicated(subset=[tbno_col], keep=False)].groupby(tbno_col):
                group = group.sort_values(by=screening_date_col)
                for i, idx in enumerate(group.index):
                    df.loc[idx, "Duplicate_check"] = f"duplicated({ordinal(i+1)})"
        else:
            dup = df.duplicated(subset=[tbno_col], keep=False)
            df.loc[dup, "Duplicate_check"] = "To recheck for duplication"

    # ---------------- Channel_category_screening ----------------
    if not df_screen.empty:
        scr_tbno = find_column(df_screen, "Registration No")
        scr_channel = find_column(df_screen, "Channel_category")

        if scr_tbno and scr_channel:
            channel_map = dict(zip(
                df_screen[scr_tbno].astype(str),
                df_screen[scr_channel]
            ))
            df["Channel_category_screening"] = df[tbno_col].astype(str).map(channel_map).fillna("")


    # ---------------- BC_screening ----------------
    if not df_screen.empty:
        scr_tbno = find_column(df_screen, "Registration No")
        scr_bc = find_column(df_screen, "Bact_confirmed_TB")

        if scr_tbno and scr_bc:
            bc_map = dict(zip(
                df_screen[scr_tbno].astype(str),
                df_screen[scr_bc]
            ))
            df["BC_screening"] = (
                df[tbno_col].astype(str)
                .map(bc_map)
                .fillna(0)
                .astype("Int64")
            )

    # ---------------- BC_results (from lab results) ----------------
    # compute here so it can be placed next to BC_screening
    bc_results_series = df.apply(BC_results, axis=1).astype("Int64")
    if "BC_screening" in df.columns:
        pos = df.columns.get_loc("BC_screening") + 1
        df.insert(pos, "BC_results", bc_results_series)
    else:
        df["BC_results"] = bc_results_series

    # ---------------- TB_detected_screening ----------------
    if not df_screen.empty:
        scr_tbno = find_column(df_screen, "Registration No")
        scr_tbdet = find_column(df_screen, "TB_detected")

        if scr_tbno and scr_tbdet:
            tbdet_map = dict(zip(
                df_screen[scr_tbno].astype(str),
                df_screen[scr_tbdet]
            ))
            df["TB_detected_screening"] = (
                df[tbno_col].astype(str)
                .map(tbdet_map)
                .fillna(0)
                .astype("Int64")
            )

    # ---------------- TB Quarter Function ----------------
    def get_tb_quarter(date):
        if pd.isna(date):
            return None

        month = date.month
        day = date.day

        if (month == 3 and day >= 21) or month in [4, 5] or (month == 6 and day <= 20):
            return "Q1"
        elif (month == 6 and day >= 21) or month in [7, 8] or (month == 9 and day <= 20):
            return "Q2"
        elif (month == 9 and day >= 21) or month in [10, 11] or (month == 12 and day <= 20):
            return "Q3"
        else:
            return "Q4"


    # ---------------- TBDT_1 ----------------
    type_col = find_column(df, "TYPE of patient")
    tb_no_col = find_column(df, "Township TB No")
    enroll_col = find_column(df, "Enrolled date")

    if type_col and tb_no_col and enroll_col:

        df[enroll_col] = pd.to_datetime(df[enroll_col], errors="coerce", dayfirst=True)

        # ---- Exclude Rif detected ----
        truenat_col = find_column(df, "TrueNat Result")
        genexpert_col = find_column(df, "Gene Xpert Result")

        def rif_excluded(row):
            truenat = norm_lower(row[truenat_col]) if truenat_col else ""
            genexpert = norm_lower(row[genexpert_col]) if genexpert_col else ""

        def rif_excluded(row):
            truenat = norm_lower(row.get(truenat_col)) if truenat_col else ""
            genexpert = norm_lower(row.get(genexpert_col)) if genexpert_col else ""

            rif_keywords = ["rif detected", "rr"]

            if any(k in truenat for k in rif_keywords):
                return 1
            if any(k in genexpert for k in rif_keywords):
                return 1

            return 0

        df["DRTB"] = df.apply(rif_excluded, axis=1)

        # Only eligible types AND not Rif detected
        df["_eligible"] = df.apply(
            lambda r: 1
            if norm_lower(r[type_col]) in {"new", "relapse"}
            and r["DRTB"] == 0
            else 0,
            axis=1
        )

        # Add TB quarter
        df["_tb_quarter"] = df[enroll_col].apply(get_tb_quarter)

        df["TBDT_1"] = ""

        eligible_df = df[(df["_eligible"] == 1) & (df[tb_no_col].notna())]

        for tb_no, group in eligible_df.groupby(tb_no_col):

            quarters = group["_tb_quarter"].unique()

            group = group.sort_values(enroll_col)

            quarters = group["_tb_quarter"].dropna().unique()

            if len(quarters) == 1:
                idx = group[enroll_col].idxmax()
                df.loc[idx, "TBDT_1"] = 1

            # DIFFERENT QUARTER → count earliest
            else:
                idx = group[enroll_col].idxmin()
                df.loc[idx, "TBDT_1"] = 1

        # clean helper columns
        df.drop(columns=["_eligible", "_tb_quarter"], inplace=True)


    # ---------------- TBDT_3c ----------------
    df["TBDT_3c"] = ""

    channel_col = find_column(df, "Channel")
    cat_col = find_column(df, "Channel_category_screening")

    if "TBDT_1" in df.columns:
        df["TBDT_3c"] = df.apply(
            lambda r: 1
            if r.get("TBDT_1") == 1
            and (
                norm_lower(r.get(channel_col)) in {"volunteer", "ichv"}
                or (cat_col and "volunteer" in norm_lower(r.get(cat_col)))
            )
            else "",
            axis=1
        )


    # ---------------- TBHIV_5 ----------------
    hiv_col = find_column(df, "HIV status")
    if hiv_col:
        df["TBHIV_5"] = df.apply(
            lambda r: 1
            if r.get("TBDT_1") == 1 and norm_lower(r.get(hiv_col)) in {"positive", "negative"}
            else "",
            axis=1
        )

    # ---------------- TBO2a_N ----------------
    df["TBO2a_N"] = ""

    outcome_col = find_column(df, "Treatment outcome")

    VALID_OUTCOMES = {
        "cure",
        "cured",
        "complete",
        "completed",
        "treatment completed"
    }

    if outcome_col and "TBO2a_D" in df.columns:
        df["TBO2a_N"] = df.apply(
            lambda r: 1
            if r.get("TBO2a_D") == 1
            and norm_lower(r.get(outcome_col)) in VALID_OUTCOMES
            else "",
            axis=1
        )


    # ---------------- BC (from lab results) ----------------
    # moved earlier next to BC_screening
    pass



    # ---------------- Tin/Refer_from_check (FINAL FIX) ----------------
    df["Tin_Refer_from"] = ""

    reg_tbno = find_column(df, "Township TB No")
    scr_tbno = find_column(df_screen, "Registration No")

    if reg_tbno and scr_tbno and not df_screen.empty:
        screening_tbnos = set(
            normalize(str(x)) for x in df_screen[scr_tbno].astype(str)
        )

        df["Tin_Refer_from"] = df[reg_tbno].astype(str).apply(
            lambda v: "Yes" if normalize(v) not in screening_tbnos else ""
        )



    return df


# =================================================
# TB OUTCOME FOLLOW-UP RULES
# =================================================

def process_tb_outcome_follow_up(df, df_dropdown, df_service, df_register):
    df = df.copy()

    dropdowns = load_dropdowns(df_dropdown)
    service_pairs = load_service_pairs(df_service)
    level_map = load_basecode_level(df_service)

    # ---------------- Dropdown checks ----------------
    dropdown_rules = [
        ("Channel", "Channel"),
        ("TYPE of patient", "TYPE of patient"),
        ("Type of disease", "Type of disease"),
        ("Treatment regimen", "Treatment regimen"),
        ("DM status", "DM status"),
        ("HIV status", "HIV status"),
        ("Treatment outcome", "Treatment outcome"),
        ("Microscopy Result_0", "Microscopy Result"),
        ("TrueNat Result_0", "TrueNat Result"),
        ("Gene Xpert Result_0", "Gene Xpert Result"),
        ("Microscopy Result_5", "Microscopy Result"),
        ("TrueNat Result_5", "TrueNat Result"),
        ("Gene Xpert Result_5", "Gene Xpert Result"),
        ("Microscopy Result_End", "Microscopy Result"),
        ("TrueNat Result_End", "TrueNat Result"),
        ("Gene Xpert Result_End", "Gene Xpert Result"),
    ]

    for label, var in dropdown_rules:
        col = find_column(df, label)
        if not col:
            continue
        allowed = dropdowns.get(var, set())
        df[f"{label.replace(' ', '_')}_Error"] = df[col].apply(
            lambda v: "" if normalize(v) in allowed else "Invalid value"
        )

    # ---------------- Dates ----------------
    for label in ["Enrolled date", "Started date"]:
        col = find_column(df, label)
        if col:
            df[f"{label.replace(' ', '_')}_Error"] = df[col].apply(
                lambda v: "" if is_valid_date(v) else f"Invalid {label.lower()}"
            )

    # ---------------- Outcome date rule ----------------
    outcome_col = find_column(df, "Treatment outcome")
    out_date_col = find_column(df, "Outcome date")

    if outcome_col and out_date_col:
        def outcome_date_check(row):
            if normalize(row.get(outcome_col)) == "":
                return ""
            if not is_valid_date(row.get(out_date_col)):
                return "Invalid outcome date"
            return ""

        df["Outcome_Date_Error"] = df.apply(outcome_date_check, axis=1)

    # ---------------- Age ----------------
    age_col = find_column(df, "Age")
    if age_col:
        df["Age_Error"] = df[age_col].apply(
            lambda v: "" if is_numeric(v) else "Invalid age"
        )
    else:
        df["Age_Error"] = "Age column missing"


    # ---------------- Gender ----------------
    def gender_check(row):
        sex = norm_lower(row.get("Gender"))
        name = row.get("Patient Name", "")
        prefix = extract_prefix(name)

        if sex not in {s.lower() for s in dropdowns.get("Gender", set())}:
            return "Invalid gender"

        if prefix in {"Ma", "Daw"} and sex not in {"female", "f"}:
            return "Gender mismatch with name"
        if prefix in {"Mg", "U", "Ko"} and sex not in {"male", "m"}:
            return "Gender mismatch with name"
        return ""

    df["Gender_Error"] = df.apply(gender_check, axis=1)

    # ---------------- Base Code ----------------
    df["Base_Code_Error"] = df.apply(
        lambda r: ""
        if (normalize(r.get("Township")), normalize(r.get("Base Code"))) in service_pairs
        else "Invalid Base Code / Township",
        axis=1
    )

    # ---------------- Level ----------------
    df["Level"] = df["Base Code"].apply(lambda x: level_map.get(normalize(x), ""))

    # ---------------- Township TB No ----------------
    tbno_col = find_column(df, "Township TB No")
    if tbno_col:
        df["Township_TB_No_Error"] = df[tbno_col].apply(
            lambda v: "Blank Township TB No" if normalize(v) == "" else ""
        )

        def ordinal(n):
            if 10 <= n % 100 <= 20:
                suffix = 'th'
            else:
                suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
            return str(n) + suffix

        enroll_col = find_column(df, "Enrolled date")
        if enroll_col:
            df[enroll_col] = pd.to_datetime(df[enroll_col], errors="coerce", dayfirst=True)
            for tbno, group in df[df.duplicated(subset=[tbno_col], keep=False)].groupby(tbno_col):
                group = group.sort_values(by=enroll_col)
                for i, idx in enumerate(group.index):
                    df.loc[idx, "Duplicate_check"] = f"duplicated({ordinal(i+1)})"
        else:
            dup = df.duplicated(subset=[tbno_col], keep=False)
            df.loc[dup, "Duplicate_check"] = "To recheck for duplication"

    # ---------------- BC from TB register ----------------
    # (moved below bact_confirmed to appear side-by-side later)
    # placeholder: actual insertion will happen after lab-derived flag
    bc_block_needed = not df_register.empty and tbno_col
    bc_block = None
    if bc_block_needed:
        reg_tbno = find_column(df_register, "Township TB No")
        reg_bc = find_column(df_register, "BC")
        if reg_tbno and reg_bc:
            bc_map = dict(zip(
                df_register[reg_tbno].astype(str),
                df_register[reg_bc]
            ))
            bc_block = df[tbno_col].astype(str).map(bc_map).fillna(0).astype("Int64")

    # ---------------- bact_confirmed (from lab results) ----------------
    df["BC_results"] = df.apply(BC_results, axis=1).astype("Int64")

    # insert BC column immediately after bact_confirmed if we computed it earlier
    if bc_block is not None:
        insert_pos = df.columns.get_loc("BC_results") + 1
        df.insert(insert_pos, "BC_TB_register", bc_block)


    # ---------------- TBDT1 consistency check ----------------
    tbno_col = find_column(df, "Township TB No")
    out_tbdt1_col = find_column(df, "TBDT1")

    reg_tbno = find_column(df_register, "Township TB No")
    reg_tbdt1 = find_column(df_register, "TBDT1")

    df["TBDT1_check"] = ""

    if tbno_col and out_tbdt1_col and reg_tbno and reg_tbdt1:
        reg_map = dict(zip(
            df_register[reg_tbno].astype(str),
            df_register[reg_tbdt1]
        ))

        def check_tbdt1(row):
            tbno = normalize(row.get(tbno_col))
            if tbno not in reg_map:
                return "TB No not found in TB register"

            reg_val = reg_map.get(tbno)
            out_val = row.get(out_tbdt1_col)

            if str(reg_val).strip() != str(out_val).strip():
                return "TBDT1 mismatch with TB register"

            return ""

        df["TBDT1_check"] = df.apply(check_tbdt1, axis=1)
        
    
    # ---------------- Outcome calculation ----------------
    def compute_outcome(row):
        # ---------------- Dates ----------------
        start_date = pd.to_datetime(row.get("Started date"), errors="coerce", dayfirst=True)
        outcome_date = pd.to_datetime(row.get("Outcome date"), errors="coerce", dayfirst=True)

        if pd.isna(start_date):
            return ""

        # ====================================================
        # Skip checks if Treatment outcome is "Died" or "Not evaluated"
        # ====================================================
        treatment_outcome = norm_lower(row.get("Treatment outcome"))
        if treatment_outcome in {"died", "not evaluated"}:
            return row.get("Treatment outcome")

        today = pd.Timestamp.today().normalize()

        days_from_start_to_today = (today - start_date).days

        def months_between(d1, d2):
            return (d2.year - d1.year) * 12 + (d2.month - d1.month)

        def days_between(d1, d2):
            return (d2 - d1).days

        # ---------------- Helper: BC-positive ----------------
        def is_bc_positive(val):
            s = norm_lower(val)
            return any(k in s for k in [
                "positive", "scanty", "1+", "2+", "3+",
                 "rif detected", "rr"
            ])

        # ====================================================
        # 1️⃣ Failed (HIGHEST PRIORITY)
        # ====================================================
        for c in [
            "Microscopy Result_5", "TrueNat Result_5", "Gene Xpert Result_5",
            "Microscopy Result_End", "TrueNat Result_End", "Gene Xpert Result_End",
        ]:
            if c in row.index and is_bc_positive(row.get(c)):
                return "Failed"

        # ====================================================
        # 2️⃣ Cure
        # ====================================================

        micro5 = norm_lower(row.get("Microscopy Result_5"))
        micro_end = norm_lower(row.get("Microscopy Result_End"))

        negative_values = {"negative", "no afb seen"}

        bc_reg = pd.to_numeric(row.get("BC_TB_register"), errors="coerce")
        bc_res = pd.to_numeric(row.get("BC_results"), errors="coerce")

        if (
            (bc_reg == 1 or bc_res == 1)
            and pd.notna(outcome_date)
            and days_between(start_date, outcome_date) >= 168
            and micro5 in negative_values
            and micro_end in negative_values
        ):
            return "Cure"

        # ====================================================
        # 3️⃣ Complete
        # ====================================================
        if (
            pd.notna(outcome_date) and
            days_between(start_date, outcome_date) >= 168
        ):
            return "Complete"

        # ====================================================
        # 4️⃣ LTFU
        # ====================================================
        if days_from_start_to_today >= 168:
            if pd.isna(outcome_date):
                return "LTFU"
            if days_between(start_date, outcome_date) < 168:
                return "LTFU"

        # ====================================================
        # 5️⃣ Blank
        # ====================================================
        return ""



    df["Outcome"] = df.apply(compute_outcome, axis=1)


    return df


# =================================================
# TPT REGISTER RULES
# =================================================

def process_tpt_register(df, df_dropdown, df_service, df_screen):
    df = df.copy()

    dropdowns = load_dropdowns(df_dropdown)
    service_pairs = load_service_pairs(df_service)
    level_map = load_basecode_level(df_service)

    # ---------------- Age ----------------
    age_col = find_column(df, "Age")
    if age_col:
        df["Age_Error"] = df[age_col].apply(
            lambda v: "" if is_numeric(v) else "Invalid age"
        )
    else:
        df["Age_Error"] = "Age column missing"

    # ---------------- Gender ----------------
    def gender_check(row):
        sex = norm_lower(row.get("Gender"))
        name = row.get("Name", "")
        prefix = extract_prefix(name)

        if sex not in {s.lower() for s in dropdowns.get("Gender", set())}:
            return "Invalid gender"

        if prefix in {"Ma", "Daw"} and sex not in {"female", "f"}:
            return "Gender mismatch with name"
        if prefix in {"Mg", "U", "Ko"} and sex not in {"male", "m"}:
            return "Gender mismatch with name"
        return ""

    df["Gender_Error"] = df.apply(gender_check, axis=1)

    # ---------------- TPT Regime ----------------
    regime_col = find_column(df, "TPT Regime")
    if regime_col:
        allowed = dropdowns.get("TPT Regime", set())
        df["TPT_Regime_Error"] = df[regime_col].apply(
            lambda v: "" if normalize(v) in allowed else "Invalid value"
        )

    # ---------------- TPT Started Date ----------------
    start_col = find_column(df, "TPT Started Date")
    if start_col:
        df["TPT_Started_Date_Error"] = df[start_col].apply(
            lambda v: "" if is_valid_date(v) else "Invalid TPT started date"
        )

    # ---------------- Discharge date ----------------
    discharge_col = find_column(df, "Discharge date")
    if discharge_col:
        df["Discharge_date_Error"] = df[discharge_col].apply(
            lambda v: "" if is_valid_date(v) else "Invalid discharge date"
        )

    # ---------------- Duration check ----------------
    if start_col and discharge_col and regime_col:
        def duration_check(row):
            start = pd.to_datetime(row.get(start_col), errors="coerce", dayfirst=True)
            discharge = pd.to_datetime(row.get(discharge_col), errors="coerce", dayfirst=True)
            regime = norm_lower(row.get(regime_col))

            if pd.isna(start) or pd.isna(discharge):
                return ""

            expected_days = {"1hp": 30, "3hp": 90, "3hr": 90, "6h": 180}.get(regime, 0)
            if expected_days == 0:
                return ""

            expected_end = start + pd.Timedelta(days=expected_days)
            today = pd.Timestamp.today()

            if today < expected_end:
                return ""  # Designated duration not yet reached, don't check

            days = (discharge - start).days
            if abs(days - expected_days) <= 7:
                return ""
            else:
                return "to recheck outcome date"

        df["Duration_Error"] = df.apply(duration_check, axis=1)

    # ---------------- Code ----------------
    code_col = find_column(df, "Code")
    if code_col:
        df["Code_Error"] = df[code_col].apply(
            lambda v: "Blank Code" if normalize(v) == "" else ""
        )

        # Duplicate check
        def ordinal(n):
            if 10 <= n % 100 <= 20:
                suffix = 'th'
            else:
                suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
            return str(n) + suffix

        df["Duplicate_check"] = ""
        for name, group in df.groupby(code_col):
            if len(group) > 1:
                # Sort group by index to maintain original order
                group = group.sort_index()
                for i, idx in enumerate(group.index):
                    df.loc[idx, "Duplicate_check"] = f"duplicated({ordinal(i+1)})"

    # ---------------- New columns from TB screening ----------------
    if not df_screen.empty and code_col:
        scr_tbno = find_column(df_screen, "Registration No")
        scr_township = find_column(df_screen, "Township")
        scr_basecode = find_column(df_screen, "Base Code")
        scr_hiv = find_column(df_screen, "HIV status")

        if scr_tbno:
            # Build maps
            township_map = {}
            basecode_map = {}
            hiv_map = {}

            for _, r in df_screen.iterrows():
                tbno = normalize(str(r.get(scr_tbno, "")))
                township = r.get(scr_township, "") if scr_township else ""
                basecode = r.get(scr_basecode, "") if scr_basecode else ""
                hiv = r.get(scr_hiv, "") if scr_hiv else ""

                township_map[tbno] = township
                basecode_map[tbno] = basecode
                hiv_map[tbno] = hiv

            df["Township_Scr"] = df[code_col].astype(str).apply(lambda x: township_map.get(normalize(x), ""))
            df["Base Code_Scr"] = df[code_col].astype(str).apply(lambda x: basecode_map.get(normalize(x), ""))
            df["HIV status_Scr"] = df[code_col].astype(str).apply(lambda x: hiv_map.get(normalize(x), ""))

    return df


# =================================================
# Combine Errors → Comment
# =================================================

def combine_errors(df, sheet_name=None):
    err_cols = [c for c in df.columns if c.endswith("_Error") or c.endswith("_check")]

    def row_comment(row):
        msgs = []

        for c in err_cols:
            val = normalize(row.get(c))
            if not val:
                continue
            if c.endswith("_check") and val.lower() in {"t", "true"}:
                continue
            msgs.append(f"{c.replace('_Error','').replace('_check','')}: {val}")

        # ✅ APPLY ONLY FOR TB SCREENING
        if sheet_name == "TB screening":
            symptom_msg = symptom_issue(row)
            if symptom_msg:
                msgs.append(symptom_msg)

        return "; ".join(msgs)

    df["Comment"] = df.apply(row_comment, axis=1)
    return df.drop(columns=err_cols, errors="ignore")


# =================================================
# Main entry
# =================================================

def check_rules(excel_file, output_file=None):
    xls = pd.ExcelFile(excel_file)

    df_screen = xls.parse("TB screening") if "TB screening" in xls.sheet_names else pd.DataFrame()
    df_register = xls.parse("TB register") if "TB register" in xls.sheet_names else pd.DataFrame()
    df_outcome = (
        xls.parse("TB outcome follow up")
        if "TB outcome follow up" in xls.sheet_names
        else pd.DataFrame()
    )
    df_dropdown = xls.parse("Dropdown") if "Dropdown" in xls.sheet_names else pd.DataFrame()
    df_service = xls.parse("Service point") if "Service point" in xls.sheet_names else pd.DataFrame()
    df_tpt = xls.parse("TPT register") if "TPT register" in xls.sheet_names else pd.DataFrame()

    # ================== TB screening ==================
    df_screen = process_tb_screening(df_screen, df_dropdown, df_service, df_register, df_tpt)
    df_screen = combine_errors(df_screen, sheet_name="TB screening")
    df_screen = clean_dates(df_screen)   # 🔴 LAST STEP

    # ================== TB register ==================
    df_register = process_tb_register(df_register, df_dropdown, df_service, df_screen)
    df_register = combine_errors(df_register, sheet_name="TB register")
    df_register = clean_dates(df_register)  # 🔴 LAST STEP

    # ================== TB outcome follow up ==================
    if not df_outcome.empty:
        df_outcome = process_tb_outcome_follow_up(
            df_outcome,
            df_dropdown,
            df_service,
            df_register
        )
        df_outcome = combine_errors(df_outcome, sheet_name="TB outcome follow up")
        df_outcome = clean_dates(df_outcome)  # 🔴 LAST STEP

    # ================== TPT register ==================
    df_tpt = process_tpt_register(df_tpt, df_dropdown, df_service, df_screen)
    df_tpt = combine_errors(df_tpt, sheet_name="TPT register")
    df_tpt = clean_dates(df_tpt)  # 🔴 LAST STEP

    # ================== RESULTS ==================
    results = {
        "TB screening": df_screen,
        "TB register": df_register,
    }

    if not df_outcome.empty:
        results["TB outcome follow up"] = df_outcome

    results["TPT register"] = df_tpt

    # ================== EXPORT ==================
    if output_file:
        with pd.ExcelWriter(output_file) as w:
            for name, df in results.items():
                df.to_excel(w, sheet_name=name[:31], index=False)

    return results


