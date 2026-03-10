from utils.helpers import normalize, normalize_col_name, norm_lower

def symptom_issue(row):
    symptom_cols = {"cough>2 weeks","fever","wt loss","night sweat","other","dm status","hiv status","tb contact"}
    NO_SET = {"no","n","uk","","negative"}

    values = []
    for c in row.index:
        if normalize_col_name(c) in symptom_cols:
            values.append(norm_lower(row.get(c)))

    if values and all(v in NO_SET for v in values):
        return "to recheck symptom"
    return ""

def combine_errors(df, sheet_name=None):
    err_cols = [c for c in df.columns if c.endswith("_Error") or c.endswith("_check")]

    def row_comment(row):
        msgs = []
        for c in err_cols:
            val = normalize(row.get(c))
            if not val:
                continue
            if c.endswith("_check") and val.lower() in {"t","true"}:
                continue
            msgs.append(f"{c.replace('_Error','').replace('_check','')}: {val}")

        if sheet_name == "TB screening":
            symptom_msg = symptom_issue(row)
            if symptom_msg:
                msgs.append(symptom_msg)

        return "; ".join(msgs)

    df["Comment"] = df.apply(row_comment, axis=1)
    return df.drop(columns=err_cols, errors="ignore")
