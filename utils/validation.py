from utils.helpers import normalize

def combine_errors(df, sheet_name=None):

    err_cols = [c for c in df.columns if c.endswith("_Error") or c.endswith("_check")]

    def row_comment(row):
        msgs = []

        for c in err_cols:
            val = normalize(row.get(c))
            if not val:
                continue
            msgs.append(f"{c.replace('_Error','').replace('_check','')}: {val}")

        return "; ".join(msgs)

    df["Comment"] = df.apply(row_comment, axis=1)

    return df.drop(columns=err_cols, errors="ignore")
