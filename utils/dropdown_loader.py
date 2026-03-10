from utils.helpers import normalize

def load_dropdowns(df):
    d = {}
    for _, r in df.iterrows():
        var = normalize(r.get("Variable"))
        val = normalize(r.get("Value"))
        if var:
            d.setdefault(var, set()).add(val)
    return d


def load_service_pairs(df):
    return set(
        (normalize(r.get("Township")), normalize(r.get("Base Code")))
        for _, r in df.iterrows()
        if normalize(r.get("Township")) and normalize(r.get("Base Code"))
    )


def load_basecode_level(df):
    return {
        normalize(r.get("Base Code")): normalize(r.get("Level"))
        for _, r in df.iterrows()
        if normalize(r.get("Base Code"))
    }
