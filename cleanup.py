import pandas as pd

def merge_broken_rows(df: pd.DataFrame) -> pd.DataFrame:
    out_rows = []
    prev = None
    for _, row in df.iterrows():
        non_null_count = row.count()
        total = len(row)
        # Consider row broken if <30% filled
        if prev is not None and (non_null_count / total) <= 0.3:
            for col in df.columns:
                val = row[col]
                if pd.isna(val):
                    continue
                if pd.isna(prev[col]):
                    prev[col] = val
                else:
                    prev[col] = str(prev[col]).rstrip() + " " + str(val).lstrip()
        else:
            if prev is not None:
                out_rows.append(prev)
            prev = row.copy()
    if prev is not None:
        out_rows.append(prev)
    new_df = pd.DataFrame(out_rows).reset_index(drop=True)
    return new_df
