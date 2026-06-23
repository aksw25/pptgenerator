import os
import pandas as pd

MAX_ROWS_PER_SHEET = 200


def read_xlsx(path: str) -> dict:
    xf = pd.ExcelFile(path)
    sheets = []
    for name in xf.sheet_names:
        df = pd.read_excel(xf, sheet_name=name, nrows=MAX_ROWS_PER_SHEET)
        df = df.dropna(how="all")

        columns = [str(c) for c in df.columns.tolist()]
        rows = df.head(MAX_ROWS_PER_SHEET).astype(str).values.tolist()

        summary = {"row_count": len(df)}
        numeric_cols = df.select_dtypes(include="number")
        for col in numeric_cols.columns:
            series = numeric_cols[col].dropna()
            if len(series):
                summary[str(col)] = {
                    "min": float(series.min()),
                    "max": float(series.max()),
                    "mean": round(float(series.mean()), 4),
                }

        sheets.append({"name": name, "columns": columns, "rows": rows, "summary": summary})

    return {"file": os.path.basename(path), "type": "xlsx", "sheets": sheets}
