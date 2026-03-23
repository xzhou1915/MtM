import pandas as pd
import numpy as np

# --- User-editable paths ---
INPUT_FILE  = "Sample.xlsx"
OUTPUT_FILE = "forward_curve_output.xlsx"

# --- Load input ---
df = pd.read_excel(INPUT_FILE)

# --- Column name aliases ---
col_map = {"ExpiryDays": "days_to_expiry", "Value": "fwd_rate"}
df = df.rename(columns=col_map)

# --- Validation ---
required_cols = {"days_to_expiry", "fwd_rate"}
if not required_cols.issubset(df.columns):
    raise ValueError(f"Input file must contain columns: {required_cols}. Found: {list(df.columns)}")

df = df[["days_to_expiry", "fwd_rate"]].dropna()

if not pd.api.types.is_numeric_dtype(df["days_to_expiry"]):
    raise TypeError("days_to_expiry must be numeric.")

df["days_to_expiry"] = df["days_to_expiry"].astype(int)

# Sort and handle duplicates: keep last occurrence after sorting.
# Rationale: the last entry is assumed to be the most recent/corrected value.
df = df.sort_values("days_to_expiry").drop_duplicates(subset="days_to_expiry", keep="last")

# --- Build full integer day grid ---
day_min, day_max = df["days_to_expiry"].min(), df["days_to_expiry"].max()
full_grid = pd.DataFrame({"days_to_expiry": np.arange(day_min, day_max + 1)})

# --- Merge known points onto grid, then interpolate ---
merged = full_grid.merge(df, on="days_to_expiry", how="left")

# Linear interpolation across the integer day axis
merged["fwd_rate_interpolated"] = merged["fwd_rate"].interpolate(method="linear")

# --- Save output ---
result = merged[["days_to_expiry", "fwd_rate_interpolated"]]
result.to_excel(OUTPUT_FILE, index=False)

print(f"Done. {len(result)} rows written to {OUTPUT_FILE}")
