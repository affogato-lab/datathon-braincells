import pandas as pd
import numpy as np

file_path = "2025 Allianz Datathon Dataset.xlsx"
df = pd.read_excel(file_path, sheet_name="Climate Data")

df["Date"] = pd.to_datetime(df[["Year", "Month", "Day"]])
cols = ["Maximum temperature (Degree C)", "Minimum temperature (Degree C)", "Rainfall amount (millimetres)"]
df = df[["Date"] + cols].replace(["M", "T", "-", "--", ""], np.nan)
df[cols] = df[cols].apply(pd.to_numeric, errors="coerce")
df = df[df["Date"] >= "2010-05-26"]

weekly = df.set_index("Date").resample("W-SUN").agg({
    "Maximum temperature (Degree C)": "mean",
    "Minimum temperature (Degree C)": "mean",
    "Rainfall amount (millimetres)": "sum"
})

temp_cols = ["Maximum temperature (Degree C)", "Minimum temperature (Degree C)"]
for col in temp_cols:
    weekly[col] = weekly[col].interpolate(method="linear")

weekly = weekly.dropna(how="all").reset_index()
weekly["Mean temperature (Degree C)"] = weekly[["Maximum temperature (Degree C)", "Minimum temperature (Degree C)"]].mean(axis=1)
weekly["Week"] = [f"Week{i+1} ({d.strftime('%Y-%m-%d')})" for i, d in enumerate(weekly["Date"])]

final = weekly[["Week"] + cols + ["Mean temperature (Degree C)"]]
final.to_excel("weekly_climate_data_final.xlsx", index=False)
final.to_csv("weekly_climate_data_final.csv", index=False)
