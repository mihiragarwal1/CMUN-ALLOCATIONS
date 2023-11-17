import pandas as pd
import openpyxl
import numpy as np

sheet = "CCC"

matrix_df = pd.read_excel("Matrices ONLY MAKE CHANGES TO THIS.xlsx", sheet_name=sheet)
round1_df = pd.read_excel("Round 1 Applications CHIREC MUN 2023 3.xlsx")
priority_df = pd.read_excel("Priority round applications.xlsx")

for name in matrix_df["Name of the Delegate"]:
    if pd.isna(name):
        pass
    else:
        if round1_df["Full Name:"].eq(name).any():
            row = round1_df.index[round1_df["Full Name:"] == name].tolist()

            matrix_df.at[
                matrix_df.index[matrix_df["Name of the Delegate"] == name].tolist()[0],
                "Address",
            ] = round1_df.at[row[0], "Please Enter Your Address:"]
matrix_df.to_excel(
    "Matrices ONLY MAKE CHANGES TO THIS.xlsx", sheet_name=sheet, index=False
)
