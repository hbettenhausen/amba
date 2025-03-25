# QEP Summary Web App (Shiny-style Python app using Streamlit)
# This will process Excel files or .rtf text reports, run ANOVA + Tukey HSD, and output clean reports with heatmaps.

import streamlit as st
import pandas as pd
import numpy as np
from statsmodels.formula.api import ols
from statsmodels.stats.multicomp import pairwise_tukeyhsd
import seaborn as sns
import matplotlib.pyplot as plt
from io import BytesIO
import base64
from collections import defaultdict
import re
import os

# Clean column name for modeling
def clean_param_name(name):
    return re.sub(r'[^a-zA-Z0-9_]', '_', name.strip())

# Assign group letters (Tukey-style)
def assign_groups(means, grouping_dict):
    treatments = list(means["Treatment"])
    groups = {}
    assigned = set()
    current_letter = 'A'
    for t in treatments:
        if t in assigned:
            continue
        group = {t}
        to_check = {t}
        while to_check:
            item = to_check.pop()
            for neighbor in grouping_dict[item]:
                if neighbor not in group:
                    group.add(neighbor)
                    to_check.add(neighbor)
        for item in group:
            groups[item] = current_letter
            assigned.add(item)
        current_letter = chr(ord(current_letter) + 1)
    return groups

# Parse RTF text and extract treatment, mean, group data
# This is a placeholder – real parsing may depend on known format structure

def parse_rtf_content(content):
    lines = content.splitlines()
    data = []
    for line in lines:
        if re.match(r"^\s*\w+\s+\d+\.\d+\s+[A-Z]+", line):
            parts = line.strip().split()
            if len(parts) >= 3:
                treatment = parts[0]
                mean = float(parts[1])
                group = parts[2]
                data.append({"Treatment": treatment, "Mean": mean, "Group": group})
    return pd.DataFrame(data)

# App interface
from io import BytesIO

def main():
    st.title("QEP Summary Stats & Visualization Tool")
    uploaded_file = st.file_uploader("Upload your QEP data file (Excel or RTF)", type=[".xlsx", ".rtf"])

    if uploaded_file:
        filetype = os.path.splitext(uploaded_file.name)[-1].lower()

        if filetype == ".xlsx":
            xls = pd.ExcelFile(uploaded_file)
            trial_sheets = [s for s in xls.sheet_names if not s.endswith("_Stats") and s != "Legend"]

            priority_params = [
                "Fri %", "FAN mg/L", "DP °L", "Extract FGDB %",
                "BG ppm", "Beta-Glucan ppm", "Protein %", "S/T %", "Wort Color SRM"
            ]

            all_results = []

            for sheet in trial_sheets:
                df = pd.read_excel(xls, sheet_name=sheet, header=[1])
                df.columns = df.columns.str.strip()
                df = df.rename(columns={df.columns[1]: "Treatment", df.columns[2]: "Site"})
                df = df.dropna(subset=["Treatment"])

                available_params = [p for p in priority_params if p in df.columns]

                for param in available_params:
                    sub_df = df[["Treatment", param]].dropna()
                    if sub_df["Treatment"].nunique() < 2 or sub_df.shape[0] < 3:
                        continue

                    safe_param = clean_param_name(param)
                    sub_df = sub_df.rename(columns={param: safe_param})
                    try:
                        model = ols(f'Q("{safe_param}") ~ C(Treatment)', data=sub_df).fit()
                        tukey = pairwise_tukeyhsd(endog=sub_df[safe_param],
                                                  groups=sub_df["Treatment"], alpha=0.05)
                    except Exception:
                        continue

                    means = sub_df.groupby("Treatment")[safe_param].mean().reset_index()
                    means.columns = ["Treatment", "Mean"]
                    tukey_summary = pd.DataFrame(data=tukey._results_table.data[1:], columns=tukey._results_table.data[0])
                    grouping_dict = defaultdict(set)
                    for _, row in tukey_summary.iterrows():
                        g1, g2, reject = row['group1'], row['group2'], row['reject']
                        if not reject:
                            grouping_dict[g1].add(g2)
                            grouping_dict[g2].add(g1)
                    group_map = assign_groups(means, grouping_dict)
                    means["Group"] = means["Treatment"].map(group_map)
                    means["Parameter"] = param
                    means["Trial"] = sheet
                    means["p_value"] = model.f_pvalue
                    all_results.append(means)

            if all_results:
                full_df = pd.concat(all_results, ignore_index=True)
                st.success("Analysis complete! Here's a preview:")
                st.dataframe(full_df.head(20))

                filtered_df = full_df[full_df["p_value"] < 0.05]
                st.subheader("Filtered by p < 0.05")
                st.dataframe(filtered_df)

                trial_choice = st.selectbox("Select a trial for heatmap:", filtered_df["Trial"].unique())
                heatmap_data = filtered_df[filtered_df["Trial"] == trial_choice]
                heatmap_pivot = heatmap_data.pivot(index="Parameter", columns="Treatment", values="Mean")
                st.subheader(f"Heatmap: Mean values for {trial_choice}")
                fig, ax = plt.subplots(figsize=(12, 6))
                sns.heatmap(heatmap_pivot, annot=True, fmt=".1f", cmap="coolwarm", ax=ax)
                st.pyplot(fig)

                def to_excel(df):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, sheet_name='QEP Results')
                    return output.getvalue()

                st.download_button(
                    label="📥 Download All Results (Excel)",
                    data=to_excel(full_df),
                    file_name="QEP_Stats_Results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        elif filetype == ".rtf":
            rtf_bytes = uploaded_file.read()
            rtf_text = rtf_bytes.decode(errors="ignore")
            df = parse_rtf_content(rtf_text)
            if df.empty:
                st.error("No readable data found in the RTF file.")
                st.subheader("Raw RTF Preview (first 30 lines):")
          st.text("\n".join(rtf_text.splitlines()[:30]))


            else:
                st.success("RTF data parsed!")
                st.dataframe(df)
                def to_excel(df):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False)
                    return output.getvalue()
                st.download_button(
                    label="📥 Download Parsed Table (Excel)",
                    data=to_excel(df),
                    file_name="QEP_RTF_Parsed.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
