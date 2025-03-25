
# QEP Summary Web App (Shiny-style Python app using Streamlit)
# Enhanced RTF handling

import streamlit as st
import pandas as pd
import numpy as np
import re
import os
from io import BytesIO
from statsmodels.formula.api import ols
from statsmodels.stats.multicomp import pairwise_tukeyhsd
import seaborn as sns
import matplotlib.pyplot as plt
from collections import defaultdict

def clean_param_name(name):
    return re.sub(r'[^a-zA-Z0-9_]', '_', name.strip())

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

def parse_rtf_content(content):
    content = re.sub(r'[{}]', '', content)                # remove braces
    content = re.sub(r'\\tab', '\t', content)          # convert 	ab
    content = re.sub(r'\\par', '\n', content)          # convert \par
    content = re.sub(r'\\[a-z]+[0-9]*', '', content)    # strip other formatting
    content = content.replace('\', '')                   # cleanup slashes
    lines = content.splitlines()
    data = []
    for line in lines:
        parts = re.split(r'[\t\s]+', line.strip())
        if len(parts) >= 3:
            try:
                treatment = parts[0]
                mean = float(parts[1])
                group = parts[2]
                data.append({"Treatment": treatment, "Mean": mean, "Group": group})
            except:
                continue
    return pd.DataFrame(data), lines

def main():
    st.title("QEP Summary Stats & Visualization Tool")
    uploaded_file = st.file_uploader("Upload your QEP data file (Excel or RTF)", type=[".xlsx", ".rtf"])

    if uploaded_file:
        filetype = os.path.splitext(uploaded_file.name)[-1].lower()

        if filetype == ".rtf":
            rtf_bytes = uploaded_file.read()
            rtf_text = rtf_bytes.decode(errors="ignore")
            df, lines = parse_rtf_content(rtf_text)
            if df.empty:
                st.error("No readable data found in the RTF file.")
                st.subheader("Raw preview (first 30 lines):")
                st.text("\n".join(lines[:30]))
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
