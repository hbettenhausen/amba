
# PATCHED RTF HANDLING VERSION

import streamlit as st
import pandas as pd
import re

def parse_rtf_content(content):
    # Clean out RTF control characters
    content = re.sub(r'[{}\\]', '', content)
    content = re.sub(r'\\[a-z]+[0-9]*', '', content)
    lines = content.splitlines()
    data = []
    for line in lines:
        # Try to find tab- or space-separated values with treatment, mean, and group
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
    st.title("Debug RTF Reader")
    uploaded_file = st.file_uploader("Upload your .rtf file", type=[".rtf"])

    if uploaded_file:
        raw_bytes = uploaded_file.read()
        text = raw_bytes.decode(errors="ignore")

        df, lines = parse_rtf_content(text)
        if df.empty:
            st.warning("Could not parse expected Treatment/Mean/Group rows.")
            st.subheader("Preview of RTF contents:")
            st.text("\n".join(lines[:30]))  # show first 30 lines for debugging
        else:
            st.success("Successfully parsed!")
            st.dataframe(df)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            st.download_button("Download Excel version", output.getvalue(), "Parsed_RTF.xlsx")

if __name__ == "__main__":
    main()
