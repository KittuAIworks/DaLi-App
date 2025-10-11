import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Governance Lineage Generator", layout="centered")

st.title("Upload Governance Model")
st.markdown("Generate data lineage document from your governance model Excel file.")

uploaded_file = st.file_uploader("Upload your Excel file (.xlsm)", type=["xlsm"])

def generate_lineage(file):
    # Load sheets with macros disabled
    xls = pd.ExcelFile(file, engine="openpyxl")
    sheets_to_keep = ["BUSINESS RULES", "BUSINESS CONDITIONS", "GOVERNANCE MAPPING", "CONTEXTS"]
    data = {sheet: xls.parse(sheet) for sheet in sheets_to_keep if sheet in xls.sheet_names}

    # BUSINESS RULES
    br = data["BUSINESS RULES"].rename(columns={"NAME": "RULE NAME", "DISPLAY NAME": "RULE DISPLAY NAME"})
    br = br[["RULE NAME", "TYPE", "DEFINITION", "RULE DISPLAY NAME", "IS ENABLED?"]]
    br = br[br["IS ENABLED?"] == "Yes"]

    # BUSINESS CONDITIONS
    bc = data["BUSINESS CONDITIONS"].rename(columns={"NAME": "CONDITION NAME", "DISPLAY NAME": "CONDITION DISPLAY NAME"})
    bc = bc[["CONDITION NAME", "MAPPED BUSINESS RULE(s)", "IMPACTED ROLES", "IMPACTED ATTRIBUTES", "IMPACTED RELATIONSHIPS", "CONDITION DISPLAY NAME", "IS ENABLED?"]]
    bc = bc[bc["IS ENABLED?"] == "Yes"]

    # GOVERNANCE MAPPING
    gm = data["GOVERNANCE MAPPING"]
    gm["FOR CONTEXT"] = gm["FOR CONTEXT"].astype(str)
    gm = gm[["ENTITY", "MAPPED BUSINESS RULE", "MAPPED BUSINESS CONDITION", "FOR CONTEXT", "IS ENABLED?"]]
    gm = gm[gm["IS ENABLED?"] == "Yes"]

    # CONTEXTS
    ctx = data["CONTEXTS"].rename(columns={
        "NAME": "CONTEXT NAME",
        "CONTEXT TYPE || CONTEXT NAME": "CONTEXT TYPE AND NAME"
    })
    ctx = ctx[["CONTEXT NAME", "CONTEXT TYPE AND NAME", "WORKFLOW ACTIVITY", "WORKFLOW ACTIVITY ACTION(s)", "WORKFLOW ACTIVITY CRITERIA"]]

    # Lineage compilation
    records = []
    for _, rule in br.iterrows():
        rule_name = rule["RULE NAME"]
        matched_conditions = bc[bc["MAPPED BUSINESS RULE(s)"].str.contains(rule_name, na=False)]
        lineage_found = False

        if not matched_conditions.empty:
            for _, cond in matched_conditions.iterrows():
                keys = [f"{cond['CONDITION NAME']}{suffix}Context" for suffix in ["", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15"]]
                matched_gm = gm[gm["FOR CONTEXT"].isin(keys)]
                for _, gm_row in matched_gm.iterrows():
                    matched_ctx = ctx[ctx["CONTEXT NAME"] == gm_row["FOR CONTEXT"]]
                    ctx_row = matched_ctx.iloc[0] if not matched_ctx.empty else pd.Series(dtype=object)
                    records.append({
                        **rule.drop("IS ENABLED?"),
                        **cond.drop("IS ENABLED?"),
                        **gm_row.drop("IS ENABLED?"),
                        **ctx_row
                    })
                    lineage_found = True

        if not lineage_found:
            keys = [f"{rule_name}{suffix}Context" for suffix in ["", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15"]]
            matched_gm = gm[gm["FOR CONTEXT"].isin(keys)]
            if not matched_gm.empty:
                for _, gm_row in matched_gm.iterrows():
                    matched_ctx = ctx[ctx["CONTEXT NAME"] == gm_row["FOR CONTEXT"]]
                    ctx_row = matched_ctx.iloc[0] if not matched_ctx.empty else pd.Series(dtype=object)
                    records.append({
                        **rule.drop("IS ENABLED?"),
                        "CONDITION NAME": "",
                        "IMPACTED ROLES": "",
                        "IMPACTED ATTRIBUTES": "",
                        "IMPACTED RELATIONSHIPS": "",
                        "CONDITION DISPLAY NAME": "",
                        **gm_row.drop("IS ENABLED?"),
                        **ctx_row
                    })
                    lineage_found = True

        if not lineage_found:
            matched_gm = gm[gm["MAPPED BUSINESS RULE"] == rule_name]
            for _, gm_row in matched_gm.iterrows():
                records.append({
                    **rule.drop("IS ENABLED?"),
                    "CONDITION NAME": "",
                    "IMPACTED ROLES": "",
                    "IMPACTED ATTRIBUTES": "",
                    "IMPACTED RELATIONSHIPS": "",
                    "CONDITION DISPLAY NAME": "",
                    "FOR CONTEXT": "",
                    "CONTEXT NAME": "",
                    "CONTEXT TYPE AND NAME": "",
                    "WORKFLOW ACTIVITY": "",
                    "WORKFLOW ACTIVITY ACTION(s)": "",
                    "WORKFLOW ACTIVITY CRITERIA": "",
                    **gm_row.drop(["IS ENABLED?", "FOR CONTEXT"])
                })

    output_df = pd.DataFrame(records)
    return output_df

if uploaded_file:
    lineage_df = generate_lineage(uploaded_file)
    st.success("Lineage document generated successfully!")

    # Provide download link
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        lineage_df.to_excel(writer, index=False, sheet_name="Lineage Output")
    st.download_button("Download Lineage Document", data=output.getvalue(), file_name="Governance_rules_lineage_output.xlsx")
