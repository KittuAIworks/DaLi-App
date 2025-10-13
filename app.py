import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Lineage Generator", layout="wide")

# --- Governance Lineage Logic ---
def generate_governance_lineage(file):
    xls = pd.ExcelFile(file, engine="openpyxl")
    sheets_to_keep = ["BUSINESS RULES", "BUSINESS CONDITIONS", "GOVERNANCE MAPPING", "CONTEXTS"]
    df_rules = pd.read_excel(xls, sheet_name="BUSINESS RULES", engine="openpyxl")
    df_conditions = pd.read_excel(xls, sheet_name="BUSINESS CONDITIONS", engine="openpyxl")
    df_mapping = pd.read_excel(xls, sheet_name="GOVERNANCE MAPPING", engine="openpyxl")
    df_contexts = pd.read_excel(xls, sheet_name="CONTEXTS", engine="openpyxl")

    # Transform BUSINESS RULES
    df_rules = df_rules.rename(columns={"NAME": "RULE NAME", "DISPLAY NAME": "RULE DISPLAY NAME"})
    df_rules = df_rules[["RULE NAME", "TYPE", "DEFINITION", "RULE DISPLAY NAME", "IS ENABLED?"]]
    df_rules = df_rules[df_rules["IS ENABLED?"] == "Yes"]

    # Transform BUSINESS CONDITIONS
    df_conditions = df_conditions.rename(columns={"NAME": "CONDITION NAME", "DISPLAY NAME": "CONDITION DISPLAY NAME"})
    df_conditions = df_conditions[["CONDITION NAME", "MAPPED BUSINESS RULE(s)", "IMPACTED ROLES", "IMPACTED ATTRIBUTES", "IMPACTED RELATIONSHIPS", "CONDITION DISPLAY NAME", "IS ENABLED?"]]
    df_conditions = df_conditions[df_conditions["IS ENABLED?"] == "Yes"]

    # Transform GOVERNANCE MAPPING
    df_mapping["FOR CONTEXT"] = df_mapping["FOR CONTEXT"].astype(str)
    df_mapping = df_mapping[["ENTITY", "MAPPED BUSINESS RULE", "MAPPED BUSINESS CONDITION", "FOR CONTEXT", "IS ENABLED?"]]
    df_mapping = df_mapping[df_mapping["IS ENABLED?"] == "Yes"]

    # Transform CONTEXTS
    df_contexts = df_contexts.rename(columns={"NAME": "CONTEXT NAME", "CONTEXT TYPE || CONTEXT NAME": "CONTEXT TYPE AND NAME"})
    df_contexts = df_contexts[["CONTEXT NAME", "CONTEXT TYPE AND NAME", "WORKFLOW ACTIVITY", "WORKFLOW ACTIVITY ACTION(s)", "WORKFLOW ACTIVITY CRITERIA"]]

    lineage_records = []

    for _, rule in df_rules.iterrows():
        rule_name = rule["RULE NAME"]
        matched_conditions = df_conditions[df_conditions["MAPPED BUSINESS RULE(s)"].str.contains(rule_name, na=False)]

        if not matched_conditions.empty:
            for _, cond in matched_conditions.iterrows():
                cond_name = cond["CONDITION NAME"]
                context_keys = [f"{cond_name}{i if i > 0 else ''}Context" for i in range(16)]
                matched_mappings = df_mapping[df_mapping["FOR CONTEXT"].isin(context_keys)]
                for _, map_row in matched_mappings.iterrows():
                    context_row = df_contexts[df_contexts["CONTEXT NAME"] == map_row["FOR CONTEXT"]]
                    context_data = context_row.iloc[0].to_dict() if not context_row.empty else {col: "" for col in df_contexts.columns}
                    record = {
                        **rule[["RULE NAME", "TYPE", "DEFINITION", "RULE DISPLAY NAME"]].to_dict(),
                        **cond[["CONDITION NAME", "IMPACTED ROLES", "IMPACTED ATTRIBUTES", "IMPACTED RELATIONSHIPS", "CONDITION DISPLAY NAME"]].to_dict(),
                        **map_row[["ENTITY", "MAPPED BUSINESS RULE", "MAPPED BUSINESS CONDITION", "FOR CONTEXT"]].to_dict(),
                        **context_data
                    }
                    lineage_records.append(record)
        else:
            context_keys = [f"{rule_name}{i if i > 0 else ''}Context" for i in range(16)]
            matched_mappings = df_mapping[df_mapping["FOR CONTEXT"].isin(context_keys)]
            if not matched_mappings.empty:
                for _, map_row in matched_mappings.iterrows():
                    context_row = df_contexts[df_contexts["CONTEXT NAME"] == map_row["FOR CONTEXT"]]
                    context_data = context_row.iloc[0].to_dict() if not context_row.empty else {col: "" for col in df_contexts.columns}
                    record = {
                        **rule[["RULE NAME", "TYPE", "DEFINITION", "RULE DISPLAY NAME"]].to_dict(),
                        CONDITION_NAME="", IMPACTED_ROLES="", IMPACTED_ATTRIBUTES="", IMPACTED_RELATIONSHIPS="", CONDITION_DISPLAY_NAME="",
                        **map_row[["ENTITY", "MAPPED BUSINESS RULE", "MAPPED BUSINESS CONDITION", "FOR CONTEXT"]].to_dict(),
                        **context_data
                    }
                    lineage_records.append(record)
            else:
                fallback_mappings = df_mapping[df_mapping["MAPPED BUSINESS RULE"].str.contains(rule_name, na=False)]
                for _, map_row in fallback_mappings.iterrows():
                    record = {
                        **rule[["RULE NAME", "TYPE", "DEFINITION", "RULE DISPLAY NAME"]].to_dict(),
                        CONDITION_NAME="", IMPACTED_ROLES="", IMPACTED_ATTRIBUTES="", IMPACTED_RELATIONSHIPS="", CONDITION_DISPLAY_NAME="",
                        **map_row[["ENTITY", "MAPPED BUSINESS RULE", "MAPPED BUSINESS CONDITION"]].to_dict(),
                        FOR CONTEXT="", CONTEXT NAME="", CONTEXT TYPE AND NAME="", WORKFLOW ACTIVITY="", WORKFLOW ACTIVITY ACTION(s)="", WORKFLOW ACTIVITY CRITERIA=""
                    }
                    lineage_records.append(record)

    df_output = pd.DataFrame(lineage_records)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_output.to_excel(writer, index=False, sheet_name="Lineage")
    output.seek(0)
    return output

# --- Dynamic Authorization Lineage Logic ---
def generate_dynamic_auth_lineage(file):
    xls = pd.ExcelFile(file, engine="openpyxl")
    df_policy = pd.read_excel(xls, sheet_name="POLICY", engine="openpyxl")
    df_mapping = pd.read_excel(xls, sheet_name="POLICY MAPPING", engine="openpyxl")
    df_permissions = pd.read_excel(xls, sheet_name="POLICY PERMISSIONS", engine="openpyxl")

    df_policy = df_policy[df_policy["ENABLED"] == "Yes"]
    df_policy = df_policy[["POLICY", "ENTITY TYPE", "CONDITION"]]

    df_mapping = df_mapping.rename(columns={"POLICY": "MAPPING POLICY", "PERMISSION SET": "MAPPING PERMISSION SET"})
    df_mapping = df_mapping[["MAPPING POLICY", "ROLE", "MAPPING PERMISSION SET"]]

    lineage_records = []

    for _, policy in df_policy.iterrows():
        policy_name = policy["POLICY"]
        matched_mappings = df_mapping[df_mapping["MAPPING POLICY"].str.contains(policy_name, na=False)]
        for _, map_row in matched_mappings.iterrows():
            perm_set = map_row["MAPPING PERMISSION SET"]
            matched_permissions = df_permissions[df_permissions["PERMISSION SET"].str.contains(perm_set, na=False)]
            for _, perm_row in matched_permissions.iterrows():
                record = {
                    **policy.to_dict(),
                    ROLE=map_row["ROLE"],
                    **perm_row[["PERMISSION SET", "ATTRIBUTE", "RELATIONSHIP", "PERMISSION"]].to_dict()
                }
                lineage_records.append(record)

    df_output = pd.DataFrame(lineage_records)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_output.to_excel(writer, index=False, sheet_name="Lineage")
    output.seek(0)
    return output

# --- UI Layout ---
col1, col2 = st.columns(2)

with col1:
    st.header("Upload Governance Model")
    st.markdown("Generate data lineage document from your Governance model Excel file.")
    gov_file = st.file_uploader("Upload Governance Excel (.xlsm)", type=["xlsm"], key="gov")
    if gov_file:
        gov_output = generate_governance_lineage(gov_file)
        st.download_button("Download Governance Lineage", gov_output, file_name="Goverance_rules_lineage_output.xlsx")

with col2:
    st.header("Upload Dynamic Authorization Model")
    st.markdown("Generate data lineage document from your Dynamic Authorization model Excel file.")
    auth_file = st.file_uploader("Upload Authorization Excel (.xlsm)", type=["xlsm"], key="auth")
    if auth_file:
        auth_output = generate_dynamic_auth_lineage(auth_file)
        st.download_button("Download Authorization Lineage", auth_output, file_name="dynamic_auth_lineage_output.xlsx")
