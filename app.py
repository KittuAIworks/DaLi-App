import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Lineage Generator", layout="wide")

# Title and description
st.title("Data Lineage Document Generator")
st.markdown("Upload your governance or dynamic authorization model Excel file to generate lineage documents.")

# Create two columns for side-by-side upload sections
col1, col2 = st.columns(2)

# ------------------- Governance Lineage Section -------------------
with col1:
    st.header("Upload Governance Model")
    st.markdown("Generate data lineage document from your Governance model Excel file.")
    gov_file = st.file_uploader("Upload Governance Excel (.xlsm)", type=["xlsm"], key="gov")

    if gov_file:
        # Load sheets
        xls = pd.ExcelFile(gov_file, engine="openpyxl")
        sheets_to_keep = ["BUSINESS RULES", "BUSINESS CONDITIONS", "GOVERNANCE MAPPING", "CONTEXTS"]
        sheets = {sheet: xls.parse(sheet) for sheet in sheets_to_keep if sheet in xls.sheet_names}

        # BUSINESS RULES
        br = sheets["BUSINESS RULES"].rename(columns={"NAME": "RULE NAME", "DISPLAY NAME": "RULE DISPLAY NAME"})
        br = br[["RULE NAME", "TYPE", "DEFINITION", "RULE DISPLAY NAME", "IS ENABLED?"]]
        br = br[br["IS ENABLED?"] == "Yes"]

        # BUSINESS CONDITIONS
        bc = sheets["BUSINESS CONDITIONS"].rename(columns={"NAME": "CONDITION NAME", "DISPLAY NAME": "CONDITION DISPLAY NAME"})
        bc = bc[["CONDITION NAME", "MAPPED BUSINESS RULE(s)", "IMPACTED ROLES", "IMPACTED ATTRIBUTES", "IMPACTED RELATIONSHIPS", "CONDITION DISPLAY NAME", "IS ENABLED?"]]
        bc = bc[bc["IS ENABLED?"] == "Yes"]

        # GOVERNANCE MAPPING
        gm = sheets["GOVERNANCE MAPPING"]
        gm["FOR CONTEXT"] = gm["FOR CONTEXT"].astype(str)
        gm = gm[["ENTITY", "MAPPED BUSINESS RULE", "MAPPED BUSINESS CONDITION", "FOR CONTEXT", "IS ENABLED?"]]
        gm = gm[gm["IS ENABLED?"] == "Yes"]

        # CONTEXTS
        ctx = sheets["CONTEXTS"].rename(columns={"NAME": "CONTEXT NAME", "CONTEXT TYPE || CONTEXT NAME": "CONTEXT TYPE AND NAME"})
        ctx = ctx[["CONTEXT NAME", "CONTEXT TYPE AND NAME", "WORKFLOW ACTIVITY", "WORKFLOW ACTIVITY ACTION(s)", "WORKFLOW ACTIVITY CRITERIA"]]

        # Lineage logic
        lineage_records = []

        for _, rule_row in br.iterrows():
            rule_name = rule_row["RULE NAME"]
            matched_conditions = bc[bc["MAPPED BUSINESS RULE(s)"].str.contains(rule_name, na=False)]

            matched = False

            if not matched_conditions.empty:
                for _, cond_row in matched_conditions.iterrows():
                    cond_name = cond_row["CONDITION NAME"]
                    context_keys = [f"{cond_name}{'' if i == 0 else i}Context" for i in range(16)]
                    matched_gm = gm[gm["FOR CONTEXT"].isin(context_keys)]

                    for _, gm_row in matched_gm.iterrows():
                        context_name = gm_row["FOR CONTEXT"]
                        ctx_row = ctx[ctx["CONTEXT NAME"] == context_name]
                        if not ctx_row.empty:
                            lineage_records.append({
                                **rule_row.drop("IS ENABLED?").to_dict(),
                                **cond_row.drop("IS ENABLED?").to_dict(),
                                **gm_row.drop("IS ENABLED?").to_dict(),
                                **ctx_row.iloc[0].to_dict()
                            })
                            matched = True
            if not matched:
                context_keys = [f"{rule_name}{'' if i == 0 else i}Context" for i in range(16)]
                matched_gm = gm[gm["FOR CONTEXT"].isin(context_keys)]
                if not matched_gm.empty:
                    for _, gm_row in matched_gm.iterrows():
                        context_name = gm_row["FOR CONTEXT"]
                        ctx_row = ctx[ctx["CONTEXT NAME"] == context_name]
                        if not ctx_row.empty:
                            lineage_records.append({
                                **rule_row.drop("IS ENABLED?").to_dict(),
                                "CONDITION NAME": "", "IMPACTED ROLES": "", "IMPACTED ATTRIBUTES": "", "IMPACTED RELATIONSHIPS": "", "CONDITION DISPLAY NAME": "",
                                **gm_row.drop("IS ENABLED?").to_dict(),
                                **ctx_row.iloc[0].to_dict()
                            })
                            matched = True
            if not matched:
                fallback_gm = gm[gm["MAPPED BUSINESS RULE"].str.contains(rule_name, na=False)]
                for _, gm_row in fallback_gm.iterrows():
                    lineage_records.append({
                        **rule_row.drop("IS ENABLED?").to_dict(),
                        "CONDITION NAME": "", "IMPACTED ROLES": "", "IMPACTED ATTRIBUTES": "", "IMPACTED RELATIONSHIPS": "", "CONDITION DISPLAY NAME": "",
                        "ENTITY": gm_row["ENTITY"], "MAPPED BUSINESS RULE": gm_row["MAPPED BUSINESS RULE"], "MAPPED BUSINESS CONDITION": gm_row["MAPPED BUSINESS CONDITION"], "FOR CONTEXT": "",
                        "CONTEXT NAME": "", "CONTEXT TYPE AND NAME": "", "WORKFLOW ACTIVITY": "", "WORKFLOW ACTIVITY ACTION(s)": "", "WORKFLOW ACTIVITY CRITERIA": ""
                    })

        if lineage_records:
            df_output = pd.DataFrame(lineage_records)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_output.to_excel(writer, index=False, sheet_name="Lineage")
            st.download_button("Download Governance Lineage Document", data=output.getvalue(), file_name="Goverance_rules_lineage_output.xlsx")

# ------------------- Dynamic Authorization Lineage Section -------------------
with col2:
    st.header("Upload Dynamic Authorization Model")
    st.markdown("Generate data lineage document from your Dynamic Authorization model Excel file.")
    auth_file = st.file_uploader("Upload Authorization Excel (.xlsm)", type=["xlsm"], key="auth")

    if auth_file:
        xls = pd.ExcelFile(auth_file, engine="openpyxl")
        sheets_to_keep = ["POLICY", "POLICY MAPPING", "POLICY PERMISSIONS"]
        sheets = {sheet: xls.parse(sheet) for sheet in sheets_to_keep if sheet in xls.sheet_names}

        # POLICY
        policy = sheets["POLICY"]
        policy = policy[["POLICY", "ENTITY TYPE", "CONDITION", "ENABLED"]]
        policy = policy[policy["ENABLED"] == "Yes"]

        # POLICY MAPPING
        mapping = sheets["POLICY MAPPING"].rename(columns={"POLICY": "MAPPING POLICY", "PERMISSION SET": "MAPPING PERMISSION SET"})
        mapping = mapping[["MAPPING POLICY", "ROLE", "MAPPING PERMISSION SET"]]

        # POLICY PERMISSIONS
        permissions = sheets["POLICY PERMISSIONS"]

        lineage_records = []

        for _, pol_row in policy.iterrows():
            pol_name = pol_row["POLICY"]
            matched_map = mapping[mapping["MAPPING POLICY"].str.contains(pol_name, na=False)]

            for _, map_row in matched_map.iterrows():
                perm_set = map_row["MAPPING PERMISSION SET"]
                matched_perm = permissions[permissions["PERMISSION SET"].str.contains(perm_set, na=False)]

                for _, perm_row in matched_perm.iterrows():
                    lineage_records.append({
                        "POLICY": pol_row["POLICY"],
                        "ENTITY TYPE": pol_row["ENTITY TYPE"],
                        "CONDITION": pol_row["CONDITION"],
                        "ROLE": map_row["ROLE"],
                        "PERMISSION SET": perm_row["PERMISSION SET"],
                        "ATTRIBUTE": perm_row["ATTRIBUTE"],
                        "RELATIONSHIP": perm_row["RELATIONSHIP"],
                        "PERMISSION": perm_row["PERMISSION"]
                    })

        if lineage_records:
            df_output = pd.DataFrame(lineage_records)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_output.to_excel(writer, index=False, sheet_name="Lineage")
            st.download_button("Download Authorization Lineage Document", data=output.getvalue(), file_name="dynamic_auth_lineage_output.xlsx")
