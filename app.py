
import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

# Set Streamlit page configuration
st.set_page_config(page_title="Lineage Generator", layout="wide")

# Page title
st.title("Lineage Document Generator")

# Create two columns for Governance and Authorization sections
col1, col2 = st.columns(2)

# Function to write Excel file with filters and editable mode
def write_excel_with_filters(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Lineage"

    # Write headers
    ws.append(list(df.columns))

    # Write data rows
    for row in df.itertuples(index=False):
        ws.append(list(row))

    # Apply filter as a table
    table_range = f"A1:{chr(65 + len(df.columns) - 1)}{len(df) + 1}"
    table = Table(displayName="LineageTable", ref=table_range)
    style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    table.tableStyleInfo = style
    ws.add_table(table)

    # Save to BytesIO
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Governance Model Section
with col1:
    st.header("Upload Governance Model")
    st.markdown("Generate data lineage document from your Governance model Excel file.")
    gov_file = st.file_uploader("Upload Governance Excel (.xlsm)", key="gov")

    if gov_file:
        xls = pd.ExcelFile(gov_file, engine='openpyxl')
        sheets_to_keep = ['BUSINESS RULES', 'BUSINESS CONDITIONS', 'GOVERNANCE MAPPING', 'CONTEXTS']
        sheets = {sheet: xls.parse(sheet) for sheet in sheets_to_keep if sheet in xls.sheet_names}

        # BUSINESS RULES
        br = sheets['BUSINESS RULES'].rename(columns={'NAME': 'RULE NAME', 'DISPLAY NAME': 'RULE DISPLAY NAME'})
        br = br[['RULE NAME', 'TYPE', 'DEFINITION', 'RULE DISPLAY NAME', 'IS ENABLED?']]
        br = br[br['IS ENABLED?'] == 'Yes']

        # BUSINESS CONDITIONS
        bc = sheets['BUSINESS CONDITIONS'].rename(columns={'NAME': 'CONDITION NAME', 'DISPLAY NAME': 'CONDITION DISPLAY NAME'})
        bc = bc[['CONDITION NAME', 'MAPPED BUSINESS RULE(s)', 'IMPACTED ROLES', 'IMPACTED ATTRIBUTES', 'IMPACTED RELATIONSHIPS', 'CONDITION DISPLAY NAME', 'IS ENABLED?']]
        bc = bc[bc['IS ENABLED?'] == 'Yes']

        # GOVERNANCE MAPPING
        gm = sheets['GOVERNANCE MAPPING']
        gm['FOR CONTEXT'] = gm['FOR CONTEXT'].astype(str)
        gm = gm[['ENTITY', 'MAPPED BUSINESS RULE', 'MAPPED BUSINESS CONDITION', 'FOR CONTEXT', 'IS ENABLED?']]
        gm = gm[gm['IS ENABLED?'] == 'Yes']

        # CONTEXTS
        ctx = sheets['CONTEXTS'].rename(columns={'NAME': 'CONTEXT NAME', 'CONTEXT TYPE || CONTEXT NAME': 'CONTEXT TYPE AND NAME'})
        ctx = ctx[['CONTEXT NAME', 'CONTEXT TYPE AND NAME', 'WORKFLOW ACTIVITY', 'WORKFLOW ACTIVITY ACTION(s)', 'WORKFLOW ACTIVITY CRITERIA']]

        lineage_records = []

        for _, rule in br.iterrows():
            rule_name = rule['RULE NAME']
            matched_conditions = bc[bc['MAPPED BUSINESS RULE(s)'].str.contains(rule_name, na=False)]

            if not matched_conditions.empty:
                for _, cond in matched_conditions.iterrows():
                    cond_name = cond['CONDITION NAME']
                    context_keys = [f"{cond_name}{i if i > 0 else ''}Context" for i in range(16)]
                    matched_gm = gm[gm['FOR CONTEXT'].isin(context_keys)]

                    if not matched_gm.empty:
                        for _, gm_row in matched_gm.iterrows():
                            matched_ctx = ctx[ctx['CONTEXT NAME'] == gm_row['FOR CONTEXT']]
                            if not matched_ctx.empty:
                                for _, ctx_row in matched_ctx.iterrows():
                                    lineage_records.append({
                                        'RULE NAME': rule['RULE NAME'],
                                        'TYPE': rule['TYPE'],
                                        'DEFINITION': rule['DEFINITION'],
                                        'RULE DISPLAY NAME': rule['RULE DISPLAY NAME'],
                                        'CONDITION NAME': cond['CONDITION NAME'],
                                        'IMPACTED ROLES': cond['IMPACTED ROLES'],
                                        'IMPACTED ATTRIBUTES': cond['IMPACTED ATTRIBUTES'],
                                        'IMPACTED RELATIONSHIPS': cond['IMPACTED RELATIONSHIPS'],
                                        'CONDITION DISPLAY NAME': cond['CONDITION DISPLAY NAME'],
                                        'ENTITY': gm_row['ENTITY'],
                                        'MAPPED BUSINESS RULE': gm_row['MAPPED BUSINESS RULE'],
                                        'MAPPED BUSINESS CONDITION': gm_row['MAPPED BUSINESS CONDITION'],
                                        'FOR CONTEXT': gm_row['FOR CONTEXT'],
                                        'CONTEXT NAME': ctx_row['CONTEXT NAME'],
                                        'CONTEXT TYPE AND NAME': ctx_row['CONTEXT TYPE AND NAME'],
                                        'WORKFLOW ACTIVITY': ctx_row['WORKFLOW ACTIVITY'],
                                        'WORKFLOW ACTIVITY ACTION(s)': ctx_row['WORKFLOW ACTIVITY ACTION(s)'],
                                        'WORKFLOW ACTIVITY CRITERIA': ctx_row['WORKFLOW ACTIVITY CRITERIA']
                                    })
            else:
                context_keys = [f"{rule_name}{i if i > 0 else ''}Context" for i in range(16)]
                matched_gm = gm[gm['FOR CONTEXT'].isin(context_keys)]

                if not matched_gm.empty:
                    for _, gm_row in matched_gm.iterrows():
                        matched_ctx = ctx[ctx['CONTEXT NAME'] == gm_row['FOR CONTEXT']]
                        if not matched_ctx.empty:
                            for _, ctx_row in matched_ctx.iterrows():
                                lineage_records.append({
                                    'RULE NAME': rule['RULE NAME'],
                                    'TYPE': rule['TYPE'],
                                    'DEFINITION': rule['DEFINITION'],
                                    'RULE DISPLAY NAME': rule['RULE DISPLAY NAME'],
                                    'CONDITION NAME': "",
                                    'IMPACTED ROLES': "",
                                    'IMPACTED ATTRIBUTES': "",
                                    'IMPACTED RELATIONSHIPS': "",
                                    'CONDITION DISPLAY NAME': "",
                                    'ENTITY': gm_row['ENTITY'],
                                    'MAPPED BUSINESS RULE': gm_row['MAPPED BUSINESS RULE'],
                                    'MAPPED BUSINESS CONDITION': gm_row['MAPPED BUSINESS CONDITION'],
                                    'FOR CONTEXT': gm_row['FOR CONTEXT'],
                                    'CONTEXT NAME': ctx_row['CONTEXT NAME'],
                                    'CONTEXT TYPE AND NAME': ctx_row['CONTEXT TYPE AND NAME'],
                                    'WORKFLOW ACTIVITY': ctx_row['WORKFLOW ACTIVITY'],
                                    'WORKFLOW ACTIVITY ACTION(s)': ctx_row['WORKFLOW ACTIVITY ACTION(s)'],
                                    'WORKFLOW ACTIVITY CRITERIA': ctx_row['WORKFLOW ACTIVITY CRITERIA']
                                })
                else:
                    fallback_gm = gm[gm['MAPPED BUSINESS RULE'].str.contains(rule_name, na=False)]
                    for _, gm_row in fallback_gm.iterrows():
                        lineage_records.append({
                            'RULE NAME': rule['RULE NAME'],
                            'TYPE': rule['TYPE'],
                            'DEFINITION': rule['DEFINITION'],
                            'RULE DISPLAY NAME': rule['RULE DISPLAY NAME'],
                            'CONDITION NAME': "",
                            'IMPACTED ROLES': "",
                            'IMPACTED ATTRIBUTES': "",
                            'IMPACTED RELATIONSHIPS': "",
                            'CONDITION DISPLAY NAME': "",
                            'ENTITY': gm_row['ENTITY'],
                            'MAPPED BUSINESS RULE': gm_row['MAPPED BUSINESS RULE'],
                            'MAPPED BUSINESS CONDITION': gm_row['MAPPED BUSINESS CONDITION'],
                            'FOR CONTEXT': "",
                            'CONTEXT NAME': "",
                            'CONTEXT TYPE AND NAME': "",
                            'WORKFLOW ACTIVITY': "",
                            'WORKFLOW ACTIVITY ACTION(s)': "",
                            'WORKFLOW ACTIVITY CRITERIA': ""
                        })

        if lineage_records:
            df_output = pd.DataFrame(lineage_records)
            excel_data = write_excel_with_filters(df_output)
            st.download_button("Download Governance Lineage Document", data=excel_data, file_name="Goverance_rules_lineage_output.xlsx")

# Dynamic Authorization Section
with col2:
    st.header("Upload Dynamic Authorization Model")
    st.markdown("Generate data lineage document from your Dynamic Authorization model Excel file.")
    auth_file = st.file_uploader("Upload Authorization Excel (.xlsm)", key="auth")

    if auth_file:
        xls = pd.ExcelFile(auth_file, engine='openpyxl')
        sheets_to_keep = ['POLICY', 'POLICY MAPPING', 'POLICY PERMISSIONS']
        sheets = {sheet: xls.parse(sheet) for sheet in sheets_to_keep if sheet in xls.sheet_names}

        # POLICY
        policy = sheets['POLICY']
        policy = policy[['POLICY', 'ENTITY TYPE', 'CONDITION', 'ENABLED']]
        policy = policy[policy['ENABLED'] == 'Yes']

        # POLICY MAPPING
        mapping = sheets['POLICY MAPPING'].rename(columns={'POLICY': 'MAPPING POLICY', 'PERMISSION SET': 'MAPPING PERMISSION SET'})
        mapping = mapping[['MAPPING POLICY', 'ROLE', 'MAPPING PERMISSION SET']]

        # POLICY PERMISSIONS
        permissions = sheets['POLICY PERMISSIONS']

        lineage_records = []

        for _, pol in policy.iterrows():
            pol_name = pol['POLICY']
            matched_map = mapping[mapping['MAPPING POLICY'].str.contains(pol_name, na=False)]

            for _, map_row in matched_map.iterrows():
                perm_set = map_row['MAPPING PERMISSION SET']
                matched_perm = permissions[permissions['PERMISSION SET'].str.contains(perm_set, na=False)]

                for _, perm_row in matched_perm.iterrows():
                    lineage_records.append({
                        'POLICY': pol['POLICY'],
                        'ENTITY TYPE': pol['ENTITY TYPE'],
                        'CONDITION': pol['CONDITION'],
                        'ROLE': map_row['ROLE'],
                        'PERMISSION SET': perm_row['PERMISSION SET'],
                        'ATTRIBUTE': perm_row['ATTRIBUTE'],
                        'RELATIONSHIP': perm_row['RELATIONSHIP'],
                        'PERMISSION': perm_row['PERMISSION']
                    })

        if lineage_records:
            df_output = pd.DataFrame(lineage_records)
            excel_data = write_excel_with_filters(df_output)
            st.download_button("Download Authorization Lineage Document", data=excel_data, file_name="dynamic_auth_lineage_output.xlsx")
