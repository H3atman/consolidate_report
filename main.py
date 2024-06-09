import streamlit as st
import pandas as pd
from io import BytesIO




def read_all_sheets(file):
    xls = pd.ExcelFile(file)
    dfs = [pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names]
    return dfs


def main():
    st.set_page_config("Excel Sheet Consolidator")
    st.title("Excel Sheet Consolidator")
    st.write("Upload multiple Excel workbooks, and this app will consolidate all sheets into one workbook and one sheet.")
    
    uploaded_files = st.file_uploader("Choose Excel files", type="xlsx", accept_multiple_files=True)
    
    if uploaded_files:
        consolidated_data = []

        for file in uploaded_files:
            sheets = read_all_sheets(file)
            for sheet in sheets:
                consolidated_data.append(sheet)
        
        if consolidated_data:
            consolidated_df = pd.concat(consolidated_data, ignore_index=True)
            st.write("Consolidation Complete! Here's a preview:")
            st.write(consolidated_df.head())

            # Convert the consolidated dataframe to Excel format
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                consolidated_df.to_excel(writer, index=False, sheet_name='Consolidated')
            output.seek(0)

            st.download_button(
                label="Download Consolidated Workbook",
                data=output,
                file_name="consolidated_workbook.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )

if __name__ == "__main__":
    main()


