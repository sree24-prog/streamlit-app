import streamlit as st
import pandas as pd

# Title of the app
st.title("Compare Two Sheets in One Excel File")

# File uploader for a single Excel file
uploaded_file = st.file_uploader("Upload the Excel file", type=["xlsx"])

# Only proceed if the file is uploaded
if uploaded_file:
    # Load the Excel file
    xls = pd.ExcelFile(uploaded_file)

    # Get all sheet names
    sheet_names = xls.sheet_names

    # Display dropdowns for sheet selection
    sheet1_name = st.selectbox("Select the first sheet (Value_SP)", sheet_names)
    sheet2_name = st.selectbox("Select the second sheet (Value_SAP)", sheet_names)

    # Load the selected sheets into DataFrames
    if sheet1_name and sheet2_name:
        df1 = pd.read_excel(uploaded_file, sheet_name=sheet1_name)
        df2 = pd.read_excel(uploaded_file, sheet_name=sheet2_name)

        # Ensure both DataFrames contain 'Attribute', 'Value_SP' in the first sheet, and 'Value_SAP' in the second sheet
        if 'Attribute' in df1.columns and 'Value_SP' in df1.columns and 'Attribute' in df2.columns and 'Value_SAP' in df2.columns:
            # Merge dataframes on 'Attribute' to compare values
            result = pd.merge(df1[['Attribute', 'Value_SP']], df2[['Attribute', 'Value_SAP']], how='outer', on='Attribute')

            # Create a 'Comparison Status' column
            result['Comparison Status'] = result.apply(lambda x: 'Pass' if x['Value_SP'] == x['Value_SAP'] else 'Fail', axis=1)

            # Search bar for filtering specific attribute (default is showing all)
            search_attribute = st.text_input("Search for a specific attribute (optional):")

            # If search is provided, filter the result
            if search_attribute:
                result = result[result['Attribute'].str.contains(search_attribute, case=False, na=False)]

            # Function to color rows based on comparison status
            def color_pass_fail(val):
                if val == 'Pass':
                    return 'background-color: lightgreen'
                elif val == 'Fail':
                    return 'background-color: lightcoral'
                else:
                    return ''

            # Apply conditional coloring to the dataframe
            styled_result = result[['Attribute', 'Value_SP', 'Value_SAP', 'Comparison Status']].style.applymap(color_pass_fail, subset=['Comparison Status'])

            # Display the comparison result with color formatting
            st.write("Comparison result between selected sheets:")
            st.dataframe(styled_result)
        else:
            st.error("Both sheets must contain 'Attribute', 'Value_SP' in the first sheet, and 'Value_SAP' in the second sheet.")
