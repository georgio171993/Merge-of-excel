import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Excel Column Processing with Validation Checks")

# Step 1: Upload Excel file
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Step 2: Read the uploaded Excel file
    df = pd.read_excel(uploaded_file)
    
    st.write("### Original Data")
    st.write(df)
    
    # Define column names for easier access
    column_1_4_10 = '1_4_10'
    column_1_5_8 = '1_5_8'
    column_1_5_9 = '1_5_9'
    column_1_5_14 = '1_5_14'
    column_1_5_14_1 = '1_5_14_1'
    column_1_5_15 = '1_5_15'
    column_1_6_3 = '1_6_3'
    column_1_6_5 = '1_6_5'
    column_1_6_4 = '1_6_4'
    column_1_6_6 = '1_6_6'
    column_1_6_7 = '1_6_7'
    column_1_6_9 = '1_6_9'
    column_1_6_8 = '1_6_8'
    column_1_6_10 = '1_6_10'
    column_1_7_11 = '1_7_11'
    column_sc_1_7_10 = 'sc_1_7_10'
    column_sc_1_7_4 = 'sc_1_7_4'
    column_1_7_3 = '1_7_3'
    column_1_7_4 = '1_7_4'
    
    # Check each condition and mark cells that don't meet the criteria
    df['Validation'] = ''  # Column to mark validation issues

    for index, row in df.iterrows():
        issues = []

        # Condition 1: If 1_5_8 - 1_5_9 > 0, 1_5_14, 1_5_14_1, and 1_5_15 should not be empty
        if (row[column_1_5_8] - row[column_1_5_9] > 0) and (pd.isna(row[column_1_5_14]) or pd.isna(row[column_1_5_14_1]) or pd.isna(row[column_1_5_15])):
            issues.append(f"{column_1_4_10} violates condition 1")

        # Condition 2: If 1_5_8 - 1_5_9 == 0, 1_5_14, 1_5_14_1, and 1_5_15 should be empty
        if (row[column_1_5_8] - row[column_1_5_9] == 0) and (not pd.isna(row[column_1_5_14]) or not pd.isna(row[column_1_5_14_1]) or not pd.isna(row[column_1_5_15])):
            issues.append(f"{column_1_4_10} violates condition 2")

        # Condition 3: If 1_6_3 has value, then 1_6_5 should not be empty
        if not pd.isna(row[column_1_6_3]) and pd.isna(row[column_1_6_5]):
            issues.append(f"{column_1_4_10} violates condition 3")

        # Condition 4: If 1_6_4 has value, then 1_6_6 should not be empty
        if not pd.isna(row[column_1_6_4]) and pd.isna(row[column_1_6_6]):
            issues.append(f"{column_1_4_10} violates condition 4")

        # Condition 5: If 1_6_7 has value, then 1_6_9 should not be empty
        if not pd.isna(row[column_1_6_7]) and pd.isna(row[column_1_6_9]):
            issues.append(f"{column_1_4_10} violates condition 5")

        # Condition 6: If 1_6_8 has value, then 1_6_10 should not be empty
        if not pd.isna(row[column_1_6_8]) and pd.isna(row[column_1_6_10]):
            issues.append(f"{column_1_4_10} violates condition 6")

        # Condition 7: If 1_7_11 has value, then sc_1_7_10 and sc_1_7_4 should not be empty
        if not pd.isna(row[column_1_7_11]) and (pd.isna(row[column_sc_1_7_10]) or pd.isna(row[column_sc_1_7_4])):
            issues.append(f"{column_1_4_10} violates condition 7")

        # Condition 8: For indirect users, if 1_7_3 is 'Yes', then 1_7_4 should not be empty
        if row[column_1_7_3] == 'Yes' and pd.isna(row[column_1_7_4]):
            issues.append(f"{column_1_4_10} violates condition 8")
        
        # Mark issues in the 'Validation' column
        if issues:
            df.at[index, 'Validation'] = "; ".join(issues)

    # Show the rows that have validation issues
    st.write("### Rows with Validation Issues")
    validation_issues = df[df['Validation'] != '']
    st.write(validation_issues)

    # Convert the DataFrame with validation highlights to an Excel file for download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Validation_Checks')
        worksheet = writer.sheets['Validation_Checks']
        format_highlight = writer.book.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        
        # Apply highlighting to cells in column 1_4_10 where there are validation issues
        for idx in validation_issues.index:
            worksheet.write(f'{column_1_4_10}{idx+2}', validation_issues.at[idx, column_1_4_10], format_highlight)
    
    # Download button
    st.download_button(
        label="Download Validation Results",
        data=output.getvalue(),
        file_name="Validation_Checks.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.write("Please upload an Excel file to proceed.")
