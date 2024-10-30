import streamlit as st
import pandas as pd

st.title("Excel Column Processing")

# Step 1: Upload Excel file
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Step 2: Read the uploaded Excel file
    df = pd.read_excel(uploaded_file)
    
    st.write("### Original Data")
    st.write(df)
    
    # Define the column to work on
    column_name = '1_4_10'
    
    # Check if the column exists in the uploaded file
    if column_name in df.columns:
        # Step 3: Fill blank cells in the specified column with the previous cell's value
        df[column_name] = df[column_name].ffill()

        st.write("### After Filling Blank Cells")
        st.write(df)
        
        # Step 4: Merge cells with the same value in the specified column for display
        merged_df = df.copy()
        merged_df['Merged'] = (df[column_name] != df[column_name].shift()).cumsum()
        
        # Group by the 'Merged' column and aggregate other columns
        display_df = merged_df.groupby('Merged').agg(lambda x: x.iloc[0] if x.nunique() == 1 else list(x))
        
        # Convert any list-type cells to strings to avoid PyArrow issues
        display_df = display_df.applymap(lambda x: ', '.join(map(str, x)) if isinstance(x, list) else x)
        
        # Drop the 'Merged' column used for grouping
        display_df.drop(columns=['Merged'], inplace=True)

        st.write("### Merged DataFrame")
        st.write(display_df)
        
        # Optional: Download the processed file
        processed_file = df.to_excel("/mnt/data/Processed_File.xlsx", index=False)
        st.download_button(
            label="Download Processed File",
            data=processed_file,
            file_name="Processed_File.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.write(f"The column '{column_name}' was not found in the uploaded file. Please check the column name and try again.")
else:
    st.write("Please upload an Excel file to proceed.")
