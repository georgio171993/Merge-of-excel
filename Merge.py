import streamlit as st
import pandas as pd

# Load the Excel file
file_path = '/mnt/data/Questionnaire Answers (4).xlsx'
df = pd.read_excel(file_path)

# Display the original DataFrame
st.write("### Original Data")
st.write(df)

# Step 1: Fill blank cells in column `1_4_10` with the previous cell's value
column_name = '1_4_10'
df[column_name] = df[column_name].ffill()

# Display the DataFrame after filling blank cells
st.write("### After Filling Blank Cells")
st.write(df)

# Step 2: Merge cells with the same value in the `1_4_10` column
# Create a new DataFrame for display purposes with grouping
merged_df = df.copy()
merged_df['Merged'] = (df[column_name] != df[column_name].shift()).cumsum()

# Group by the 'Merged' column and aggregate other columns
display_df = merged_df.groupby('Merged').agg(lambda x: x.iloc[0] if x.nunique() == 1 else x.tolist())

# Drop the 'Merged' column used for grouping
display_df.drop(columns=['Merged'], inplace=True)

# Display the merged DataFrame
st.write("### Merged DataFrame")
st.write(display_df)
