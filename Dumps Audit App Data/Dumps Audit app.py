import asyncio
import streamlit as st
import pandas as pd
import os
import time

async def run_comparison(file_01, file_02):
    # Load both files into pandas dataframes
    file_01 = pd.ExcelFile(file_01)
    file_02 = pd.ExcelFile(file_02)

    # Get the sheet names from the file handlers, excluding the first two sheets
    sheet_names_01 = file_01.sheet_names[2:]
    sheet_names_02 = file_02.sheet_names[2:]

    # Set the primary key column
    primary_key = 'MOI'

    # Identify the rows that have differences between the two sheets
    diff_rows = []
    num_sheets = len(sheet_names_01)
    
    my_bar = st.progress(0)
    for i, sheet_name in enumerate(sheet_names_01):
        
        df1 = pd.read_excel(file_01, sheet_name=sheet_name, skiprows=[1,2,3,4])
        df2 = pd.read_excel(file_02, sheet_name=sheet_name, skiprows=[1,2,3,4])
        # Exclude the first column from common_columns
        common_columns = list(set(df1.columns[1:]).intersection(set(df2.columns[1:])))
        for j, row1 in df1.iterrows():
            key = row1[primary_key]
            row2 = df2.loc[df2[primary_key] == key, common_columns].iloc[0]       
            for col in common_columns:
                if row1[col] != row2[col]:
                    diff_rows.append([sheet_name, key, col, row1[col], row2[col]])
        sheet_progress = (i + 1) / num_sheets
        my_bar.progress(sheet_progress, text=f"Processing sheet '{sheet_name}'. Please wait...")

    my_bar.progress(100, text=f"Processing sheet '{sheet_name}'." + " All Done!")

    # Create a dataframe with the different rows
    diff_df = pd.DataFrame(diff_rows, columns=['sheet_name', primary_key, 'parameter', 'prev', 'present'])

    # Keep only the relevant columns
    output_df = diff_df[['sheet_name', primary_key, 'parameter', 'prev', 'present']]

    # Pivot the dataframe to create separate columns for previous and present values
    try:
        #output_df = output_df.pivot_table(index=['sheet_name', primary_key, 'parameter'], values=['prev', 'present'])
        output_df = output_df.groupby(['sheet_name', primary_key, 'parameter']).agg({'prev': 'first', 'present': 'last'}).reset_index()
    except ValueError as e:
        print(f"Failed to pivot table for column '{e.args[0]}'")

    # Drop the columns where the previous and present values are the same
    output_df = output_df.loc[:, (output_df != output_df.shift()).any()]
    # Filter out blank values from the "prev" and "present" columns
    output_df = output_df.dropna(subset=['prev', 'present']).reset_index()

    # Get username
    username = os.getlogin()

    # Offer the user the option to choose the location to save the output file
    output_folder = f"C:/Users/{username}/Desktop/"

    # Save the output dataframe to a file
    output_file = os.path.join(output_folder, 'Output.xlsx')
    output_df.to_excel(output_file)

    return output_file
    
async def main():
    # Create a title for the app
    st.title("ZTE Dumps Audit App")

    # Upload the files to be compared
    file_01 = st.file_uploader("Upload old file", type="xlsx")
    file_02 = st.file_uploader("Upload new file", type="xlsx")

    # Compare the files and save the output to a new file
    if st.button("Compare"):
        output_file = await run_comparison(file_01, file_02)

        # Offer the option to download the output file
        with open(output_file, "rb") as f:
            bytes_data = f.read()
            st.download_button(
                label="Download output file",
                data=bytes_data,
                file_name= "Comparison_Result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

if __name__ == '__main__':
    asyncio.run(main())

