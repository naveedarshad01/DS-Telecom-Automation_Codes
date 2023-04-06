import zipfile
import os
import pandas as pd
from glob import glob
import streamlit as st
import datetime

def process_data(working_directory, progress_bar):
    # Extract all zip files in the working directory
    zip_files = [f for f in working_directory if zipfile.is_zipfile(f)]
    for i, file in enumerate(zip_files):
        with zipfile.ZipFile(file) as item:
            item.extractall()
        prog=0
        progress_bar.progress(prog+10)
    # Read and concatenate all CSV files in the working directory
    df0 = sorted(glob('*.csv'))
    df1=pd.concat((pd.read_csv(file,skiprows=0,skipfooter=0,engine='python',\
                                      na_values=['NIL','/0']\
                                     ,usecols=['Start Time','ManagedElement Name','Cell Name',\
                                               'eNodeB Name', \
                                               'User DL Average Throughput(Mbps)_SLA','[FDD]DL PRB Utilization Rate','Mean Number of RRC Connection User','4G PS TRAFFIC (GB)'])\
                                                for file in df0),ignore_index=True).fillna(0)
    progress_bar.progress(prog+20)
    # Extract date and time from Start Time column
    df1['Date'] = pd.to_datetime(df1['Start Time']).dt.date
    df1['Time'] = pd.to_datetime(df1['Start Time']).dt.time

    # Convert columns to appropriate data types
    df1['User DL Average Throughput(Mbps)_SLA'] = df1['User DL Average Throughput(Mbps)_SLA'].astype('float64')
    df1['Mean Number of RRC Connection User'] = df1['Mean Number of RRC Connection User'].astype('float64')
    df1['[FDD]DL PRB Utilization Rate'] = df1['[FDD]DL PRB Utilization Rate'].str.rstrip('%').astype('float') 
    df1['4G PS TRAFFIC (GB)'] = df1['4G PS TRAFFIC (GB)'].astype('float') 
    progress_bar.progress(prog+20)
    # Filter high utilization cells
    HU_Cells=df1[(df1['User DL Average Throughput(Mbps)_SLA']<=3)&(df1['Mean Number of RRC Connection User']>=37)&(df1['[FDD]DL PRB Utilization Rate']>=85)]

    # Group by Cell Name and Date, and count the number of HU hours
    df2=HU_Cells.groupby('Cell Name')['Date'].value_counts().unstack().reset_index()
    df2['Count of HU Hours']=df2.iloc[0:,1:9].sum(axis=1)

    # Fill missing values with 0
    df2.fillna(0)
    progress_bar.progress(prog+100)
    # Get username
    username = os.getlogin()
    # Generate a timestamp for the output file name
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    # Offer the user the option to choose the location to save the output file
    output_folder = f"C:/Users/{username}/Desktop/HU Output Folder"
    output_file = os.path.join(output_folder, f'High_Utilization_cells_data {timestamp}.xlsx')

    # Save the output dataframe to a file
    df2.to_excel(output_file)
    
    return output_file

# Define the Streamlit app
st.title("High Utilization Cells Data")

# File selector for choosing the directory
working_directory = st.file_uploader(
    "Choose a directory", type="zip", accept_multiple_files=True
)

# Button for processing the data
if st.button("Process Data"):
    if working_directory is not None:
        progress_bar = st.progress(0)
        save_path = process_data(working_directory, progress_bar)
        st.write("Processed data is saved to:", save_path)
    else:
        st.warning("Please choose a directory before processing the data.")
