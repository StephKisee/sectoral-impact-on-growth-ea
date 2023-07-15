# This module contains functions to clean the World Bank data
# and to create a dataframe with the data of interest.

import pandas as pd


def process_data(file_name):
    """
    Process data from an Excel file, converting it from wide to long format and storing it in a new sheet.

    Parameters:
        file_name (str): The path and name of the Excel file.git

    Returns:
        DataFrame: The processed data in a long format.

    Raises:
        PermissionError: If there are insufficient permissions to write to the specified file or directory.
    """

    # Read the Excel file and drop empty rows
    df = pd.read_excel(file_name, sheet_name='Data', skipfooter=2, na_values='..').dropna(how='all')

    # Remove unnecessary columns
    df.drop(['Country Code', 'Series Code'], axis='columns', inplace=True)

    # Convert the data from wide to long format
    df = df.melt(id_vars=['Country Name', 'Series Name'], var_name='Year', value_name='Value')

    # Pivot the table to have Series Names as columns
    df = df.pivot_table(index=['Country Name', 'Year'], columns='Series Name', values='Value').reset_index()

    # Extract the year from the 'Year' column and convert it to datetime format
    df['Year'] = df['Year'].astype(str).str[:4]
    df['Year'] = pd.to_datetime(df['Year']).dt.year

    # Set the index to 'Year' and 'Country Name'
    df.set_index('Year', inplace=True)

    try:
        # Write the processed data to a new sheet in the same Excel file
        with pd.ExcelWriter(file_name) as writer:
            df.to_excel(writer, sheet_name='Processed Data')
        print("Data written successfully to Excel file.")
    except PermissionError as e:
        print("PermissionError:", e)
        print("Make sure you have the necessary permissions to write to the specified file or directory.")

    df.to_csv('data/processed_data.csv')

    return df

#%%
