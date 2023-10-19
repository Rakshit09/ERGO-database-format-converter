'''
@Author: Rakshit Joshi
@Company: Gallagher Re, Munich

This is a Python script that performs conversion of modelling results into the ERGO specific database format
1. Import Necessary Libraries:
  - The script starts by importing various Python libraries, including `pandas`, `datetime`, `re`, `numpy`, and `xlwings`.
2. Function Definitions:
  - The code defines three functions: `process_data_48`, `process_data_50`, and `process_data_65`. These functions are 
     used to process data from Excel sheets with specific column lengths (48, 50, and more than 64). 
        *For results with Average Annual Loss only, the column length is 48.
        *For results with Average Annual Loss+ Standard Deviation+ Coefficient of Variation, the column length is 50.
        *For results that have Net Pre Cat data, the column length is 65.
        NOTE: The column length is dependednt on the row number of the string "AAL Checks"
3. Main Execution (`if __name__ == "__main__":`):
 - The code enters the main execution block when run as a standalone script. It sets various variables such 
     as `source`, `destination`, and `year`, specifying the source and destination file paths and the target year for 
     data processing.
 - The code defines two dictionaries, `column_mapping` and `columns_in_table`, which are used to rename and select 
     specific columns in the resulting DataFrame.
 - It opens the source Excel file using `xlwings` and prepares the destination workbook.
 - The code iterates through sheets with names ending in "AEP" or "OEP" in the source workbook. For each sheet, it reads 
     data into a DataFrame, determines the column length of the data, and calls one of the processing functions 
     (`process_data_48`, `process_data_50`, or `process_data_65`) based on the column length.
 - The processed data is then appended to the "Processed Data" sheet in the destination workbook. If the sheet is empty, 
      it starts writing from cell A1; otherwise, it appends data below the existing content.

 - Additionally, it copies the "Cover Page" and "Disclaimers" sheets from the source workbook to the destination workbook.
 - Finally, it saves the destination workbook and quits the Excel application.

'''

import pandas as pd
from datetime import datetime
import re
import numpy as np
import xlwings as xw
import os
import time
# Rows in excel sheet are 2 ahead.
#Idea1:  Aks user to input row numbers for Business unit (3), comments(23) exposure (31). The last is determined by column length  

def process_data_48(dataframe, year_to_find, setup):

    def contains_year(cell, year_to_find):
        return isinstance(cell, datetime) and cell.year == year_to_find

    values_dict_column_b = {}
    values_dict_column_d = {}
    column_b = dataframe.iloc[:, 1].astype(str)
    column_d = dataframe.iloc[:, 3].astype(str)

    year_indices = []
    row_number=26
    for column_index, column in enumerate(dataframe.columns):
        if dataframe[column].apply(contains_year, year_to_find=year_to_find).any():
            year_indices.append(column_index-1)
            for index, value in enumerate(column_b):
                if not pd.isna(value) and 3 <= index <= 25:
                    variable_name = re.sub(r'\W+', '_', value)
                    values_dict_column_b.setdefault(variable_name, []).append(dataframe.iloc[index, column_index])
            for index, value in enumerate(column_d):
                if not pd.isna(value) and 31 <= index <= 45:
                    variable_name = re.sub(r'\W+', '_', value)
                    values_dict_column_d.setdefault(variable_name, []).append(dataframe.iloc[index, column_index])

    
    #Portfolio = (dataframe.iloc[[row_number], year_indices]).values.flatten()                 
    Portfolio = (df.iloc[[row_number], year_indices]).fillna(method='ffill', axis=1).values.flatten() #fills an empty value with the last value
    
    result_dataframe = pd.DataFrame.from_dict(values_dict_column_b, orient='index').T
    result_dataframe['Portfolio'] = Portfolio
    
    lines_of_businesses = len(result_dataframe)

    result_dataframe_column_d = pd.DataFrame.from_dict(values_dict_column_d, orient='index').T
    if setup == 'AEP':
        desired_order = ['1000', '500', '250', '200', '100', '50', '25', '10', '5']
    else: 
        desired_order = ['1000', '500', '250', '200', '100', '50', '25', '10', '5', 'Exposure', 'Modelled_Exposure', 'Average_Annual_Loss']
    
    column_d_reordered = {key: values_dict_column_d[key] for key in desired_order}

    repeat_count = len(column_d_reordered)

    repeated_data = [result_dataframe.loc[[index]].reindex([index] * repeat_count) for index in result_dataframe.index]
    repeated_dataframe = pd.concat(repeated_data, ignore_index=True)

    Return_Period = list(column_d_reordered.keys()) * len(result_dataframe)
    Value = np.array(list(column_d_reordered.values())).T.ravel()

    if len(Return_Period) != len(Value):
        raise ValueError("Error: Length of column Return Period is not equal to length of Value")
    if len(Return_Period) != len(column_d_reordered) * len(result_dataframe):
        raise ValueError("Error: Length of variables does not match the number of lines of businesses")
    if len(repeated_dataframe) != len(Value):
        raise ValueError("Error: Length of repeated data and return periods don't match")

    dataframe_final = repeated_dataframe.copy()
    dataframe_final['Return Period'] = Return_Period
    dataframe_final['Value'] = Value
    dataframe_final.to_csv('output.txt', sep='\t', index=False)


    dataframe_final_renamed = dataframe_final.rename(columns=column_mapping)
    dataframe_final_renamed2 = dataframe_final_renamed[columns_in_table]
    return dataframe_final_renamed2

def process_data_50(dataframe, year_to_find, setup):
    
    def contains_year(cell, year_to_find):
        return isinstance(cell, datetime) and cell.year == year_to_find

    values_dict_column_b = {}
    values_dict_column_d = {}
    column_b = dataframe.iloc[:, 1].astype(str)
    column_d = dataframe.iloc[:, 3].astype(str)

    year_indices = []
    row_number=26
    
    for column_index, column in enumerate(dataframe.columns):
        if dataframe[column].apply(contains_year, year_to_find=year_to_find).any():
            year_indices.append(column_index-1)
            for index, value in enumerate(column_b):
                if not pd.isna(value) and 3 <= index <= 25:
                    variable_name = re.sub(r'\W+', '_', value)
                    values_dict_column_b.setdefault(variable_name, []).append(dataframe.iloc[index, column_index])
            for index, value in enumerate(column_d):
                if not pd.isna(value) and 31 <= index <= 47:
                    variable_name = re.sub(r'\W+', '_', value)
                    values_dict_column_d.setdefault(variable_name, []).append(dataframe.iloc[index, column_index])

    
    Portfolio = (df.iloc[[row_number], year_indices]).fillna(method='ffill', axis=1).values.flatten()
    result_dataframe = pd.DataFrame.from_dict(values_dict_column_b, orient='index').T
    result_dataframe['Portfolio'] = Portfolio
    
    lines_of_businesses = len(result_dataframe)

    result_dataframe_column_d = pd.DataFrame.from_dict(values_dict_column_d, orient='index').T
    
    if setup == 'AEP':
        desired_order = ['1000', '500', '250', '200', '100', '50', '25', '10', '5']
    else: 
        desired_order = ['1000', '500', '250', '200', '100', '50', '25', '10', '5', 'Exposure', 'Modelled_Exposure', 'Average_Annual_Loss', 'Standard_Deviation', 'Coefficient_of_Variation']

    

    column_d_reordered = {key: values_dict_column_d[key] for key in desired_order}

    repeat_count = len(column_d_reordered)

    repeated_data = [result_dataframe.loc[[index]].reindex([index] * repeat_count) for index in result_dataframe.index]
    repeated_dataframe = pd.concat(repeated_data, ignore_index=True)

    Return_Period = list(column_d_reordered.keys()) * len(result_dataframe)
    Value = np.array(list(column_d_reordered.values())).T.ravel()

    if len(Return_Period) != len(Value):
        raise ValueError("Error: Length of column Return Period is not equal to length of Value")
    if len(Return_Period) != len(column_d_reordered) * len(result_dataframe):
        raise ValueError("Error: Length of variables does not match the number of lines of businesses")
    if len(repeated_dataframe) != len(Value):
        raise ValueError("Error: Length of repeated data and return periods don't match")

    dataframe_final = repeated_dataframe.copy()
    dataframe_final['Return Period'] = Return_Period
    dataframe_final['Value'] = Value

    
    dataframe_final_renamed = dataframe_final.rename(columns=column_mapping)
    dataframe_final_renamed2 = dataframe_final_renamed[columns_in_table]
    return dataframe_final_renamed2



def process_data_65(dataframe, year_to_find, setup):
    processed_data1 = process_data_48(df, year_to_find, setup)
    df2 = df.copy() 
    df2.iloc[35:46] = df.iloc[55:66] 
    processed_data2 = process_data_48(df2, year_to_find, setup)
    processed_data2['Measure'] = 'Net Pre CAT'

    concatenated_data = pd.concat([processed_data1, processed_data2], ignore_index=True)
    return  concatenated_data



if __name__ == "__main__":
    source = r"C:\Users\rajoshi\Desktop\Modelling_Results\2024\Gallagher Re - ERGO 2024 renewal modelling - results PL.xlsx"
    destination = r"C:\Users\rajoshi\Desktop\Modelling_Results\2024\test.xlsx"
    year = 2023

    global column_mapping
    column_mapping = {
        'Business_Unit_BU_': 'Business Unit',
        'incl_Subperil': 'incl Subperil',
        'Country_modelled_': 'Country modelled',
        'Date_of_Portfolio': 'Date of Portfolio',
        'Measure_Perspective': 'Perspective',
        'Exchange_Rate': 'Exchange Rate',
        'Data_Supplier': 'Data Supplier',
        'NatCat_Model': 'NatCat Model',
        'Model_Version': 'Model Version',
        'Post_loss_amplification': 'Post Loss Amplification',
        'Original_adjusted': 'original/adjusted'}

    global columns_in_table
    columns_in_table= ['Business Unit', 'Peril', 'incl Subperil', 'Portfolio',  'original/adjusted', 'Modelling_ID', 'Country modelled', 'Date of Portfolio',
                        'Perspective', 'Measure', 'Return Period', 'Value', 'Currency', 'Exchange Rate',
                      'Data Supplier', 'Modeler', 'NatCat Model', 'Model Version', 'Post Loss Amplification', 'Comments']
    
    wb1 = xw.Book(source)

    # Define the destination path and ensure the destination file is removed if it exists
    if os.path.exists(destination):
        os.remove(destination)

    # Create a new destination workbook
    wb2 = xw.Book()

    # Create a single destination sheet for all processed data
    destination_sheet = wb2.sheets.add()
    destination_sheet.name = "Processed Data"  # You can name it as needed

    # List of sheet names to copy from source to destination
    sheets_to_copy = [sheet.name for sheet in wb1.sheets if (sheet.visible and (sheet.name.rstrip().endswith("AEP") or sheet.name.rstrip().endswith("OEP")))]

# Iterate through the sheets and append processed data to the destination sheet
    year = 2023
    c=0
    length =0
    for sheet_name in sheets_to_copy:
        
        if (sheet_name.rstrip().endswith("AEP")):
            sheet_type='AEP'    
        else:    
            sheet_type='OEP'       

        print("Processing " + sheet_name + "....")
        search_word = "AAL"
        df = pd.read_excel(source, sheet_name)  # Read sheet data into a DataFrame
        row_number = df[df.apply(lambda row: row.astype(str).str.contains(search_word).any(), axis=1)].index.max()+1
        column_d_length = len(df.iloc[:, 3].astype(str))
        if row_number == 48:
            processed_data = process_data_48(df, year_to_find=year, setup = sheet_type)
            print("48")
            length+=len(processed_data)
            print("Done....")
            c+=1
        elif row_number == 50:
            processed_data = process_data_50(df, year_to_find=year, setup = sheet_type)
            print("50")
            length+=len(processed_data)
            print("Done....")
            c+=1
        elif row_number == 68:
            processed_data = process_data_65(df, year_to_find=year, setup = sheet_type)
            print("68")
            length+=len(processed_data)
            print("Done....")
            c+=1
        
        if destination_sheet.range('A1').value is None:
            # If the first cell is empty, start writing data from A1
                destination_sheet.range('A1').options(index=False).value = processed_data
        else:
                last_row = len(destination_sheet.range('A1').expand('table').value)
                processed_data.columns = [" "] * len(processed_data.columns)
                destination_sheet.range((last_row + 1, 1)).options(index=False).value = processed_data.rename_axis(None, axis=1)
     
       # time.sleep(0.5) 
    if (c==len(sheets_to_copy)):
            print("All sheets processed :)")
     
    print(length)                          
                              
    # Copy a sheet with a name starting with "Cover" to the second position (1-based index) in the target workbook
    for sheet in wb1.sheets:
        if re.match(r'^Cover', sheet.name):
            ws1 = sheet
            break

    if ws1 is not None:
        ws1.api.Copy(Before=wb2.sheets(1).api)

    # Copy a sheet with a name starting with "Disclaim" to the last position in the target workbook
    for sheet in wb1.sheets:
        if re.match(r'^Disclaim', sheet.name):
            sheet.api.Copy(Before=wb2.sheets[-1].api)     
    
    # Save the destination workbook
    wb2.save(destination)
    wb2.app.quit()