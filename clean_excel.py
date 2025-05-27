import pandas as pd
import re
import os
import numpy as np

def clean_excel_file(input_file):
    # Read the Excel file
    df = pd.read_excel(input_file, skiprows=1)
    print(df.columns)
    
    df.drop(df[df.index >= 6][df[df.index >= 6]['Loop Component'].isna() | (df[df.index >= 6]['Loop Component'] == '')].index, inplace=True)
    #df_cleaned = df.drop_duplicates(subset=['Instrument Identification'], keep='first')
    #print(df.columns)
    print(df.shape)
    df.to_excel("output_clean2.xlsx", index=False)
    print(df)
    return df


# def merge_rows_from_index(df, start_index=3):
#     """
#     Merge pairs of rows starting from specified index
#     Handles null values properly - merges even when Loop Component is null but other columns have values
#     """
#     # Check if df is None or empty
#     if df is None:
#         raise ValueError("DataFrame is None. Please check your data loading.")
    
#     if df.empty:
#         raise ValueError("DataFrame is empty.")
    
#     # Keep the first rows unchanged (before start_index)
#     result_rows = []
    
#     # Add unchanged rows
#     for i in range(start_index):
#         if i < len(df):
#             result_rows.append(df.iloc[i].copy())
    
#     # Process pairs starting from start_index
#     i = start_index
#     while i + 1 < len(df):
#         # Get the two consecutive rows to merge
#         row1 = df.iloc[i]      # First row of the pair
#         row2 = df.iloc[i + 1]  # Second row of the pair
        
#         # Create merged row
#         merged_row = row1.copy()
        
#         # Merge each column
#         for col in df.columns:
#             # Handle NaN values properly
#             val1 = row1[col] if pd.notna(row1[col]) else None
#             val2 = row2[col] if pd.notna(row2[col]) else None
            
#             # Convert to string for concatenation, but preserve None for empty values
#             str_val1 = str(val1).strip() if val1 is not None else ''
#             str_val2 = str(val2).strip() if val2 is not None else ''
            
#             # Remove 'nan' strings that might appear
#             str_val1 = '' if str_val1.lower() == 'nan' else str_val1
#             str_val2 = '' if str_val2.lower() == 'nan' else str_val2
            
#             # Merge logic
#             if str_val1 and str_val2:
#                 # Both have values - concatenate with separator
#                 merged_row[col] = f"{str_val1} - {str_val2}"
#             elif str_val1:
#                 # Only first row has value
#                 merged_row[col] = str_val1
#             elif str_val2:
#                 # Only second row has value
#                 merged_row[col] = str_val2
#             else:
#                 # Both are empty/null - keep as NaN
#                 merged_row[col] = np.nan
        
#         result_rows.append(merged_row)
        
#         # Move to next pair (skip both rows we just merged)
#         i += 2
    
#     # If there's an unpaired row at the end, keep it as is
#     if i < len(df):
#         result_rows.append(df.iloc[i].copy())
    
#     # Create new DataFrame
#     result_df = pd.DataFrame(result_rows, columns=df.columns)
#     result_df.reset_index(drop=True, inplace=True)
    
#     return result_df

def merge_rows_from_index(df, start_index=3):
    """
    Merge pairs of rows starting from specified index
    Handles null values properly - merges even when Loop Component is null but other columns have values
    """
    # Check if df is None or empty
    if df is None:
        raise ValueError("DataFrame is None. Please check your data loading.")
    
    if df.empty:
        raise ValueError("DataFrame is empty.")
    
    # Keep the first rows unchanged (before start_index)
    result_rows = []
    
    # Add unchanged rows
    for i in range(start_index):
        if i < len(df):
            result_rows.append(df.iloc[i].copy())
    
    # Process pairs starting from start_index
    i = start_index
    while i + 1 < len(df):
        # Get the two consecutive rows to merge
        row1 = df.iloc[i]      # First row of the pair
        row2 = df.iloc[i + 1]  # Second row of the pair
        
        # Additional check: If row2 has 'Loop Component' value, 
        # then 'Unnamed: 2' should be empty, otherwise skip merge
        if (pd.notna(row2['Loop Component']) and 
            row2['Loop Component'] != '' and 
            pd.notna(row2.get('Unnamed: 2', '')) and 
            str(row2.get('Unnamed: 2', '')).strip() != ''):
            # Skip merge, move to next single row
            result_rows.append(row1.copy())
            i += 1
            continue
        
        # Create merged row
        merged_row = row1.copy()
        
        # Merge each column
        for col in df.columns:
            # Handle NaN values properly
            val1 = row1[col] if pd.notna(row1[col]) else None
            val2 = row2[col] if pd.notna(row2[col]) else None
            
            # Convert to string for concatenation, but preserve None for empty values
            str_val1 = str(val1).strip() if val1 is not None else ''
            str_val2 = str(val2).strip() if val2 is not None else ''
            
            # Remove 'nan' strings that might appear
            str_val1 = '' if str_val1.lower() == 'nan' else str_val1
            str_val2 = '' if str_val2.lower() == 'nan' else str_val2
            
            # Merge logic
            if str_val1 and str_val2:
                # Both have values - concatenate with separator
                merged_row[col] = f"{str_val1} - {str_val2}"
            elif str_val1:
                # Only first row has value
                merged_row[col] = str_val1
            elif str_val2:
                # Only second row has value
                merged_row[col] = str_val2
            else:
                # Both are empty/null - keep as NaN
                merged_row[col] = np.nan
        
        result_rows.append(merged_row)
        
        # Move to next pair (skip both rows we just merged)
        i += 2
    
    # If there's an unpaired row at the end, keep it as is
    if i < len(df):
        result_rows.append(df.iloc[i].copy())
    
    # Create new DataFrame
    result_df = pd.DataFrame(result_rows, columns=df.columns)
    result_df.reset_index(drop=True, inplace=True)
    
    return result_df

# Usage:
# df_merged = merge_rows_from_index(df, start_index=4)

input_file = r"C:\Users\isarivelan.mani\repo\rag-chatbot\backend\adnoc\0751\task03\merged_format_1.xlsx"
df = clean_excel_file(input_file)
count = df['Loop Component'].notna().sum()
print(count)
df.to_excel("check.xlsx", index=False)
result_df = merge_rows_from_index(df)
print(result_df.shape)
final_df = result_df[~result_df['Instrument Identification'].str.contains('Instrument Identification - Loop Tag', na=False)]
final_df.reset_index(drop=True, inplace=True)
final_df.to_excel("final_output6.xlsx", index=False)