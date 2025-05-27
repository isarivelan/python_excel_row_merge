# Excel Data Processor

A Python script for cleaning and merging Excel data files with customizable business rules and conditional row merging logic.

## Overview

This script performs two main operations:
1. **Data Cleaning**: Removes empty rows and invalid entries from Excel files
2. **Row Merging**: Intelligently merges consecutive rows based on specific business rules

## Features

- Cleans Excel files by removing rows with empty key columns
- Merges consecutive rows while preserving data integrity
- Handles null values and empty cells appropriately
- Applies conditional logic for selective row merging
- Exports processed data to new Excel files

## Requirements

```python
pandas
numpy
openpyxl  # For Excel file handling
```

## Installation

```bash
pip install pandas numpy openpyxl
```

## Usage

### Basic Usage

```python
from excel_processor import clean_excel_file, merge_rows_from_index

# Process your Excel file
input_file = "path/to/your/excel_file.xlsx"
df = clean_excel_file(input_file)
result_df = merge_rows_from_index(df, start_index=3)
```

### Complete Example

```python
import pandas as pd
import numpy as np

# Load and clean the Excel file
input_file = r"C:\path\to\your\data_file.xlsx"
df = clean_excel_file(input_file)

# Check primary column count
count = df['Primary_Column'].notna().sum()
print(f"Non-null Primary Column entries: {count}")

# Merge rows starting from index 3
result_df = merge_rows_from_index(df, start_index=3)

# Filter out header rows
final_df = result_df[~result_df['Identifier_Column'].str.contains('Header - Tag Pattern', na=False)]
final_df.reset_index(drop=True, inplace=True)

# Save final result
final_df.to_excel("final_output.xlsx", index=False)
```

## Functions

### `clean_excel_file(input_file)`

Cleans the input Excel file by:
- Skipping the first row during import
- Removing rows where the primary column is null or empty (starting from row 6)
- Exporting cleaned data to "output_clean2.xlsx"

**Parameters:**
- `input_file` (str): Path to the input Excel file

**Returns:**
- `pandas.DataFrame`: Cleaned DataFrame

### `merge_rows_from_index(df, start_index=3)`

Merges consecutive rows starting from a specified index with intelligent business logic.

**Parameters:**
- `df` (pandas.DataFrame): Input DataFrame to process
- `start_index` (int): Index from which to start merging rows (default: 3)

**Returns:**
- `pandas.DataFrame`: DataFrame with merged rows

**Merging Logic:**
1. Rows before `start_index` remain unchanged
2. Consecutive rows are merged in pairs
3. **Special Condition**: If the second row has a primary column value AND a non-empty secondary column value, merging is skipped
4. Column values are concatenated with " - " separator when both rows have data
5. Null values are preserved appropriately

## Data Structure

The script expects Excel files with customizable column structure. Key columns used in the example:
- `Primary_Column`: Main identifier for filtering and conditional logic
- `Identifier_Column`: Used for final filtering operations  
- `Secondary_Column`: Used in conditional merging logic
- `Description_Column`: Contains descriptive text data

## Processing Rules

### Cleaning Rules
- Rows with empty primary column (from row 6 onwards) are removed
- First row is skipped during import (assumed to be headers)

### Merging Rules
- Rows are merged in consecutive pairs
- Merging is conditional based on primary and secondary column values
- If row2 has both primary column data AND non-empty secondary column, merge is skipped
- Values from both rows are concatenated with " - " separator
- Empty/null values are handled gracefully

## Output Files

The script generates several output files:
- `output_clean2.xlsx`: Cleaned data after initial processing
- `final_output.xlsx`: Final processed data (user-defined name)

## Error Handling

The script includes error handling for:
- Empty or null DataFrames
- Missing columns
- File I/O operations

## Example Data Transformation

**Before Merging:**
```
Row 1: Primary_Column: "PC001", Description_Column: "Main Process"
Row 2: Primary_Column: null, Description_Column: "Sub Process"
```

**After Merging:**
```
Merged: Primary_Column: "PC001", Description_Column: "Main Process - Sub Process"
```

**Conditional Skip Example:**
```
Row 1: Primary_Column: "PC001", Secondary_Column: "empty"
Row 2: Primary_Column: "PC002", Secondary_Column: "data_value"
Result: No merge (skipped due to both columns having values in row2)
```

## Customization

To adapt this script for your specific use case:

1. **Update Column Names**: Replace the column references in the code:
   ```python
   # Change these column names to match your data
   PRIMARY_COLUMN = 'Your_Primary_Column_Name'
   SECONDARY_COLUMN = 'Your_Secondary_Column_Name'  
   IDENTIFIER_COLUMN = 'Your_Identifier_Column_Name'
   ```

2. **Modify Filtering Logic**: Adjust the cleaning conditions:
   ```python
   # Customize the filtering condition
   df.drop(df[df.index >= 6][df[df.index >= 6][PRIMARY_COLUMN].isna()].index, inplace=True)
   ```

3. **Adjust Conditional Logic**: Modify the merging conditions:
   ```python
   # Customize the conditional merging logic
   if (pd.notna(row2[PRIMARY_COLUMN]) and 
       pd.notna(row2[SECONDARY_COLUMN])):
       # Skip merge logic
   ```

## Configuration Options

You can customize several aspects of the processing:

- **Start Index**: Change where row merging begins
- **Skip Rows**: Modify how many initial rows to skip
- **Filter Conditions**: Adjust what constitutes an "empty" row
- **Merge Separator**: Change the " - " separator to any string
- **Output File Names**: Customize the output file names

## Notes

- The script uses generic column processing that can be adapted to various data formats
- Column names are case-sensitive and should match your Excel file exactly
- The conditional merging logic can be modified to suit different business requirements
- Make sure your Excel files have consistent column structure

## Troubleshooting

1. **KeyError for columns**: Update column names to match your Excel file structure
2. **Empty DataFrame**: Check if your input file has data beyond the header rows
3. **Merge not working**: Verify the conditional logic matches your data requirements
4. **Wrong data types**: Ensure your columns contain the expected data types

## Contributing

When modifying the script for your use case:
- Update column name constants at the top of the file
- Ensure null value handling remains consistent
- Test the conditional logic with your specific data
- Update output file names to avoid conflicts
- Document any custom business rules you implement
