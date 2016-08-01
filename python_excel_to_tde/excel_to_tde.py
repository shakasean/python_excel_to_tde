# Import modules
import sys, os, time, datetime, locale
import pandas as pd
from tableausdk import *
from tableausdk.Extract import *

# Validate Excel input
if len(sys.argv) < 2:
    raise NameError('Excel filename argument missing.')
excel_file = sys.argv[1]
tde_file = excel_file.split('.')[0] + '.tde'

# Read and process Excel file
df = pd.read_excel(excel_file)

type_obj = []
for i in df.select_dtypes(include=['object']).columns:
    type_obj.append(i)

type_non_obj = []
for i in df.select_dtypes(exclude=['object']).columns:
    type_non_obj.append(i)

obj_dict = {x: None for x in type_obj}
non_obj_dict = {x: None for x in type_non_obj}

df = df.replace(obj_dict,'na')
df = df.replace(non_obj_dict,0)

# STEP 1 - Initialize the Extract API
ExtractAPI.initialize()

# STEP 2 - Initialize a new extract by calling the Extract() constructor
if os.path.isfile(tde_file):
    os.remove(tde_file)

td_extract = Extract(tde_file)

# STEP 3 - Create a table definition
table_definition = TableDefinition()

# STEP 4 - Create new column definition
for i in range(len(df.columns)):
    if df.dtypes[i] == 'object':
        table_definition.addColumn(df.columns[i], Type.UNICODE_STRING)
    elif df.dtypes[i] == 'float64':
        table_definition.addColumn(df.columns[i], Type.DOUBLE)
    elif df.dtypes[i] == 'int64':
        table_definition.addColumn(df.columns[i], Type.INTEGER)
    elif df.dtypes[i] == 'datetime64[ns]':
        table_definition.addColumn(df.columns[i], Type.DATE)

# STEP 5 - Initialize a new table in the extract with the addTable() method
new_table = td_extract.addTable('Extract', table_definition)

# STEP 6 - Create a new row with the Row() constructor
new_row = Row(table_definition)

# STEP 7 - Populate each new row
for j in range(0, df.shape[0]):
    for i in range(len(df.columns)):
        if df.dtypes[i] == 'object':
            new_row.setString(i, df.iloc[j,i])
        if df.dtypes[i] == 'float64':
            new_row.setDouble(i, df.iloc[j,i])
        elif df.dtypes[i] == 'int64':
            new_row.setInteger(i, df.iloc[j,i])
        elif df.dtypes[i] == 'datetime64[ns]':
            new_row.setDate(i, df.iloc[j,i].year, df.iloc[j,i].month, df.iloc[j,i].day)
    new_table.insert(new_row)

# STEP 8 - Save table, extract, and performs cleanup
td_extract.close()

# STEP 9 - Release the Extract API. ONLY if ExtractAPI.initialize() was used
ExtractAPI.cleanup()


