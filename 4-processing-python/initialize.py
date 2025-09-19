# set modules path
import sys
sys.path.append('modules')

# Import Custom Modules
from modules.run         import initialize_id_model_from_sql
from modules.sql         import query
from modules.credentials import target_db
import argparse
import os
import re

# suppress warinings from pandas
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

def main():
    #
    # Set debugging flag
    is_debugging = "1"
    os.environ["IS_DEBUGGING"] = is_debugging
    #
    print("Python: Initializing processing environment...")
    #
    # Extract Model ID from Environment Variable
    id_model = initialize_id_model_from_sql(is_debugging)
    #
    # Build SQL Statement to Extract list of Transforma
    tx_query  = f"SELECT MAX([ni_process_group]) AS [ni_process_group]\n"
    tx_query += f"FROM [dta].[process_group] AS pgp\n"
    tx_query += f"WHERE pgp.[id_model]       = '{id_model}'\n"
    tx_query += f"AND   pgp.[nm_tgt_schema] != 'mdm'\n"
    tx_query += f"AND   pgp.[nm_tgt_table]  != 'meta_attributes'"
    if (is_debugging == "1"):
        print(f"Query to extract max process group:\n{tx_query}")
    #
    # Execute the query and store result in todo
    ni_process_group = query(target_db(), tx_query).iloc[0]['ni_process_group']
    #
    # Write ni_process_group to control file
    fp_control_max_process_group = f"C:/Temp/control_max_process_group.txt"
    with open(fp_control_max_process_group, "w") as file:
        file.write(str(ni_process_group))   
    if (is_debugging == "1"):
        print(f"Control file for max process group written to: {fp_control_max_process_group} with value: {ni_process_group}")

if __name__ == "__main__":
    main()