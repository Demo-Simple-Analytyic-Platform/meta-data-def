import sys
sys.path.append('modules')

# Import Custom Modules
from modules.run         import initialize_id_model_from_sql

import argparse
import os
import re
import time
import random

# Import Custom Modules
from modules.run         import data_pipeline
from modules.sql         import query
from modules.credentials import target_db


def main():
    #
    # Set debugging flag
    is_debugging = os.getenv("IS_DEBUGGING", "0")
    print("Python: Starting parallel processing session...")
    #
    # Parse command line arguments
    try:
        parser = argparse.ArgumentParser(description="Parallel session runner")
        parser.add_argument("--ni-process-group", type=int, required=True, help="# Process Group")
        parser.add_argument("--ni-sessions",      type=int, required=True, help="Total # parallel sessions")
        parser.add_argument("--id-session",       type=int, required=True, help="Unique ID for this session (1 to ni-sessions)")
        args = parser.parse_args()
        #
        # Validate arguments
        if (args.ni_sessions < 1) or (args.id_session > args.ni_sessions):
            raise ValueError("id-session must be between 1 and ni-sessions (inclusive)")
        #
        # Set local variables
        ni_process_group = args.ni_process_group
        ni_sessions      = args.ni_sessions
        id_session       = args.id_session
    
    except Exception as e:
        ni_sessions = 1
        id_session = 1
        print(f"Error parsing arguments, using defaults ni-sessions={ni_sessions}, id-session={id_session}")    
    #
    # Run the session
    print(f"Running session {id_session} of {ni_sessions}")
    #
    # Extract Model ID from Environment Variable
    id_model = initialize_id_model_from_sql(is_debugging)
    #
    # Build SQL Statement for fetching the list of target schema/table to be processed by this session
    tx_query  = f"SELECT pgp.nm_target_schema, pgp.nm_target_table\n"
    tx_query += f"FROM [dta].[parallel_proces_group] ('{id_model}', {ni_process_group}, {ni_sessions}, {id_session}) AS pgp\n"
    if (is_debugging == "1"):
        print(f"Query to extract max process group:\n{tx_query}")
    #
    # Execute the query and store result in todo
    todo = query(target_db(), tx_query)
    #
    # Loop though the list of schema/table to be processed
    for dst in todo.itertuples(index=False):
        #
        # Extract schema/table to be processed
        nm_target_schema = dst.nm_target_schema
        nm_target_table  = dst.nm_target_table
        print("")
        print(f"Processing {nm_target_schema}.{nm_target_table}...")
        #
        # Run the processing for this schema/table
        result = data_pipeline(id_model, nm_target_schema, nm_target_table, is_debugging)
        #
        # Success or failure message
        if result:
            print(f"Successfully processed {nm_target_schema}.{nm_target_table}.")
        else:
            print(f"Failed to process {nm_target_schema}.{nm_target_table}.")
        print("")

    # End for loop
    print(f"Session {id_session} of {ni_sessions} completed.")
    
    #
    # Remove control file for this session if it exists
    fp_control = f"C:/Temp/control_parallel_{id_session}.txt"
    if os.path.exists(fp_control):
        os.remove(fp_control)
    #
    # All Done
    
if __name__ == "__main__":
    main()