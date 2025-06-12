# Import Custom Modules
from modules import credentials as sa
from modules import run         as run

import pandas as pd
from sqlalchemy import create_engine

def load_tsl(
    
    # Input Parameters
    df_source_dataset,  # DataFrame
    nm_tsl_schema,   # Target schema name
    nm_tsl_table,    # Target table name
    
    # Debugging
    is_debugging = "0"
    
):

    # Truncate Target Table
    run.truncate_table(sa.target_db, nm_tsl_schema, nm_tsl_table)
    
    # Load Source DataFrame to SQL Schema / Table
    engine = run.engine(sa.target_db)
    result = df_source_dataset.to_sql(nm_tsl_table, con=engine, schema=nm_tsl_schema, if_exists='replace', index=False)

    # Show Input Parameter(s)
    if (is_debugging == "1"):
        print(f"nm_target_schema : '{nm_tsl_schema}'")
        print(f"nm_target_table  : '{nm_tsl_table}'")
        print(f"ni_ingested      : # {str(result)}")
        
    # return the result
    return result
        