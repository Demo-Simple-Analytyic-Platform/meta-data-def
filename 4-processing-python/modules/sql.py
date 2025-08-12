import pyodbc
import pandas as pd
from urllib.parse import quote  
from sqlalchemy import create_engine, text


def truncate_table(credentials_db, nm_schema, nm_table):
    
    # Build SQL Statement
    tx_sql_statement = f"TRUNCATE TABLE {nm_schema}.{nm_table}"
    
    # Execute SQL Statement
    result = execute_sql(credentials_db, tx_sql_statement)

    # Done
    return result        

def connection_string(credentials_db):
    
    # Define the connection string
    return (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"TrustServerCertificate=no;"
        f"Encrypt=no;"
        f"SERVER={credentials_db['server']};"
        f"DATABASE={credentials_db['database']};"
        f"UID={credentials_db['username']};"
        f"PWD={credentials_db['password']}"
    )

def query(credentials_db, tx_sql_statement):

    # Define the connection string
    conn_str = connection_string(credentials_db)

    # Establish the connection
    conn = pyodbc.connect(conn_str)

    # Load data into a pandas DataFrame
    df = pd.read_sql(tx_sql_statement, conn)

    # Close the connection
    conn.close()

    return df

def engine(credentials_db):

    driver   = r"ODBC Driver 17 for SQL Server"
    server   = credentials_db['server']
    database = credentials_db['database']
    username = credentials_db['username']
    password = quote(credentials_db['password'])
    encrypt  = "no"
    trustedservercertificate = "no"
    
    conn_str = f"mssql+pyodbc://{username}:{password}@{server}/{database}?driver={driver}&encrypt={encrypt}&trustedservercertificate={trustedservercertificate}"

    return create_engine(conn_str)

# This function "executes" SQL against the "Database"
def execute_sql(credentials_db, tx_sql_statement, is_debugging = "0"):
        
    with engine(credentials_db).connect() as connection:
           
        # Execute the stored procedure
        result = connection.execute(text(tx_sql_statement))
        
        if ("INSERT INTO" in tx_sql_statement) or ("UPDATE" in tx_sql_statement) or ("DELETE" in tx_sql_statement):
            # Commit the transaction if it's an INSERT, UPDATE, or DELETE statement
            connection.commit()
            connection.close()
            result = None
        
        if (is_debugging == "1") : # Show excuted "procedure"
            print(f"SQL Executed : {tx_sql_statement}")

    # Fetch results if the stored procedure returns data
    return result

# Function to execute a stored procedure
def execute_procedure(credentials_db, nm_procedure, **params):

    # Check if debugging is enabled
    if params.get('ip_is_debugging') == "1":
        print(f"Executing stored procedure: {nm_procedure}")
        print("Parameters:")
        for key, value in params.items():
            print(f"{key}: '{value}'")

    # Build the stored procedure call with parameters
    param_list = ", ".join([f"@{key} = '{value}'" for key, value in params.items()])
    stored_procedure = f"EXEC {nm_procedure} {param_list}"
    if params.get('ip_is_debugging') == "1":
        print(f"Stored Procedure Call: {stored_procedure}")
    
    # Execute the stored procedure
    with engine(credentials_db).connect() as connection:
        with connection.connection.cursor() as cursor:
            result = connection.execute(text(stored_procedure))
            # result = cursor.execute(stored_procedure)
            
    # Done
    return result

def execute_procedure2(credentials_db, nm_procedure, **params):
    conn_str = connection_string(credentials_db)
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    # Remove special keys not meant for the procedure
    proc_params = {k: v for k, v in params.items() if k.startswith('ip_')}

    # Build parameter placeholders and values
    placeholders = ', '.join(['?' for _ in proc_params])
    values = list(proc_params.values())

    # Build the procedure call string
    if placeholders:
        call_str = f"{{CALL {nm_procedure}({placeholders})}}"
    else:
        call_str = f"{{CALL {nm_procedure}}}"

    if params.get('ip_is_debugging') == "1":
        print(f"Executing stored procedure: {nm_procedure}")
        print("Parameters:", proc_params)
        print(f"Call string: {call_str}")

    cursor.execute(call_str, values)

    # If the procedure returns results
    try:
        result = cursor.fetchall()
    except pyodbc.ProgrammingError:
        result = None

    conn.commit()
    cursor.close()
    conn.close()
    return result