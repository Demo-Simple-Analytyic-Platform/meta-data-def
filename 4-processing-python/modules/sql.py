from multiprocessing import connection
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

def odbc_connection_string(credentials_db):
    return f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={credentials_db['server']};DATABASE={credentials_db['database']};UID={credentials_db['username']};PWD={credentials_db['password']};TrustServerCertificate=no;Encrypt=no;"

def connection_string(credentials_db):
    
    # Enhanced connection string with better settings for log file operations
    return (
        r"DRIVER={ODBC Driver 17 for SQL Server};"
        f"TrustServerCertificate=yes;"  # Changed to yes for better compatibility
        f"Encrypt=no;"                  # Keep as no for internal networks
        f"SERVER={credentials_db['server']};"
        f"DATABASE={credentials_db['database']};"
        f"UID={credentials_db['username']};"
        f"PWD={credentials_db['password']};"
        f"APP=PythonDataPipeline;"      # Application name for monitoring
        f"ConnectionTimeout=30;"        # Connection timeout
        f"CommandTimeout=300;"          # Command timeout (5 minutes)
        f"AutoTranslate=no;"           # Prevent data type issues
        f"QuotedId=yes;"               # Enable quoted identifiers
        f"ANSI_NULLS=yes;"             # Enable ANSI NULL handling
        f"Isolation=READ_COMMITTED;"
    )

def query(credentials_db, tx_sql_statement):

    # SQLAlchemy engine using pyodbc
    engine = create_engine(
        f"mssql+pyodbc://{credentials_db['username']}:{quote(credentials_db['password'])}@{credentials_db['server']}/{credentials_db['database']}?driver=ODBC+Driver+17+for+SQL+Server&Encrypt=no&TrustServerCertificate=yes"
    )

    conn_str = connection_string(credentials_db)

    # Establish the connection
    #conn = pyodbc.connect(conn_str)

    # Load data into a pandas DataFrame
    #df = pd.read_sql(tx_sql_statement, conn)
    df = pd.read_sql(tx_sql_statement, engine)

    # Close the connection
    #conn.close()

    return df

def engine(credentials_db):

    driver   = r"ODBC Driver 17 for SQL Server"
    server   = credentials_db['server']
    database = credentials_db['database']
    username = credentials_db['username']
    password = quote(credentials_db['password'])
    
    # Enhanced connection parameters for better transaction management
    params = {
        'driver': driver,
        'TrustServerCertificate': 'yes',
        'Encrypt': 'no',
        'APP': 'PythonDataPipeline',
        'ConnectionTimeout': '30',
        'CommandTimeout': '300',
        'AutoTranslate': 'no',
        'QuotedId': 'yes',
        'ANSI_NULLS': 'yes',
        'TransactionIsolation': 'READ_COMMITTED'  # Added parameter for transaction isolation level
    }
    
    # Build parameter string
    param_str = '&'.join([f"{k}={v}" for k, v in params.items()])
    
    conn_str = f"mssql+pyodbc://{username}:{password}@{server}/{database}?{param_str}"
    
    # Create engine with enhanced settings for transaction management
    return create_engine(
        conn_str #, 
        #pool_size=5,            # Connection pool size
        #max_overflow=10,        # Max overflow connections
        #pool_timeout=30,        # Pool timeout
        #pool_recycle=3600,      # Recycle connections every hour
        #echo=False,             # Set to True for SQL debugging
        #isolation_level="READ_COMMITTED"  # Set isolation level
    )

# This function "executes" SQL against the "Database"
def execute_sql(credentials_db, tx_sql_statement, is_debugging = "0"):
        
    try:

        # Build SQL Connection String for pyodbc
        tx_connections = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={credentials_db['server']};DATABASE={credentials_db['database']};UID={credentials_db['username']};PWD={credentials_db['password']};TrustServerCertificate=no;Encrypt=no;"

        # Create a new connection
        conn = pyodbc.connect(tx_connections,autocommit=True)

        # Create a new cursor
        cursor = conn.cursor()

        # Execute the stored procedure
        result = cursor.execute(tx_sql_statement)

        # Close the connection
        conn.close()
                
    except Exception as e:
        error_msg = f"Error executing SQL: {e}"

        if (is_debugging == "1"):
            print(error_msg)
        raise Exception(error_msg)

    # Fetch results if the SQL statement returns data
    return result

# Function to execute a stored procedure
def execute_procedure(credentials_db, nm_procedure, **params):

    # Check if debugging is enabled
    if params['is_debugging'] == "1":
        print(f"Executing stored procedure: {nm_procedure}")
        print("Parameters:")
        for key, value in params.items():
            print(f"{key}: '{value}'")

    # Build the stored procedure call with parameters, exclude is_debugging, using pyodbc formatting
    param_list = ", ".join([f"@{key} = ?" for key in params if key != 'is_debugging'])
    param_values = [params[key] for key in params if key != 'is_debugging']
    stored_procedure = "{" + f"CALL {nm_procedure}({param_list})" + "}"

    # Build SQL Connection String for pyodbc
    tx_connections = odbc_connection_string(credentials_db)

    # Create a new connection
    conn = pyodbc.connect(tx_connections,autocommit=True)

    # Create a new cursor
    cursor = conn.cursor()

    # Execute the stored procedure
    result = cursor.execute(stored_procedure, param_values)

    # Close the connection
    conn.close()

    # Done
    return result
