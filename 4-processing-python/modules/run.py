# Import Custom Modules
from modules import credentials as sa
from modules import source      as src
from modules import target      as tgt
from modules import run         as run

from azure.storage.blob import BlobServiceClient
from datetime           import datetime as dt
from sqlalchemy         import create_engine, text
from urllib.parse       import quote

import pandas as pd
import pyodbc
import os
from azure.storage.blob import ContentSettings


def folder_exists(folder_path):
    if os.path.isdir(folder_path):
        print(f"The folder '{folder_path}' exists.")
        return True
    else:
        print(f"The folder '{folder_path}' does not exist.")
        return False

def create_folder(folder_path):
    try:
        os.makedirs(folder_path, exist_ok=True)  # `exist_ok=True` prevents errors if the folder already exists
        print(f"Folder '{folder_path}' created successfully.")
    except Exception as e:
        print(f"Error creating folder '{folder_path}': {e}")


def update_dataset(ds_external_reference_id, id_dataset, is_ingestion, nm_procedure, nm_tsl_schema, nm_tsl_table, is_debugging):
    
    # Local Vairables
    result = False

    try:
    
        # If "Ingestion" first extract "source" data
        if is_ingestion == 1:

            # for "Ingestion" the run must be started, if "Transformation" the run is started in the "procedure" itself.
            start(id_dataset, is_debugging, ds_external_reference_id)

            # Get the parameters
            params = get_parameters(id_dataset)
            
            # Paramters
            cd_parameter_group = params.loc[0]['cd_parameter_group']

            # Switch for cd_parameter_group
            if cd_parameter_group == 'web_table_anonymous_web':

                # Get Ingestion specific parameters
                wtb_1_any_ds_url   = get_param_value('wtb_1_any_ds_url', params)
                wtb_2_any_ds_path  = get_param_value('wtb_2_any_ds_path', params)
                wtb_3_any_ni_index = get_param_value('wtb_3_any_ni_index', params)

                # load source to dataframe
                source_df = src.web_table_anonymous_web(wtb_1_any_ds_url, wtb_2_any_ds_path, wtb_3_any_ni_index, is_debugging)

            elif cd_parameter_group == 'abs_sas_url_csv':

                # Get Ingestion specific parameters
                abs_1_csv_nm_account         = get_param_value('abs_1_csv_nm_account', params)
                abs_2_csv_nm_secret          = get_param_value('abs_2_csv_nm_secret', params)
                abs_3_csv_nm_container       = get_param_value('abs_3_csv_nm_container', params)
                abs_4_csv_ds_folderpath      = get_param_value('abs_4_csv_ds_folderpath', params)
                abs_5_csv_ds_filename        = get_param_value('abs_5_csv_ds_filename', params)
                abs_6_csv_nm_decode          = get_param_value('abs_6_csv_nm_decode', params)
                abs_7_csv_is_1st_header      = get_param_value('abs_7_csv_is_1st_header', params)
                abs_8_csv_cd_delimiter_value = get_param_value('abs_8_csv_cd_delimiter_value', params)
                abs_9_csv_cd_delimter_text   = get_param_value('abs_9_csv_cd_delimter_text', params)

                # load source to dataframe
                source_df = src.abs_sas_url_csv(abs_1_csv_nm_account, abs_2_csv_nm_secret, abs_3_csv_nm_container, abs_4_csv_ds_folderpath, abs_5_csv_ds_filename, abs_6_csv_nm_decode, abs_7_csv_is_1st_header, abs_8_csv_cd_delimiter_value, abs_9_csv_cd_delimter_text, is_debugging)

            elif cd_parameter_group == 'abs_sas_url_xls':

                # Get Ingestion specific parameters
                abs_1_xls_nm_account           = get_param_value('abs_1_xls_nm_account', params)
                abs_2_xls_nm_secret            = get_param_value('abs_2_xls_nm_secret', params)
                abs_3_xls_nm_container         = get_param_value('abs_3_xls_nm_container', params)
                abs_4_xls_ds_folderpath        = get_param_value('abs_4_xls_ds_folderpath', params)
                abs_5_xls_ds_filename          = get_param_value('abs_5_xls_ds_filename', params)
                abs_6_xls_nm_sheet             = get_param_value('abs_6_xls_nm_sheet', params)
                abs_7_xls_is_first_header      = get_param_value('abs_7_xls_is_first_header', params)
                abs_8_xls_cd_top_left_cell     = get_param_value('abs_8_xls_cd_top_left_cell', params)
                abs_9_xls_cd_bottom_right_cell = get_param_value('abs_9_xls_cd_bottom_right_cell', params)

                # load source to dataframe
                source_df = src.abs_sas_url_xls(abs_1_xls_nm_account, abs_2_xls_nm_secret, abs_3_xls_nm_container, abs_4_xls_ds_folderpath, abs_5_xls_ds_filename, abs_6_xls_nm_sheet, abs_7_xls_is_first_header, abs_8_xls_cd_top_left_cell, abs_9_xls_cd_bottom_right_cell, is_debugging)

            elif cd_parameter_group == 'sql_user_password':

                # Get Ingestion specific parameters
                sql_1_nm_server   = get_param_value('sql_1_nm_server', params)
                sql_2_nm_username = get_param_value('sql_2_nm_username', params)
                sql_3_nm_secret   = get_param_value('sql_3_nm_secret', params)
                sql_4_nm_database = get_param_value('sql_4_nm_database', params)
                sql_5_tx_query    = get_param_value('sql_5_tx_query', params)

                # load source to dataframe
                source_df = src.sql_user_password(sql_1_nm_server, sql_2_nm_username, sql_3_nm_secret, sql_4_nm_database, sql_5_tx_query, is_debugging)
                
            else:
                raise ValueError(f"Unsupported cd_parameter_group: {cd_parameter_group}")
            
            # Load "Source"-dataframe to "Temporal Staging Landing"-table.
            tgt.load_tsl(source_df, nm_tsl_schema, nm_tsl_table, is_debugging)
            
            # Start sql procedure specific for the "Target"-dataset on database side.
            usp_dataset_ingestion(nm_procedure, is_debugging)
        
        # If "Transformation" start the run and the procedure
        else:
            usp_dataset_transformation(nm_procedure, ds_external_reference_id)
    
        # If everything is done, return True
        result = True

    except Exception as e:

        print(f"Error occurred: {e}")
        result = False
    
    
        # All is well
        return result

def data_pipeline(nm_target_schema, nm_target_table, is_debugging):
    
    # Build SQL for Query
    tx_query = f"SELECT ni_process_group, id_dataset, is_ingestion, nm_procedure, nm_tsl_schema, nm_tsl_table, nm_tgt_schema, nm_tgt_table "\
             + f"FROM dta.process_group "\
             + f"WHERE nm_tgt_schema = '{nm_target_schema}'"\
             + f"AND   nm_tgt_table  = '{nm_target_table}'"\
             + f"ORDER BY ni_process_group ASC"
    
    # fetch all dataset tobe processed
    todo = query(sa.target_db, tx_query)

    # External Reference ID
    ds_external_reference_id = 'python-'+todo.loc[0]['id_dataset']+dt.now().strftime('%Y%m%d%H%M%S')

    # Parameter for "update_dataset"
    id_dataset    = todo.loc[0]['id_dataset']  
    is_ingestion  = todo.loc[0]['is_ingestion'] 
    nm_procedure  = todo.loc[0]['nm_procedure'] 
    nm_tsl_schema = todo.loc[0]['nm_tsl_schema'] 
    nm_tsl_table  = todo.loc[0]['nm_tsl_table']
    nm_tgt_schema = todo.loc[0]['nm_tgt_schema'] 
    nm_tgt_table  = todo.loc[0]['nm_tgt_table']

    if (is_debugging == "1"): # Show what dataset is being processed
        print("--- " + ("Ingestion ----" if (is_ingestion == 1) else "Transformation ") + "--------------------------------------")
        print(f"ds_external_reference_id : '{ds_external_reference_id}'")
        print(f"id_dataset               : '{id_dataset}'")
        print(f"nm_tgt_schema            : '{nm_tgt_schema}'")
        print(f"nm_tgt_table             : '{nm_tgt_table}'")
        print(f"nm_procedure             : '{nm_procedure}'")  
        print(f"nm_tsl_schema            : '{nm_tsl_schema}'")
        print(f"nm_tsl_table             : '{nm_tsl_table}'")
        print("")

    # Update dataset "NVIDIA Corporation (NVDA)"
    attempt = 0
    result  = False
    while (result == False and attempt < 3):
        
        if (is_debugging == "1"):
            print(f"Attempt {attempt+1} to update dataset...")
        
        # Call the function to update the dataset
        result = update_dataset(ds_external_reference_id, id_dataset, is_ingestion, nm_procedure, nm_tsl_schema, nm_tsl_table, is_debugging)    
        
        # Add 1 to the attempt counter
        attempt += 1

    # export documentation for dataset
    documentation = export_documentation(id_dataset, is_debugging)
    
    print("all done")
    
def export_documentation(id_dataset, is_debugging):
    
    # Build SQL for Query
    tx_query  = f"SELECT f.ds_file_path"
    tx_query += f"\n     , f.nm_file_name"
    tx_query += f"\n     , t.ni_line"
    tx_query += f"\n     , tx_line"
    tx_query += f"\nFROM mdm.html_file_name AS f" 
    tx_query += f"\nJOIN mdm.html_file_text AS t ON t.id_dataset= f.id_dataset" 
    tx_query += f"\nWHERE f.id_dataset = '{id_dataset}'"
    
    # Fetch the data
    df = query(sa.target_db, tx_query)
    tx = df['tx_line']

    # Generate HTML content
    tx_content_data = "\n".join(tx)
    cd_content_type = "text/html"

    # Define file path and name
    ds_filepath_blob  = df.loc[0]['ds_file_path'] + df.loc[0]['nm_file_name']
    ds_filepath_blob  = ds_filepath_blob.replace('\\', r'/')
    ds_temp_folder    = "C:/Temp"
    ds_filepath_local = f"{ds_temp_folder}/{ds_filepath_blob}"
    ds_folderpath_local = ds_filepath_local.replace("/" + df.loc[0]['nm_file_name'], "")

    # check if folders exist
    if folder_exists(ds_temp_folder) == False:
        create_folder(ds_temp_folder)
        
    # check if folders exist
    if folder_exists(ds_folderpath_local) == False:
        create_folder(ds_folderpath_local)
        
    # Write HTML content to a file
    with open(ds_filepath_local, "w", encoding="utf-8") as file:
        file.write(tx_content_data)

    # Upload the file to Azure Blob Storage
    abs_1_nm_account   = sa.blob_documentation['account']
    abs_2_cd_accesskey = get_secret(sa.blob_documentation['secret'], is_debugging)
    abs_3_nm_container = sa.blob_documentation['container']

    # Define the connection string and the blob details
    tx_connection_string = f"DefaultEndpointsProtocol=https;AccountName={abs_1_nm_account};AccountKey={abs_2_cd_accesskey};EndpointSuffix=core.windows.net"

    # Create the BlobServiceClient object
    blob_service_client = BlobServiceClient.from_connection_string(tx_connection_string)

    # Create the BlobClient object
    blob_client = blob_service_client.get_blob_client(container=abs_3_nm_container, blob=ds_filepath_blob)

    with open(ds_filepath_local, "rb") as data:
        blob_client.upload_blob(
            data,
            overwrite=True,
            content_settings=ContentSettings(content_type=cd_content_type)
        )

    if is_debugging == "1":
        print(f"HTML file '{ds_filepath_blob}' uploaded to Azure Blob Storage container '{abs_3_nm_container}'.")

def get_parameters(id_dataset):

    # Define the query
    tx_sql_statement  = f"SELECT * FROM rdp.tvf_get_parameters('{id_dataset}')\n"

    # Load data into a pandas DataFrame
    return query(sa.target_db, tx_sql_statement)

def get_secret(nm_secret, is_debugging):

    # Show input Parameter(s)
    if (is_debugging == "1"):
        print("nm_secret : '" + nm_secret + "'")

    # Build SQL Statement
    tx_query = f"SELECT ds_secret FROM dbo.secrets WHERE nm_secret = '{nm_secret}'"
    
    # Run SQL query
    df = query(sa.secret_db, tx_query)

    # Show the result
    return None if len(df) == 0 else df['ds_secret'].iloc[0]

def get_param_value(nm_parameter_value, params):
    return params.loc[params['nm_parameter_value'] == nm_parameter_value].values[0][3]

def start(id_dataset, is_debugging, ds_external_reference_id):
      
    # Execute "rdp.run_start"
    return execute_procedure(sa.target_db, 'rdp.run_start',\
        ip_id_dataset_or_dq_control = id_dataset,\
        ip_ds_external_reference_id = ds_external_reference_id,\
        ip_is_debugging             = is_debugging\
    )

def usp_dataset_ingestion(nm_procedure, is_debugging):

    # Build the stored procedure call with parameters
    stored_procedure = f"EXEC {nm_procedure}"

    # Show excuted "procedure"
    if (is_debugging == "1") :
        print(f"Executing stored procedure: {nm_procedure}")

    # Execute the stored procedure
    with engine(sa.target_db).connect() as connection:
        with connection.connection.cursor() as cursor:
            result = cursor.execute(stored_procedure)
            
    # Done
    return result

def usp_dataset_transformation(nm_procedure, ds_external_reference_id):

    return execute_procedure(sa.target_db, nm_procedure, ip_ds_external_reference_id = ds_external_reference_id)

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

    driver   = r"ODBC+Driver+17+for+SQL+Server"
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

    # Execute the stored procedure
    with engine(credentials_db).connect() as connection:
        with connection.connection.cursor() as cursor:
            result = cursor.execute(stored_procedure)
            
    # Done
    return result
