# Import Custom Modules
from modules import credentials as sa
from modules import source      as src
from modules.fso import folder_exists, create_folder
from modules.sql import query, execute_procedure, engine, truncate_table, execute_sql

# Import Libraries
from azure.storage.blob import BlobServiceClient
from datetime           import datetime as dt
import pandas as pd

from azure.storage.blob import ContentSettings

def data_pipeline(id_model, nm_target_schema, nm_target_table, is_debugging):
    
    # Build SQL for Query
    tx_query = f"SELECT ni_process_group, id_dataset, is_ingestion, nm_procedure, nm_tsl_schema, nm_tsl_table, nm_tgt_schema, nm_tgt_table "\
             + f"FROM dta.process_group "\
             + f"WHERE nm_tgt_schema = '{nm_target_schema}'"\
             + f"AND   nm_tgt_table  = '{nm_target_table}'"\
             + f"AND   id_model      = '{id_model}' "\
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
        print(f"id_model                 : '{id_model}'")
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
        result = update_dataset(id_model, ds_external_reference_id, id_dataset, is_ingestion, nm_procedure, nm_tsl_schema, nm_tsl_table, is_debugging)    
        
        # Add 1 to the attempt counter
        attempt += 1

    # export documentation for dataset
    documentation = export_documentation(id_dataset, is_debugging)
    
    print("all done")
    
def update_dataset(id_model, ds_external_reference_id, id_dataset, is_ingestion, nm_procedure, nm_tsl_schema, nm_tsl_table, is_debugging):
    
    # Local Vairables
    result = False

    try:
    
        # If "Ingestion" first extract "source" data
        if is_ingestion == 1:

            # for "Ingestion" the run must be started, if "Transformation" the run is started in the "procedure" itself.
            start(id_model, id_dataset, ds_external_reference_id, is_debugging)

            # Get the parameters
            params = get_parameters(id_model, id_dataset)
            
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
                sql_3_nm_database = get_param_value('sql_3_nm_database', params)
                sql_6_nm_secret   = get_param_value('sql_6_nm_secret', params)              
                sql_5_tx_query    = get_param_value('sql_7_tx_query', params)

                # load source to dataframe
                source_df = src.sql_user_password(sql_1_nm_server, sql_2_nm_username, sql_6_nm_secret, sql_3_nm_database, sql_5_tx_query, is_debugging)
                
            else:
                raise ValueError(f"Unsupported cd_parameter_group: {cd_parameter_group}")
            
            # Load "Source"-dataframe to "Temporal Staging Landing"-table.
            load_tsl(source_df, nm_tsl_schema, nm_tsl_table, is_debugging)
            
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

def load_tsl(
    
    # Input Parameters
    df_source_dataset,  # DataFrame
    nm_tsl_schema,   # Target schema name
    nm_tsl_table,    # Target table name
    
    # Debugging
    is_debugging = "0"
    
):

    # Truncate Target Table
    truncate_table(sa.target_db, nm_tsl_schema, nm_tsl_table)
    
    # Load Source DataFrame to SQL Schema / Table
    sql_engine = engine(sa.target_db)
    result = df_source_dataset.to_sql(nm_tsl_table, con=sql_engine, schema=nm_tsl_schema, if_exists='replace', index=False)

    # Show Input Parameter(s)
    if (is_debugging == "1"):
        print(f"nm_target_schema : '{nm_tsl_schema}'")
        print(f"nm_target_table  : '{nm_tsl_table}'")
        print(f"ni_ingested      : # {str(result)}")
        
    # return the result
    return result

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

def get_parameters(id_model, id_dataset):

    # Define the query
    tx_sql_statement  = f"SELECT * FROM rdp.tvf_get_parameters('{id_model}', '{id_dataset}')\n"

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
   
def start(id_model, ip_id_dataset_or_dq_control, ds_external_reference_id, is_debugging = "0"):
    
    # /* Local Variables. */
    dt_run_started = dt.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
    
    # /* Local Varaibles for "Starting" run of "Dataset" or "DQ Control". */
    id_run        = query(sa.target_db, f"SELECT id_run        = LOWER(CONVERT(CHAR(32),HASHBYTES('MD5',CONCAT(CONVERT(NVARCHAR(MAX),''), '|', '{id_model}', '|', '{ip_id_dataset_or_dq_control}', '|', '{dt_run_started}', '|')), 2))").iloc[0]['id_run']
    id_dataset    = query(sa.target_db, f"SELECT id_dataset    = ISNULL((SELECT id_dataset    FROM dta.dataset    WHERE meta_is_active = 1 AND id_dataset    = '{ip_id_dataset_or_dq_control}'), 'n/a')").iloc[0]['id_dataset']
    id_dq_control = query(sa.target_db, f"SELECT id_dq_control = ISNULL((SELECT id_dq_control FROM dqm.dq_control WHERE meta_is_active = 1 AND id_dq_control = '{ip_id_dataset_or_dq_control}'), 'n/a')").iloc[0]['id_dq_control']
    
    # /* Local Variables for "Extraction" or " Processing Infromation". */
    nm_target_schema = query(sa.target_db, f"SELECT nm_target_schema FROM dta.dataset WHERE meta_is_active = 1 AND id_dataset = '{ip_id_dataset_or_dq_control}'").iloc[0]['nm_target_schema']
    nm_target_table  = query(sa.target_db, f"SELECT nm_target_table  FROM dta.dataset WHERE meta_is_active = 1 AND id_dataset = '{ip_id_dataset_or_dq_control}'").iloc[0]['nm_target_table']

	# /* Local Variables for "Previous Stand". */
    dt_previous_stand = '1970-01-01 00:00:00.000'

    # -------------------
	# -- "Start" run. --
    # -------------------
	
    if (1==1): # /* Finish "runs" that are NOT "finished". */
        
        # /* Build SQL Statement to "Update" "run" that are NOT "finished". */
        tx_sql  = ""   + f"UPDATE rdp.run SET"
        tx_sql += "\n" + f"  dt_run_finished      = dt_run_started,"
        tx_sql += "\n" + f"  id_processing_status = gnc_commen.id_processing_status('{id_model}', 'Unfinished')"
        tx_sql += "\n" + f"WHERE id_model         = '{id_model}'"
        tx_sql += "\n" + f"AND   id_dataset       = '{id_dataset}'"
        tx_sql += "\n" + f"AND   id_dq_control    = '{id_dq_control}'"
        tx_sql += "\n" + f"AND   ISNULL(dt_run_finished, CONVERT(DATETIME, '9999-12-31')) >= CONVERT(DATETIME, '9999-12-31')"
        
        # /* Execute SQL Statement to "Insert" new "run". */
        if (is_debugging == "1"):
            print(f"SQL Statement: {tx_sql}")

        # Execute the SQL statement
        execute_sql(sa.target_db, tx_sql)
    
    # end if
    
    ni_run = query(sa.target_db, f"SELECT ni_run = COUNT(*) FROM rdp.run WHERE id_run = '{id_run}'").iloc[0]['ni_run']
    while ni_run > 0: # /* Check if @id_run is Unique */
        
        # Show Info on invalid id_run
        if (is_debugging == "1"):
            print(f"The value of `id_run` `{id_run}` was not unique! Hashed value was `CONCAT(CONVERT(NVARCHAR(MAX),''),'|', '{id_model}', '|', '{ip_id_dataset_or_dq_control}', '|', '{dt_run_started}', '|')`.")
            print(f"ip_id_dataset_or_dq_control : `{ip_id_dataset_or_dq_control}`")
            print(f"dt_run_started              : `{dt_run_started}`")
            
        # Determine new dt_run_started, id_run and ni_run
        dt_run_started = dt.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        id_run = query(sa.target_db, f"SELECT id_run = LOWER(CONVERT(CHAR(32),HASHBYTES('MD5',CONCAT(CONVERT(NVARCHAR(MAX),''), '|', '{id_model}', '|', '{ip_id_dataset_or_dq_control}', '|', '{dt_run_started}', '|')), 2))").iloc[0]['id_run']
        ni_run = query(sa.target_db, f"SELECT ni_run = COUNT(*) FROM rdp.run WHERE id_run = '{id_run}'").iloc[0]['ni_run']

    # end while
    
    if (1==1): #/* Create ##dt to prevent "warning" in SQL parsing of project. */) BEGIN
    
        # /* Build and Execute SQL Statment to "Create" ##dt. */
        tx_sql  = ""   + f"SELECT MAX(u.dt) AS dt_previous_stand FROM ("
        tx_sql += "\n" + f"    SELECT MAX(meta_dt_valid_from)         AS dt FROM {nm_target_schema}.{nm_target_table} UNION"
        tx_sql += "\n" + f"    SELECT MAX(meta_dt_valid_till)         AS dt FROM {nm_target_schema}.{nm_target_table} WHERE meta_dt_valid_till < CONVERT(DATE, '9999-12-31') UNION"
        tx_sql += "\n" + f"    SELECT CONVERT(DATETIME, '1970-01-01') AS dt"
        tx_sql += "\n" + f") AS u WHERE dt IS NOT NULL"
        
        # /* Fetch dt.previous_stand */
        dt_previous_stand = query(sa.target_db, tx_sql).iloc[0]['dt_previous_stand']

    # end if

    if (1==1): # /* Insert new "run". */
        
        # /* Build SQL Statement to "Insert" new "run". */
        tx_sql  = ""   + f"INSERT INTO rdp.run ("
        tx_sql += "\n" + f"    id_run,"
        tx_sql += "\n" + f"    id_model,"
        tx_sql += "\n" + f"    id_dataset,"
        tx_sql += "\n" + f"    id_dq_control,"
        tx_sql += "\n" + f"    ds_external_reference_id,"
        tx_sql += "\n" + f"    dt_previous_stand,"
        tx_sql += "\n" + f"    dt_current_stand,"
        tx_sql += "\n" + f"    ni_previous_epoch,"
        tx_sql += "\n" + f"    ni_current_epoch,"
        tx_sql += "\n" + f"    id_processing_status,"
        tx_sql += "\n" + f"    dt_run_started,"
        tx_sql += "\n" + f"    dt_run_finished"
        tx_sql += "\n" + f")"
        tx_sql += "\n" + f"SELECT"
        tx_sql += "\n" + f"    id_run                   = '{id_run}',"
        tx_sql += "\n" + f"    id_model                 = '{id_model}',"
        tx_sql += "\n" + f"    id_dataset               = '{id_dataset}',"
        tx_sql += "\n" + f"    id_dq_control            = '{id_dq_control}',"
        tx_sql += "\n" + f"    ds_external_reference_id = '{ds_external_reference_id}',"
        tx_sql += "\n" + f"    dt_previous_stand        = '{dt_previous_stand}',"
        tx_sql += "\n" + f"    dt_current_stand         = '{dt_run_started}',"
        tx_sql += "\n" + f"    ni_previous_epoch        = DATEDIFF(SECOND, CONVERT(DATETIME, '1970-01-01'), CONVERT(DATETIME, '{dt_previous_stand}')),"
        tx_sql += "\n" + f"    ni_current_epoch         = DATEDIFF(SECOND, CONVERT(DATETIME, '1970-01-01'), CONVERT(DATETIME, '{dt_run_started}')),"
        tx_sql += "\n" + f"    id_processing_status     = gnc_commen.id_processing_status('{id_model}', 'Started'),"
        tx_sql += "\n" + f"    dt_run_started           = CONVERT(DATETIME, '{dt_run_started}'),"
        tx_sql += "\n" + f"    dt_run_finished          = CONVERT(DATETIME, '9999-12-31')"
        tx_sql += "\n" + f"FROM (" # /* make "recordset" of @dt_run_started to ensure there is a record in de SELECT. */
        tx_sql += "\n" + f"    SELECT dt_current_stand  = CONVERT(DATETIME, '{dt_run_started}'),"
        tx_sql += "\n" + f"           dt_previous_stand = CONVERT(DATETIME, '{dt_previous_stand}')"
        tx_sql += "\n" + f") AS std LEFT JOIN rdp.run AS run"
        tx_sql += "\n" + f"ON  run.id_dataset     = '{id_dataset}'"
        tx_sql += "\n" + f"AND run.id_dq_control  = '{id_dq_control}'"
        tx_sql += "\n" + f"AND run.dt_run_started = (" # /* Find the "Previous" run that NOT ended in "Failed"-status. */
        tx_sql += "\n" + f"    SELECT MAX(dt_run_started) FROM rdp.run"
        tx_sql += "\n" + f"    WHERE id_model             = '{id_model}'"
        tx_sql += "\n" + f"    AND   id_dataset           = '{id_dataset}'"
        tx_sql += "\n" + f"    AND   id_dq_control        = '{id_dq_control}'"
        tx_sql += "\n" + f"    AND   id_processing_status = gnc_commen.id_processing_status('{id_model}', 'Finished')"
        tx_sql += "\n" + f")"
        
        # /* Execute SQL Statement to "Insert" new "run". */
        if (is_debugging == "1"):
            print(f"SQL Statement: {tx_sql}")

        # Execute the SQL statement
        execute_sql(sa.target_db, tx_sql)
        
    # end if
    
    # /* All is Well, return "new" ID. */
    if (is_debugging == "1"):
        print(f"Run started with id_run: {id_run}")