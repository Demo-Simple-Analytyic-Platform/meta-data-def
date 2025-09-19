# Description: Python module to load data from various sources into a Spark DataFrame.
# - abs_sas_url_csv   : Load a CSV file from Azure Blob Storage into a Spark DataFrame using a SAS token.
# - abs_sas_url_xls   : Load an Excel file from Azure Blob Storage into a Spark DataFrame using a SAS token.
# - sql_user_password : Load data from a SQL Server database into a Spark DataFrame.
# - web_table_anonymous_web : Load a table from a webpage into a Spark DataFrame.
# - dropbox_access_token : Access token for Dropbox API.

# Import Custom Modules
from modules import run         as rn

# Import for Blob Storage Account access in Azure
from azure.storage.blob import BlobServiceClient

# Import for Dropbox access vai Access Token
import dropbox

# Import for web_table_anonymous_web
import pandas as pd
import time
from selenium.webdriver.common.by import By as by
from selenium                     import webdriver
from bs4                          import BeautifulSoup
from io                           import StringIO

def abs_sas_url_csv(

    # Input Parameteres
    abs_1_csv_nm_account,           # Account          | Name of Azure Storage Account (abs).
    abs_2_csv_nm_secret,            # Secret           | Name of the Secret in the Azure Key Vault.
    abs_3_csv_nm_container,         # Container        | Name of Container where the "CSV" file can be found.
    abs_4_csv_ds_folderpath,        # Folderpath       | Folderpath to the "CSV"-file in the Container.
    abs_5_csv_ds_filename,          # Has Header       | Filename of the "CSV"-file.
    abs_6_csv_nm_decode,            # Encoding         | Encoding of the file.
    abs_7_csv_is_1st_header,        # Has Header       | Is first record Header.
    abs_8_csv_cd_delimiter_value,   # Delimiter Value  | Character of the Delimiter for Values.
    abs_9_csv_cd_delimter_text,     # Delimter Text    | Character of the Delimiter for Text.

    # Debugging
    is_debugging
):
    # Handling None values for top left and bottom right cells
    abs_4_csv_ds_folderpath    = "" if abs_4_csv_ds_folderpath    is None else abs_4_csv_ds_folderpath
    abs_9_csv_cd_delimter_text = "" if abs_9_csv_cd_delimter_text is None else abs_9_csv_cd_delimter_text

    # Show input Parameter(s)
    if (is_debugging == "1"):
        print("abs_1_csv_nm_account         : '" + abs_1_csv_nm_account + "'")
        print("abs_2_csv_nm_secret          : '" + abs_2_csv_nm_secret + "'")
        print("abs_3_csv_nm_container       : '" + abs_3_csv_nm_container + "'")
        print("abs_4_csv_ds_folderpath      : '" + abs_4_csv_ds_folderpath + "'")
        print("abs_5_csv_ds_filename        : '" + abs_5_csv_ds_filename + "'")
        print("abs_6_csv_nm_decode          : '" + abs_6_csv_nm_decode + "'")
        print("abs_7_csv_is_1st_header      : '" + abs_7_csv_is_1st_header + "'")
        print("abs_8_csv_cd_delimiter_value : '" + abs_8_csv_cd_delimiter_value + "'")
        print("abs_9_csv_cd_delimter_text   : '" + abs_9_csv_cd_delimter_text + "'")

    # Helper SAS Token URL
    is_header         = 0 if (abs_7_csv_is_1st_header == "1") else None
    tx_accesskey      = rn.get_secret(abs_2_csv_nm_secret, is_debugging)
    ds_filepath_local = f"C:/temp/{abs_5_csv_ds_filename}"
    ds_filepath_blob  = "" if len(abs_4_csv_ds_folderpath) == 0 else abs_4_csv_ds_folderpath + "/" 
    ds_filepath_blob += f"{abs_5_csv_ds_filename}"    

    # Define the connection string and the blob details
    tx_connection_string = f"DefaultEndpointsProtocol=https;AccountName={abs_1_csv_nm_account};AccountKey={tx_accesskey};EndpointSuffix=core.windows.net"

    # Create the BlobServiceClient object
    blob_service_client = BlobServiceClient.from_connection_string(tx_connection_string)

    # Create the BlobClient object
    blob_client = blob_service_client.get_blob_client(container=abs_3_csv_nm_container, blob=ds_filepath_blob)
    
    # Download the blob to a local file
    with open(ds_filepath_local, "wb") as download_file:
        download_file.write(blob_client.download_blob().readall())
    
    # Read the CSV file with the specified parameters
    df = pd.read_csv(ds_filepath_local, header=is_header, delimiter=abs_8_csv_cd_delimiter_value, encoding=abs_6_csv_nm_decode, quotechar=abs_9_csv_cd_delimter_text)

    # Clean up the DataFrame (optional, depending on your needs)
    df = cleanup_headers(df, is_debugging)
    
    # All done
    return df

def abs_sas_url_xls(
        
    # Input Parameters
    abs_1_xls_nm_account,             # Account           | Name of Azure Storage Account (abs).
    abs_2_xls_nm_secret,              # Secret            | Name of the Secret in the Azure Key Vault.
    abs_3_xls_nm_container,           # Container         | Name of Container where the "Excel" file can be found.
    abs_4_xls_ds_folderpath,          # Folderpath        | Folderpath where the "Excel" file can be found.
    abs_5_xls_ds_filename,            # Filename          | Filename of the "Excel"-file.
    abs_6_xls_nm_sheet,               # Sheetname         | Sheetname within the "Excel" where the dataset is to be found.
    abs_7_xls_is_first_header,        # Has Header        | Is first record Header.
    abs_8_xls_cd_top_left_cell,       # Top Left Cell     | If provided the cells marked between the "Top Left Cell"  and "Bottom Right Cell"  are set as range.
    abs_9_xls_cd_bottom_right_cell,   # Bottom Right Cell | If provided the cells marked between the "Top Left Cell"  and "Bottom Right Cell"  are set as range.

    # Debugging
    is_debugging
):
    # Handling None values for top left and bottom right cells
    abs_8_xls_cd_top_left_cell     = "" if abs_8_xls_cd_top_left_cell     is None else abs_8_xls_cd_top_left_cell
    abs_9_xls_cd_bottom_right_cell = "" if abs_9_xls_cd_bottom_right_cell is None else abs_9_xls_cd_bottom_right_cell

    # Show input Parameter(s)
    if (is_debugging == "1"):
        print("abs_1_xls_nm_account           : '" + abs_1_xls_nm_account           + "'")
        print("abs_2_xls_nm_secret            : '" + abs_2_xls_nm_secret            + "'")
        print("abs_3_xls_nm_container         : '" + abs_3_xls_nm_container         + "'")
        print("abs_4_xls_ds_folderpath        : '" + abs_4_xls_ds_folderpath        + "'")
        print("abs_5_xls_ds_filename          : '" + abs_5_xls_ds_filename          + "'")
        print("abs_6_xls_nm_sheet             : '" + abs_6_xls_nm_sheet             + "'")
        print("abs_7_xls_is_first_row_header  : '" + abs_7_xls_is_first_header      + "'")
        print("abs_8_xls_cd_top_left_cell     : '" + abs_8_xls_cd_top_left_cell     + "'")
        print("abs_9_xls_cd_bottom_right_cell : '" + abs_9_xls_cd_bottom_right_cell + "'")

    # Determine local variables
    is_header         = 0 if (abs_7_xls_is_first_header == "1") else None
    tx_accesskey      = rn.get_secret(abs_2_xls_nm_secret, is_debugging)
    ds_filepath_local = f"C:/temp/{abs_5_xls_ds_filename}"
    ds_filepath_blob  = "" if len(abs_4_xls_ds_folderpath) == 0 else abs_4_xls_ds_folderpath + "/" 
    ds_filepath_blob +=  f"{abs_5_xls_ds_filename}"
    cd_range          = f"{abs_8_xls_cd_top_left_cell}:{abs_9_xls_cd_bottom_right_cell}" if (abs_8_xls_cd_top_left_cell != "" and abs_9_xls_cd_bottom_right_cell != "") else None

    # Define the connection string and the blob details
    tx_connection_string = f"DefaultEndpointsProtocol=https;AccountName={abs_1_xls_nm_account};AccountKey={tx_accesskey};EndpointSuffix=core.windows.net"

    # Create the BlobServiceClient object
    blob_service_client = BlobServiceClient.from_connection_string(tx_connection_string)

    # Create the BlobClient object
    blob_client = blob_service_client.get_blob_client(container=abs_3_xls_nm_container, blob=ds_filepath_blob)

    # Download the blob to a local file
    with open(ds_filepath_local, "wb") as download_file:
        download_file.write(blob_client.download_blob().readall())

    # Load the Excel file into a Pandas DataFrame
    df = pd.read_excel(ds_filepath_local, engine='openpyxl', sheet_name=abs_6_xls_nm_sheet, header=is_header, usecols=cd_range)

    # Clean up the DataFrame (optional, depending on your needs)
    df = cleanup_headers(df, is_debugging)

    # All done
    return df

def sql_user_password(

    # Input Parameters:
    sql_1_nm_server,
    sql_2_nm_username,
    sql_3_nm_secret,
    sql_4_nm_database,
    sql_5_tx_query,

    # Debugging
    is_debugging
):
    
    # Helper SAS Token URL
    sql_6_cd_password = rn.get_secret(sql_3_nm_secret, is_debugging)

    # Database credentials
    credentials_db = {
        "server"   : sql_1_nm_server,
        "database" : sql_3_nm_secret,
        "username" : sql_2_nm_username,
        "password" : sql_6_cd_password
    }

    # Run SQL query
    df = rn.query(credentials_db, sql_5_tx_query)

    # Show input Parameter(s)
    if (is_debugging == "1"):
        print("sql_1_nm_server   : '" + sql_1_nm_server   + "'")
        print("sql_2_nm_username : '" + sql_2_nm_username + "'")
        print("sql_3_nm_secret   : '" + sql_3_nm_secret   + "'")
        print("sql_4_nm_database : '" + sql_4_nm_database + "'")
        print("sql_5_tx_query    : '" + sql_5_tx_query    + "'")
        print("DataFrame:")
        df.head(10)

    # Clean up the DataFrame (optional, depending on your needs)
    df = cleanup_headers(df, is_debugging)
            
    # Show the result
    return df

def web_table_anonymous_web(
    
        # Input Parameters":
        wtb_1_any_ds_url,
        wtb_2_any_ds_path,
        wtb_3_any_ni_index,
        
        # Debugging
        is_debugging
    ):

    # If is Debugging then show imput parameters
    if (is_debugging == 1):
        print("wtb_1_any_ds_url   : '" + wtb_1_any_ds_url + "'")
        print("wtb_2_any_ds_path  : '" + wtb_2_any_ds_path + "'")
        print("wtb_3_any_ni_index : '" + wtb_3_any_ni_index + "'")

    # Initialize the WebDriver (e.g., Chrome)
    driver = webdriver.Chrome()

    # Open the webpage
    driver.get(wtb_1_any_ds_url + wtb_2_any_ds_path)

    # Wait for the page to load (you might need to adjust the sleep time)
    time.sleep(5)

    try: # Find and click the "Accept Cookies" button (adjust the selector as needed)
        accept_button = driver.find_element(by.XPATH, '//button[text()="Alles accepteren"]')
        accept_button.click()

        # Wait for the page to load after accepting cookies
        time.sleep(5)

    except Exception as e:
        # Code to handle any other exceptions
        print(f"An unexpected error occurred: {e}")        

    # Get the page source after accepting cookies
    page_source = driver.page_source

    # Close the browser
    driver.quit()
    
    # Parse the page source with BeautifulSoup
    soup = BeautifulSoup(page_source, 'html.parser')
    
    # Create a StringIO object
    table = StringIO()

    # find the 2nd table in the webpage (you might need to adjust the selector based on the webpage structure)
    table.write(str(soup.find_all('table')))
    
    # Read the table into a pandas DataFrame
    df = pd.read_html(StringIO(table.getvalue()))[int(wtb_3_any_ni_index)]

    # Handle MultiIndex columns if present
    if isinstance(df.columns, pd.MultiIndex):
        # Flatten MultiIndex columns to single level
        df.columns = ['_'.join(str(level) for level in col if str(level) != 'nan' and str(level) != '') 
                            for col in df.columns.values]
        # Clean up any empty or double underscores
        df.columns = [col.replace('__', '_').strip('_') for col in df.columns]

    # Handle MultiIndex rows if present
    if isinstance(df.index, pd.MultiIndex):
        df = df.reset_index()

    # Clean up the DataFrame (optional, depending on your needs)
    df = cleanup_headers(df, is_debugging)

    # Return the webtable as a DataFrame
    return df

def dbx_acc_tkn_csv(

    # Input Parameteres
    dbx_1_csv_nm_secret,            # Secret           | Name of the Secret
    dbx_2_csv_ds_folderpath,        # Folderpath       | Folderpath to the "CSV"-file in the Container.
    dbx_3_csv_ds_filename,          # Has Header       | Filename of the "CSV"-file.
    dbx_4_csv_nm_encoding,          # Encoding         | Encoding of the file.
    dbx_5_csv_is_1st_header,        # Has Header       | Is first record Header.
    dbx_6_csv_cd_delimiter_value,   # Delimiter Value  | Character of the Delimiter for Values.
    dbx_7_csv_cd_delimter_text,     # Delimter Text    | Character of the Delimiter for Text.

    # Debugging
    is_debugging

):
    # Handling None values for top left and bottom right cells
    dbx_7_csv_cd_delimter_text  = "" if dbx_7_csv_cd_delimter_text is None else dbx_7_csv_cd_delimter_text

    # Show Parameters if in Debugging mode
    if (is_debugging == "1"):
        print("Debugging information:")
        print(f"Secret Name     : {dbx_1_csv_nm_secret}")
        print(f"Folder Path     : {dbx_2_csv_ds_folderpath}")
        print(f"File Name       : {dbx_3_csv_ds_filename}")
        print(f"Encoding        : {dbx_4_csv_nm_encoding}")
        print(f"Is First Header : {dbx_5_csv_is_1st_header}")
        print(f"Delimiter Value : {dbx_6_csv_cd_delimiter_value}")
        print(f"Delimiter Text  : {dbx_7_csv_cd_delimter_text}")
    
    # Extract Dropbox Access Token from Secrets-database
    tx_access_token = rn.get_secret(dbx_1_csv_nm_secret, is_debugging)

    # Initialize Dropbox client
    dbx = dropbox.Dropbox(tx_access_token)

    # Construct full path
    if (dbx_2_csv_ds_folderpath == ""):
        fp_dropbox = f"/{dbx_3_csv_ds_filename}"
    else:
        fp_dropbox = f"/{dbx_2_csv_ds_folderpath}/{dbx_3_csv_ds_filename}"

    try:
        # Download file
        metadata, res = dbx.files_download(fp_dropbox)
        file_content = res.content.decode(dbx_4_csv_nm_encoding)

        # Load into pandas DataFrame
        if (dbx_7_csv_cd_delimter_text == ""):
            df = pd.read_csv(
                StringIO(file_content), 
                delimiter=dbx_6_csv_cd_delimiter_value, 
                header=0 if dbx_5_csv_is_1st_header == "1" else None
            )
        else:
            df = pd.read_csv(
                StringIO(file_content), 
                delimiter=dbx_6_csv_cd_delimiter_value,
                header=0 if dbx_5_csv_is_1st_header == "1" else None, 
                quoting=1, quotechar=dbx_7_csv_cd_delimter_text
            )

        # Clean up the DataFrame (optional, depending on your needs)
        df = cleanup_headers(df, is_debugging)

        # All Done
        return df

    except dropbox.exceptions.ApiError as err:
        print(f"Dropbox API error: {err}")
        return None

def cleanup_headers(df, is_debugging="0"):
    """
    Cleanup DataFrame headers by removing leading/trailing spaces and replacing spaces with underscores.
    """
    if is_debugging == "1":
        print("Cleaning up DataFrame headers...")

    # Clean up the DataFrame (optional, depending on your needs)
    df.columns = df.columns.str.strip()  # Remove leading/trailing spaces from column names
    df.columns = df.columns.str.replace(' ', '_')   # Replace spaces with underscores in column names
    df.columns = df.columns.str.replace('[', '_')   # Replace [ with _ in column names
    df.columns = df.columns.str.replace(']', '_')   # Replace ] with _ in column names 
    df.columns = df.columns.str.replace('(', '_')   # Replace ( with _ in column names
    df.columns = df.columns.str.replace(')', '_')   # Replace ) with _ in column names
    df.columns = df.columns.str.replace(',', '_')   # Replace , with _ in column names
    df.columns = df.columns.str.replace('\\', '_')  # Remove \ from column names
    df.columns = df.columns.str.replace('/', '_')   # Remove / from column names
    df.columns = df.columns.str.replace('\'', '_')  # Remove ' from column names
    df.columns = df.columns.str.replace('\"', '_')  # Remove " from column names   
    df.columns = df.columns.str.replace('?', '_')   # Remove ? from column names
    df.columns = df.columns.str.replace('!', '_')   # Remove ! from column names
    df.columns = df.columns.str.replace('=', '_')   # Remove = from column names
    df.columns = df.columns.str.replace('|', '_')   # Remove | from column names
    df.columns = df.columns.str.replace('.', '_')   # Remove | from column names
    df.columns = df.columns.str.replace('__', '_')  # Remove | from column names
    df.columns = df.columns.str.replace('__', '_')  # Remove | from column names   
    df.columns = df.columns.str.replace('__', '_')  # Remove | from column names   
    df.columns = df.columns.str.rstrip('_') # Remove trailing underscores from column names
    df.columns = df.columns.str.lstrip('_') # Remove leading underscores from column names
    
    # Show the DataFrame columns if debugging is enabled
    if (is_debugging == "1"):
        print("Column Name:")
        print("------------------------------------")
        for idx, col_name in enumerate(df.columns):
            print(f"{col_name}")
        print("------------------------------------")
        print(f"")
        print(f"# Columns : {len(df.columns)}")
        print(f"# Records : {df.shape[0]}")
        print(f"")
        print("------------------------------------")
    
    # Return the cleaned DataFrame
    return df