# Source Module Documentation (`source.py`)

The `source.py` module provides comprehensive data ingestion capabilities for the Demo Simple Analytic Platform. It handles multiple data source types including web tables, Azure Blob Storage files, and SQL databases, converting them into pandas DataFrames for further processing.

## Overview

This module serves as the primary data extraction layer, supporting various source systems and formats:
- Anonymous web table extraction
- Azure Blob Storage CSV and Excel files
- SQL Server databases with authentication
- Secure credential management integration

## Dependencies

```python
import pandas as pd
import requests
from io import StringIO, BytesIO
from azure.storage.blob import BlobServiceClient
import pyodbc
from modules import secrets
```

## Core Functions

### Web Data Sources

#### `web_table_anonymous_web(wtb_1_any_ds_url, wtb_2_any_ds_path, wtb_3_any_ni_index, is_debugging="0")`
**Purpose**: Extracts tabular data from web pages without authentication requirements.

**Parameters**:
- `wtb_1_any_ds_url` (str): Base URL of the website
- `wtb_2_any_ds_path` (str): Specific path to the page containing the table
- `wtb_3_any_ni_index` (str/int): Table index on the page (0-based)
- `is_debugging` (str): Debug mode flag ("0" or "1")

**Returns**: `pandas.DataFrame` - Extracted table data

**Functionality**:
1. Constructs full URL by combining base URL and path
2. Sends HTTP GET request to retrieve page content
3. Uses `pandas.read_html()` to parse HTML tables
4. Extracts specific table by index
5. Provides debug output when enabled

**Error Handling**:
- HTTP request failures
- HTML parsing errors
- Invalid table indices
- Network connectivity issues

**Example Usage**:
```python
df = web_table_anonymous_web(
    "https://example.com", 
    "/data/tables", 
    "0", 
    "1"
)
```

### Azure Blob Storage Sources

#### `abs_sas_url_csv(abs_1_csv_nm_account, abs_2_csv_nm_secret, abs_3_csv_nm_container, abs_4_csv_ds_folderpath, abs_5_csv_ds_filename, abs_6_csv_nm_decode, abs_7_csv_is_1st_header, abs_8_csv_cd_delimiter_value, abs_9_csv_cd_delimter_text, is_debugging="0")`
**Purpose**: Extracts CSV data from Azure Blob Storage using SAS authentication.

**Parameters**:
- `abs_1_csv_nm_account` (str): Azure Storage account name
- `abs_2_csv_nm_secret` (str): Secret name for SAS token retrieval
- `abs_3_csv_nm_container` (str): Container name
- `abs_4_csv_ds_folderpath` (str): Folder path within container
- `abs_5_csv_ds_filename` (str): CSV filename
- `abs_6_csv_nm_decode` (str): Text encoding (e.g., 'utf-8', 'latin-1')
- `abs_7_csv_is_1st_header` (str): Whether first row contains headers ("1" or "0")
- `abs_8_csv_cd_delimiter_value` (str): Delimiter code
- `abs_9_csv_cd_delimter_text` (str): Actual delimiter character
- `is_debugging` (str): Debug mode flag

**Returns**: `pandas.DataFrame` - CSV data

**Functionality**:
1. Retrieves SAS token from secrets management system
2. Constructs Azure Blob Storage connection string
3. Creates BlobServiceClient and downloads file content
4. Determines appropriate delimiter (comma, semicolon, tab, pipe)
5. Parses CSV with specified encoding and header settings
6. Handles various CSV formatting scenarios

**Delimiter Mapping**:
- `"comma"` → `","`
- `"semicolon"` → `";"`
- `"tab"` → `"\t"`
- `"pipe"` → `"|"`

#### `abs_sas_url_xls(abs_1_xls_nm_account, abs_2_xls_nm_secret, abs_3_xls_nm_container, abs_4_xls_ds_folderpath, abs_5_xls_ds_filename, abs_6_xls_nm_sheet, abs_7_xls_is_first_header, abs_8_xls_cd_top_left_cell, abs_9_xls_cd_bottom_right_cell, is_debugging="0")`
**Purpose**: Extracts Excel data from Azure Blob Storage with precise cell range specification.

**Parameters**:
- `abs_1_xls_nm_account` (str): Azure Storage account name
- `abs_2_xls_nm_secret` (str): Secret name for SAS token retrieval
- `abs_3_xls_nm_container` (str): Container name
- `abs_4_xls_ds_folderpath` (str): Folder path within container
- `abs_5_xls_ds_filename` (str): Excel filename (.xlsx, .xls)
- `abs_6_xls_nm_sheet` (str): Worksheet name
- `abs_7_xls_is_first_header` (str): Whether first row contains headers
- `abs_8_xls_cd_top_left_cell` (str): Starting cell (e.g., "A1")
- `abs_9_xls_cd_bottom_right_cell` (str): Ending cell (e.g., "Z100")
- `is_debugging` (str): Debug mode flag

**Returns**: `pandas.DataFrame` - Excel data

**Functionality**:
1. Retrieves SAS token from secrets management
2. Downloads Excel file from Azure Blob Storage
3. Converts cell range notation to pandas usecols format
4. Reads specific worksheet and cell range
5. Handles header row detection and processing
6. Supports both .xlsx and .xls formats

**Cell Range Processing**:
- Converts Excel notation (e.g., "A1:Z100") to column indices
- Calculates skiprows and nrows based on cell range
- Handles dynamic range specifications

### SQL Database Sources

#### `sql_user_password(sql_1_nm_server, sql_2_nm_username, sql_6_nm_secret, sql_3_nm_database, sql_5_tx_query, is_debugging="0")`
**Purpose**: Executes SQL queries against SQL Server databases using username/password authentication.

**Parameters**:
- `sql_1_nm_server` (str): SQL Server instance name or IP
- `sql_2_nm_username` (str): Database username
- `sql_6_nm_secret` (str): Secret name for password retrieval
- `sql_3_nm_database` (str): Database name
- `sql_5_tx_query` (str): SQL query to execute
- `is_debugging` (str): Debug mode flag

**Returns**: `pandas.DataFrame` - Query results

**Functionality**:
1. Retrieves password from secrets management system
2. Constructs SQL Server connection string with ODBC Driver 17
3. Establishes database connection
4. Executes provided SQL query
5. Returns results as pandas DataFrame
6. Handles connection cleanup automatically

**Connection String Format**:
```
DRIVER={ODBC Driver 17 for SQL Server};SERVER={server};DATABASE={database};UID={username};PWD={password}
```

## Utility Functions

### `excel_range_to_pandas(range_str)`
**Purpose**: Converts Excel cell range notation to pandas-compatible parameters.

**Parameters**:
- `range_str` (str): Excel range (e.g., "A1:Z100")

**Returns**: `tuple` - (usecols, skiprows, nrows)

**Functionality**:
1. Parses Excel range notation
2. Converts column letters to numeric indices
3. Calculates row offsets and counts
4. Returns parameters suitable for `pandas.read_excel()`

### `column_letter_to_number(col_letter)`
**Purpose**: Converts Excel column letters to numeric indices.

**Parameters**:
- `col_letter` (str): Column letter(s) (e.g., "A", "AB")

**Returns**: `int` - Zero-based column index

**Algorithm**: Implements base-26 conversion for Excel column notation

## Security Features

### Credential Management
- Integrates with `secrets` module for secure credential storage
- Retrieves passwords and SAS tokens without exposing them in logs
- Supports encrypted credential storage and retrieval

### Connection Security
- Uses ODBC Driver 17 for SQL Server connections
- Supports SAS token authentication for Azure Storage
- Implements secure connection string construction

## Error Handling

The module implements comprehensive error handling for:

### Network Issues
- HTTP request timeouts and failures
- Azure Blob Storage connectivity problems
- SQL Server connection failures

### Data Format Issues
- Invalid CSV delimiters
- Malformed Excel ranges
- Encoding problems
- Missing files or containers

### Authentication Issues
- Invalid credentials
- Expired SAS tokens
- Permission denied scenarios

## Debug Mode Features

When `is_debugging="1"`:
- Displays connection parameters (excluding sensitive data)
- Shows data extraction progress
- Reports DataFrame shapes and sample data
- Logs processing steps and timing information

## Integration Points

### Secrets Management
```python
from modules import secrets
password = secrets.read_secret(secret_name)
```

### Target Database Loading
The extracted DataFrames are typically passed to:
- `run.load_tsl()` for temporal staging
- Database loading procedures
- Transformation pipelines

## Data Type Handling

### Automatic Type Inference
- CSV files: Uses pandas automatic type detection
- Excel files: Preserves Excel data types
- SQL queries: Maintains database column types

### Encoding Support
- UTF-8, Latin-1, and other character encodings
- Automatic encoding detection where possible
- Configurable encoding parameters

## Performance Considerations

### Memory Management
- Streams large files when possible
- Uses BytesIO for in-memory file processing
- Implements efficient DataFrame construction

### Network Optimization
- Single download operations for blob storage
- Connection pooling for database queries
- Appropriate timeout settings

## Usage Examples

### Web Table Extraction
```python
df = source.web_table_anonymous_web(
    "https://finance.yahoo.com",
    "/quote/AAPL/holders",
    "0",
    "1"
)
```

### Azure CSV Processing
```python
df = source.abs_sas_url_csv(
    "mystorageaccount",
    "csv_sas_token_secret",
    "data-container",
    "financial/stocks",
    "stock_prices.csv",
    "utf-8",
    "1",
    "comma",
    ",",
    "1"
)
```

### SQL Data Extraction
```python
df = source.sql_user_password(
    "myserver.database.windows.net",
    "readonly_user",
    "sql_password_secret",
    "analytics_db",
    "SELECT * FROM daily_prices WHERE date >= '2024-01-01'",
    "1"
)
```

## File Format Support

### CSV Files
- Multiple delimiter types (comma, semicolon, tab, pipe)
- Various encodings (UTF-8, Latin-1, etc.)
- Configurable header detection
- Large file handling

### Excel Files (.xlsx/.xls)
- Multi-worksheet support
- Precise cell range extraction
- Header row detection
- Formula value extraction (not formulas themselves)

### Web Tables
- HTML table parsing
- Multi-table page support
- Index-based table selection
- Automatic column detection

## Return Data Structure

All functions return `pandas.DataFrame` objects with:
- Consistent column naming
- Appropriate data types
- Preserved source formatting where possible
- Empty DataFrames for failed extractions (with error logging)

This module serves as the foundation for all data ingestion workflows in the platform, providing reliable, secure, and flexible data extraction capabilities across multiple