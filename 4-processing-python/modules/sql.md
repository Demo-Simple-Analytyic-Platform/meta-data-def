# SQL Module Documentation

The sql.py module provides comprehensive database connectivity and SQL execution functionality for the Demo Simple Analytic Platform. It serves as the primary interface for all SQL Server database operations, offering both direct SQL execution and stored procedure calling capabilities.

## Overview

This module provides functionality for:
- SQL Server database connections using ODBC and SQLAlchemy
- SQL query execution and result retrieval
- Stored procedure execution with parameter handling
- Database table operations (truncate, etc.)
- Connection string management
- Engine creation for pandas DataFrame operations

## Dependencies

```python
import pyodbc
import pandas as pd
from urllib.parse import quote  
from sqlalchemy import create_engine, text
```

## Core Functions

### Connection Management

#### `connection_string(credentials_db)`
**Purpose**: Generates ODBC connection string for SQL Server database connections.

**Parameters**:
- `credentials_db` (dict): Database credentials dictionary containing:
  - `server`: SQL Server instance name
  - `database`: Database name
  - `username`: Database username
  - `password`: Database password

**Returns**: `str` - Formatted ODBC connection string

**Connection String Format**:
```
DRIVER={ODBC Driver 17 for SQL Server};
TrustServerCertificate=no;
Encrypt=no;
SERVER={server};
DATABASE={database};
UID={username};
PWD={password}
```

**Security Features**:
- Uses ODBC Driver 17 for SQL Server
- Encryption disabled for local development
- Server certificate validation disabled

---

#### `engine(credentials_db)`
**Purpose**: Creates SQLAlchemy engine for pandas DataFrame operations and advanced SQL execution.

**Parameters**:
- `credentials_db` (dict): Database credentials dictionary

**Returns**: `sqlalchemy.Engine` - SQLAlchemy engine object

**Functionality**:
- URL-encodes password for special characters using `quote()`
- Creates `mssql+pyodbc` connection string
- Configures encryption and certificate settings
- Returns engine suitable for pandas `to_sql()` operations

**Connection String Format**:
```
mssql+pyodbc://{username}:{password}@{server}/{database}?driver={driver}&encrypt={encrypt}&trustedservercertificate={trustedservercertificate}
```

### Query Operations

#### `query(credentials_db, tx_sql_statement)`
**Purpose**: Executes SELECT queries and returns results as pandas DataFrame.

**Parameters**:
- `credentials_db` (dict): Database credentials dictionary
- `tx_sql_statement` (str): SQL SELECT statement to execute

**Returns**: `pandas.DataFrame` - Query results

**Functionality**:
1. Creates ODBC connection using `connection_string()`
2. Executes SQL query using `pd.read_sql()`
3. Returns DataFrame with query results
4. Automatically closes connection

**Usage Examples**:
```python
# Simple query
df = query(credentials, "SELECT * FROM users")

# Parameterized query (string formatting)
df = query(credentials, f"SELECT * FROM orders WHERE user_id = {user_id}")
```

---

#### `execute_sql(credentials_db, tx_sql_statement, is_debugging="0")`
**Purpose**: Executes any SQL statement (INSERT, UPDATE, DELETE, DDL) with transaction management.

**Parameters**:
- `credentials_db` (dict): Database credentials dictionary
- `tx_sql_statement` (str): SQL statement to execute
- `is_debugging` (str): Debug mode flag ("0" or "1")

**Returns**: `sqlalchemy.Result` or `None`

**Functionality**:
1. Creates SQLAlchemy engine connection
2. Executes SQL using `text()` wrapper
3. Auto-commits for DML operations (INSERT, UPDATE, DELETE)
4. Provides debug logging when enabled
5. Handles connection cleanup automatically

**Transaction Management**:
- Automatically detects DML operations
- Commits transactions for data modification statements
- Returns None for committed operations

### Table Operations

#### `truncate_table(credentials_db, nm_schema, nm_table)`
**Purpose**: Truncates specified database table.

**Parameters**:
- `credentials_db` (dict): Database credentials dictionary
- `nm_schema` (str): Schema name
- `nm_table` (str): Table name

**Returns**: Result from `execute_sql()`

**Functionality**:
- Constructs `TRUNCATE TABLE {schema}.{table}` statement
- Executes via `execute_sql()` for consistent transaction handling
- Provides fast table clearing operation

### Stored Procedure Execution

#### `execute_procedure(credentials_db, nm_procedure, **params)`
**Purpose**: Executes stored procedures with named parameter support.

**Parameters**:
- `credentials_db` (dict): Database credentials dictionary
- `nm_procedure` (str): Stored procedure name
- `**params`: Named parameters for the procedure

**Returns**: `sqlalchemy.Result` - Procedure execution result

**Functionality**:
1. Builds parameter list with `@parameter = 'value'` format
2. Constructs `EXEC procedure_name @param1 = 'value1', @param2 = 'value2'` statement
3. Executes using SQLAlchemy text execution
4. Provides comprehensive debug logging
5. Handles cursor-based execution

**Debug Features**:
- Parameter logging when `ip_is_debugging = "1"`
- Complete procedure call string display
- Parameter value inspection

**Usage Example**:
```python
result = execute_procedure(
    credentials, 
    "usp_update_user",
    ip_user_id="123",
    ip_username="john_doe",
    ip_is_debugging="1"
)
```

---

#### `execute_procedure2(credentials_db, nm_procedure, **params)`
**Purpose**: Alternative stored procedure execution using direct ODBC cursor approach.

**Parameters**:
- `credentials_db` (dict): Database credentials dictionary
- `nm_procedure` (str): Stored procedure name
- `**params`: Named parameters for the procedure

**Returns**: `list` or `None` - Procedure results or None if no results

**Functionality**:
1. Uses direct pyodbc connection and cursor
2. Filters parameters to only include those starting with 'ip_'
3. Uses ODBC `{CALL procedure_name(?)}` syntax
4. Handles parameter placeholders automatically
5. Attempts result fetching with error handling
6. Manual transaction commit and connection cleanup

**Key Differences from `execute_procedure`**:
- Direct ODBC approach vs SQLAlchemy
- Parameter filtering for 'ip_' prefix
- ODBC CALL syntax vs EXEC syntax
- Result fetching with exception handling
- Manual connection management

## Error Handling

The module implements several error handling strategies:

### Connection Errors
- Automatic connection cleanup in all functions
- Context manager usage for SQLAlchemy connections
- Explicit connection closing for pyodbc

### Parameter Handling
- URL encoding for special characters in passwords
- Parameter filtering in `execute_procedure2`
- Safe string formatting for SQL construction

### Result Handling
- Exception handling for procedures that don't return results
- Graceful handling of empty result sets
- Proper resource cleanup on failures

## Security Considerations

### Strengths
- Uses parameterized procedure calls in `execute_procedure2`
- URL encoding for password special characters
- Automatic connection cleanup

### Areas for Improvement
- String formatting in some functions could lead to SQL injection
- Credentials passed as plain text dictionaries
- No connection pooling or timeout management
- Debug mode may log sensitive information

## Integration Points

The module integrates with:
- **Credentials Module**: Receives database credentials
- **Pandas**: Returns DataFrames for query results
- **Run Module**: Provides database operations for pipeline execution
- **Source Module**: Supports data loading operations

## Performance Considerations

- **Connection Management**: Creates new connections for each operation
- **No Connection Pooling**: Each function creates fresh connections
- **Transaction Scope**: Appropriate commit handling for DML operations
- **Resource Cleanup**: Proper connection closure prevents leaks

## Usage Patterns

### Query Data
```python
# Load configuration data
config_df = query(target_db, "SELECT * FROM config WHERE active = 1")

# Get processing metadata
metadata_df = query(target_db, f"SELECT * FROM process_group WHERE id_model = '{model_id}'")
```

### Execute Procedures
```python
# Start processing run
execute_procedure(
    target_db, 
    "usp_start_run",
    ip_model_id=model_id,
    ip_dataset_id=dataset_id,
    ip_reference_id=reference_id
)

# Data transformation
execute_procedure(
    target_db,
    "usp_transform_data",
    ip_source_table="staging.raw_data",
    ip_target_table="processed.clean_data"
)
```

### Table Operations
```python
# Clear staging table
truncate_table(target_db, "staging", "temp_data")

# Load DataFrame to database
engine_obj = engine(target_db)
df.to_sql("target_table", con=engine_obj, schema="staging", if_exists="replace")
```

This module provides the foundational database connectivity layer for the entire platform, enabling robust and flexible SQL Server interactions across all data processing workflows.