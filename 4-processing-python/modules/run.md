<h1>run.py Module Documentation</h1>

The run.py module is the core orchestration component of the Demo Simple Analytic Platform's data processing pipeline. It manages data ingestion, transformation, and documentation generation workflows.

## Overview

This module provides comprehensive functionality for:
- Data pipeline orchestration and execution
- Multi-source data ingestion (web tables, Azure Blob Storage, SQL databases)
- Database transformation processing
- Run lifecycle management
- HTML documentation generation and Azure Blob Storage deployment

## Dependencies

```python
from modules import credentials as sa
from modules import source as src
from modules.fso import folder_exists, create_folder
from modules.sql import query, execute_procedure, engine, truncate_table, execute_sql
from azure.storage.blob import BlobServiceClient, ContentSettings
from datetime import datetime as dt
import pandas as pd
```

## Core Functions

### Initialization

#### `initialize(is_debugging="0", ip_is_override="0")`
**Purpose**: Initialize the module with necessary database credentials and configurations.

**Parameters**:
- `is_debugging`: Debug mode flag ("0" or "1")
- `ip_is_override`: Override flag for credential initialization

**Functionality**:
- Initializes credentials for secrets database via `sa.credentials`
- Initializes credentials for target database
- Sets up global credential objects for subsequent operations

### Pipeline Orchestration

#### `data_pipeline(id_model, nm_target_schema, nm_target_table, is_debugging)`
**Purpose**: Main orchestration function that executes the complete data pipeline for a specific dataset.

**Parameters**:
- `id_model`: Model identifier
- `nm_target_schema`: Target schema name
- `nm_target_table`: Target table name
- `is_debugging`: Debug mode flag

**Workflow**:
1. Queries `dta.process_group` to get processing metadata
2. Generates external reference ID with timestamp
3. Executes dataset update with retry logic (up to 3 attempts)
4. Exports documentation via `export_documentation`

#### `update_dataset(id_model, ds_external_reference_id, id_dataset, is_ingestion, nm_procedure, nm_tsl_schema, nm_tsl_table, is_debugging)`
**Purpose**: Core dataset processing function handling both ingestion and transformation workflows.

**Parameters**:
- `id_model`: Model identifier
- `ds_external_reference_id`: External reference for tracking
- `id_dataset`: Dataset identifier
- `is_ingestion`: Flag indicating ingestion (1) vs transformation (0)
- `nm_procedure`: Stored procedure name to execute
- `nm_tsl_schema`: Temporal staging landing schema
- `nm_tsl_table`: Temporal staging landing table
- `is_debugging`: Debug mode flag

**Ingestion Workflow**:
1. Starts run tracking via `start`
2. Retrieves parameters via `get_parameters`
3. Routes to appropriate source connector based on parameter group:
   - `web_table_anonymous_web`: `src.web_table_anonymous_web`
   - `abs_sas_url_csv`: `src.abs_sas_url_csv`
   - `abs_sas_url_xls`: `src.abs_sas_url_xls`
   - `sql_user_password`: `src.sql_user_password`
4. Loads data to temporal staging via `load_tsl`
5. Executes ingestion procedure via `usp_dataset_ingestion`

**Transformation Workflow**:
- Executes transformation procedure via `usp_dataset_transformation`

### Data Loading

#### `load_tsl(df_source_dataset, nm_tsl_schema, nm_tsl_table, is_debugging="0")`
**Purpose**: Loads source DataFrame into temporal staging landing table.

**Parameters**:
- `df_source_dataset`: Source pandas DataFrame
- `nm_tsl_schema`: Target schema name
- `nm_tsl_table`: Target table name
- `is_debugging`: Debug mode flag

**Functionality**:
1. Truncates target table via `truncate_table`
2. Creates SQL engine connection via `engine`
3. Loads DataFrame using pandas `to_sql` with replace strategy

### Documentation Management

#### `export_documentation(id_dataset, is_debugging)`
**Purpose**: Generates and deploys HTML documentation for datasets to Azure Blob Storage.

**Parameters**:
- `id_dataset`: Dataset identifier ("-1" for main page)
- `is_debugging`: Debug mode flag

**Workflow**:
1. Queries documentation content from `mdm.html_file_name` and `mdm.html_file_text`
2. Generates HTML content from query results
3. Creates local temporary file structure
4. Writes HTML content to local file
5. Uploads to Azure Blob Storage with proper content type settings

### Run Management

#### `start(id_model, ip_id_dataset_or_dq_control, ds_external_reference_id, is_debugging="0")`
**Purpose**: Comprehensive run lifecycle initialization and tracking setup.

**Parameters**:
- `id_model`: Model identifier
- `ip_id_dataset_or_dq_control`: Dataset or DQ control identifier
- `ds_external_reference_id`: External reference for tracking
- `is_debugging`: Debug mode flag

**Functionality**:
1. **Run ID Generation**: Creates unique MD5-based run identifier
2. **Entity Resolution**: Determines if processing dataset or DQ control
3. **Previous Run Cleanup**: Marks unfinished runs as 'Unfinished'
4. **Uniqueness Validation**: Ensures run ID uniqueness with retry logic
5. **Previous Stand Calculation**: Determines last successful processing timestamp
6. **Run Record Creation**: Inserts comprehensive run tracking record in `rdp.run`

### Utility Functions

#### `get_parameters(id_model, id_dataset)`
**Purpose**: Retrieves dataset-specific parameters for processing.

**Returns**: DataFrame containing parameter definitions from `rdp.tvf_get_parameters`

#### `get_secret(nm_secret, is_debugging)`
**Purpose**: Securely retrieves secrets from the secrets database.

**Parameters**:
- `nm_secret`: Secret name identifier
- `is_debugging`: Debug mode flag

**Returns**: Decrypted secret value or None if not found

#### `get_param_value(nm_parameter_value, params)`
**Purpose**: Extracts specific parameter value from parameters DataFrame.

**Parameters**:
- `nm_parameter_value`: Parameter name to retrieve
- `params`: Parameters DataFrame

**Returns**: Parameter value from the fourth column (index 3)

### Stored Procedure Execution

#### `usp_dataset_ingestion(nm_procedure, is_debugging)`
**Purpose**: Executes ingestion-specific stored procedures.

**Parameters**:
- `nm_procedure`: Procedure name to execute
- `is_debugging`: Debug mode flag

**Functionality**: Direct SQL procedure execution using raw cursor connection

#### `usp_dataset_transformation(nm_procedure, ds_external_reference_id)`
**Purpose**: Executes transformation procedures with external reference tracking.

**Parameters**:
- `nm_procedure`: Procedure name to execute
- `ds_external_reference_id`: External reference for tracking

**Functionality**: Uses `execute_procedure` with parameter passing

## Integration Points

The module integrates with several key components:

- **Source Connectors**: `modules.source` for data extraction
- **SQL Operations**: `modules.sql` for database interactions
- **Credentials Management**: `modules.credentials` for secure access
- **File System**: `modules.fso` for folder operations
- **Process Groups**: `dta.process_group` for processing metadata
- **Run Tracking**: `rdp.run` for execution monitoring

## Error Handling

The module implements comprehensive error handling:
- Retry logic for dataset updates (up to 3 attempts)
- Exception capture with detailed error reporting
- Graceful failure handling with boolean return values
- Debug mode logging for troubleshooting

## Usage Examples

```python
# Initialize the module
run.initialize(is_debugging="1")

# Execute full pipeline for specific dataset
run.data_pipeline("model_id", "target_schema", "target_table", "1")

# Export documentation only
run.export_documentation("dataset_id", "1")
```

This module serves as the central orchestration layer for the platform's data processing capabilities, providing robust workflow management, comprehensive tracking, and automated documentation generation.