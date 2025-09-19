# This is the main script to run the data pipeline for the PSA and DTA datasets.
# It is the main entry point for the data pipeline and is used to run the data pipeline for all datasets.
# It is also used to run the data pipeline for a specific dataset.
import sys
import os
import re
sys.path.append('modules')

# Import Custom Modules
import modules.run as run
import modules.sql as sql
import modules.credentials as crd

# Set Debugging to "1" => true
is_debugging = "1"

# Initialize id_model from SQL definition file
id_model = run.initialize_id_model_from_sql(is_debugging)

run.data_pipeline(id_model, 'psa_revolut', 'account_statements', is_debugging)

# rebuild html documentation for main pagepip
#run.export_documentation('-1', is_debugging)

# Process Ingestion Datasets (currently the development status is still in development)
#run.process_ingestion_datasets(id_model, 'all', is_debugging)

# Process Ingestion Datasets
#run.process_transformation_datasets(id_model, is_debugging)

# All Done
print("All Done.")
