# This is the main script to run the data pipeline for the PSA and DTA datasets.
# It is the main entry point for the data pipeline and is used to run the data pipeline for all datasets.
# It is also used to run the data pipeline for a specific dataset.
import sys
sys.path.append('modules')

# Import Custom Modules
import modules.run as run

# Set Debugging to "1" => true
is_debugging = "1"
id_model     = "5f4a1942465c575a1f5a5a575d1e191c" # was id_model was updated by the initialization

# rebuild html documentation for main page (you must setup azure storage accoutn with static web option activated and store the secret in the "Secrets"-database)
# run.export_documentation('-1', is_debugging) 

# Process all datasets
run.data_pipeline(id_model, 'psa_yahoo_dividends', 'o4q', is_debugging)


print("all done")