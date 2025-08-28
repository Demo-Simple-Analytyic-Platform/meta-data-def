# This is the main script to run the data pipeline for the PSA and DTA datasets.
# It is the main entry point for the data pipeline and is used to run the data pipeline for all datasets.
# It is also used to run the data pipeline for a specific dataset.
import sys
sys.path.append('modules')

# Import Custom Modules
import modules.run as run
import modules.sql as sql
import modules.credentials as crd

# Set Debugging to "1" => true
is_debugging = "0"

# Assumtions: stuff a overarching procedure shoudl extract, but for our example we will hardcode it
id_model = '5f4a1942465c575a1f5a5a575d1e191c' # was id_model was updated by the initialization
id_development_status_ahc = '010408050302010500060b0207190003' # --> Ad-Hoc
id_development_status_oos = '06010b090001080103040f070e011504' # --> Out-of-Scope
id_development_status_dev = '06010b0900010908010d0e0404021503' # --> Development
id_development_status_uat = '01040805030201000104090406190800' # --> User Acceptance Testing
id_development_status_prd = '06030d080400090702000c0502001500' # --> Production
id_development_status     = id_development_status_prd

# rebuild html documentation for main pagepip
run.export_documentation('-1', is_debugging)

# Process Ingestion Datasets
run.process_ingestion_datasets(id_model, id_development_status, is_debugging)

# Process Ingestion Datasets
run.process_transformation_datasets(id_model, is_debugging)

# All Done
print("All Done.")
