# This is the main script to run the data pipeline for the PSA and DTA datasets.
# It is the main entry point for the data pipeline and is used to run the data pipeline for all datasets.
# It is also used to run the data pipeline for a specific dataset.
import sys
sys.path.append('modules')

# Import Custom Modules
import modules.run as run

# Set Debugging to "1" => true
is_debugging = "1"

# Assumtions: stuff a overarching procedure shoudl extract, but for our example we will hardcode it
id_model          = "<id_model>" # was id_model was updated by the initialization
id_dataset        = "<id_dataset>" # was id_dataset was updated by the initialization
ds_external_reference_id = "ds_external_reference_id" # was ds_external_reference_id was updated by the initialization
nm_target_scehme  = '<nm_target_schema>'
nm_target_table   = '<nm_target_table>'    

# Extraction of metadata for the desired model + dataset
run.data_pipeline(id_model, nm_target_scehme, nm_target_table, is_debugging)



# rebuild html documentation for main pagepip
# run.export_documentation('-1', is_debugging)

# Process all datasets

print("all done")