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

# rebuild html documentation for main pagepip
run.export_documentation('-1', is_debugging)

# Assumtions: stuff a overarching procedure shoudl extract, but for our example we will hardcode it
id_model = '5f4a1942465c575a1f5a5a575d1e191c' # was id_model was updated by the initialization
id_development_status_ahc = '010408050302010500060b0207190003' # --> Ad-Hoc
id_development_status_oos = '06010b090001080103040f070e011504' # --> Out-of-Scope
id_development_status_dev = '06010b0900010908010d0e0404021503' # --> Development
id_development_status_uat = '01040805030201000104090406190800' # --> User Acceptance Testing
id_development_status_prd = '06030d080400090702000c0502001500' # --> Production

id_development_status     = id_development_status_prd

# Build SQL Statement to Extract list of Ingestions
tx_query  = f"SELECT nm_target_schema\n"
tx_query += f"     , nm_target_table\n"
tx_query += f"FROM dta.dataset\n"
tx_query += f"WHERE id_development_status = '{id_development_status}'\n"
tx_query += f"AND   id_model              = '{id_model}'\n"
tx_query += f"AND   nm_target_schema     != 'mdm'\n"
tx_query += f"AND   meta_is_active        = 1\n"
tx_query += f"AND   is_ingestion          = 1"

# Execute the query and store result in todo
todo = sql.query(crd.target_db(), tx_query)

# Loop through the todo results and call run.data_pipeline for each
print("\n---------------------------------------------------------------------")
print(f"Processing {len(todo)} datasets from todo list...")
print("---------------------------------------------------------------------")
print("")
i = 0
m = len(todo)
while i < m:

    # Extract schema and table names
    nm_target_schema = todo.iloc[i]['nm_target_schema']
    nm_target_table  = todo.iloc[i]['nm_target_table']

    # Show progress which dataset
    print("---------------------------------------------------------------------")
    print(f"Processing {i + 1}/{len(todo)}: {nm_target_schema}.{nm_target_table}")

    try: # Execute Data Pipeline for dataset.
        run.data_pipeline(id_model, nm_target_schema, nm_target_table, is_debugging)
        print(f"   ✓ Successfully processed: {nm_target_schema}.{nm_target_table}")

    except Exception as e: # Continue with next item even if one fails
        print(f"   ✗ Error processing {nm_target_schema}.{nm_target_table}: {str(e)}")
        continue
    
    # Empty line
    print("")

    # Next Dataset
    i += 1

# Process all1 datasets
print("all done")