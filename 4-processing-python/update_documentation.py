import sys
sys.path.append('modules')

# Import Custom Modules
import modules.run as run
import modules.credentials as sa

# Set Debugging to "1" => true
is_debugging = "1"

# update main page
run.export_documentation('-1', is_debugging)
tx_sql_statement = "SELECT id_dataset from dta.dataset where meta_is_active = 1 and nm_target_schema NOT IN ('dqm', 'dq_totals', 'dq_result', 'dta_dq_aggregates')"
dst = run.query(sa.target_db, tx_sql_statement)

# update all datasets
print("Updating all documentation of datasets")
for i in range(0, len(dst)):
  for row in dst:
    id_dataset = dst.loc[i]['id_dataset']
    run.export_documentation(id_dataset, is_debugging)

# update all datasets   
print("Done all documentation of datasets")