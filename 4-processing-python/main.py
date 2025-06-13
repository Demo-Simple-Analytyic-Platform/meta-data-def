# This is the main script to run the data pipeline for the PSA and DTA datasets.
# It is the main entry point for the data pipeline and is used to run the data pipeline for all datasets.
# It is also used to run the data pipeline for a specific dataset.
import sys
sys.path.append('modules')

# Import Custom Modules
import modules.run as run

# Set Debugging to "1" => true
is_debugging = "1"

# rebuild html documentation for main pagepip
run.export_documentation('-1', is_debugging)

# Process all datasets

# run.data_pipeline('psa_references', 'currency', is_debugging)
# run.data_pipeline('psa_references', 'stock', is_debugging)
#
#run.data_pipeline('psa_yahoo_exchange_rate', 'eur_x_cad', is_debugging)
#run.data_pipeline('psa_yahoo_exchange_rate', 'eur_x_usd', is_debugging)
#
run.data_pipeline('psa_yahoo_dividends', '04q', is_debugging)
#run.data_pipeline('psa_yahoo_dividends', 'abnas', is_debugging)
#run.data_pipeline('psa_yahoo_dividends', 'arr', is_debugging)
#run.data_pipeline('psa_yahoo_dividends', 'bamnbas', is_debugging)
#run.data_pipeline('psa_yahoo_dividends', 'bce', is_debugging)
#run.data_pipeline('psa_yahoo_dividends', 'fro', is_debugging)
#run.data_pipeline('psa_yahoo_dividends', 'gogl', is_debugging)
#run.data_pipeline('psa_yahoo_dividends', 'ingaas', is_debugging)
#run.data_pipeline('psa_yahoo_dividends', 'kpnas', is_debugging)
#run.data_pipeline('psa_yahoo_dividends', 'nvidia', is_debugging)
#run.data_pipeline('psa_yahoo_dividends', 'o', is_debugging)
#run.data_pipeline('psa_yahoo_dividends', 'ohi', is_debugging)
#run.data_pipeline('psa_yahoo_dividends', 'pfe', is_debugging)
#run.data_pipeline('psa_yahoo_dividends', 'rf', is_debugging)
#run.data_pipeline('psa_yahoo_dividends', 'stwd', is_debugging)
#run.data_pipeline('psa_yahoo_dividends', 'td', is_debugging)
#run.data_pipeline('psa_yahoo_dividends', 'trmd', is_debugging)
#
run.data_pipeline('psa_yahoo_stocks', '04q', is_debugging)
#run.data_pipeline('psa_yahoo_stocks', 'abnas', is_debugging)
#run.data_pipeline('psa_yahoo_stocks', 'arr', is_debugging)
#run.data_pipeline('psa_yahoo_stocks', 'bamnbas', is_debugging)
#run.data_pipeline('psa_yahoo_stocks', 'bce', is_debugging)
#run.data_pipeline('psa_yahoo_stocks', 'fro', is_debugging)
#run.data_pipeline('psa_yahoo_stocks', 'gogl', is_debugging)
#run.data_pipeline('psa_yahoo_stocks', 'ingaas', is_debugging)
#run.data_pipeline('psa_yahoo_stocks', 'kpnas', is_debugging)
#run.data_pipeline('psa_yahoo_stocks', 'nvidia', is_debugging)
#run.data_pipeline('psa_yahoo_stocks', 'o', is_debugging)
#run.data_pipeline('psa_yahoo_stocks', 'ohi', is_debugging)
#run.data_pipeline('psa_yahoo_stocks', 'pfe', is_debugging)
#run.data_pipeline('psa_yahoo_stocks', 'rf', is_debugging)
#run.data_pipeline('psa_yahoo_stocks', 'stwd', is_debugging)
#run.data_pipeline('psa_yahoo_stocks', 'td', is_debugging)
#run.data_pipeline('psa_yahoo_stocks', 'trmd', is_debugging)

print("all done")