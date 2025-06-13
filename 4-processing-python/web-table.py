import modules.source as src

# Input Parameters":
wtb_1_any_ds_url   = 'https://finance.yahoo.com/quote/'
wtb_2_any_ds_path  = 'RF/history/?period1=1742277579&period2=1742709170&filter=history&frequency=1d'
wtb_3_any_ni_index = '0'

# Debugging
is_debugging = "1"

df = src.web_table_anonymous_web(wtb_1_any_ds_url, wtb_2_any_ds_path, wtb_3_any_ni_index, is_debugging)
df.columns