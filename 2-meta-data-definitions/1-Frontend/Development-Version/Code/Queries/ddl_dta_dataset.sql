-- Query: ddl_dta_dataset
-- Created: 2025-07-03 23:24:10
-- Type: Select Query
-- SQL Statement:
SELECT dta_dataset.id_dataset AS id, [fn_dataset] & " (" & [nm_target_schema] & "." & [nm_target_table] & ")" AS display, dta_dataset.id_model
FROM dta_dataset;

