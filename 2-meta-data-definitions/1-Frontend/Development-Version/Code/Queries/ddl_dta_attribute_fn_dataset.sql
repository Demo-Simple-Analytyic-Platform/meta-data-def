-- Query: ddl_dta_attribute_fn_dataset
-- Created: 2025-07-03 23:24:09
-- Type: Select Query
-- SQL Statement:
SELECT dta_attribute.id_attribute AS id, dta_dataset.fn_dataset AS display, dta_attribute.id_model
FROM dta_dataset LEFT JOIN dta_attribute ON dta_dataset.id_dataset = dta_attribute.id_dataset;

