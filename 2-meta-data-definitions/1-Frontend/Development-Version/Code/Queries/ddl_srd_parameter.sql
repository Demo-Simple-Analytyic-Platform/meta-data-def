-- Query: ddl_srd_parameter
-- Created: 2025-07-03 23:24:10
-- Type: Select Query
-- Parameters:
--   [Forms]![dta_dataset]![id_model] (Text)
-- SQL Statement:
SELECT srd_parameter.id_parameter AS id, srd_parameter.nm_parameter AS display, srd_parameter.id_model
FROM srd_parameter
WHERE (((srd_parameter.id_model)=[Forms]![dta_dataset]![id_model]));

