-- Query: ddl_srd_parameter_group
-- Created: 2025-07-03 23:24:10
-- Type: Select Query
-- SQL Statement:
SELECT srd_parameter_group.id_parameter_group AS id, [cd_parameter_group] & " - " & [fn_parameter_group] AS display, srd_parameter_group.id_model
FROM srd_parameter_group;

