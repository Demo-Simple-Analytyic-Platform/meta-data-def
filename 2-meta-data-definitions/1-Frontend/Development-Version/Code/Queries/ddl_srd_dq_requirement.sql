-- Query: ddl_srd_dq_requirement
-- Created: 2025-07-03 23:24:10
-- Type: Select Query
-- SQL Statement:
SELECT dqm_dq_requirement.id_dq_requirement AS id, [cd_dq_requirement] & " - " & [fn_dq_requirement] AS display, dqm_dq_requirement.id_model
FROM dqm_dq_requirement;

