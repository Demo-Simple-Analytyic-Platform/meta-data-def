-- Query: ddl_srd_dq_risk_level
-- Created: 2025-07-03 23:24:10
-- Type: Select Query
-- SQL Statement:
SELECT srd_dq_risk_level.id_dq_risk_level AS id, [fn_dq_risk_level] & " / " & [fn_dq_status] AS display, srd_dq_risk_level.id_model
FROM srd_dq_risk_level;

