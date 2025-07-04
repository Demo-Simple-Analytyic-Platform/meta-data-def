-- Query: dqm_dq_control_list
-- Created: 2025-07-03 23:24:10
-- Type: Select Query
-- SQL Statement:
SELECT dqc.*, mdl.[nm_repository], dqr.fn_dq_requirement, dqd.fn_dq_dimension, dvs.nm_development_status
FROM srd_development_status AS dvs INNER JOIN (srd_dq_dimension AS dqd INNER JOIN (dqm_dq_requirement AS dqr INNER JOIN (dta_model AS mdl INNER JOIN dqm_dq_control AS dqc ON mdl.id_model = dqc.id_model) ON (dqr.id_model = dqc.id_model) AND (dqr.id_dq_requirement = dqc.id_dq_requirement)) ON (dqd.id_model = dqc.id_model) AND (dqd.id_dq_dimension = dqc.id_dq_dimension)) ON (dvs.id_model = dqc.id_model) AND (dvs.id_development_status = dqc.id_development_status);

