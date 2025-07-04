-- Query: dta_dataset_list
-- Created: 2025-07-03 23:24:10
-- Type: Select Query
-- SQL Statement:
SELECT dst.*, mdl.[nm_repository], dvs.nm_development_status, grp.fn_group
FROM srd_development_status AS dvs INNER JOIN (ohg_group AS grp INNER JOIN (dta_model AS mdl INNER JOIN dta_dataset AS dst ON mdl.id_model = dst.id_model) ON (grp.id_model = dst.id_model) AND (grp.id_group = dst.id_group)) ON (dvs.id_model = dst.id_model) AND (dvs.id_development_status = dst.id_development_status);

