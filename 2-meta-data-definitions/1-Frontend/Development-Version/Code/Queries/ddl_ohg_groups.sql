-- Query: ddl_ohg_groups
-- Created: 2025-07-03 23:24:10
-- Type: Select Query
-- SQL Statement:
SELECT id_group AS id, ohg_group.fn_group AS display, id_model
FROM ohg_group
ORDER BY ohg_group.fn_group;

