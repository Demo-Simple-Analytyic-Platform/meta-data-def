-- Query: ddl_group_x_dataset
-- Created: 2025-07-03 23:24:10
-- Type: Select Query
-- SQL Statement:
SELECT dta_dataset.id_dataset AS id, [fn_group] & " | " & [fn_dataset] AS display
FROM ohg_group INNER JOIN dta_dataset ON ohg_group.id_group = dta_dataset.id_group
ORDER BY [fn_group] & " | " & [fn_dataset];

