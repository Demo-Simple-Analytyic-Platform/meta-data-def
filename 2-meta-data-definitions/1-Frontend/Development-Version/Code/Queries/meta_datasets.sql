-- Query: meta_datasets
-- Created: 2025-07-03 23:24:10
-- Type: Select Query
-- SQL Statement:
SELECT Left([Name],3) AS nm_schema, Mid([Name],5.123) AS nm_table
FROM MSysObjects
WHERE (((Left([Name],3)) In ("srd","ohg","dta","dqm")) AND ((Left([Name],4))<>"MSys") AND ((MSysObjects.ParentId)=251658241));

