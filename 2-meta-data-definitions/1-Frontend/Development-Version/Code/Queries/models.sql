-- Query: models
-- Created: 2025-07-03 23:24:10
-- Type: Select Query
-- SQL Statement:
SELECT dta_model.id_model, dta_model.nm_repository, tx_repo_folderpath([nm_repository]) AS tx_repo_folderpath, tx_repo_folderpath_exists([nm_repository]) AS tx_repo_folderpath_exists, IIf([id_model]=id_model_default(),-1,0) AS is_current_model
FROM dta_model;

