-- Table: repo
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [repo] (
    [nm_repo] VARCHAR(16),
    [tx_folderpath] TEXT
);

-- Primary Key: PrimaryKey
ALTER TABLE [repo] ADD CONSTRAINT [PrimaryKey] PRIMARY KEY ([nm_repo]);

