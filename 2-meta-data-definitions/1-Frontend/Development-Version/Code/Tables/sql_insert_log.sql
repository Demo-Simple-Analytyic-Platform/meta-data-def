-- Table: sql_insert_log
-- Created: 2025-07-03 21:57:04
-- Records: 1

CREATE TABLE [sql_insert_log] (
    [id_log] IDENTITY(1,1) INT,
    [nm_schema] VARCHAR(255),
    [nm_table] VARCHAR(255),
    [tx_sql] TEXT,
    [dt_log] DATETIME DEFAULT Now()
);

-- Index: ID
CREATE INDEX [ID] ON [sql_insert_log] ([id_log]);

-- Primary Key: PrimaryKey
ALTER TABLE [sql_insert_log] ADD CONSTRAINT [PrimaryKey] PRIMARY KEY ([id_log], [id_log]);

