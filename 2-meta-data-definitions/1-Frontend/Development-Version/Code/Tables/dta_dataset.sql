-- Table: dta_dataset
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [dta_dataset] (
    [id_model] VARCHAR(32),
    [id_dataset] VARCHAR(255),
    [id_development_status] VARCHAR(255),
    [id_group] VARCHAR(32),
    [is_ingestion] BIT,
    [fn_dataset] VARCHAR(128),
    [fd_dataset] TEXT,
    [nm_target_schema] VARCHAR(128),
    [nm_target_table] VARCHAR(128),
    [tx_source_query] TEXT
);

-- Index: dta_datasetid_group
CREATE INDEX [dta_datasetid_group] ON [dta_dataset] ([id_group]);

-- Index: dta_modeldta_dataset
CREATE INDEX [dta_modeldta_dataset] ON [dta_dataset] ([id_group], [id_group], [id_model]);

-- Index: id_dataset
CREATE INDEX [id_dataset] ON [dta_dataset] ([id_group], [id_group], [id_model], [id_group], [id_model], [id_dataset]);

-- Index: id_development_status
CREATE INDEX [id_development_status] ON [dta_dataset] ([id_group], [id_group], [id_model], [id_group], [id_model], [id_dataset], [id_group], [id_model], [id_dataset], [id_development_status]);

