-- Table: dta_ingestion_etl
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [dta_ingestion_etl] (
    [id_model] VARCHAR(255),
    [id_dataset] VARCHAR(255),
    [id_ingestion_etl] VARCHAR(32),
    [nm_processing_type] VARCHAR(255),
    [tx_sql_for_meta_dt_valid_from] TEXT,
    [tx_sql_for_meta_dt_valid_till] VARCHAR(255)
);

-- Unique Index: dta_datasetdta_ingestion_etl
CREATE UNIQUE INDEX [dta_datasetdta_ingestion_etl] ON [dta_ingestion_etl] ([id_model], [id_dataset]);

