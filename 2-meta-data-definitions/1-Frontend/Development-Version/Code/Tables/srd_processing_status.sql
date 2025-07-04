-- Table: srd_processing_status
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [srd_processing_status] (
    [id_model] VARCHAR(32),
    [id_processing_status] VARCHAR(128),
    [ni_processing_status] INT,
    [fn_processing_status] VARCHAR(128),
    [fd_processing_status] TEXT
);

-- Index: dta_modelsrd_processing_status
CREATE INDEX [dta_modelsrd_processing_status] ON [srd_processing_status] ([id_model]);

