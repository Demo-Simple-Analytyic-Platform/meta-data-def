-- Table: srd_processing_step
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [srd_processing_step] (
    [id_model] VARCHAR(32),
    [id_processing_step] VARCHAR(128),
    [ni_processing_step] INT,
    [fn_processing_step] VARCHAR(128),
    [fd_processing_step] TEXT
);

-- Index: dta_modelsrd_processing_step
CREATE INDEX [dta_modelsrd_processing_step] ON [srd_processing_step] ([id_model]);

