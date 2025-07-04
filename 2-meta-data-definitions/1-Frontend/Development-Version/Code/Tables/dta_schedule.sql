-- Table: dta_schedule
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [dta_schedule] (
    [id_model] VARCHAR(32),
    [id_dataset] VARCHAR(32),
    [id_schedule] VARCHAR(32),
    [cd_frequency] VARCHAR(255),
    [ni_LONGerval] INT,
    [dt_start] DATETIME,
    [dt_end] DATETIME
);

-- Unique Index: dta_datasetdta_schedule
CREATE UNIQUE INDEX [dta_datasetdta_schedule] ON [dta_schedule] ([id_model], [id_dataset]);

