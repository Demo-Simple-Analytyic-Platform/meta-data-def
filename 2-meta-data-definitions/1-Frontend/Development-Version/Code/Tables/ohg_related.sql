-- Table: ohg_related
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [ohg_related] (
    [id_model] VARCHAR(32),
    [id_related] VARCHAR(255),
    [id_dataset] VARCHAR(32),
    [id_group] VARCHAR(32)
);

-- Index: dta_datasetohg_related
CREATE INDEX [dta_datasetohg_related] ON [ohg_related] ([id_model], [id_dataset]);

