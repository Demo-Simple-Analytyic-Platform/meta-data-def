-- Table: ohg_group
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [ohg_group] (
    [id_model] VARCHAR(32),
    [id_group] VARCHAR(255),
    [fn_group] VARCHAR(128),
    [fd_group] TEXT
);

-- Index: dta_modelohg_group
CREATE INDEX [dta_modelohg_group] ON [ohg_group] ([id_model]);

-- Index: id_group
CREATE INDEX [id_group] ON [ohg_group] ([id_model], [id_model], [id_group]);

