-- Table: ohg_hierarchy
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [ohg_hierarchy] (
    [id_model] VARCHAR(32),
    [id_hierarchy] VARCHAR(255),
    [id_group] VARCHAR(255),
    [id_hierarchy_parent] VARCHAR(32)
);

-- Index: dta_modelohg_hierarchy
CREATE INDEX [dta_modelohg_hierarchy] ON [ohg_hierarchy] ([id_model]);

-- Index: id_group
CREATE INDEX [id_group] ON [ohg_hierarchy] ([id_model], [id_model], [id_group]);

-- Index: id_hierarchy
CREATE INDEX [id_hierarchy] ON [ohg_hierarchy] ([id_model], [id_model], [id_group], [id_model], [id_group], [id_hierarchy]);

