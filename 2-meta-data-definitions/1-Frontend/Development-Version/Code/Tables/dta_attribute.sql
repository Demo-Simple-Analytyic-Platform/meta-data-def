-- Table: dta_attribute
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [dta_attribute] (
    [id_model] VARCHAR(32),
    [id_attribute] VARCHAR(255),
    [id_dataset] VARCHAR(32),
    [id_datatype] VARCHAR(32),
    [fn_attribute] VARCHAR(128),
    [fd_attribute] TEXT,
    [ni_ordering] INT,
    [nm_target_column] VARCHAR(128),
    [is_businesskey] BIT,
    [is_nullable] BIT
);

-- Index: dta_attributeid_dataset
CREATE INDEX [dta_attributeid_dataset] ON [dta_attribute] ([id_dataset]);

-- Index: dta_attributeid_datatype
CREATE INDEX [dta_attributeid_datatype] ON [dta_attribute] ([id_dataset], [id_dataset], [id_datatype]);

