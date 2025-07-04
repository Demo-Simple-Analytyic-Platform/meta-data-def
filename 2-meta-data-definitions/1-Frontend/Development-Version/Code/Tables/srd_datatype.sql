-- Table: srd_datatype
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [srd_datatype] (
    [id_model] VARCHAR(32),
    [id_datatype] VARCHAR(255),
    [fn_datatype] VARCHAR(128),
    [fd_datatype] TEXT,
    [cd_target_datatype] VARCHAR(32),
    [cd_prefix_column_name] VARCHAR(32),
    [cd_symbol_functional_naam] VARCHAR(32)
);

-- Index: dta_modelsrd_datatype
CREATE INDEX [dta_modelsrd_datatype] ON [srd_datatype] ([id_model]);

-- Index: id_datatype
CREATE INDEX [id_datatype] ON [srd_datatype] ([id_model], [id_model], [id_datatype]);

