-- Table: srd_parameter
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [srd_parameter] (
    [id_model] VARCHAR(32),
    [id_parameter] VARCHAR(32),
    [id_parameter_group] VARCHAR(32),
    [nm_parameter] VARCHAR(128),
    [fn_parameter] VARCHAR(128),
    [fd_parameter] VARCHAR(255)
);

-- Index: dta_modelsrd_parameter
CREATE INDEX [dta_modelsrd_parameter] ON [srd_parameter] ([id_model]);

