-- Table: srd_parameter_group
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [srd_parameter_group] (
    [id_model] VARCHAR(32),
    [id_parameter_group] VARCHAR(32),
    [cd_parameter_group] VARCHAR(32),
    [fn_parameter_group] VARCHAR(128),
    [fd_parameter_group] VARCHAR(255)
);

-- Index: dta_modelsrd_parameter_group
CREATE INDEX [dta_modelsrd_parameter_group] ON [srd_parameter_group] ([id_model]);

