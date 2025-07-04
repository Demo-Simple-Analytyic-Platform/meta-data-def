-- Table: srd_development_status
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [srd_development_status] (
    [id_model] VARCHAR(32),
    [id_development_status] VARCHAR(255),
    [cd_development_status] VARCHAR(255),
    [nm_development_status] VARCHAR(255)
);

-- Index: dta_modelsrd_development_status
CREATE INDEX [dta_modelsrd_development_status] ON [srd_development_status] ([id_model]);

-- Index: id_development_status
CREATE INDEX [id_development_status] ON [srd_development_status] ([id_model], [id_model], [id_development_status]);

