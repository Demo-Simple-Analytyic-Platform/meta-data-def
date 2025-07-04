-- Table: dta_database
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [dta_database] (
    [id_model] VARCHAR(32),
    [id_database] VARCHAR(32),
    [id_environment] VARCHAR(32),
    [nm_server] VARCHAR(128),
    [nm_database] VARCHAR(128),
    [nm_username] VARCHAR(128),
    [nm_secret] VARCHAR(128)
);

-- Index: dta_modeldta_database
CREATE INDEX [dta_modeldta_database] ON [dta_database] ([id_model]);

-- Index: id_database
CREATE INDEX [id_database] ON [dta_database] ([id_model], [id_model], [id_database]);

