-- Table: hlp_dta_database
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [hlp_dta_database] (
    [id_model] VARCHAR(32),
    [id_database] VARCHAR(32),
    [id_environment] VARCHAR(32),
    [nm_server] VARCHAR(128),
    [nm_database] VARCHAR(128),
    [nm_username] VARCHAR(128),
    [nm_secret] VARCHAR(128)
);

-- Index: id_database
CREATE INDEX [id_database] ON [hlp_dta_database] ([id_database]);

-- Index: id_model
CREATE INDEX [id_model] ON [hlp_dta_database] ([id_database], [id_database], [id_model]);

