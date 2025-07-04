-- Table: srd_dq_risk_level
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [srd_dq_risk_level] (
    [id_model] VARCHAR(32),
    [id_dq_risk_level] VARCHAR(32),
    [cd_dq_risk_level] VARCHAR(255),
    [fn_dq_risk_level] VARCHAR(128),
    [fd_dq_risk_level] TEXT,
    [cd_dq_status] VARCHAR(32),
    [fn_dq_status] VARCHAR(128),
    [fd_dq_status] TEXT
);

-- Index: dta_modelsrd_dq_risk_level
CREATE INDEX [dta_modelsrd_dq_risk_level] ON [srd_dq_risk_level] ([id_model]);

