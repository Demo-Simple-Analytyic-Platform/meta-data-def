-- Table: srd_dq_dimension
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [srd_dq_dimension] (
    [id_model] VARCHAR(32),
    [id_dq_dimension] VARCHAR(255),
    [fn_dq_dimension] VARCHAR(128),
    [fd_dq_dimension] TEXT
);

-- Index: dta_modelsrd_dq_dimension
CREATE INDEX [dta_modelsrd_dq_dimension] ON [srd_dq_dimension] ([id_model]);

