-- Table: srd_dq_result_status
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [srd_dq_result_status] (
    [id_model] VARCHAR(32),
    [id_dq_result_status] VARCHAR(255),
    [fn_dq_result_status] VARCHAR(128),
    [fd_dq_result_status] TEXT
);

-- Index: dta_modelsrd_dq_result_status
CREATE INDEX [dta_modelsrd_dq_result_status] ON [srd_dq_result_status] ([id_model]);

