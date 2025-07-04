-- Table: dqm_dq_threshold
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [dqm_dq_threshold] (
    [id_model] VARCHAR(32),
    [id_dq_threshold] VARCHAR(255),
    [id_dq_risk_level] VARCHAR(32),
    [id_dq_control] VARCHAR(32),
    [nr_dq_threshold_from] VARCHAR(255),
    [nr_dq_threshold_till] VARCHAR(255),
    [dt_valid_from] DATETIME,
    [dt_valid_till] DATETIME
);

-- Index: dqm_dq_controldqm_dq_threshold
CREATE INDEX [dqm_dq_controldqm_dq_threshold] ON [dqm_dq_threshold] ([id_model], [id_dq_control]);

