-- Table: dqm_dq_control
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [dqm_dq_control] (
    [id_model] VARCHAR(255),
    [id_dq_control] VARCHAR(255),
    [id_dq_requirement] VARCHAR(32),
    [id_development_status] VARCHAR(255),
    [id_dq_dimension] VARCHAR(32),
    [id_dataset] VARCHAR(32),
    [cd_dq_control] VARCHAR(32),
    [fn_dq_control] VARCHAR(128),
    [fd_dq_control] TEXT,
    [tx_dq_control_query] TEXT,
    [dt_valid_from] DATETIME,
    [dt_valid_till] DATETIME
);

-- Index: dqm_dq_contolid_dataset
CREATE INDEX [dqm_dq_contolid_dataset] ON [dqm_dq_control] ([id_dataset]);

-- Index: dqm_dq_contolid_dq_dimension
CREATE INDEX [dqm_dq_contolid_dq_dimension] ON [dqm_dq_control] ([id_dataset], [id_dataset], [id_dq_dimension]);

-- Index: dqm_dq_contolid_dq_requirement
CREATE INDEX [dqm_dq_contolid_dq_requirement] ON [dqm_dq_control] ([id_dataset], [id_dataset], [id_dq_dimension], [id_dataset], [id_dq_dimension], [id_dq_requirement]);

