-- Table: dqm_dq_requirement
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [dqm_dq_requirement] (
    [id_model] VARCHAR(32),
    [id_dq_requirement] VARCHAR(255),
    [id_development_status] VARCHAR(255),
    [cd_dq_requirement] VARCHAR(32),
    [fn_dq_requirement] VARCHAR(128),
    [fd_dq_requirement] TEXT,
    [dt_valid_from] DATETIME,
    [dt_valid_till] DATETIME
);

-- Index: dta_modeldqm_dq_requirement
CREATE INDEX [dta_modeldqm_dq_requirement] ON [dqm_dq_requirement] ([id_model]);

-- Index: id_development_status
CREATE INDEX [id_development_status] ON [dqm_dq_requirement] ([id_model], [id_model], [id_development_status]);

-- Index: id_dq_requirement
CREATE INDEX [id_dq_requirement] ON [dqm_dq_requirement] ([id_model], [id_model], [id_development_status], [id_model], [id_development_status], [id_dq_requirement]);

