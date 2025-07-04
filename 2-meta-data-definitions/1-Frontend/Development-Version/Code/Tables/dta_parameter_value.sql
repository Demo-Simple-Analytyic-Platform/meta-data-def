-- Table: dta_parameter_value
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [dta_parameter_value] (
    [id_model] VARCHAR(255),
    [id_parameter_value] VARCHAR(32),
    [id_dataset] VARCHAR(32),
    [id_parameter] VARCHAR(32),
    [tx_parameter_value] TEXT,
    [ni_parameter_value] INT
);

-- Index: dta_datasetdta_parameter_value
CREATE INDEX [dta_datasetdta_parameter_value] ON [dta_parameter_value] ([id_model], [id_dataset]);

