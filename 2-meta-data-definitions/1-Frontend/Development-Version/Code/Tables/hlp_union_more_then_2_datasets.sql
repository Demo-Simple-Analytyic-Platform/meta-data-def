-- Table: hlp_union_more_then_2_datasets
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [hlp_union_more_then_2_datasets] (
    [id_dataset] VARCHAR(32),
    [id_dataset_tobe_unioned] VARCHAR(32)
);

-- Primary Key: PrimaryKey
ALTER TABLE [hlp_union_more_then_2_datasets] ADD CONSTRAINT [PrimaryKey] PRIMARY KEY ([id_dataset], [id_dataset_tobe_unioned]);

