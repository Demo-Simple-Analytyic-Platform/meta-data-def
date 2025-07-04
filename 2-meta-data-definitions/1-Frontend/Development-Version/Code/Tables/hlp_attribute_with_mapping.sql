-- Table: hlp_attribute_with_mapping
-- Created: 2025-07-03 23:24:09
-- Records: 0

CREATE TABLE [hlp_attribute_with_mapping] (
    [id] IDENTITY(1,1) INT,
    [id_dataset_target] VARCHAR(255),
    [id_attribute_target] VARCHAR(32),
    [id_dataset_source] VARCHAR(255),
    [id_attribute_source] VARCHAR(32),
    [tx_mapping_source] VARCHAR(255)
);

-- Index: dta_attributeid_dataset
CREATE INDEX [dta_attributeid_dataset] ON [hlp_attribute_with_mapping] ([id_attribute_target]);

-- Index: id
CREATE INDEX [id] ON [hlp_attribute_with_mapping] ([id_attribute_target], [id_attribute_target], [id]);

-- Index: id_attribute
CREATE INDEX [id_attribute] ON [hlp_attribute_with_mapping] ([id_attribute_target], [id_attribute_target], [id], [id_attribute_target], [id], [id_dataset_target]);

-- Index: id_dataset_taget
CREATE INDEX [id_dataset_taget] ON [hlp_attribute_with_mapping] ([id_attribute_target], [id_attribute_target], [id], [id_attribute_target], [id], [id_dataset_target], [id_attribute_target], [id], [id_dataset_target], [id_dataset_source]);

