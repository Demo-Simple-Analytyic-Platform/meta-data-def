/* -------------------------------------------------------------------------- */
/* Definitions for `Dataset` and `related`-objects like `attributes`,         */
/* `DQ Controls`, `DQ Thresholds` and `related Group(s)`.                     */
/* -------------------------------------------------------------------------- */
/*                                                                            */
/* ID Dataset : `dummy-dummy-dummy-dummy-dummy-01`                            */
/*                                                                            */
/* -------------------------------------------------------------------------- */
BEGIN

  /* --------------------- */
  /* `Dataset`-definitions */
  /* --------------------- */
  INSERT INTO tsa_dta.tsa_dataset (id_model, id_development_status, id_dataset, id_group, is_ingestion, fn_dataset, fd_dataset, nm_target_schema, nm_target_table, tx_source_query) 
  SELECT mdl.id_model                       AS id_model
       , '06010b090001080103040f070e011504' AS id_development_status
       , 'dummy-dummy-dummy-dummy-dummy-01' AS id_dataset
       , NULL                               AS id_group
       , '0'                                AS is_ingestion
       , 'Dummy'                            AS fn_dataset
       , '<div>Dummy Dataset</div>'         AS fd_dataset
       , 'mdm'                              AS nm_target_schema
       , 'meta_attributes'                  AS nm_target_table
       , 'SELECT 1 AS [dummy]'              AS tx_source_query
  FROM mdm.current_model as mdl;
  
  /* ----------------------- */
  /* `Attribute`-definitions */
  /* ----------------------- */
  INSERT INTO tsa_dta.tsa_attribute (id_model, id_attribute, id_dataset, id_datatype, fn_attribute, fd_attribute, ni_ordering, nm_target_column, is_businesskey, is_nullable) 
  SELECT mdl.id_model                       AS id_model
       , 'dummy-dummy-attribute-0123456789' AS id_attribute
       , 'dummy-dummy-dummy-dummy-dummy-01' AS id_dataset
       , '000e0b00050008010800000102140a0c' AS id_datatype
       , 'dummy'                            AS id_datatype
       , '<div>Dummy</div>'                 AS fn_attribute
       , '1'                                AS ni_ordering
       , 'dummy'                            AS nm_target_column
       , '1'                                AS is_businesskey
       , '0'                                AS is_nullable
  FROM mdm.current_model as mdl;

  /* ------------------------------ */
  /* `Parameter Values`-definitions */
  /* ------------------------------ */
  -- No Defintions for `Parameter Values`

  /* ------------------------------ */
  /* `SQL for ETL`-definitions      */
  /* ------------------------------ */
  INSERT INTO tsa_dta.tsa_ingestion_etl (id_model, id_ingestion_etl, id_dataset, nm_processing_type, tx_sql_for_meta_dt_valid_from, tx_sql_for_meta_dt_valid_till) 
  SELECT mdl.id_model, 'dummy-dummy-ingestion_etl-456789', 'dummy-dummy-dummy-dummy-dummy-01', NULL, NULL, NULL
  FROM mdm.current_model as mdl;
  /* ------------------------------ */
  /* `Schedule`-definitions         */
  /* ------------------------------ */
  -- n/a

  /* -------------------------------- */
  /* `Related (Group(s))`-definitions */
  /* -------------------------------- */
  -- n/a

  /* ------------------------ */
  /* `DQ Control`-definitions */
  /* ------------------------ */
  -- n/a

  /* -------------------------- */
  /* `DQ Threshold`-definitions */
  /* -------------------------- */
  -- n/a
  
END
GO

