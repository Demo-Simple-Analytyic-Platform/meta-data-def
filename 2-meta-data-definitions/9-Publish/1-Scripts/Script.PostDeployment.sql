SET NOCOUNT ON;

PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Create all "Temp"-tables.                                                 */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
BEGIN EXEC deployment.usp_create_tmp; END
GO

PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Load all "Definitions" into " temp"-tables.                               */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
:r ..\..\2-Definitions\insert_definition_into_temp_tables.sql
GO

/* Add "Dummy" record so dataset can easily be created */
:r .\dummy.sql -- !!! DO NOT REMOVE THIS FILE !!!

GO

PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Execute "##usp_add_all_meta_attributes".                                  */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
BEGIN
  EXEC deployment.usp_add_all_meta_attributes;
END
GO
PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Execute "##usp_insert_dq_controls_as_datasets".                           */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
BEGIN
  EXEC deployment.usp_insert_dq_controls_as_datasets; 
END
GO
PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Execute "##usp_insert_all_current_aggregate_dq_datasets".                 */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
BEGIN
  EXEC deployment.usp_insert_all_current_aggregate_dq_datasets; 
END
GO
PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Execute "##usp_execute_all_usp".                                          */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
BEGIN
  EXEC deployment.usp_execute_all_usp @ip_is_only_metadata = 1; 
END
GO

PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Add all "DQ Aggregate" of "DQ Result" as "Dataset/Attribute"-definitions. */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
BEGIN 
  EXEC deployment.usp_delete_all_current_aggregate_dq_datasets;
END
GO

PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Add all "DQ Aggregate" of "DQ Result" as "Dataset/Attribute"-definitions. */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
BEGIN 
  EXEC deployment.usp_dq_aggregates_as_datasets_for_results;
END
GO

PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Add all "DQ Aggregate" of "DQ Result" as "Dataset/Attribute"-definitions. */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
BEGIN 
  EXEC deployment.usp_dq_aggregates_as_datasets_for_totals;
END
GO

PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Add all "DQ Aggregate" of "DQ Result" as "Dataset/Attribute"-definitions. */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
BEGIN 
  EXEC deployment.usp_execute_all_usp @ip_is_only_metadata = 1; 
END
GO

PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Extract all "Transformation"-parts from "Source"-queries in definitions.  */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
BEGIN EXEC deployment.usp_transformation_part_all; END
GO

PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Extract all "Transformation"-dataset from "Source"-query in definitions.  */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
BEGIN EXEC deployment.usp_transformation_dataset_all; END
GO

PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Extract all "Transformation"-mapping from "Source"-query in definitions.  */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
BEGIN EXEC deployment.usp_transformation_mapping_all; END
GO

PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Extract all "Transformation"-attribute from "Source"-query in definitions.  */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
BEGIN EXEC deployment.usp_transformation_attribute_all; END
GO

PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Execute all "Update"-procedure for the "Transformation"-definitions.      */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
BEGIN EXEC deployment.usp_execute_all_usp @ip_is_only_metadata = 0; END
GO

PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Execute all "Update"-documentation for the "Transformation"-definitions.  */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
BEGIN EXEC deployment.build_html_file_dataset_all @ip_is_debugging = 0; END
GO


PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Execute all "Update"-documentation for the "main"-page.                   */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
BEGIN EXEC mdm.usp_build_html_page_main @ip_is_debugging = 0; END
GO

PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Execute all "Validation" of "metadata"-definitions.                       */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
BEGIN EXEC deployment.check_for_validation_issues @ip_is_debugging = 0; END;
GO

PRINT('/* ------------------------------------------------------------------------- */');
PRINT('/* Execute all "Validation" of "metadata"-definitions.                       */');
PRINT('/* ------------------------------------------------------------------------- */');
GO
BEGIN EXEC deployment.deploy_datasets @ip_is_debugging = 0; END;
GO
