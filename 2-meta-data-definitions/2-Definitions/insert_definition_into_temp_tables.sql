
/* Static Reference Data */
:r ".\1-Static-Reference-Data\datatype.sql"
:r ".\1-Static-Reference-Data\development_status.sql"
:r ".\1-Static-Reference-Data\dq_dimension.sql"
:r ".\1-Static-Reference-Data\dq_result_status.sql"
:r ".\1-Static-Reference-Data\dq_review_status.sql"
:r ".\1-Static-Reference-Data\dq_risk_level.sql"
:r ".\1-Static-Reference-Data\processing_status.sql"
:r ".\1-Static-Reference-Data\processing_step.sql"
:r ".\1-Static-Reference-Data\parameter.sql"
:r ".\1-Static-Reference-Data\parameter_group.sql"

/* Organization, Hierarchies and Groups */
:r ".\2-Organization-Hierarchies-and-Groups\group.sql"
:r ".\2-Organization-Hierarchies-and-Groups\hierarchy.sql"

/* Data Quality Model */
:r ".\4-Data-Quality-Model\dq_requirement.sql"

/* All Model(s), Database(s), Dataset(s) */
:r ".\3-Data-Transformation-Area\model.sql"
:r ".\3-Data-Transformation-Area\database.sql"
:r ".\3-Data-Transformation-Area\datasets.sql"

BEGIN /* Name of Git Repository / Current Model */
  
  DELETE FROM mdm.current_model; INSERT INTO mdm.current_model ( id_model, nm_repository ) SELECT
    id_model      = CONVERT(CHAR(32),      '<id_model>'),
    nm_repository = CONVERT(NVARCHAR(128), '<nm_repository>');
    
END
GO