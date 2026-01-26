Attribute VB_Name = "vsp_create_ddl_schema_and_table"
Option Compare Database
Option Explicit
'
' Get cd_dataset_type
Public Function get_cd_dataset_type(ip_nm_target_schema As String, ip_is_ingestion As Boolean) As String
  get_cd_dataset_type = IIf(Left(ip_nm_target_schema, 2) = "dq", "DQControl", IIf(ip_is_ingestion, "Ingestion", "Transformation"))
End Function
'
' Get full of relative path for "Schema"-folder
Public Function get_fp_schema(ip_nm_schema As String, ip_cd_dataset_type As String, relative As Boolean) As String
  '
  ' Build relative Folderpath Schema
  If (ip_cd_dataset_type = "Ingestion") Then get_fp_schema = fp_ingestions(relative) & ip_nm_schema
  If (ip_cd_dataset_type = "DQControl") Then get_fp_schema = fp_dqcontrols(relative) & ip_nm_schema
  If (ip_cd_dataset_type = "Transformation") Then get_fp_schema = fp_transformations(relative) & ip_nm_schema
  If (relative = True) Then get_fp_schema = Replace(get_fp_schema, ".\", "")
  get_fp_schema = get_fp_schema & "\"
  '
End Function


Public Sub add_schema(ip_nm_schema As String, ip_cd_dataset_type As String, Optional is_debugging As Boolean = False)
  '
  ' Local Variables
  Dim fso As FileSystemObject: Set fso = New FileSystemObject
  '
  ' ensure main folder with subfolder s exists
  Call build_folder_structure
  '
  ' determin schema folderpath
  Dim rp_schema As String: rp_schema = get_fp_schema(ip_nm_schema, ip_cd_dataset_type, True)
  Dim fp_schema As String: fp_schema = get_fp_schema(ip_nm_schema, ip_cd_dataset_type, False)
  '
  ' Create Folder if not exists
  Call create_folder_if_not_exists(fp_schema)
  '
  ' Add Folder to Visual Studio Project
  Call AddFolderToSqlProj(rp_schema, "None", is_debugging)
  '
  If (ip_nm_schema <> "dqm" And ip_nm_schema <> "tsa_dqm" And ip_nm_schema <> "tsl_dqm") Then ' Create SQL File
    Dim sql As String: sql = "CREATE SCHEMA [" & ip_nm_schema & "]"
    Dim fil As TextStream: Set fil = fso.OpenTextFile(fp_schema & "\" & ip_nm_schema & ".sql", ForWriting, True, TristateTrue)
    fil.Write sql: fil.Close
    '
    ' Add the SQL-file to the Project as well
    Call AddSqlFileToSqlProj(rp_schema & ip_nm_schema & ".sql", "Build", is_debugging)
    '
  End If
  '
  ' Create Underlaying Folders of Tables, Indexes and Procedures
  Call create_folder_if_not_exists(fp_schema & "Tables"):     Call AddFolderToSqlProj(rp_schema & "Tables", "None", is_debugging)
  If (Left(ip_nm_schema, 3) <> "tsa" And Left(ip_nm_schema, 3) <> "tsl") Then
    Call create_folder_if_not_exists(fp_schema & "Indexes"):    Call AddFolderToSqlProj(rp_schema & "Indexes", "None", is_debugging)
    Call create_folder_if_not_exists(fp_schema & "Procedures"): Call AddFolderToSqlProj(rp_schema & "Procedures", "None", is_debugging)
  End If
  '
End Sub
'
' Add All "Target"-schemas to the appropiate folders
Public Sub add_all_schemas(Optional is_debugging As Boolean = False)
  '
  Dim cd_dataset_type As String
  Dim sql As String: sql = "SELECT nm_target_schema, is_ingestion FROM dta_dataset GROUP BY nm_target_schema, is_ingestion"
  Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset(sql): Do Until rst.EOF
    '
    If (is_debugging) Then Debug.Print "Schema:" & rst!nm_target_schema
    '
    ' Add Schemas from the Datasets to the Project
    cd_dataset_type = get_cd_dataset_type(rst!nm_target_schema, rst!is_ingestion)
    add_schema Trim("    " & rst!nm_target_schema), cd_dataset_type, is_debugging
    '
    If (rst!nm_target_schema <> "dqm") Then ' Create SQL File
      add_schema Trim("tsa_" & rst!nm_target_schema), cd_dataset_type, is_debugging
      add_schema Trim("tsl_" & rst!nm_target_schema), cd_dataset_type, is_debugging
    End If
    '
  rst.MoveNext: Loop
  '
End Sub
'
' Create Create Table from metadata
Public Sub add_or_update_create_table_sql_file(ip_id_dataset As String): On Error GoTo errHandle
  '
  ' Local Variables
  Dim fso As FileSystemObject: Set fso = New FileSystemObject
  Dim emp As String: emp = ""
  Dim nwl As String: nwl = vbNewLine
  Dim col As String: col = ""
  Dim mda As String: mda = ""
  Dim sql As String
  Dim nam As String
  Dim fil As TextStream
  Dim rp_schema As String
  Dim fp_schema As String
  '
  ' Fetch All Attributes of Dataset
  sql = "SELECT att.is_businesskey" & _
  nwl & "     , att.nm_target_column" & _
  nwl & "     , dtp.cd_target_datatype" & _
  nwl & "     , att.is_nullable" & _
  nwl & "FROM dta_attribute AS att INNER" & _
  nwl & "JOIN srd_datatype  AS dtp" & _
  nwl & "ON  (att.id_datatype = dtp.id_datatype)" & _
  nwl & "AND (att.id_model    = dtp.id_model)" & _
  nwl & "WHERE (att.id_model   = '" & id_model_default() & "')" & _
  nwl & "AND   (att.id_dataset = '" & ip_id_dataset & "')" & _
  nwl & "ORDER BY att.ni_ordering ASC"
  Dim att As Recordset: Set att = CurrentDb.OpenRecordset(sql)
  '
  ' Fetch Dataset information
  sql = "SELECT dst.*, Iif(Left(dst.nm_target_schema,2)='dq','DQControl',Iif(is_ingestion,'Ingestion','Transformation')) AS cd_dataset_type FROM dta_dataset AS dst WHERE dst.id_dataset = '" & ip_id_dataset & "'"
  Dim dst As Recordset: Set dst = CurrentDb.OpenRecordset(sql): Do Until dst.EOF
    '
    ' Adding the Attrbibutes
    col = col & nwl & "  "
    col = col & nwl & "  /* Data Attribute(s) */"
    Do Until att.EOF: col = col & nwl & "  [" & att!nm_target_column & "] " & att!cd_target_datatype & IIf(att!is_nullable, "", " NOT") & " NULL" & IIf(att!is_businesskey = True, " /* BK */", ""): att.MoveNext: col = col & IIf(att.EOF, "", ","): Loop
    '
    ' Add "Metadata"-attributes
    mda = "," & nwl & "  "
    mda = mda & nwl & "  /* Metadata Attributes */"
    mda = mda & nwl & "  [meta_dt_valid_from] DATETIME NOT NULL,"
    mda = mda & nwl & "  [meta_dt_valid_till] DATETIME NOT NULL,"
    mda = mda & nwl & "  [meta_is_active]     BIT      NOT NULL,"
    mda = mda & nwl & "  [meta_ch_rh]         CHAR(32) NOT NULL,"
    mda = mda & nwl & "  [meta_ch_bk]         CHAR(32) NOT NULL,"
    mda = mda & nwl & "  [meta_ch_pk]         CHAR(32) NOT NULL,"
    mda = mda & nwl & "  [meta_dt_created]    DATETIME NOT NULL DEFAULT GETDATE()"
    '
    If (dst!nm_target_schema <> "dqm") Then ' Add "Target"-table (psa or dta)
      '
      ' Build SQL "CREATE TABLE"-Statement for in SQL file.
      sql = emp & emp & "CREATE TABLE [" & dst!nm_target_schema & "].[" & dst!nm_target_table & "] ("
      sql = sql & emp & col & mda
      sql = sql & nwl & "  "
      sql = sql & nwl & ");"
      sql = sql & nwl & "GO"
      sql = sql & nwl & ""
      '
      ' Create SQL File
      nam = dst!nm_target_table & ".sql"
      rp_schema = get_fp_schema(dst!nm_target_schema, dst!cd_dataset_type, True) & "Tables\"
      fp_schema = get_fp_schema(dst!nm_target_schema, dst!cd_dataset_type, False) & "Tables\"
      Set fil = fso.OpenTextFile(fp_schema & nam, ForWriting, True, TristateTrue): fil.Write sql: fil.Close
      Call AddSqlFileToSqlProj(rp_schema & nam, "Build")
      '
      ' Build Index-file
      sql = emp & emp & "CREATE CLUSTERED COLUMNSTORE" ' [C]LUSTERED [C]OLUMN[S]TORE => CCS
      sql = sql & nwl & "INDEX [idx_ccs_" & dst!nm_target_schema & "_" & dst!nm_target_table & "]"
      sql = sql & nwl & "ON [" & dst!nm_target_schema & "].[" & dst!nm_target_table & "];"
      sql = sql & nwl & "GO"
      sql = sql & nwl & ""
      '
      ' Create SQL File
      nam = "idx_ccs_" & dst!nm_target_schema & "_" & dst!nm_target_table & ".sql"
      rp_schema = get_fp_schema(dst!nm_target_schema, dst!cd_dataset_type, True) & "Indexes\"
      fp_schema = get_fp_schema(dst!nm_target_schema, dst!cd_dataset_type, False) & "Indexes\"
      Set fil = fso.OpenTextFile(fp_schema & nam, ForWriting, True, TristateTrue): fil.Write sql: fil.Close
      Call AddSqlFileToSqlProj(rp_schema & nam, "Build")
      '
    End If
    '
    If (dst!nm_target_schema <> "dqm") Then ' Add "Temporal Staging Area"-table (tsa)
      '
      ' Get "Schema"-folder full and relative
      rp_schema = get_fp_schema("tsa_" & dst!nm_target_schema, dst!cd_dataset_type, True) & "Tables\"
      fp_schema = get_fp_schema("tsa_" & dst!nm_target_schema, dst!cd_dataset_type, False) & "Tables\"
      '
      ' Create Folder if not exists, Add Folder to Visual Studio Project
      Call create_folder_if_not_exists(fp_schema)
      Call AddFolderToSqlProj(rp_schema)
      '
      ' Set Name file, Added it to Visual Studio Project
      nam = "tsa_" & dst!nm_target_table & ".sql"
      '
      ' Build SQL "CREATE TABLE"-Statement for in SQL file.
      sql = emp & emp & "CREATE TABLE [tsa_" & dst!nm_target_schema & "].[tsa_" & dst!nm_target_table & "] ("
      sql = sql & emp & col & mda
      sql = sql & nwl & "  "
      sql = sql & nwl & ");"
      sql = sql & nwl & "GO"
      sql = sql & nwl & ""
      '
      ' Create SQL File
      Set fil = fso.OpenTextFile(fp_schema & nam, ForWriting, True, TristateTrue): fil.Write sql: fil.Close
      Call AddSqlFileToSqlProj(rp_schema & nam, "Build")
      '
    End If
    '
    If (dst!nm_target_schema <> "dqm") Then ' Add "Temporal Staging Landing"-table (tsa)
      '
      ' Get "Schema"-folder full and relative
      rp_schema = get_fp_schema("tsl_" & dst!nm_target_schema, dst!cd_dataset_type, True) & "Tables\"
      fp_schema = get_fp_schema("tsl_" & dst!nm_target_schema, dst!cd_dataset_type, False) & "Tables\"
      '
      ' Create Folder if not exists, Add Folder to Visual Studio Project
      Call create_folder_if_not_exists(fp_schema)
      Call AddFolderToSqlProj(rp_schema)
      '
      ' Set Name file, Added it to Visual Studio Project
      nam = "tsl_" & dst!nm_target_table & ".sql"
      '
      ' Build SQL "CREATE TABLE"-Statement for in SQL file.
      sql = emp & emp & "CREATE TABLE [tsl_" & dst!nm_target_schema & "].[tsl_" & dst!nm_target_table & "] ("
      sql = sql & emp & col
      sql = sql & nwl & "  "
      sql = sql & nwl & ");"
      sql = sql & nwl & "GO"
      sql = sql & nwl & ""
      '
      ' Create SQL File
      Set fil = fso.OpenTextFile(fp_schema & nam, ForWriting, True, TristateTrue): fil.Write sql: fil.Close
      Call AddSqlFileToSqlProj(rp_schema & nam, "Build")
      '
    End If
    '
  dst.MoveNext: Loop
  Exit Sub
errHandle:
  Debug.Print "--- Error ------------------------------------------------"
  Debug.Print "Number      : " & CStr(Err.Number)
  Debug.Print "Description : " & Err.Description
  Stop
  Resume
  '
  '
End Sub