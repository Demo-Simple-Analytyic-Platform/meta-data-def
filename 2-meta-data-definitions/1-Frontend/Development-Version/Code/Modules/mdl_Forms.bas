Option Compare Database
Option Explicit
'
Public Sub set_buttons(ByRef ip_form As Form)
    '
    ' Local Variables
    Dim cd As String: cd = get_cd_development_status(Nz(ip_form.id_development_status, "-1"))
    Dim nc As Boolean: nc = IIf(ip_form.id_model = id_model_default() And ip_form.Name = "dta_dataset", True, False) 'nc = "Not Current Model"
    '
    ' Close all button by default
    ip_form.btn_Save.Enabled = False
    ip_form.btn_Back_to_Development.Enabled = False
    ip_form.btn_Deploy_to_Acceptance.Enabled = False
    ip_form.btn_Deploy_to_Production.Enabled = False
    ip_form.btn_Move_to_Ad_Hoc.Enabled = False
    ip_form.btn_Move_to_Out_of_Scope.Enabled = False
    ip_form.btn_Delete.Enabled = False
    '
    ' Open buttons on "Developemnt"-status
    If (cd = "DEV") Then
        ip_form.btn_Save.Enabled = nc
        ip_form.btn_Delete.Enabled = nc
        ip_form.btn_Move_to_Ad_Hoc.Enabled = nc
        ip_form.btn_Deploy_to_Acceptance.Enabled = nc
        ip_form.btn_Move_to_Out_of_Scope.Enabled = nc
        Exit Sub
    End If
    '
    ' Open buttons on "Acceptance"-status
    If (cd = "UAT") Then
        ip_form.btn_Back_to_Development.Enabled = nc
        ip_form.btn_Deploy_to_Production.Enabled = nc
        Exit Sub
    End If
    '
    ' Open buttons on "Production, Adhoc or Out-of-Scope"-status
    If (cd = "PRD") Or (cd = "AHC") Or (cd = "OOS") Then
        ip_form.btn_Back_to_Development.Enabled = nc
        Exit Sub
    End If
    '
    ' If "refernced dataset" for form "dta_dataset" all buttons should be disabled, regardless of which "status" the dataset has.
    Dim rd As Boolean
    If (ip_form.Name = "dta_dataset") Then
        If (is_referenced_dataset(ip_form.Controls("id_model"), ip_form.Controls("id_dataset")) = True) Then
            ip_form.btn_Save.Enabled = True
            ip_form.btn_Back_to_Development.Enabled = False
            ip_form.btn_Deploy_to_Acceptance.Enabled = False
            ip_form.btn_Deploy_to_Production.Enabled = False
            ip_form.btn_Move_to_Ad_Hoc.Enabled = False
            ip_form.btn_Move_to_Out_of_Scope.Enabled = False
            ip_form.btn_Delete.Enabled = True
        End If
    End If
    '
End Sub

Public Function get_id_development_status(cd_development_status As String) As String
    Dim sql As String: sql = "SELECT id_development_status FROM srd_development_status WHERE cd_development_status = '" & cd_development_status & "'"
    Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset(sql)
    get_id_development_status = rst.fields("id_development_status")
End Function
Public Function get_cd_development_status(id_development_status As String) As String
    Dim sql As String: sql = "SELECT cd_development_status FROM srd_development_status WHERE id_development_status = '" & id_development_status & "'"
    Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset(sql)
    If (rst.EOF = True) Then: get_cd_development_status = "DEV": Exit Function
    get_cd_development_status = Nz(rst.fields("cd_development_status"), "DEV")
End Function

Public Sub btn_delete_Click(ob_form As Form, nm_form As String, nm_table As String): DoCmd.OpenForm nm_form, acNormal, , "id_" & nm_table & " = '" & ob_form.Controls("id_" & nm_table) & "' AND id_model = '" & ob_form.Controls("id_model") & "'", acFormReadOnly, acWindowNormal: End Sub
Public Sub btn_update_Click(ob_form As Form, nm_form As String, nm_table As String): DoCmd.OpenForm nm_form, acNormal, , "id_" & nm_table & " = '" & ob_form.Controls("id_" & nm_table) & "' AND id_model = '" & ob_form.Controls("id_model") & "'", acFormEdit, acWindowNormal: End Sub
Public Sub btn_insert_Click(ob_form As Form, nm_form As String, nm_table As String): DoCmd.OpenForm nm_form, acNormal, , "id_" & nm_table & " = '-'", acFormEdit, acWindowNormal: End Sub
'
Public Sub btn_Save_Click_for_dataset(ob_form As Form, nm_form_list As String, nm_table As String)
    '
    ' Save "Dataset"  to Export.
    If (Nz(ob_form.Controls("id_" & nm_table), "-1") = "-1") Then: Exit Sub
    DoCmd.RunCommand acCmdSaveRecord
    Select Case nm_form_list
        '
        Case "dta_dataset_list"
            Form_dta_dataset_list.Recalc
            '
        Case "dqm_dq_control_list"
            Form_dqm_dq_control_list.Recalc
            '
        Case "dqm_dq_requirement_list"
            Form_dqm_dq_requirement_list.Recalc
            '
        Case "ohg_related"
            Form_ohg_related.Recalc
            '
    End Select
    '
    ' Update Repository file(s)
    Call export_dataset_and_related_definitions(ob_form.Controls("id_dataset"))
    Call build_sql_file_dataset
    '
End Sub
Public Sub btn_delete_Click_for_dataset(ob_form As Form, nm_form_list As String, nm_schema As String, nm_table As String)
    '
    ' Delete Record in Access DB
    Dim nm_form As String: nm_form = ob_form.Name
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim fnm As String: fnm = ob_form.Controls("id_" & nm_table) & ".sql"
    Dim sql As String: sql = "DELETE * FROM " & nm_schema & "_" & nm_table & " WHERE id_" & nm_table & " = '" & ob_form.Controls("id_" & nm_table) & "'"
    If (MsgBox("This will remove all related to then `" & Replace(nm_table, "_", " ") & "`, clikc `Yes` to continue.", vbYesNo, "Delete") = vbYes) Then
        '
        ' "Delete"-record and reload list.
        DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
        '
        ' Remove the file as well
        If nm_table = "dataset" Then
            Call fso.DeleteFile(mdl_Folders.dta() & fnm)
        End If
        '
        ' Update Repository file(s)
        If (nm_schema & "_" & nm_table <> "dta_dataset") Then
            Call export_dataset_and_related_definitions(ob_form.Controls("id_dataset"))
        End If
        Call build_sql_file_dataset
        '
        ' Close "Current"-form.
        DoCmd.Close acForm, nm_form, acSaveNo
        '
        Select Case nm_form_list
            '
            Case "dta_dataset_list"
                Form_dta_dataset_list.Recalc
                Form_dta_dataset_list.Refresh
                '
            Case "dqm_dq_control_list"
                Form_dqm_dq_control_list.Recalc
                Form_dqm_dq_control_list.Refresh
                '
            Case "dqm_dq_requirement_list"
                Form_dqm_dq_requirement_list.Recalc
                Form_dqm_dq_requirement_list.Refresh
                '
        End Select
        '
    End If
    '
End Sub
Public Sub update_development_status(ob_form As Form, cd_development_status As String)
    '
    ' Update "Development"-status to "Development"
    ob_form.id_development_status = get_id_development_status(cd_development_status)
    DoCmd.RunCommand acCmdSaveRecord
    Call export_dataset_and_related_definitions(ob_form.id_dataset)
    Call build_sql_file_dataset
    Call set_buttons(ob_form)
    Call set_attributes(ob_form)
    '
End Sub
Public Sub set_attributes(ob_form As Form)
    '
    ' Local Variables
    Dim nc As Boolean:         nc = IIf(ob_form.id_model = id_model_default() And ob_form.Name = "dta_dataset", False, True)
    Dim nmf As String:        nmf = ob_form.Name
    Dim sts As String:        sts = get_cd_development_status(Nz(ob_form.id_development_status, ""))
    Dim shm As String:        shm = Left(nmf, 3)
    Dim tbl As String:        tbl = Mid(Replace(nmf, "_from_dq_requirement", ""), 5)
    Dim sql As String:        sql = "SELECT * FROM " & shm & "_" & tbl & " WHERE 1=2"
    Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset(sql)
    Dim lck As Boolean:       lck = IIf(sts = "DEV", nc, True)
    Dim fld As Field
    Dim ctr As Control
    '
    ' If "refernced dataset" for form "dta_dataset" all fields should be disabled, regardless of which "status" the dataset has.
    Dim rd As Boolean
    If (ob_form.Name = "dta_dataset") Then
        If (is_referenced_dataset(ob_form.Controls("id_model"), ob_form.Controls("id_dataset")) = True) Then
            lck = True
        End If
    End If
    '
    ' Loop through "Attributes"
    For Each fld In rst.fields: ob_form.Controls(fld.Name).Locked = lck: Next fld
    '
    ' Loop through "Controls"
    For Each ctr In ob_form.Controls
        If (ctr.Name = "ohg_related") _
        Or (ctr.Name = "dta_attribute") _
        Or (ctr.Name = "dqm_dq_threshold") _
        Or (ctr.Name = "dta_ingestion_etl") _
        Or (ctr.Name = "dta_parameter_value") _
        Or (ctr.Name = "dta_schedule") Then
            ctr.Locked = lck
        End If
    Next ctr
    '
End Sub
Public Sub btn_copy_dataset_Click(id_dataset As String, my As Form)
    '
    ' Generate new_id_dataset
    Dim new_id_dataset As String: new_id_dataset = get_random_code(32)
    '
    ' Copy the following table record of provided id_dataset and give theme new id`s.
    Call copy_table_record(id_model_default, id_dataset, new_id_dataset, "dta_dataset")
    Call copy_table_record(id_model_default, id_dataset, new_id_dataset, "dta_attribute")
    Call copy_table_record(id_model_default, id_dataset, new_id_dataset, "dta_ingestion_etl")
    Call copy_table_record(id_model_default, id_dataset, new_id_dataset, "dta_parameter_value")
    Call copy_table_record(id_model_default, id_dataset, new_id_dataset, "dta_schedule")
    Call copy_table_record(id_model_default, id_dataset, new_id_dataset, "ohg_related")
    '
    ' Have the form open the "new" dataset.
    Call DoCmd.OpenForm(my.Name, acNormal, , "id_dataset = '" & new_id_dataset & "'", acFormEdit, acWindowNormal)
    '
End Sub
Public Sub copy_table_record(ip_id_model As String, id_dataset As String, new_id_dataset As String, nm_table As String)
    '
    ' Open recordset for "Dataset
    Dim ent As String:        ent = Mid(nm_table, 5)         ' entity name
    Dim idf As String:        idf = "id_" & Mid(nm_table, 5) ' id-field name
    Dim sql As String:        sql = "SELECT * FROM " & nm_table & " WHERE id_dataset = '" & id_dataset & "' AND id_model = '" & ip_id_model & "'"
    Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset(sql)
    Dim grp As Recordset: Set grp = CurrentDb.OpenRecordset("SELECT TOP 1 id_group FROM ohg_group WHERE id_model = '" & id_model_default & "'")
    Dim cln As Recordset
    Dim fld As Field
    '
    ' clone record
    Set cln = CurrentDb.OpenRecordset(sql & " AND 1=2")
    '
    ' loop through all records and fields and copy the data to new records, if an id-field of nm_table the generate new id.
    Do While Not rst.EOF
        '
        ' clone record
        cln.AddNew
        '
        ' copy all fields
        For Each fld In rst.fields
            '
            ' Copy field
            cln.fields(fld.Name) = fld
            '
            ' generate new id
            If fld.Name = "id_dataset" Then
                cln.fields(fld.Name).Value = new_id_dataset
            End If
            If fld.Name = idf And Not (ent = "dataset") Then
                cln.fields(fld.Name).Value = get_random_code(32)
            End If
            '
            ' set development status to "Development"
            If fld.Name = "id_development_status" Then
                cln.fields(fld.Name) = "06010b0900010908010d0e0404021503"
            End If
            '
            ' If "Entity" is "Dataset" then append "Copy" to the fn_dataset.
            If ent = "dataset" Then
                If fld.Name = "fn_dataset" Then
                    cln.fields(fld.Name).Value = fld & " - Copy"
                End If
                If fld.Name = "nm_target_table" Then
                    cln.fields(fld.Name).Value = fld & "_copy"
                End If
                If fld.Name = "tx_source_query" Then
                    cln.fields(fld.Name).Value = Replace(fld, rst.fields("nm_target_table"), rst.fields("nm_target_table") & "_copy")
                End If
                If fld.Name = "id_group" Then
                    cln.fields(fld.Name).Value = grp.fields("id_group")
                End If
                
            End If
            '
        Next fld
        '
        ' set Id_model to "Current"-model
        cln.fields("id_model") = id_model_default
        '
        ' Save new record
        cln.Update
        '
    rst.MoveNext: Loop
    '
    ' Close clone recordset
    cln.Close
    '
End Sub
Public Function id_model_default() As String
    id_model_default = mdl_Folders.id_model(mdl_Folders.nm_repository())
End Function
    
Public Function is_referenced_dataset(ip_id_model As String, ip_id_dataset As String) As Boolean
    Dim sql As String: sql = "SELECT DISTINCT pvl.id_model, pvl.id_dataset, pgp.cd_parameter_group, Left([cd_parameter_group],3) AS [Check]" & vbNewLine _
                           & "FROM (srd_parameter_group AS pgp RIGHT" & vbNewLine _
                           & "JOIN srd_parameter AS pmt ON (pgp.id_parameter_group = pmt.id_parameter_group) AND (pgp.id_model = pmt.id_model)) RIGHT" & vbNewLine _
                           & "JOIN dta_parameter_value AS pvl ON (pmt.id_parameter = pvl.id_parameter) AND (pmt.id_model = pvl.id_model)" & vbNewLine _
                           & "WHERE (((Left([cd_parameter_group],3)) In ('rds','rdo'))) AND (pvl.id_model = '" & ip_id_model & "') AND (pvl.id_dataset = '" & ip_id_dataset & "')"
    Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset(sql)
    is_referenced_dataset = Not rst.EOF
    rst.Close
    '
End Function