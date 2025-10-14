Attribute VB_Name = "sql_create_view"
Option Compare Database
Option Explicit

Sub CreateView(nm_view As String, tx_sql As String)
    Dim db As DAO.Database: Set db = CurrentDb
    Dim qdf As DAO.QueryDef
    Dim strSQL As String

    ' Get the current database
    

    ' Check if the query already exists, and delete it if it does
    On Error Resume Next
    db.QueryDefs.Delete nm_view
    On Error GoTo 0

    ' Create a new query (acts like a view)
    Set qdf = db.CreateQueryDef(nm_view, tx_sql)

    ' Clean up
    Set qdf = Nothing
    Set db = Nothing

End Sub

