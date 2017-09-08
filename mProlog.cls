VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mProLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function fleLogin(clsUserID As String, txtRtnName As String, txtDatIn As Date, txtTmeIn As Date)
    Dim db6 As Database, rstLogfle As Recordset
    Set wrktemp = DBEngine.Workspaces(0)
    Set db6 = wrktemp.OpenDatabase(DBpath, True)
    Set rstLogfle = db6.OpenRecordset("LogFle", dbOpenDynaset)
    rstLogfle.AddNew
    If clsUserID = "" Then
        txtUserID = "***"
    Else
        txtUserID = clsUserID
    End If
    rstLogfle!userid = txtUserID
    rstLogfle!UserName = mvUserName
    rstLogfle!RtnName = Left(txtRtnName, 60)
    rstLogfle!DatIn = txtDatIn
    rstLogfle!TmeIn = txtTmeIn
    rstLogfle.Update
    rstLogfle.Close
End Function


