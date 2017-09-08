VERSION 5.00
Begin VB.Form frmAlog 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Tracking System"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   8685
   Begin VB.CommandButton cmdGrad 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Archive Domiciled Mails"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton CmdArchClr 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clear All Archive Records"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton CmdArch 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Archive Log Records"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton cmdSLClr 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Archive In/Out Mails"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton cmdSTLog 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Archive Deleted Mails"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdClr 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clear User  Log"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdViewR 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Log View By Routine"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmdViewS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Log View By User"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000FF00&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton cmdViewT 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Log View  By Time"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label fhdr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Archive and Log Management"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   0
      Width           =   8295
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   240
      Picture         =   "frmALog.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmAlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db1 As Database, wrktemp As Workspace
Dim rstUserFle As Recordset
Dim GenClass As New mProLog
Dim rstLog As Recordset

Private Sub CmdArchClr_Click()
On Error GoTo PError
MsgBox "Note, this process is irreversible. Press Okay to continue.", vbExclamation, "Archives and Logs"
Screen.MousePointer = vbHourglass
Dim i As Integer
    i = ConfirmLBin()
    If i = vbYes Then
        db1.Execute "DELETE * FROM Archmmast"
        db1.Execute "DELETE * FROM archdomm"
        db1.Execute "DELETE * FROM Archdelm"
        db1.Execute "DELETE * FROM Archlog"
        MsgBox "All archives have been obliterated", vbExclamation, "Archives and Logs"
    Else
        MsgBox "Archives delete Cancelled", vbExclamation, "Archives and Logs"
    End If
PError:
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdclr_Click()
On Error GoTo PError
MsgBox "Note, this process is irreversible. Press Okay to continue.", vbExclamation, "Archives and Logs"
Screen.MousePointer = vbHourglass
Dim i As Integer
    i = ConfirmLBin()
    If i = vbYes Then
        db1.Execute "DELETE * FROM logfle"
        MsgBox "User Activities Log binned", vbExclamation, "User Archives and Logs"
    Else
        MsgBox "Log bin Cancelled", vbExclamation, "User Archives and Logs"
    End If
PError:
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdGrad_Click()
On Error GoTo PError
MsgBox "Note, this process is irreversible. Press Okay to continue.", vbExclamation, "Archives and Logs"
Screen.MousePointer = vbHourglass
Dim i As Integer
    i = cont()
    If i = vbYes Then
        db1.Execute " INSERT INTO archDomM SELECT * FROM [dommails];"
        db1.Execute "DELETE * FROM dommails"
        MsgBox "All Domiciled Mails have been Archived", vbExclamation, "Archives and Logs"
    Else
        MsgBox "Archiving Cancelled", vbExclamation, "Archives and Logs"
    End If
PError:
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSLClr_Click()
On Error GoTo PError
MsgBox "Note, this process is irreversible. Press Okay to continue.", vbExclamation, "Archives and Logs"
Screen.MousePointer = vbHourglass
Dim i As Integer
    i = cont()
    If i = vbYes Then
        db1.Execute " INSERT INTO autommast SELECT * FROM [mailmast];"
        db1.Execute "DELETE * FROM mailmast"
        MsgBox "All Incoming and outgoing Mails have been Archived", vbExclamation, "Archives and Logs"
    Else
        MsgBox "Archiving Cancelled", vbExclamation, "Archives and Logs"
    End If
PError:
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSTLog_Click()
On Error GoTo PError
MsgBox "Note, this process is irreversible. Press Okay to continue.", vbExclamation, "Archives and Logs"
Screen.MousePointer = vbHourglass
Dim i As Integer
    i = cont()
    If i = vbYes Then
        db1.Execute " INSERT INTO autodelm SELECT * FROM [delmails];"
        db1.Execute "DELETE * FROM delmails"
        MsgBox "All Deleted Mails have been Archived", vbExclamation, "Archives and Logs"
    Else
        MsgBox "Archiving Cancelled", vbExclamation, "Archives and Logs"
    End If
PError:
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdViewR_Click()
On Error GoTo PError
     mvSql = "Select userid,username,rtnname,datin,tmein From logfle order by RTNNAME;"
     Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
     If (rst.BOF And rst.EOF) Then
         MsgBox "No activities", vbExclamation, "User Archives and Logs"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "USERS ACTIVITY SORTED BY EVENT CARRIED OUT"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 2800
         frmBrowse.grdGeneral.Columns(2).Width = 4500
         frmBrowse.grdGeneral.Columns(3).Width = 1200
         frmBrowse.grdGeneral.Columns(4).Width = 1200
         Col0.Caption = "User Id"
         Col1.Caption = "User Name"
         Col2.Caption = "Event Carried Out"
         Col3.Caption = "Date"
         Col4.Caption = "Time"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdViewS_Click()
On Error GoTo PError
     mvSql = "Select userid,username,rtnname,datin,tmein From logfle order by userid;"
     Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
     If (rst.BOF And rst.EOF) Then
         MsgBox "No activities", vbExclamation, "User Archives and Logs"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "USERS ACTIVITY SORTED BY STAFF IDENTIFICATION"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 2800
         frmBrowse.grdGeneral.Columns(2).Width = 4500
         frmBrowse.grdGeneral.Columns(3).Width = 1200
         frmBrowse.grdGeneral.Columns(4).Width = 1200
         Col0.Caption = "User Id"
         Col1.Caption = "User Name"
         Col2.Caption = "Event Carried Out"
         Col3.Caption = "Date"
         Col4.Caption = "Time"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdViewT_Click()
On Error GoTo PError
     mvSql = "Select userid,username,rtnname,datin,tmein From logfle order by datin,tmein;"
     Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
     If (rst.BOF And rst.EOF) Then
         MsgBox "No activities", vbExclamation, "User Archives and Logs"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "USERS ACTIVITY SORTED BY DATE AND TIME"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 2800
         frmBrowse.grdGeneral.Columns(2).Width = 4500
         frmBrowse.grdGeneral.Columns(3).Width = 1200
         frmBrowse.grdGeneral.Columns(4).Width = 1200
         Col0.Caption = "User Id"
         Col1.Caption = "User Name"
         Col2.Caption = "Event Carried Out"
         Col3.Caption = "Date"
         Col4.Caption = "Time"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub CmdArch_Click()
On Error GoTo PError
Dim i As Integer
    i = cont()
    If i = vbYes Then
        db1.Execute " INSERT INTO archlog SELECT * FROM [logfle];"
        db1.Execute "DELETE * FROM logfle"
        MsgBox "Log Reports have been Archived", vbExclamation, "Archives and Logs"
    Else
        MsgBox "Archiving Cancelled", vbExclamation, "Archives and Logs"
    End If
PError:
End Sub

Private Sub Form_Load()
    On Error GoTo PError
    Call FormCentreMDI(Me)
    Set wrktemp = DBEngine.Workspaces(0)
    Set db1 = wrktemp.OpenDatabase(DBpath, True)
    Set rstLog = db1.OpenRecordset("logfle", dbOpenDynaset)
    Me.Caption = MVCoyname
    GenClass.fleLogin mvUserid, "Accessed User Archives and Logs File", Date, Time
PError:
End Sub

