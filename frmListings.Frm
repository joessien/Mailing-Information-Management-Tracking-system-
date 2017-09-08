VERSION 5.00
Begin VB.Form FrmListings 
   BackColor       =   &H00B19983&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Tracking System"
   ClientHeight    =   6630
   ClientLeft      =   1305
   ClientTop       =   1380
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   8280
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdDDept 
      Caption         =   "Domiciled Mails by Domiciled Date"
      Height          =   495
      Left            =   4320
      TabIndex        =   16
      Top             =   5400
      Width           =   3375
   End
   Begin VB.CommandButton cmfHold 
      Caption         =   "Mails on Hold"
      Height          =   495
      Left            =   480
      TabIndex        =   13
      Top             =   5400
      Width           =   3375
   End
   Begin VB.CommandButton cmdOSubj 
      Caption         =   "Sent Mails by Subject"
      Height          =   495
      Left            =   4320
      TabIndex        =   12
      Top             =   3240
      Width           =   3375
   End
   Begin VB.CommandButton cmdOref 
      Caption         =   "Sent Mails by Reference Number"
      Height          =   495
      Left            =   4320
      TabIndex        =   11
      Top             =   2760
      Width           =   3375
   End
   Begin VB.CommandButton cmdOdate 
      Caption         =   "Sent Mails by Out Date"
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   2280
      Width           =   3375
   End
   Begin VB.CommandButton cmdOOrig 
      Caption         =   "Sent Mails by Originator"
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   1800
      Width           =   3375
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Sent Mails by Receiving Department"
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CommandButton cmdOBen 
      Caption         =   "Sent Mails by Beneficiary"
      Height          =   495
      Left            =   4320
      TabIndex        =   8
      Top             =   840
      Width           =   3375
   End
   Begin VB.CommandButton cmdISubj 
      Caption         =   "Recieved by Subject"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   3240
      Width           =   3375
   End
   Begin VB.CommandButton cmdIRef 
      Caption         =   "Received Mails by Reference Number"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   2760
      Width           =   3375
   End
   Begin VB.CommandButton cmdIDATE 
      Caption         =   "Received Mails by In Date"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   2280
      Width           =   3375
   End
   Begin VB.CommandButton cmdIOrig 
      Caption         =   "Received Mails by Originator"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   3375
   End
   Begin VB.CommandButton cmdODept 
      Caption         =   "Received Mails by Originating Department"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CommandButton cmdIben 
      Caption         =   "Received Mails by Beneficiary"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.CommandButton cmdRDate 
      Caption         =   "Removed Mails by In Date"
      Height          =   495
      Left            =   480
      TabIndex        =   15
      Top             =   4920
      Width           =   3375
   End
   Begin VB.CommandButton cmdRRef 
      Caption         =   "Removed Mails by Reference"
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   4440
      Width           =   3375
   End
   Begin VB.CommandButton cmdDDate 
      Caption         =   "Domiciled Mails by In Date"
      Height          =   495
      Left            =   4320
      TabIndex        =   18
      Top             =   4920
      Width           =   3375
   End
   Begin VB.CommandButton cmdDRef 
      Caption         =   " Domiciled Mails by Reference"
      Height          =   495
      Left            =   4320
      TabIndex        =   17
      Top             =   4440
      Width           =   3375
   End
   Begin VB.Label fhdr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sorted Mail Listings"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "FrmListings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim db1 As Database, wrktemp As Workspace
Dim PicName As String, DateBuff, AgeBuff As String

Dim rstState, RegTab, rstTitle As Recordset
Dim rstsClass, rstBUnitD As Recordset
Dim rstBUnitR, rstFac, rstDept As Recordset
Dim rstPrior, rstSubj As Recordset
Dim rstMMast, rstSelData As Recordset

Dim AddFlag As Boolean, PBuff As Variant
Dim EditFlag As Boolean, MaxErrors As Integer
Dim Fldchk As Boolean, Insql As String
Dim strpos As Integer, tempBuff As String
Dim FirstTime As Integer, FirstPass As Integer
Dim GenClass As New mProLog, buff As String
Dim BType As String

Private Sub cmdDDate_Click()
On Error GoTo PError
    mvSql = "Select * From dommailquery order by indate"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF DOMICILED MAILS - Domiciled Mails by In Date"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         frmBrowse.grdGeneral.Columns(0).Width = 1000
         frmBrowse.grdGeneral.Columns(1).Width = 1000
         frmBrowse.grdGeneral.Columns(2).Width = 1000
         frmBrowse.grdGeneral.Columns(3).Width = 3000
         frmBrowse.grdGeneral.Columns(4).Width = 2000
         frmBrowse.grdGeneral.Columns(5).Width = 2000
         frmBrowse.grdGeneral.Columns(6).Width = 2000
         frmBrowse.grdGeneral.Columns(7).Width = 2800
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2000
         Col0.Caption = "Mail Ref."
         Col1.Caption = "Date In       "
         Col2.Caption = "Date Dom."
         Col3.Caption = "Reason for Domiciling"
         Col4.Caption = "Mail Details"
         Col5.Caption = "Deparment"
         Col6.Caption = "Subjects"
         Col7.Caption = "Mail Sub Classification"
         Col8.Caption = "Beneficiary"
         Col9.Caption = "Authorized By"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdDDept_Click()
On Error GoTo PError
    mvSql = "Select * From dommailquery order by outdate"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF DOMICILED MAILS - Domiciled Mails by Domiciled Date"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         frmBrowse.grdGeneral.Columns(0).Width = 1000
         frmBrowse.grdGeneral.Columns(1).Width = 1000
         frmBrowse.grdGeneral.Columns(2).Width = 1000
         frmBrowse.grdGeneral.Columns(3).Width = 3000
         frmBrowse.grdGeneral.Columns(4).Width = 2000
         frmBrowse.grdGeneral.Columns(5).Width = 2000
         frmBrowse.grdGeneral.Columns(6).Width = 2000
         frmBrowse.grdGeneral.Columns(7).Width = 2800
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2000
         Col0.Caption = "Mail Ref."
         Col1.Caption = "Date In       "
         Col2.Caption = "Date Dom."
         Col3.Caption = "Reason for Domiciling"
         Col4.Caption = "Mail Details"
         Col5.Caption = "Deparment"
         Col6.Caption = "Subjects"
         Col7.Caption = "Mail Sub Classification"
         Col8.Caption = "Beneficiary"
         Col9.Caption = "Authorized By"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdDRef_Click()
On Error GoTo PError
    mvSql = "Select * From dommailquery order by mailref"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF DOMICILED MAILS - Domiciled Mails by Reference"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         frmBrowse.grdGeneral.Columns(0).Width = 1000
         frmBrowse.grdGeneral.Columns(1).Width = 1000
         frmBrowse.grdGeneral.Columns(2).Width = 1000
         frmBrowse.grdGeneral.Columns(3).Width = 3000
         frmBrowse.grdGeneral.Columns(4).Width = 2000
         frmBrowse.grdGeneral.Columns(5).Width = 2000
         frmBrowse.grdGeneral.Columns(6).Width = 2000
         frmBrowse.grdGeneral.Columns(7).Width = 2800
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2000
         Col0.Caption = "Mail Ref."
         Col1.Caption = "Date In       "
         Col2.Caption = "Date Dom."
         Col3.Caption = "Reason for Domiciling"
         Col4.Caption = "Mail Details"
         Col5.Caption = "Deparment"
         Col6.Caption = "Subjects"
         Col7.Caption = "Mail Sub Classification"
         Col8.Caption = "Beneficiary"
         Col9.Caption = "Authorized By"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdIben_Click()
On Error GoTo PError
    mvSql = "Select * From inmailquery order by bennames"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF INCOMING MAILS - Received Mails by Beneficiary"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         Set col10 = frmBrowse.grdGeneral.Columns(10)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 800
         frmBrowse.grdGeneral.Columns(2).Width = 1400
         frmBrowse.grdGeneral.Columns(3).Width = 2000
         frmBrowse.grdGeneral.Columns(4).Width = 2200
         frmBrowse.grdGeneral.Columns(5).Width = 1000
         frmBrowse.grdGeneral.Columns(6).Width = 1000
         frmBrowse.grdGeneral.Columns(7).Width = 2800
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2200
         frmBrowse.grdGeneral.Columns(10).Width = 3000
         Col0.Caption = "Reference"
         Col1.Caption = "     "
         Col2.Caption = "Originated By"
         Col3.Caption = "Department"
         Col4.Caption = "Beneficiary"
         Col5.Caption = "Indate"
         Col6.Caption = "Outdate"
         Col7.Caption = "Subject"
         Col8.Caption = "Classification"
         Col9.Caption = "Received By"
         col10.Caption = "Mail Description"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdIDATE_Click()
On Error GoTo PError
    mvSql = "Select * From inmailquery order by indate"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF INCOMING MAILS - Received Mails by In Date"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         Set col10 = frmBrowse.grdGeneral.Columns(10)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 800
         frmBrowse.grdGeneral.Columns(2).Width = 1400
         frmBrowse.grdGeneral.Columns(3).Width = 2000
         frmBrowse.grdGeneral.Columns(4).Width = 2200
         frmBrowse.grdGeneral.Columns(5).Width = 1000
         frmBrowse.grdGeneral.Columns(6).Width = 1000
         frmBrowse.grdGeneral.Columns(7).Width = 2800
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2200
         frmBrowse.grdGeneral.Columns(10).Width = 3000
         Col0.Caption = "Reference"
         Col1.Caption = "     "
         Col2.Caption = "Originated By"
         Col3.Caption = "Department"
         Col4.Caption = "Beneficiary"
         Col5.Caption = "Indate"
         Col6.Caption = "Outdate"
         Col7.Caption = "Subject"
         Col8.Caption = "Classification"
         Col9.Caption = "Received By"
         col10.Caption = "Mail Description"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdIFac_Click()
On Error GoTo PError
    mvSql = "Select * From inmailquery order by cmbfac"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF INCOMING MAILS"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         Set col10 = frmBrowse.grdGeneral.Columns(10)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 800
         frmBrowse.grdGeneral.Columns(2).Width = 1400
         frmBrowse.grdGeneral.Columns(3).Width = 2000
         frmBrowse.grdGeneral.Columns(4).Width = 2200
         frmBrowse.grdGeneral.Columns(5).Width = 1000
         frmBrowse.grdGeneral.Columns(6).Width = 1000
         frmBrowse.grdGeneral.Columns(7).Width = 2800
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2200
         frmBrowse.grdGeneral.Columns(10).Width = 3000
         Col0.Caption = "Reference"
         Col1.Caption = "     "
         Col2.Caption = "Originated By"
         Col3.Caption = "Department"
         Col4.Caption = "Beneficiary"
         Col5.Caption = "Indate"
         Col6.Caption = "Outdate"
         Col7.Caption = "Subject"
         Col8.Caption = "Classification"
         Col9.Caption = "Received By"
         col10.Caption = "Mail Description"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdIOrig_Click()
On Error GoTo PError
    mvSql = "Select * From inmailquery order by osurname"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF INCOMING MAILS - Received Mails by Originator"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         Set col10 = frmBrowse.grdGeneral.Columns(10)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 800
         frmBrowse.grdGeneral.Columns(2).Width = 1400
         frmBrowse.grdGeneral.Columns(3).Width = 2000
         frmBrowse.grdGeneral.Columns(4).Width = 2200
         frmBrowse.grdGeneral.Columns(5).Width = 1000
         frmBrowse.grdGeneral.Columns(6).Width = 1000
         frmBrowse.grdGeneral.Columns(7).Width = 2800
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2200
         frmBrowse.grdGeneral.Columns(10).Width = 3000
         Col0.Caption = "Reference"
         Col1.Caption = "     "
         Col2.Caption = "Originated By"
         Col3.Caption = "Department"
         Col4.Caption = "Beneficiary"
         Col5.Caption = "Indate"
         Col6.Caption = "Outdate"
         Col7.Caption = "Subject"
         Col8.Caption = "Classification"
         Col9.Caption = "Received By"
         col10.Caption = "Mail Description"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdIRef_Click()
On Error GoTo PError
    mvSql = "Select * From inmailquery order by mailref"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF INCOMING MAILS - Received Mails by Reference Number"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         Set col10 = frmBrowse.grdGeneral.Columns(10)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 800
         frmBrowse.grdGeneral.Columns(2).Width = 1400
         frmBrowse.grdGeneral.Columns(3).Width = 2000
         frmBrowse.grdGeneral.Columns(4).Width = 2200
         frmBrowse.grdGeneral.Columns(5).Width = 1000
         frmBrowse.grdGeneral.Columns(6).Width = 1000
         frmBrowse.grdGeneral.Columns(7).Width = 2800
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2200
         frmBrowse.grdGeneral.Columns(10).Width = 3000
         Col0.Caption = "Reference"
         Col1.Caption = "     "
         Col2.Caption = "Originated By"
         Col3.Caption = "Department"
         Col4.Caption = "Beneficiary"
         Col5.Caption = "Indate"
         Col6.Caption = "Outdate"
         Col7.Caption = "Subject"
         Col8.Caption = "Classification"
         Col9.Caption = "Received By"
         col10.Caption = "Mail Description"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdISubj_Click()
On Error GoTo PError
    mvSql = "Select * From inmailquery order by DefSubJ_RDesc"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF INCOMING MAILS - Recieved by Subject"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         Set col10 = frmBrowse.grdGeneral.Columns(10)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 800
         frmBrowse.grdGeneral.Columns(2).Width = 1400
         frmBrowse.grdGeneral.Columns(3).Width = 2000
         frmBrowse.grdGeneral.Columns(4).Width = 2200
         frmBrowse.grdGeneral.Columns(5).Width = 1000
         frmBrowse.grdGeneral.Columns(6).Width = 1000
         frmBrowse.grdGeneral.Columns(7).Width = 2800
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2200
         frmBrowse.grdGeneral.Columns(10).Width = 3000
         Col0.Caption = "Reference"
         Col1.Caption = "     "
         Col2.Caption = "Originated By"
         Col3.Caption = "Department"
         Col4.Caption = "Beneficiary"
         Col5.Caption = "Indate"
         Col6.Caption = "Outdate"
         Col7.Caption = "Subject"
         Col8.Caption = "Classification"
         Col9.Caption = "Received By"
         col10.Caption = "Mail Description"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdOBen_Click()
On Error GoTo PError
    mvSql = "Select * From outmailquery order by bennames"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF OUTGOING MAILS - Sent Mails by Beneficiary"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         Set col10 = frmBrowse.grdGeneral.Columns(10)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 800
         frmBrowse.grdGeneral.Columns(2).Width = 1400
         frmBrowse.grdGeneral.Columns(3).Width = 2000
         frmBrowse.grdGeneral.Columns(4).Width = 1000
         frmBrowse.grdGeneral.Columns(5).Width = 1000
         frmBrowse.grdGeneral.Columns(6).Width = 1400
         frmBrowse.grdGeneral.Columns(7).Width = 2000
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2200
         frmBrowse.grdGeneral.Columns(10).Width = 3000
         Col0.Caption = "Mail Ref."
         Col1.Caption = "     "
         Col2.Caption = "Originated By"
         Col3.Caption = "Beneficiary"
         Col4.Caption = "In Date"
         Col5.Caption = "Out Date"
         Col6.Caption = "Unit Sent"
         Col7.Caption = "Subject"
         Col8.Caption = "Clasification"
         Col9.Caption = "Received By"
         col10.Caption = "Mail Description"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdOdate_Click()
On Error GoTo PError
    mvSql = "Select * From outmailquery order by outdate"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF OUTGOING MAILS - Sent Mails by Out Date"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         Set col10 = frmBrowse.grdGeneral.Columns(10)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 800
         frmBrowse.grdGeneral.Columns(2).Width = 1400
         frmBrowse.grdGeneral.Columns(3).Width = 2000
         frmBrowse.grdGeneral.Columns(4).Width = 1000
         frmBrowse.grdGeneral.Columns(5).Width = 1000
         frmBrowse.grdGeneral.Columns(6).Width = 1400
         frmBrowse.grdGeneral.Columns(7).Width = 2000
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2200
         frmBrowse.grdGeneral.Columns(10).Width = 3000
         Col0.Caption = "Mail Ref."
         Col1.Caption = "     "
         Col2.Caption = "Originated By"
         Col3.Caption = "Beneficiary"
         Col4.Caption = "In Date"
         Col5.Caption = "Out Date"
         Col6.Caption = "Unit Sent"
         Col7.Caption = "Subject"
         Col8.Caption = "Clasification"
         Col9.Caption = "Received By"
         col10.Caption = "Mail Description"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdODept_Click()
On Error GoTo PError
    mvSql = "Select * From inmailquery order by Defdept_rdesc"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF INCOMING MAILS - Received Mails by Originating Department"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         Set col10 = frmBrowse.grdGeneral.Columns(10)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 800
         frmBrowse.grdGeneral.Columns(2).Width = 1400
         frmBrowse.grdGeneral.Columns(3).Width = 2000
         frmBrowse.grdGeneral.Columns(4).Width = 2200
         frmBrowse.grdGeneral.Columns(5).Width = 1000
         frmBrowse.grdGeneral.Columns(6).Width = 1000
         frmBrowse.grdGeneral.Columns(7).Width = 2800
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2200
         frmBrowse.grdGeneral.Columns(10).Width = 3000
         Col0.Caption = "Reference"
         Col1.Caption = "     "
         Col2.Caption = "Originated By"
         Col3.Caption = "Department"
         Col4.Caption = "Beneficiary"
         Col5.Caption = "Indate"
         Col6.Caption = "Outdate"
         Col7.Caption = "Subject"
         Col8.Caption = "Classification"
         Col9.Caption = "Received By"
         col10.Caption = "Mail Description"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdOFac_Click()
On Error GoTo PError
    mvSql = "Select * From outmailquery order by cmbfac"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF OUTGOING MAILS"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         Set col10 = frmBrowse.grdGeneral.Columns(10)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 800
         frmBrowse.grdGeneral.Columns(2).Width = 1400
         frmBrowse.grdGeneral.Columns(3).Width = 2000
         frmBrowse.grdGeneral.Columns(4).Width = 1000
         frmBrowse.grdGeneral.Columns(5).Width = 1000
         frmBrowse.grdGeneral.Columns(6).Width = 1400
         frmBrowse.grdGeneral.Columns(7).Width = 2000
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2200
         frmBrowse.grdGeneral.Columns(10).Width = 3000
         Col0.Caption = "Mail Ref."
         Col1.Caption = "     "
         Col2.Caption = "Originated By"
         Col3.Caption = "Beneficiary"
         Col4.Caption = "In Date"
         Col5.Caption = "Out Date"
         Col6.Caption = "Unit Sent"
         Col7.Caption = "Subject"
         Col8.Caption = "Clasification"
         Col9.Caption = "Received By"
         col10.Caption = "Mail Description"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdOOrig_Click()
On Error GoTo PError
    mvSql = "Select * From outmailquery order by osurname"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF OUTGOING MAILS - Sent Mails by Originator"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         Set col10 = frmBrowse.grdGeneral.Columns(10)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 800
         frmBrowse.grdGeneral.Columns(2).Width = 1400
         frmBrowse.grdGeneral.Columns(3).Width = 2000
         frmBrowse.grdGeneral.Columns(4).Width = 1000
         frmBrowse.grdGeneral.Columns(5).Width = 1000
         frmBrowse.grdGeneral.Columns(6).Width = 1400
         frmBrowse.grdGeneral.Columns(7).Width = 2000
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2200
         frmBrowse.grdGeneral.Columns(10).Width = 3000
         Col0.Caption = "Mail Ref."
         Col1.Caption = "     "
         Col2.Caption = "Originated By"
         Col3.Caption = "Beneficiary"
         Col4.Caption = "In Date"
         Col5.Caption = "Out Date"
         Col6.Caption = "Unit Sent"
         Col7.Caption = "Subject"
         Col8.Caption = "Clasification"
         Col9.Caption = "Received By"
         col10.Caption = "Mail Description"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdOref_Click()
On Error GoTo PError
    mvSql = "Select * From outmailquery order by mailref"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF OUTGOING MAILS - Sent Mails by Reference Number"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         Set col10 = frmBrowse.grdGeneral.Columns(10)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 800
         frmBrowse.grdGeneral.Columns(2).Width = 1400
         frmBrowse.grdGeneral.Columns(3).Width = 2000
         frmBrowse.grdGeneral.Columns(4).Width = 1000
         frmBrowse.grdGeneral.Columns(5).Width = 1000
         frmBrowse.grdGeneral.Columns(6).Width = 1400
         frmBrowse.grdGeneral.Columns(7).Width = 2000
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2200
         frmBrowse.grdGeneral.Columns(10).Width = 3000
         Col0.Caption = "Mail Ref."
         Col1.Caption = "     "
         Col2.Caption = "Originated By"
         Col3.Caption = "Beneficiary"
         Col4.Caption = "In Date"
         Col5.Caption = "Out Date"
         Col6.Caption = "Unit Sent"
         Col7.Caption = "Subject"
         Col8.Caption = "Clasification"
         Col9.Caption = "Received By"
         col10.Caption = "Mail Description"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdOSubj_Click()
On Error GoTo PError
    mvSql = "Select * From outmailquery order by DefSubJ_RDesc"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF OUTGOING MAILS - Sent Mails by Subject"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         Set col10 = frmBrowse.grdGeneral.Columns(10)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 800
         frmBrowse.grdGeneral.Columns(2).Width = 1400
         frmBrowse.grdGeneral.Columns(3).Width = 2000
         frmBrowse.grdGeneral.Columns(4).Width = 1000
         frmBrowse.grdGeneral.Columns(5).Width = 1000
         frmBrowse.grdGeneral.Columns(6).Width = 1400
         frmBrowse.grdGeneral.Columns(7).Width = 2000
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2200
         frmBrowse.grdGeneral.Columns(10).Width = 3000
         Col0.Caption = "Mail Ref."
         Col1.Caption = "     "
         Col2.Caption = "Originated By"
         Col3.Caption = "Beneficiary"
         Col4.Caption = "In Date"
         Col5.Caption = "Out Date"
         Col6.Caption = "Unit Sent"
         Col7.Caption = "Subject"
         Col8.Caption = "Clasification"
         Col9.Caption = "Received By"
         col10.Caption = "Mail Description"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdRDate_Click()
On Error GoTo PError
    mvSql = "Select * From binmailquery order by indate"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF BIN MAILS - Removed Mails by In Date"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         frmBrowse.grdGeneral.Columns(0).Width = 1000
         frmBrowse.grdGeneral.Columns(1).Width = 1000
         frmBrowse.grdGeneral.Columns(2).Width = 2400
         frmBrowse.grdGeneral.Columns(3).Width = 2400
         frmBrowse.grdGeneral.Columns(4).Width = 2000
         frmBrowse.grdGeneral.Columns(5).Width = 2000
         frmBrowse.grdGeneral.Columns(6).Width = 2000
         frmBrowse.grdGeneral.Columns(7).Width = 2800
         frmBrowse.grdGeneral.Columns(8).Width = 2000
         Col0.Caption = "Mail Ref."
         Col1.Caption = "In Date     "
         Col2.Caption = "Reason for Deleting"
         Col3.Caption = "Mail Details"
         Col4.Caption = "Deparment"
         Col5.Caption = "Subjects"
         Col6.Caption = "Mail Sub Classification"
         Col7.Caption = "Beneficiary"
         Col8.Caption = "Removal Authorized By"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdRRef_Click()
On Error GoTo PError
    mvSql = "Select * From binmailquery order by mailref"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF BIN MAILS - Removed Mails by Reference"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         frmBrowse.grdGeneral.Columns(0).Width = 1000
         frmBrowse.grdGeneral.Columns(1).Width = 1000
         frmBrowse.grdGeneral.Columns(2).Width = 2400
         frmBrowse.grdGeneral.Columns(3).Width = 2400
         frmBrowse.grdGeneral.Columns(4).Width = 2000
         frmBrowse.grdGeneral.Columns(5).Width = 2000
         frmBrowse.grdGeneral.Columns(6).Width = 2000
         frmBrowse.grdGeneral.Columns(7).Width = 2800
         frmBrowse.grdGeneral.Columns(8).Width = 2000
         Col0.Caption = "Mail Ref."
         Col1.Caption = "In Date     "
         Col2.Caption = "Reason for Deleting"
         Col3.Caption = "Mail Details"
         Col4.Caption = "Deparment"
         Col5.Caption = "Subjects"
         Col6.Caption = "Mail Sub Classification"
         Col7.Caption = "Beneficiary"
         Col8.Caption = "Removal Authorized By"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmfHold_Click()
On Error GoTo PError
    mvSql = "Select * From finmailquery order by mailref"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF INCOMING MAILS - Mails on Hold"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         Set col10 = frmBrowse.grdGeneral.Columns(10)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 800
         frmBrowse.grdGeneral.Columns(2).Width = 1400
         frmBrowse.grdGeneral.Columns(3).Width = 2000
         frmBrowse.grdGeneral.Columns(4).Width = 2200
         frmBrowse.grdGeneral.Columns(5).Width = 1000
         frmBrowse.grdGeneral.Columns(6).Width = 1000
         frmBrowse.grdGeneral.Columns(7).Width = 2800
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2200
         frmBrowse.grdGeneral.Columns(10).Width = 3000
         Col0.Caption = "Reference"
         Col1.Caption = "     "
         Col2.Caption = "Originated By"
         Col3.Caption = "Department"
         Col4.Caption = "Beneficiary"
         Col5.Caption = "Indate"
         Col6.Caption = "Outdate"
         Col7.Caption = "Subject"
         Col8.Caption = "Classification"
         Col9.Caption = "Received By"
         col10.Caption = "Mail Description"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub Command11_Click()
On Error GoTo PError
    mvSql = "Select * From outmailquery order by Rdesc"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF OUTGOING MAILS - Sent Mails by Receiving Department"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         Set col10 = frmBrowse.grdGeneral.Columns(10)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 800
         frmBrowse.grdGeneral.Columns(2).Width = 1400
         frmBrowse.grdGeneral.Columns(3).Width = 2000
         frmBrowse.grdGeneral.Columns(4).Width = 1000
         frmBrowse.grdGeneral.Columns(5).Width = 1000
         frmBrowse.grdGeneral.Columns(6).Width = 1400
         frmBrowse.grdGeneral.Columns(7).Width = 2000
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2200
         frmBrowse.grdGeneral.Columns(10).Width = 3000
         Col0.Caption = "Mail Ref."
         Col1.Caption = "     "
         Col2.Caption = "Originated By"
         Col3.Caption = "Beneficiary"
         Col4.Caption = "In Date"
         Col5.Caption = "Out Date"
         Col6.Caption = "Unit Sent"
         Col7.Caption = "Subject"
         Col8.Caption = "Clasification"
         Col9.Caption = "Received By"
         col10.Caption = "Mail Description"
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub Form_Load()
On Error GoTo PError
     GenClass.fleLogin mvUserid, "Accessed Update outgoing mail", Date, Time
     Dim AppStatBuff As String
     Call FormCentreMDI(Me)
     Set wrktemp = DBEngine.Workspaces(0)
     Set db1 = wrktemp.OpenDatabase(DBpath, True)
     
     Set rstState = db1.OpenRecordset("defstate", dbOpenDynaset)
     Set RegTab = db1.OpenRecordset("dproreg", dbOpenDynaset)
     Set rstsClass = db1.OpenRecordset("defsubclass", dbOpenDynaset)
     Set rstDept = db1.OpenRecordset("Defdept", dbOpenDynaset)
     Set rstBUnitD = db1.OpenRecordset("defbunitd", dbOpenDynaset)
     Insql = "Select * from mailmast Where outmail = true"
     Set rstMMast = db1.OpenRecordset(Insql, dbOpenDynaset)
     Set rstBUnitR = db1.OpenRecordset("defbunitr", dbOpenDynaset)
     Set rstFac = db1.OpenRecordset("deffac", dbOpenDynaset)
     Set rstPrior = db1.OpenRecordset("defprior", dbOpenDynaset)
     Set rstSubj = db1.OpenRecordset("defsubj", dbOpenDynaset)
     Set rstTitle = db1.OpenRecordset("deftitle", dbOpenDynaset)
     
     Dim flag As Integer, mtable As String
     Me.Caption = MVCoyname
PError:

End Sub
