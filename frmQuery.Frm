VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmDquery 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Tracking System"
   ClientHeight    =   6795
   ClientLeft      =   1305
   ClientTop       =   1380
   ClientWidth     =   10095
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   10095
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H00C0C0C0&
      Height          =   6735
      Index           =   0
      Left            =   0
      ScaleHeight     =   6675
      ScaleWidth      =   10035
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   8520
         TabIndex        =   27
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton cmdPhone 
         Caption         =   "Search"
         Height          =   495
         Left            =   7080
         TabIndex        =   26
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton cmdOStaff 
         Caption         =   "Search"
         Height          =   495
         Index           =   0
         Left            =   7080
         TabIndex        =   25
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton cmdBen 
         Caption         =   "Search"
         Height          =   495
         Left            =   7080
         TabIndex        =   24
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton cmdSubj 
         Caption         =   "Search"
         Height          =   495
         Left            =   7080
         TabIndex        =   23
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton cmdRBU 
         Caption         =   "Search"
         Height          =   495
         Left            =   7080
         TabIndex        =   22
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   "Search"
         Height          =   495
         Left            =   8520
         TabIndex        =   21
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdDept 
         Caption         =   "Search"
         Height          =   495
         Left            =   7080
         TabIndex        =   20
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton CmdRef 
         Caption         =   "Search"
         Height          =   495
         Left            =   5880
         TabIndex        =   19
         Top             =   1560
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Phone 
         Height          =   375
         Left            =   4680
         TabIndex        =   18
         Top             =   5880
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin VB.TextBox osurname 
         Height          =   375
         Left            =   4680
         MaxLength       =   20
         TabIndex        =   17
         Top             =   5280
         Width           =   2295
      End
      Begin VB.TextBox Ben 
         Height          =   375
         Left            =   4680
         MaxLength       =   20
         TabIndex        =   16
         Top             =   4680
         Width           =   2295
      End
      Begin VB.ComboBox cmbSubj 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   4080
         Width           =   3735
      End
      Begin VB.ComboBox cmbRBU 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3480
         Width           =   3735
      End
      Begin VB.ComboBox cmbDept 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2880
         Width           =   3735
      End
      Begin MSMask.MaskEdBox stDate 
         Height          =   375
         Left            =   4680
         TabIndex        =   12
         Top             =   2280
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.TextBox Mailref 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Data datgenTab 
         Caption         =   "Gentab"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   7320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.PictureBox SSPanel3 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   915
         Left            =   2520
         Picture         =   "frmQuery.frx":0000
         ScaleHeight     =   855
         ScaleWidth      =   5715
         TabIndex        =   1
         Top             =   120
         Width           =   5775
      End
      Begin MSComDlg.CommonDialog GetFile 
         Left            =   360
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "Bmp"
         FilterIndex     =   2
      End
      Begin MSMask.MaskEdBox eDate 
         Height          =   375
         Left            =   7080
         TabIndex        =   30
         Top             =   2280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Enter Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   32
         Top             =   6000
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "End Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   31
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Enter Surname:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   29
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Enter Surname:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   28
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Start Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Find Mail by Phone Number:"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   6000
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Find Mail by Originating Staff:"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   5400
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Find Mail by Beneficiary:"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   4800
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Find Mail by Subject:"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   4200
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Find Mail by Department Forwarded To:"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Find Mail by Originating Department:"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Find Mail by In Date Range:"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Find Mail by Reference:"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Image Image2 
         Height          =   855
         Left            =   960
         Picture         =   "frmQuery.frx":10C62
         Stretch         =   -1  'True
         Top             =   120
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   960
         Stretch         =   -1  'True
         Top             =   120
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmDquery"
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
Dim FldChk As Boolean, Insql As String
Dim strpos As Integer, tempBuff As String
Dim FirstTime As Integer, FirstPass As Integer
Dim GenClass As New mProLog, buff As String
Dim BType As String

Private Sub cmdBen_Click()
  Dim sql As String, digitflag As Integer
   Dim buff As String, ans As String
   On Error GoTo FindError1
   ans = UCase(Ben)
   If Len(ans) = 0 Then
      MsgBox "The mail beneficiary given is invalid, Retry", vbInformation, "Mail Query"
      ClearData
      Exit Sub
   Else
      mvSql = "Select * From Lmailquery where bennames LIKE " & Chr$(34) & ans & "*" & Chr$(34) & "order by bennames"
      Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
      QList
   End If
FindError0:
   Exit Sub

FindError1:
   MsgBox "ERROR:   " + Error$ + " has occurred", vbCritical, Me.Caption
   On Error GoTo 0

End Sub


Private Sub cmdDate_Click()
  Dim sql As String, digitflag As Integer
   Dim buff As String, ans As String
   On Error GoTo FindError1
   If CDate(stDate) > CDate(eDate) Then
      MsgBox "Start date cannot be after the End date, Retry", vbInformation, "Mail Query"
      stDate.SetFocus
      Exit Sub
   End If
   If Len(stDate) = 0 Or Len(eDate) = 0 Then
      MsgBox "Start date or End date cannot be blank, Retry", vbInformation, "Mail Query"
      stDate.SetFocus
      Exit Sub
   End If
   
   mvSql = "SELECT * FROM Lmailquery WHERE INDate >= #" & stDate & "# and INDate <= #" & eDate & "# order by bennames"
   Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
   QList
   
FindError0:
   Exit Sub
FindError1:
   MsgBox "ERROR:   " + Error$ + " has occurred", vbCritical, Me.Caption
   On Error GoTo 0
End Sub

Private Sub cmdDept_Click()
  Dim sql As String, digitflag As Integer
   Dim buff As String, ans As String
   'On Error GoTo FindError1
   strpos = InStr(1, cmbDept, ",", 1)
   ans = Trim(Left(cmbDept, strpos - 1))
   If Len(ans) = 0 Then
      MsgBox "The originating department given is invalid, Retry", vbInformation, "Mail Query"
      ClearData
      Exit Sub
   Else
      mvSql = "Select * From Lmailquery where cmbdept = '" & ans & "' order by bennames"
      Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
      QList
   End If
FindError0:
   Exit Sub

FindError1:
   MsgBox "ERROR:   " + Error$ + " has occurred", vbCritical, Me.Caption
   On Error GoTo 0

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOStaff_Click(Index As Integer)
  Dim sql As String, digitflag As Integer
   Dim buff As String, ans As String
   On Error GoTo FindError1
   ans = UCase(osurname)
   If Len(ans) = 0 Then
      MsgBox "The mail originator given is invalid, Retry", vbInformation, "Mail Query"
      ClearData
      Exit Sub
   Else
      mvSql = "Select * From Lmailquery where osurname LIKE " & Chr$(34) & ans & "*" & Chr$(34) & "order by osurname"
      Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
      QList
   End If
FindError0:
   Exit Sub

FindError1:
   MsgBox "ERROR:   " + Error$ + " has occurred", vbCritical, Me.Caption
   On Error GoTo 0

End Sub

Private Sub cmdPhone_Click()
  Dim sql As String, digitflag As Integer
   Dim buff As String, ans As String
   On Error GoTo FindError1
   ans = UCase(Phone)
   If Len(ans) = 0 Then
      MsgBox "The contact phone number given is invalid, Retry", vbInformation, "Mail Query"
      ClearData
      Exit Sub
   Else
      mvSql = "Select * From Lmailquery where cphone LIKE " & Chr$(34) & ans & "*" & Chr$(34) & "order by bennames"
      Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
      QList
   End If
FindError0:
   Exit Sub

FindError1:
   MsgBox "ERROR:   " + Error$ + " has occurred", vbCritical, Me.Caption
   On Error GoTo 0

End Sub

Private Sub cmdRBU_Click()
  Dim sql As String, digitflag As Integer
  Dim buff As String, ans As String
  'On Error GoTo FindError1
  strpos = InStr(1, cmbRBU, ",", 1)
  ans = Trim(Left(cmbRBU, strpos - 1))
  If Len(ans) = 0 Then
      MsgBox "The Receiving Business Unit given is invalid, Retry", vbInformation, "Mail Query"
      ClearData
      Exit Sub
  Else
      mvSql = "Select * From outmailquery where cmbrec = '" & ans & "' order by bennames"
      Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
  End If
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
FindError0:
   Exit Sub

FindError1:
   MsgBox "ERROR:   " + Error$ + " has occurred", vbCritical, Me.Caption
   On Error GoTo 0

End Sub

Private Sub CmdRef_Click()
  Dim sql As String, digitflag As Integer
   Dim buff As String, ans As String
   On Error GoTo FindError1
   ans = UCase(Mailref)
   If Len(ans) = 0 Then
      MsgBox "The mail reference given is invalid, Retry", vbInformation, "Mail Query"
      ClearData
      Exit Sub
   Else
      mvSql = "Select * From Lmailquery where mailref = '" & ans & "'"
      Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
      QList
   End If
FindError0:
   Exit Sub

FindError1:
   MsgBox "ERROR:   " + Error$ + " has occurred", vbCritical, Me.Caption
   On Error GoTo 0

End Sub

Private Sub Command6_Click()

End Sub

Private Sub cmdSubj_Click()
   Dim sql As String
   Dim ans As String
   On Error GoTo FindError1
   strpos = InStr(1, cmbSubj, ",", 1)
   ans = Trim(Left(cmbSubj, strpos - 1))
   If Len(ans) = 0 Then
      MsgBox "The entered mail subject is invalid, Retry", vbInformation, "Mail Query"
      ClearData
      Exit Sub
   Else
      mvSql = "Select * From Lmailquery where cmbSubj = '" & ans & "'order by bennames"
      Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
      QList
   End If
FindError0:
   Exit Sub

FindError1:
   MsgBox "ERROR:   " + Error$ + " has occurred", vbCritical, Me.Caption
   On Error GoTo 0
End Sub

Private Sub Form_Load()
'On Error GoTo PError
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

    If Not (rstBUnitD.BOF And rstBUnitD.EOF) Then
       rstBUnitD.Requery
       Do While Not rstBUnitD.EOF
          cmbRBU.AddItem rstBUnitD!RCode + ", " + rstBUnitD!RDesc
          rstBUnitD.MoveNext
       Loop
       cmbRBU.ListIndex = 0
    End If
    If Not (rstSubj.BOF And rstSubj.EOF) Then
       rstSubj.Requery
       Do While Not rstSubj.EOF
          cmbSubj.AddItem rstSubj!RCode + ", " + rstSubj!RDesc
          rstSubj.MoveNext
       Loop
       cmbSubj.ListIndex = 0
    End If
    If Not (rstDept.BOF And rstDept.EOF) Then
       rstDept.Requery
       Do While Not rstDept.EOF
          cmbDept.AddItem rstDept!RCode + ", " + rstDept!RDesc
          rstDept.MoveNext
       Loop
       cmbDept.ListIndex = 0
    End If
    ''__________________________________
   ClearData
End Sub
Public Sub ClearData()
On Error GoTo PError
        Mailref = ""
        stDate = Date
        eDate = Date
        Ben = ""
        osurname = ""
        Phone = ""
PError:
End Sub
Private Sub QList()
On Error GoTo PError
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Query Review"
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
         Col0.Caption = "Mail Ref."
         Col1.Caption = "     "
         Col2.Caption = "Originated By"
         Col3.Caption = "Department"
         Col4.Caption = "Beneficiary"
         Col5.Caption = "Indate"
         Col6.Caption = "Outdate"
         Col7.Caption = "Subject"
         Col8.Caption = "Classification"
         Col9.Caption = "Received By"
         col10.Caption = "Mail Descrption"
   End If
   frmBrowse.Show
PError:
End Sub
