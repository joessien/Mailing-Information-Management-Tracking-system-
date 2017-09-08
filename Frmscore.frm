VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Frmscore 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datapro"
   ClientHeight    =   5775
   ClientLeft      =   1890
   ClientTop       =   1260
   ClientWidth     =   6765
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6765
   Begin Threed.SSPanel SSPanel2 
      Height          =   1095
      Left            =   240
      TabIndex        =   17
      Top             =   4680
      Width           =   6735
      _Version        =   65536
      _ExtentX        =   11880
      _ExtentY        =   1931
      _StockProps     =   15
      Caption         =   " "
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdadd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmddel 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   3480
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   5400
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   4575
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   6735
      _Version        =   65536
      _ExtentX        =   11880
      _ExtentY        =   8070
      _StockProps     =   15
      Caption         =   " "
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      BevelInner      =   1
      Begin VB.ComboBox cmbProg 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   1800
         Width           =   3555
      End
      Begin VB.ComboBox cmbAppid 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   1440
         Width           =   3555
      End
      Begin MSMask.MaskEdBox Edate 
         Height          =   315
         Left            =   2040
         TabIndex        =   6
         Top             =   3900
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox SDate 
         Height          =   315
         Left            =   2040
         TabIndex        =   5
         Top             =   3480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin VB.Data datGentab 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5220
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3660
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cmbSubcode 
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Top             =   2700
         Width           =   3135
      End
      Begin MSMask.MaskEdBox SMarks 
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Top             =   2280
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.TextBox Remarks 
         Height          =   285
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   3120
         Width           =   3135
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   3375
         Left            =   5280
         TabIndex        =   19
         Top             =   120
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   5953
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CommandButton cmdbrow 
            Caption         =   "Bro&wse"
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdtop 
            Caption         =   "&Top"
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdbott 
            Caption         =   "&Bottom"
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdnext 
            Caption         =   "&Next"
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdprev 
            Caption         =   "&Previous"
            Height          =   375
            Left            =   240
            TabIndex        =   13
            Top             =   1200
            Width           =   975
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   735
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "Examination Result Input"
         BackColor       =   8421376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   4
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         Caption         =   "Program:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Personal Id:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   360
         TabIndex        =   26
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "End Date:"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   3900
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Start Date:"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   3480
         Width           =   1515
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000C&
         Caption         =   "Subject Code:"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   2700
         Width           =   1515
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000C&
         Caption         =   "Score:"
         Height          =   315
         Left            =   360
         TabIndex        =   22
         Top             =   2280
         Width           =   1515
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000C&
         Caption         =   "Remarks:"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   3120
         Width           =   1515
      End
   End
End
Attribute VB_Name = "Frmscore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EditFlag As Boolean, MaxErrors As Integer
Dim AType As String, ValBuff As String
Dim rstexamsub As Recordset, AddFlag As Boolean
Dim db6 As Database, wrktemp As Workspace
Dim rstExamRecs As Recordset, mvAppid As String
Dim FirstTime As Integer, FirstPass As Integer
Dim GenClass As New DProLog, strpos As Integer
Dim rstCrsTab As Recordset, rstTrain As Recordset
Dim rstGenRecs As Recordset, tempBuff As Variant

Private Sub cmbAppid_Click()
    Dspinfo
End Sub

Private Sub cmbAppid_GotFocus()
    cmbAppid.SelStart = 0
    cmbAppid.SelLength = Len(CStr(cmbAppid))

End Sub

Private Sub cmbAppid_LostFocus()
    Dspinfo
End Sub
Public Sub Dspinfo()
    On Error GoTo perror1
    If Len(cmbAppid) = 0 Then
       Exit Sub
    End If
    strpos = InStr(1, cmbAppid, ",", 1)
    If strpos <> 0 Then
        mvClid = Left(cmbAppid, strpos - 1)
    End If
    If strpos <= 0 Then
        tempBuff = "appid = '" & cmbAppid & "'"
        rstGenRecs.Requery
        rstGenRecs.FindFirst tempBuff
        If rstGenRecs.NOMATCH Then
           Beep
           MsgBox "Staff Number not found.", vbInformation, "Exams Input"
           Exit Sub
        Else
            cmbAppid = rstGenRecs!Appid & ", " & rstGenRecs!fullname
            mvAppid = rstGenRecs!Appid
            rstExamRecs.FindFirst tempBuff
            If rstExamRecs.NOMATCH Then
                ClearData
            Else
                showdata
            End If
        End If
    Else
        strpos = InStr(1, cmbAppid, ",", 1)
        mvAppid = Left(cmbAppid, strpos - 1)
        tempBuff = "appid = '" & mvAppid & "'"
        rstGenRecs.Requery
        rstGenRecs.FindFirst tempBuff
        If rstGenRecs.NOMATCH Then
           Beep
           MsgBox "Staff Number not found.", vbInformation, "Input"
           Exit Sub
        Else
            cmbAppid = rstGenRecs!Appid & ", " & rstGenRecs!fullname
            mvAppid = rstGenRecs!Appid
            rstExamRecs.FindFirst tempBuff
            If rstExamRecs.NOMATCH Then
                ClearData
            Else
                showdata
            End If
        End If
    End If
perror1:
End Sub


Private Sub CmdAdd_Click()
 If cmdadd.Caption = "&Add" Then
    rstexamsub.AddNew
    AddFlag = True
    ClearData
    cmbAppid.SetFocus
    cmdSave.Enabled = True
    AddFlag = True
    CmdDisab
    cmdedit.Enabled = False
    cmdadd.Enabled = True
    cmdadd.Caption = "&Cancel"
Else
    If Not rstexamsub.EOF Then
       rstexamsub.MoveFirst
    End If
   AddFlag = False
   cmdadd.Caption = "&Add"
   CmdEnab
   cmdedit.Enabled = True
   cmdSave.Enabled = False
   AddFlag = False
   ClearData
End If
End Sub

Private Sub CmdBrow_Click()
    Dim GenView As New frmBrowse
    rstExamRecs.Requery
    Set GenView.datGeneral.Recordset = rstExamRecs
    GenView.Caption = Me.Caption & " - Records View"
    GenView.Show

End Sub

Private Sub CmdDel_Click()
   Dim i As Integer
   If (rstExamRecs.BOF And rstExamRecs.EOF) Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
    i = DeleteCheck()
    If i = vbYes Then
        rstExamRecs.Delete
        If Not rstExamRecs.BOF Then
          rstExamRecs.MovePrevious
          showdata
        Else
          ClearData
        End If
    End If
End Sub

Private Sub CmdEdit_Click()
 If cmdedit.Caption = "&Edit" Then
   If rstexamsub.EOF Then
       MsgBox ("Empty Table")
       Exit Sub
   End If
   EditFlag = True
   cmdSave.Enabled = True
   rstexamsub.Edit
   SMarks.SetFocus
   cmdedit.Caption = "&Cancel"
   cmbAppid.Enabled = False
   CmdDisab
Else
   cmdedit.Caption = "&Edit"
   cmbAppid.Enabled = True
   CmdEnab
   EditFlag = False
   cmdSave.Enabled = False
   showdata
End If
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
  Dim i As Integer
    On Error GoTo PError
    If EditFlag = True Then
        rstExamRecs.Edit
    End If
    If AddFlag = True Then
        rstExamRecs.AddNew
        rstExamRecs("appid") = mvAppid
    End If
    rstExamRecs("score") = SMarks
    rstExamRecs("cremarks") = Remarks
    rstExamRecs("edate") = CDate(Edate)
    rstExamRecs("sdate") = CDate(SDate)
    strpos = InStr(1, cmbSubcode, ",", 1)
    rstExamRecs("examcode") = Left(cmbSubcode, strpos - 1)
    strpos = InStr(1, cmbProg, ",", 1)
    rstExamRecs("prog") = Left(cmbProg, strpos - 1)
    rstExamRecs.Update
    rstExamRecs.MoveLast
    AddFlag = False
    cmdadd.Caption = "&Add"
    CmdEnab
    cmdedit.Enabled = True
    cmdSave.Enabled = False
    cmdedit.Caption = "&Edit"
    cmbAppid.Enabled = True
    showdata
    EditFlag = False
    AddFlag = False
PError0:
    Exit Sub
    
PError:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0
End Sub


Private Sub Form_Load()
    GenClass.fleLogin mvUserid, "Accessed Job Title", Date, Time
    Call FormCentreMDI(Me)
    Set wrktemp = DBEngine.Workspaces(0)
    Set db6 = wrktemp.OpenDatabase(DproDBpath, True)
    Set rstexamsub = db6.OpenRecordset("examsub", dbOpenDynaset)
    Set rstExamRecs = db6.OpenRecordset("ExamRecs", dbOpenDynaset)
    Set rstGenRecs = db6.OpenRecordset("GenRecs", dbOpenDynaset)
    Set rstCrsTab = db6.OpenRecordset("crstab", dbOpenDynaset)
    Set rstTrain = db6.OpenRecordset("training", dbOpenDynaset)
    Dim flag As Integer, mtable As String
    If Not rstexamsub.EOF Then
       rstexamsub.MoveFirst
    End If
    ''__________________________________
    datGenTab.DatabaseName = DproDBpath
    datGenTab.RecordSource = "examsub"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       cmbSubcode.AddItem datGenTab.Recordset("code") + ", " + datGenTab.Recordset("Desc")
       datGenTab.Recordset.MoveNext
    Loop
    'cmbSClass.ListIndex = 0
    datGenTab.DatabaseName = DproDBpath
    datGenTab.RecordSource = "genrecs"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       cmbAppid.AddItem datGenTab.Recordset("appid") + ", " + datGenTab.Recordset("fullname")
       datGenTab.Recordset.MoveNext
    Loop
    'cmbSClass.ListIndex = 0
    ''__________________________________
    datGenTab.DatabaseName = DproDBpath
    datGenTab.RecordSource = "crstab"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       cmbProg.AddItem datGenTab.Recordset("code") + ", " + datGenTab.Recordset("Desc")
       datGenTab.Recordset.MoveNext
    Loop
    showdata
    EditFlag = False
    AddFlag = False
    Edate = Date
    SDate = Date
    cmdSave.Enabled = False
End Sub

Public Sub showdata()
    Dim strbuff As Variant
    strbuff = "appid = '" & mvAppid & "'" & " and  Cstatus = 'X'"
    rstTrain.FindFirst strbuff
    If rstTrain.NOMATCH Then
        cmbProg = ""
    Else
        strbuff = "Code = '" & rstTrain!CrsCode & "'"
        rstCrsTab.FindFirst strbuff
        cmbProg = rstCrsTab!code & "," & rstCrsTab!Desc
     End If
     If Not (rstExamRecs.BOF And rstExamRecs.EOF) Then
        If mvAppid = "" Then
           Exit Sub
        End If
        tempBuff = "appid = '" & mvAppid & "'"
        rstGenRecs.Requery
        rstGenRecs.FindFirst tempBuff
        If rstGenRecs.NOMATCH Then
           Beep
           MsgBox "Staff Number not found.", vbInformation, "Input"
           Exit Sub
        Else
            cmbAppid = rstGenRecs!Appid & ", " & rstGenRecs!fullname
        End If
        SMarks = rstExamRecs("score")
        Remarks = rstExamRecs("cremarks")
        Edate = rstExamRecs("edate")
        SDate = rstExamRecs("sdate")
        tempBuff = "Code = '" & rstExamRecs("examcode") & "'"
        rstexamsub.FindFirst tempBuff
        cmbSubcode = rstexamsub!code & "," & rstexamsub!Desc
        tempBuff = "appid = '" & rstGenRecs!Appid & "'" & " and  Cstatus = 'X'"
        rstTrain.FindFirst tempBuff
        If rstTrain.NOMATCH Then
            cmbProg = ""
        Else
            tempBuff = "Code = '" & rstTrain!CrsCode & "'"
            rstCrsTab.FindFirst tempBuff
            cmbProg = rstCrsTab!code & "," & rstCrsTab!Desc
         End If
    End If
End Sub

Public Sub ClearData()
    SMarks = ""
End Sub
Public Sub CmdEnab()
    cmdadd.Enabled = True
    cmddel.Enabled = True
    cmdBrow.Enabled = True
    cmdtop.Enabled = True
    cmdbott.Enabled = True
    cmdnext.Enabled = True
    cmdprev.Enabled = True
End Sub

Public Sub CmdDisab()
    cmdadd.Enabled = False
    cmddel.Enabled = False
    cmdBrow.Enabled = False
    cmdtop.Enabled = False
    cmdbott.Enabled = False
    cmdnext.Enabled = False
    cmdprev.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    db6.Close
End Sub
Private Sub CmdNext_Click()
    Dim flag As Integer
    On Error GoTo NextError
    If rstExamRecs.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstExamRecs.EOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEPAST)
        rstExamRecs.MoveLast
    Else
        rstExamRecs.MoveNext
        If rstExamRecs.EOF Then
            rstExamRecs.MoveLast
        End If
    End If
NextClear:
    showdata
    Exit Sub
NextError:
    ErrorMessages (NOMOVEPAST)
    rstExamRecs.MoveLast
    On Error GoTo 0
    Resume NextClear

End Sub

Private Sub CmdPrev_Click()
 Dim flag As Integer
    On Error GoTo PrevError
    If rstExamRecs.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstExamRecs.BOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEFRONT)
        rstExamRecs.MoveFirst
    Else
        rstExamRecs.MovePrevious
        If rstExamRecs.BOF Then
            rstExamRecs.MoveFirst
        End If
    End If
PrevClear:
    showdata
    Exit Sub
PrevError:
    ErrorMessages (NOMOVEFRONT)
    ''rstexamrecs.Requery
    rstExamRecs.MoveFirst
    On Error GoTo 0
    Resume PrevClear
End Sub
Private Sub CmdTop_Click()
    If rstExamRecs.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstExamRecs.MoveFirst
    showdata
End Sub
Private Sub CmdBott_Click()
   Dim Count As Long
    If rstExamRecs.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstExamRecs.MoveLast
    showdata
End Sub

