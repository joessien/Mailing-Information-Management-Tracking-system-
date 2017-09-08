VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Frmexsub 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datapro"
   ClientHeight    =   4935
   ClientLeft      =   1890
   ClientTop       =   1260
   ClientWidth     =   6765
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6765
   Begin Threed.SSPanel SSPanel2 
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   3840
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
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmddel 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2400
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   5400
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   3735
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6735
      _Version        =   65536
      _ExtentX        =   11880
      _ExtentY        =   6588
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
      Begin VB.Data datGentab 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   3600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cmbSClass 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2280
         Width           =   3015
      End
      Begin MSMask.MaskEdBox OMarks 
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   3000
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   "_"
      End
      Begin VB.TextBox semister 
         Height          =   285
         Left            =   2040
         TabIndex        =   20
         Top             =   2640
         Width           =   2895
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   3375
         Left            =   5280
         TabIndex        =   6
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
            TabIndex        =   12
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
            TabIndex        =   10
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdnext 
            Caption         =   "&Next"
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdprev 
            Caption         =   "&Previous"
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   1200
            Width           =   975
         End
      End
      Begin VB.TextBox xDesc 
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox xCode 
         Height          =   285
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   0
         Top             =   1560
         Width           =   375
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   735
         Left            =   360
         TabIndex        =   18
         Top             =   480
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "Examination Subjects"
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
      Begin VB.Label Label4 
         BackColor       =   &H8000000C&
         Caption         =   "Subject Class:"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000C&
         Caption         =   "Marks Obtainable:"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000C&
         Caption         =   "Semister:"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   "Subject Description:"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label StaffNo 
         BackColor       =   &H00808080&
         Caption         =   "Subject Code:"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Frmexsub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EditFlag As Boolean, MaxErrors As Integer
Dim AType As String, ValBuff As String
Dim rstexamsub As Recordset, AddFlag As Boolean
Dim db6 As Database, wrktemp As Workspace
Dim FirstTime As Integer, FirstPass As Integer
Dim GenClass As New DProLog, strpos As Integer
Dim rstSubClass As Recordset, tempBuff As Variant

Private Sub CmdAdd_Click()
 If cmdadd.Caption = "&Add" Then
    rstexamsub.AddNew
    AddFlag = True
    ClearData
    xCode.SetFocus
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
   showdata
End If
End Sub

Private Sub CmdBott_Click()
   Dim Count As Long
    If rstexamsub.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstexamsub.MoveLast
    showdata
End Sub

Private Sub CmdBrow_Click()
    Dim GenView As New frmBrowse
    Set GenView.datGeneral.Recordset = rstexamsub
    GenView.Caption = Me.Caption & " - Records View"
    GenView.Show

End Sub

Private Sub CmdDel_Click()
   Dim i As Integer
   If (rstexamsub.BOF And rstexamsub.EOF) Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
    i = DeleteCheck()
    If i = vbYes Then
        rstexamsub.Delete
        If Not rstexamsub.BOF Then
          rstexamsub.MovePrevious
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
   xDesc.SetFocus
   cmdedit.Caption = "&Cancel"
   xCode.Enabled = False
   CmdDisab
Else
   cmdedit.Caption = "&Edit"
   xCode.Enabled = True
   CmdEnab
   EditFlag = False
   cmdSave.Enabled = False
   showdata
End If
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdNext_Click()
    Dim flag As Integer
    On Error GoTo NextError
    If rstexamsub.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstexamsub.EOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEPAST)
        rstexamsub.MoveLast
    Else
        rstexamsub.MoveNext
        If rstexamsub.EOF Then
            rstexamsub.MoveLast
        End If
    End If
NextClear:
    showdata
    Exit Sub
NextError:
    ErrorMessages (NOMOVEPAST)
    ''rstexamsub.Requery
    rstexamsub.MoveLast
    On Error GoTo 0
    Resume NextClear

End Sub

Private Sub CmdPrev_Click()
 Dim flag As Integer
    On Error GoTo PrevError
    If rstexamsub.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstexamsub.BOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEFRONT)
        rstexamsub.MoveFirst
    Else
        rstexamsub.MovePrevious
        If rstexamsub.BOF Then
            rstexamsub.MoveFirst
        End If
    End If
PrevClear:
    showdata
    Exit Sub
PrevError:
    ErrorMessages (NOMOVEFRONT)
    ''rstexamsub.Requery
    rstexamsub.MoveFirst
    On Error GoTo 0
    Resume PrevClear
End Sub

Private Sub CmdSave_Click()
  Dim i As Integer
    On Error GoTo perror
    
    If EditFlag = True Then
        rstexamsub.Edit
    End If
    If AddFlag = True Then
        rstexamsub.AddNew
        rstexamsub("code") = "X" & xCode
    End If
    rstexamsub("desc") = xDesc
    rstexamsub("semister") = semister
    rstexamsub("omarks") = OMarks
    strpos = InStr(1, cmbSClass, ",", 1)
    rstexamsub("Sclass") = Left(cmbSClass, strpos - 1)
    rstexamsub.Update
    rstexamsub.MoveLast
    AddFlag = False
    cmdadd.Caption = "&Add"
    CmdEnab
    cmdedit.Enabled = True
    cmdSave.Enabled = False
    cmdedit.Caption = "&Edit"
    xCode.Enabled = True
    showdata
    EditFlag = False
    AddFlag = False
PError0:
    Exit Sub
    
perror:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0
End Sub

Private Sub CmdTop_Click()
  
    If rstexamsub.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstexamsub.MoveFirst
    showdata
End Sub
Private Sub Form_Load()
    GenClass.fleLogin mvUserid, "Accessed Job Title", Date, Time
    Call FormCentreMDI(Me)
    Set wrktemp = DBEngine.Workspaces(0)
    Set db6 = wrktemp.OpenDatabase(DproDBpath, True)
    Set rstexamsub = db6.OpenRecordset("examsub", dbOpenDynaset)
    Set rstSubClass = db6.OpenRecordset("subclass", dbOpenDynaset)
    Dim flag As Integer, mtable As String
    If Not rstexamsub.EOF Then
       rstexamsub.MoveFirst
    End If
    ''__________________________________
    datGenTab.DatabaseName = DproDBpath
    datGenTab.RecordSource = "subclass"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       cmbSClass.AddItem datGenTab.Recordset("code") + ", " + datGenTab.Recordset("Desc")
       datGenTab.Recordset.MoveNext
    Loop
    'cmbSClass.ListIndex = 0
    showdata
    EditFlag = False
    AddFlag = False
    cmdSave.Enabled = False
    cmddel.Visible = False
End Sub

Public Sub showdata()
     If Not (rstexamsub.BOF And rstexamsub.EOF) Then
        xCode = Right(rstexamsub("code"), 3)
        xDesc = rstexamsub("desc")
        semister = rstexamsub("semister")
        OMarks = rstexamsub("omarks")
        tempBuff = "Code = '" & rstexamsub("Sclass") & "'"
        rstSubClass.FindFirst tempBuff
        cmbSClass = rstSubClass!Code & "," & rstSubClass!Desc
    End If
End Sub

Public Sub ClearData()
    xCode = ""
    xDesc = ""
    semister = ""
    OMarks = 0
End Sub
Public Sub CmdEnab()
    cmdadd.Enabled = True
    cmddel.Enabled = True
    cmdbrow.Enabled = True
    cmdtop.Enabled = True
    cmdbott.Enabled = True
    cmdnext.Enabled = True
    cmdprev.Enabled = True
End Sub

Public Sub CmdDisab()
    cmdadd.Enabled = False
    cmddel.Enabled = False
    cmdbrow.Enabled = False
    cmdtop.Enabled = False
    cmdbott.Enabled = False
    cmdnext.Enabled = False
    cmdprev.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    db6.Close
End Sub
