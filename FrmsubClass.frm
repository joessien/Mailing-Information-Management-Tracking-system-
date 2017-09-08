VERSION 5.00
Begin VB.Form FrmSubclass 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Tracking System"
   ClientHeight    =   3210
   ClientLeft      =   1755
   ClientTop       =   1395
   ClientWidth     =   6720
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6720
   Begin VB.ComboBox cmbGrp 
      Height          =   315
      Left            =   1920
      TabIndex        =   16
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H0080FF80&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox RDesc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox RCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   0
      Top             =   960
      Width           =   795
   End
   Begin VB.CommandButton cmdbrow 
      Caption         =   "Bro&wse"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdbott 
      Caption         =   "&Bottom"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdprev 
      Caption         =   "&Previous"
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdtop 
      Caption         =   "&Top"
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Group:"
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
      Left            =   360
      TabIndex        =   17
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label fhdr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Subclassification Definition"
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
      TabIndex        =   5
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label DeptName 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label StaffNo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C"
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
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   375
   End
End
Attribute VB_Name = "FrmSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EditFlag As Boolean, MaxErrors As Integer
Dim rstDefTab As Recordset, AddFlag As Boolean
Dim db1 As Database, wrktemp As Workspace
Dim AType As String, ValBuff As String
Dim FirstTime As Integer, FirstPass As Integer
Dim Fldchk As Boolean
Dim GenClass As New mProLog

Private Sub cmbGrp_Change()
On Error GoTo PError:
   cmbEval
PError:
End Sub

Private Sub CmdAdd_Click()
 If cmdAdd.Caption = "&Add" Then
    rstDefTab.AddNew
    AddFlag = True
    FldEnab
    ClearData
    RCode.SetFocus
    cmdSave.Enabled = True
    AddFlag = True
    CmdDisab
    cmdEdit.Enabled = False
    cmdAdd.Enabled = True
    cmdAdd.Caption = "&Cancel"
Else
    If Not rstDefTab.EOF Then
       rstDefTab.MoveFirst
    End If
   AddFlag = False
   cmdAdd.Caption = "&Add"
   CmdEnab
   EditFlag = False
   cmdEdit.Enabled = True
   cmdSave.Enabled = False
   showdata
   FldDisab
End If
End Sub

Private Sub CmdBott_Click()
    Dim Count As Long
    If rstDefTab.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstDefTab.MoveLast
    showdata
End Sub

Private Sub CmdBrow_Click()
    Dim GenView As New frmBrowse
    Set GenView.datGeneral.Recordset = rstDefTab
    GenView.Caption = Me.Caption
    GenView.Show

    End Sub

Private Sub CmdDel_Click()
   Dim i As Integer
   If rstDefTab.RecordCount < 1 Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
    i = DeleteCheck()
    If i = vbYes Then
        rstDefTab.Delete
        If Not rstDefTab.BOF Then
          rstDefTab.MovePrevious
          showdata
        Else
          ClearData
        End If
    End If

End Sub

Private Sub CmdEdit_Click()
 If cmdEdit.Caption = "&Edit" Then
     If IsEmpty("rstDefTab") Then
         MsgBox ("Empty Table")
         Exit Sub
     Else
         If rstDefTab.EOF Then
            rstDefTab.MovePrevious
         End If
     End If
   EditFlag = True
   cmdSave.Enabled = True
   rstDefTab.Edit
   RDesc.Enabled = True
   cmbGrp.Enabled = True
   RDesc.SetFocus
   cmdEdit.Caption = "&Cancel"
   RCode.Enabled = False
   CmdDisab
Else
   cmdEdit.Caption = "&Edit"
   RCode.Enabled = False
   cmbGrp.Enabled = False
   CmdEnab
   EditFlag = False
   cmdSave.Enabled = False
   showdata
   FldDisab
End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdNext_Click()
        Dim flag As Integer
    
    On Error GoTo NextError
    If rstDefTab.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstDefTab.EOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEPAST)
        rstDefTab.MoveLast
    Else
        rstDefTab.MoveNext
        If rstDefTab.EOF Then
            rstDefTab.MoveLast
        End If
    End If
NextClear:
    showdata
    Exit Sub
NextError:
    ErrorMessages (NOMOVEPAST)
    ''rstDefTab.Requery
    rstDefTab.MoveLast
    On Error GoTo 0
    Resume NextClear

End Sub

Private Sub CmdPrev_Click()
  Dim flag As Integer
    On Error GoTo PrevError
    If rstDefTab.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstDefTab.BOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEFRONT)
        rstDefTab.MoveFirst
    Else
        rstDefTab.MovePrevious
        If rstDefTab.BOF Then
            rstDefTab.MoveFirst
        End If
    End If
PrevClear:
    showdata
    Exit Sub
PrevError:
    ErrorMessages (NOMOVEFRONT)
    rstDefTab.Requery
    rstDefTab.MoveFirst
    On Error GoTo 0
    Resume PrevClear
End Sub

Private Sub CmdSave_Click()
  Dim i As Integer
    'On Error GoTo PError
    If Len(Trim(RCode)) = 0 Then
        MsgBox "Invalid Code, Re-Enter.", vbInformation, "Mail Records"
        RCode.SetFocus
        Exit Sub
    End If
    If Len(Trim(RDesc)) = 0 Then
        MsgBox "Invalid Description, Re-Enter.", vbInformation, "Mail Records"
        RDesc.SetFocus
        Exit Sub
    End If
    Fldchk = True
    cmbEval
    If Fldchk = False Then
        Exit Sub
    End If
    If EditFlag = True Then
        rstDefTab.Edit
    End If
    If AddFlag = True Then
        rstDefTab.AddNew
        rstDefTab("rCode") = "C" & UCase(RCode)
    End If
    rstDefTab("rDesc") = RDesc
    strpos = InStr(1, cmbGrp, ",", 1)
    rstDefTab("subjg") = Trim(Left(cmbGrp, strpos - 1))
    rstDefTab.Update
       AddFlag = False
     cmdAdd.Caption = "&Add"
     CmdEnab
     cmdEdit.Enabled = True
     cmdSave.Enabled = False
    cmdEdit.Caption = "&Edit"
    RCode.Enabled = False
    cmbGrp.Enabled = False
    If AddFlag = True Then rstDefTab.MoveLast
    showdata
    FldDisab
    EditFlag = False
    AddFlag = False
PError0:
    Exit Sub
PError:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0
End Sub

Private Sub CmdTop_Click()
    If rstDefTab.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstDefTab.MoveFirst
    showdata
End Sub

Private Sub Form_Load()
    Dim flag As Integer, mtable As String
    GenClass.fleLogin mvUserid, "Accessed Subject Classification Definition", Date, Time
    Call FormCentreMDI(Me)
    Set wrktemp = DBEngine.Workspaces(0)
    Set db1 = wrktemp.OpenDatabase(DBpath, True)
    Set rstDefTab = db1.OpenRecordset("Defsubclass", dbOpenDynaset)
    Set rst = db1.OpenRecordset("Defsubj", dbOpenDynaset)
    If Not (rst.BOF And rst.EOF) Then
       rst.Requery
       Do While Not rst.EOF
          cmbGrp.AddItem rst!RCode + ", " + rst!RDesc
          rst.MoveNext
       Loop
       cmbGrp.ListIndex = 0
    End If
    Me.Caption = MVCoyname
    If Not rstDefTab.EOF Then
       rstDefTab.MoveFirst
    End If
    EditFlag = False
    AddFlag = False
    showdata
    FldDisab
    cmdDel.Visible = False
    cmdSave.Enabled = False
    Me.Caption = MVCoyname
End Sub

Public Sub showdata()
     If rstDefTab.RecordCount >= 1 Then
        RCode = Right(rstDefTab("rCode"), 3)
        RDesc = rstDefTab("rDesc")
        
        mvBuff = "rCode = '" & rstDefTab("subjg") & "'"
        rst.FindFirst mvBuff
        If rst.NOMATCH Then
           cmbGrp = " "
        Else
           cmbGrp = rst!RCode & "," & rst!RDesc
        End If
    End If
End Sub

Public Sub ClearData()
    RCode = ""
    RDesc = ""
End Sub
Public Sub CmdEnab()
    cmdAdd.Enabled = True
    cmdDel.Enabled = True
    cmdbrow.Enabled = True
    cmdtop.Enabled = True
    cmdbott.Enabled = True
    cmdnext.Enabled = True
    cmdprev.Enabled = True
End Sub

Public Sub CmdDisab()
    cmdAdd.Enabled = False
    cmdDel.Enabled = False
    cmdbrow.Enabled = False
    cmdtop.Enabled = False
    cmdbott.Enabled = False
    cmdnext.Enabled = False
    cmdprev.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    db1.Close
End Sub
Private Sub FldEnab()
    RCode.Enabled = True
    RDesc.Enabled = True
    cmbGrp.Enabled = True
End Sub

Private Sub FldDisab()
    RCode.Enabled = False
    RDesc.Enabled = False
    cmbGrp.Enabled = False
End Sub

Public Sub cmbEval()
    strpos = InStr(1, cmbGrp, ",", 1)
    If strpos = 0 Then
       mvBuff = cmbGrp
    Else
      mvBuff = Left(cmbGrp, strpos - 1)
    End If
    If Len(Trim(cmbGrp)) = 0 Then
       MsgBox "Invalid reference entered.", vbInformation, Me.Caption
       Fldchk = False
       RDesc.SetFocus
       Exit Sub
    End If
    mvSql = "rcode = '" & Trim(mvBuff) & "'"
    rst.FindFirst mvSql
    If rst.NOMATCH Then
        MsgBox "Invalid reference entered.", vbInformation, Me.Caption
        cmbGrp.SetFocus
        Fldchk = False
        Exit Sub
    Else
        cmbGrp = rst!RCode + ", " + rst!RDesc
    End If
End Sub



