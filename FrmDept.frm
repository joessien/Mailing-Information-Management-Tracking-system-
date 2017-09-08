VERSION 5.00
Begin VB.Form FrmDept 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Tracking System"
   ClientHeight    =   3540
   ClientLeft      =   1755
   ClientTop       =   1395
   ClientWidth     =   6765
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6765
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2760
      Width           =   975
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Top             =   2760
      Width           =   855
   End
   Begin VB.ComboBox cmbGrp 
      Height          =   315
      Left            =   1680
      TabIndex        =   12
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox RDesc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox RCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1080
      Width           =   795
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdbrow 
      Caption         =   "Bro&wse"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdbott 
      Caption         =   "&Bottom"
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdprev 
      Caption         =   "&Previous"
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdtop 
      Caption         =   "&Top"
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Faculty:"
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
      Left            =   240
      TabIndex        =   18
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label fhdr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Department Definition"
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
      Width           =   5055
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
      Left            =   240
      TabIndex        =   4
      Top             =   1680
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
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "D"
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
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   375
   End
End
Attribute VB_Name = "FrmDept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EditFlag As Boolean, MaxErrors As Integer
Dim rstDefTab As Recordset, AddFlag As Boolean
Dim db1 As Database, wrktemp As Workspace
Dim AType As String, ValBuff As String
Dim Fldchk As Boolean
Dim FirstTime As Integer, FirstPass As Integer
Dim GenClass As New mProLog

Private Sub cmbGrp_lostfocus()
On Error GoTo PError:
   cmbEval
PError:
End Sub

Private Sub CmdAdd_Click()
On Error GoTo PError
 If cmdadd.Caption = "&Add" Then
    rstDefTab.AddNew
    AddFlag = True
    FldEnab
    ClearData
    RCode.SetFocus
    cmdsave.Enabled = True
    AddFlag = True
    CmdDisab
    cmdedit.Enabled = False
    cmdadd.Enabled = True
    cmdadd.Caption = "&Cancel"
Else
    If Not rstDefTab.EOF Then
       rstDefTab.MoveFirst
    End If
   AddFlag = False
   cmdadd.Caption = "&Add"
   CmdEnab
   EditFlag = False
   cmdedit.Enabled = True
   cmdsave.Enabled = False
   showdata
   FldDisab
PError:
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
On Error GoTo PError
 If cmdedit.Caption = "&Edit" Then
     If IsEmpty("rstDefTab") Then
         MsgBox ("Empty Table")
         Exit Sub
     Else
         If rstDefTab.EOF Then
            rstDefTab.MovePrevious
         End If
     End If
   EditFlag = True
   cmdsave.Enabled = True
   rstDefTab.Edit
   RDesc.Enabled = True
   cmbGrp.Enabled = True
   RDesc.SetFocus
   cmdedit.Caption = "&Cancel"
   CmdDisab
Else
   cmdedit.Caption = "&Edit"
   RCode.Enabled = False
   cmbGrp.Enabled = False
   CmdEnab
   EditFlag = False
   cmdsave.Enabled = False
   showdata
   FldDisab
End If
PError:
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
  Dim sql As String, digitflag As Integer, length As Integer
   Dim buff As String, AN As String

   On Error GoTo FindError1

   AN = InputBox$("Enter the Department Code to find:")
   length = Len(AN)
   AN = UCase(AN)
   If length = 0 Then
      ErrorMessages (NOINPUT)
      Exit Sub
   End If
   If length = 1 Then
      buff = "rcode LIKE " & Chr$(34) & AN & "*" & Chr$(34)
   Else
      buff = "rcode LIKE " & Chr$(34) & AN & "*" & Chr$(34)
   End If
   rstDefTab.FindFirst buff
   If rstDefTab.NOMATCH Then
      ErrorMessages (NOMATCH)
      Exit Sub
   Else
      showdata
   End If

FindError0:
   Exit Sub

FindError1:
   MsgBox "ERROR:   " + Error$ + " has occurred", vbCritical, Me.Caption
   On Error GoTo 0

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
    On Error GoTo PError
    
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
        rstDefTab("rCode") = "D" & UCase(RCode)
    End If
    rstDefTab("rDesc") = RDesc
    strpos = InStr(1, cmbGrp, ",", 1)
    rstDefTab("facc") = Trim(Left(cmbGrp, strpos - 1))
    rstDefTab.Update
       AddFlag = False
    cmdadd.Caption = "&Add"
    CmdEnab
    cmdedit.Enabled = True
    cmdsave.Enabled = False
    cmdedit.Caption = "&Edit"
    RCode.Enabled = True
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
On Error GoTo PError
    GenClass.fleLogin mvUserid, "Accessed Department Definition", Date, Time
    Call FormCentreMDI(Me)
    Set wrktemp = DBEngine.Workspaces(0)
    Set db1 = wrktemp.OpenDatabase(DBpath, True)
    Set rstDefTab = db1.OpenRecordset("DefDept", dbOpenDynaset)
    Dim flag As Integer, mtable As String
    Set rst = db1.OpenRecordset("Deffac", dbOpenDynaset)
    If Not (rst.BOF And rst.EOF) Then
       rst.Requery
       Do While Not rst.EOF
          cmbGrp.AddItem rst!RCode + ", " + rst!RDesc
          rst.MoveNext
       Loop
       cmbGrp.ListIndex = 0
    End If
    If Not rstDefTab.EOF Then
       rstDefTab.MoveFirst
    End If
    EditFlag = False
    AddFlag = False
    showdata
    FldDisab
    cmddel.Visible = False
    cmdsave.Enabled = False
    Me.Caption = MVCoyname
PError:
End Sub

Public Sub showdata()
     If rstDefTab.RecordCount >= 1 Then
        RCode = Right(rstDefTab("rCode"), 3)
        RDesc = rstDefTab("rDesc")
        
        mvBuff = "rCode = '" & rstDefTab("facc") & "'"
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
    cmdadd.Enabled = True
    cmddel.Enabled = True
    cmdbrow.Enabled = True
    cmdtop.Enabled = True
    cmdbott.Enabled = True
    cmdnext.Enabled = True
    cmdprev.Enabled = True
    cmdFind.Enabled = True
End Sub

Public Sub CmdDisab()
    cmdadd.Enabled = False
    cmddel.Enabled = False
    cmdbrow.Enabled = False
    cmdtop.Enabled = False
    cmdbott.Enabled = False
    cmdnext.Enabled = False
    cmdprev.Enabled = False
    cmdFind.Enabled = False
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




