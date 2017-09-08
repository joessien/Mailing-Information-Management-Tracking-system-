VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmSGrading 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DataPro"
   ClientHeight    =   6210
   ClientLeft      =   1755
   ClientTop       =   1395
   ClientWidth     =   8730
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   8730
   Begin VB.CheckBox chkFGrd 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2280
      TabIndex        =   27
      Top             =   2880
      Width           =   255
   End
   Begin VB.Data datGeneral 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data datgenTab 
      Caption         =   "Gentab"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox cmbSlev 
      Height          =   315
      Left            =   2280
      TabIndex        =   25
      Text            =   "cmbSlev"
      Top             =   2400
      Width           =   4095
   End
   Begin Threed.SSPanel Cmsg 
      Height          =   375
      Left            =   720
      TabIndex        =   24
      Top             =   4920
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   661
      _StockProps     =   15
      BackColor       =   65280
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   11.99
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1455
      Left            =   720
      TabIndex        =   18
      Top             =   3360
      Width           =   5700
      _Version        =   65536
      _ExtentX        =   10054
      _ExtentY        =   2566
      _StockProps     =   15
      Caption         =   "Score Range"
      ForeColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
      Alignment       =   6
      Begin Threed.SSCommand cmdChk 
         Height          =   855
         Left            =   4440
         TabIndex        =   23
         Top             =   480
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "Check"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
      End
      Begin MSMask.MaskEdBox Lscore 
         Height          =   375
         Left            =   2760
         TabIndex        =   21
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0.0000;(##0.0000)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Uscore 
         Height          =   375
         Left            =   2760
         TabIndex        =   22
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0.0000;(##0.0000)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lower Boundary"
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
         Left            =   840
         TabIndex        =   20
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Upper Boundary"
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
         Left            =   840
         TabIndex        =   19
         Top             =   600
         Width           =   1815
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   5400
      Width           =   8655
      _Version        =   65536
      _ExtentX        =   15266
      _ExtentY        =   1296
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H00C0C0FF&
         Caption         =   "E&xit"
         Height          =   375
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H0000FF00&
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
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmddel 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4080
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
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
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   1
      Top             =   1800
      Width           =   4095
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
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   0
      Top             =   1320
      Width           =   435
   End
   Begin Threed.SSPanel SSPanel5 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   8655
      _Version        =   65536
      _ExtentX        =   15266
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "Scores Grading Definition "
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   17.99
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   4335
      Left            =   6960
      TabIndex        =   6
      Top             =   960
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   7646
      _StockProps     =   15
      Caption         =   " "
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
         Left            =   360
         TabIndex        =   7
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdbott 
         Caption         =   "&Bottom"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdtop 
         Caption         =   "&Top"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fail Grade:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   28
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "School Level:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   26
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label DeptName 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label StaffNo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Code:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   375
   End
End
Attribute VB_Name = "FrmSGrading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim EditFlag As Boolean, MaxErrors As Integer
Dim rstDefTab As Recordset, AddFlag As Boolean
Dim db1 As Database, wrktemp As Workspace
Dim AType As String, ValBuff, MSlev As String
Dim firstTime As Integer, FirstPass As Integer
Dim GenClass As New DProLog, EFlag As Boolean
Dim RSTGen, rstSLev As Recordset

Private Sub cmbSlev_Click()
    strpos = InStr(cmbSlev, ",")
    MSlev = UCase(Trim(Left(cmbSlev, strpos - 1)))
End Sub

Private Sub CmdAdd_Click()
 If cmdadd.Caption = "&Add" Then
    CmdDisab
    cmdChk.Enabled = True
    AddFlag = True
    FldEnab
    ClearData
    RCode.SetFocus
    cmdsave.Enabled = False
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
   cmdChk.Enabled = False
   showdata
   FldDisab
End If
End Sub

Private Sub cmdASave_Click()
  Dim i As Integer
    'On Error GoTo PError
    If EditFlag = True Then
        GTBuff = "code = '" & RCode & "'"
        rstDefTab.FindFirst GTBuff
        rstDefTab.Edit
    End If
    If AddFlag = True Then
        rstDefTab.AddNew
        rstDefTab("Code") = UCase(RCode)
    End If
    rstDefTab("Desc") = UCase(RDesc)
    rstDefTab("Lscore") = UCase(Lscore)
    rstDefTab("UScore") = UCase(Uscore)
    rstDefTab("FGrade") = chkFGrd
    strpos = InStr(cmbSlev, ",")
    rstDefTab("slev") = UCase(Trim(Left(cmbSlev, strpos - 1)))
    
    rstDefTab.Update
    AddFlag = False
    cmdadd.Caption = "&Add"
    CmdEnab
    cmdedit.Enabled = True
    cmdsave.Enabled = False
    cmdedit.Caption = "&Edit"
    RCode.Enabled = True
    If rstDefTab.EOF Then
      rstDefTab.MoveLast
    End If
    showdata
    Cmsg = ""
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
On Error GoTo PError
     Screen.MousePointer = vbHourglass
     mvSql = "Select * From defscgrd Order By slev , code;"
     Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
     If (rst.BOF And rst.EOF) Then
         MsgBox "No Grades Defined", vbExclamation, "Grade Definition"
         Set frmBrowse.datGeneral.Recordset = rst
         grdGeneral.Caption = Me.Caption
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 1800
         frmBrowse.grdGeneral.Columns(2).Width = 1500
         frmBrowse.grdGeneral.Columns(3).Width = 1500
         frmBrowse.grdGeneral.Columns(4).Width = 2000
         frmBrowse.grdGeneral.Columns(5).Width = 2000
         Col0.Caption = "Grade Key"
         Col1.Caption = "Description"
         Col2.Caption = "Lower Score"
         Col3.Caption = "Upper Score"
         Col4.Caption = "School Level"
         Col5.Caption = "Pass/Fail "
   End If
PError:
    Screen.MousePointer = vbDefault
    frmBrowse.Show
    End Sub

Private Sub cmdChk_Click()
On Error GoTo PError
   strpos = InStr(cmbSlev, ",")
   MSlev = UCase(Trim(Left(cmbSlev, strpos - 1)))
   EFlag = False
   cmdsave.Enabled = False
   rstDefTab.Requery
   If RCode = "" Or RDesc = "" Then
      MsgBox "Invalid code or desscription: Cannot be blank", vbCritical, "School Terms"
      Cmsg.BackColor = &HC0C0FF
      Cmsg.Caption = "Invalid Code or Description , Re-Enter"
      EFlag = True
       Exit Sub
   End If
   If (rstDefTab.EOF And rstDefTab.EOF) Then
      EFlag = False
      cmdsave.Enabled = True
      Cmsg.BackColor = &H80FF80
      Cmsg.Caption = "Range is accepted, Proceed to Save"
   Else
      If Val(Lscore) >= Val(Uscore) Then
         MsgBox "Invalid Range: Lower Score cannot be more than or equal to Upper Score", vbCritical, "Score Grading"
         Cmsg.BackColor = &HC0C0FF
         Cmsg.Caption = "Range is Invalid, Re-Enter"
         EFlag = True
         Exit Sub
      End If
      If Val(Lscore) < 0 Or Val(Lscore) > 99 Then
         MsgBox "Invalid Lower boundary: Lower Score cannot be less than zero or greater than 99", vbCritical, "Score Grading"
         Cmsg.BackColor = &HC0C0FF
         Cmsg.Caption = "Range is Invalid, Re-Enter"
         EFlag = True
         Exit Sub
      End If
      If Val(Uscore) < 1 Or Val(Uscore) > 100 Then
         MsgBox "Invalid Upper boundary: Upper Score cannot be less than one or greater than 100", vbCritical, "Score Grading"
         Cmsg.BackColor = &HC0C0FF
         Cmsg.Caption = "Range is Invalid, Re-Enter"
         EFlag = True
         Exit Sub
      End If
      rstDefTab.MoveFirst
      Do While Not rstDefTab.EOF
         'If EditFlag = False And rstDefTab!code <> RCode Then
            If (Val(Lscore) >= rstDefTab!Lscore And Val(Lscore) <= rstDefTab!Uscore And rstDefTab!sleV = MSlev) Then
               MsgBox "Invalid Lower boundary: Score is already nested within Grade " & rstDefTab!code, vbCritical, "Score Grading"
               Cmsg.BackColor = &HC0C0FF
               Cmsg.Caption = "Range is Invalid, Re-Enter"
               EFlag = True
               Exit Sub
            End If
         'End If
         If (Val(Uscore) >= rstDefTab!Lscore And Val(Uscore) <= rstDefTab!Uscore And rstDefTab!sleV = MSlev) Then
            MsgBox "Invalid Upper boundary: Score is already nested within Grade " & rstDefTab!code, vbCritical, "Score Grading"
            Cmsg.BackColor = &HC0C0FF
            Cmsg.Caption = "Range is Invalid, Re-Enter"
            EFlag = True
            Exit Sub
         End If
         If (UCase(RCode) = rstDefTab!code And AddFlag = True And rstDefTab!sleV = MSlev) Then
            MsgBox "Invalid Grade: Grade is already defined", vbCritical, "Score Grading"
            Cmsg.BackColor = &HC0C0FF
            Cmsg.Caption = "Invalid Grade, Re-Enter"
            EFlag = True
            Exit Sub
         End If
         If (UCase(RDesc) = rstDefTab!Desc And AddFlag = True And rstDefTab!sleV = MSlev) Then
            MsgBox "Invalid Description: Description already used for Grade " & rstDefTab!code, vbCritical, "Score Grading"
            Cmsg.BackColor = &HC0C0FF
            Cmsg.Caption = "Invalid Description, Re-Enter"
            EFlag = True
            Exit Sub
         End If
         rstDefTab.MoveNext
      Loop
      If EFlag = True Then
         Cmsg.BackColor = &HC0C0FF
         Cmsg.Caption = "Range is Invalid, Re-Enter"
      Else
         Cmsg.BackColor = &H80FF80
         Cmsg.Caption = "Range is accepted, Proceed to Save"
         cmdsave.Enabled = True
      End If
   End If
    Exit Sub
PError:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
   
End Sub

Private Sub CmdDel_Click()
   Dim i As Integer
   If rstDefTab.BOF And rstDefTab.EOF Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   Else
      i = DeleteCheck()
      If i = vbYes Then
         rstDefTab.Delete
         If Not rstDefTab.BOF Then
             rstDefTab.Requery
             showdata
         Else
             ClearData
         End If
      End If
      Exit Sub
   End If
PError:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
End Sub

Private Sub CmdEdit_Click()
 If cmdedit.Caption = "&Edit" Then
     If (rstDefTab.EOF And rstDefTab.EOF) Then
         MsgBox ("Empty Table")
         Exit Sub
     Else
         If rstDefTab.EOF Then
            rstDefTab.MovePrevious
         End If
     End If
   CmdDisab
   EditFlag = True
   cmdChk.Enabled = True
   cmdsave.Enabled = False
   RCode.Enabled = False
   RDesc.Enabled = True
   Lscore.Enabled = True
   Uscore.Enabled = True
   RDesc.SetFocus
   cmdedit.Caption = "&Cancel"
Else
   cmdedit.Caption = "&Edit"
   RCode.Enabled = True
   CmdEnab
   EditFlag = False
   cmdsave.Enabled = False
   cmdChk.Enabled = False
   showdata
   FldDisab
End If
End Sub

Private Sub cmdexit_Click()
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
    If EditFlag = True Then
        GTBuff = "code = '" & RCode & "'"
        rstDefTab.FindFirst GTBuff
        rstDefTab.Edit
    End If
    If AddFlag = True Then
        rstDefTab.AddNew
        rstDefTab("Code") = UCase(RCode)
    End If
    rstDefTab("Desc") = UCase(RDesc)
    rstDefTab("Lscore") = UCase(Lscore)
    rstDefTab("UScore") = UCase(Uscore)
    rstDefTab("FGrade") = chkFGrd
    strpos = InStr(cmbSlev, ",")
    rstDefTab("slev") = UCase(Trim(Left(cmbSlev, strpos - 1)))
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
    Cmsg = ""
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
    GenClass.fleLogin mvUserid, "Accessed Scores Grading Definition", Date, Time
    Call FormCentreMDI(Me)
    Set wrktemp = DBEngine.Workspaces(0)
    Set db1 = wrktemp.OpenDatabase(DBpath, True)
    Set rstDefTab = db1.OpenRecordset("DefSCGrd", dbOpenDynaset)
    Set rstSLev = db1.OpenRecordset("defslev", dbOpenDynaset)
    Me.Caption = MVCoyname
    If Not rstDefTab.EOF Then
       rstDefTab.MoveFirst
    End If
    EditFlag = False
    AddFlag = False
             ''__________________________________
    datgenTab.DatabaseName = DBpath
    datgenTab.RecordSource = "defslev"
    datgenTab.ReadOnly = False
    datgenTab.Exclusive = False
    datgenTab.Refresh
    Do While Not datgenTab.Recordset.EOF
       cmbSlev.AddItem datgenTab.Recordset("code") + ", " + datgenTab.Recordset("Desc")
       datgenTab.Recordset.MoveNext
    Loop
    If Not (datgenTab.Recordset.BOF And datgenTab.Recordset.EOF) Then
       cmbSlev.ListIndex = 0
    End If
    showdata
    FldDisab
    cmdChk.Enabled = False
    cmddel.Visible = True
    cmdsave.Enabled = False
    EFlag = False
    Me.Caption = MVCoyname
End Sub

Public Sub showdata()
     Dim MSlev As String
     If rstDefTab.RecordCount = 1 Then
        rstDefTab.MoveFirst
     End If
     If rstDefTab.BOF And rstDefTab.EOF Then
        ClearData
     Else
        RCode = rstDefTab("Code")
        RDesc = rstDefTab("Desc")
        Lscore = rstDefTab("LScore")
        Uscore = rstDefTab("UScore")
        MSlev = rstDefTab("slev")
        If rstDefTab!FGrade = True Then
            chkFGrd.Value = 1
        Else
            chkFGrd.Value = 0
        End If
        If IsNull(rstDefTab!sleV) = True Then
           cmbSlev.ListIndex = 0
        Else
           GTBuff = "code = '" & rstDefTab("slev") & "'"
           rstSLev.FindFirst GTBuff
           cmbSlev = rstSLev!code & "," & rstSLev!Desc
        End If
    End If
End Sub

Public Sub ClearData()
    RCode = ""
    RDesc = ""
    Lscore = 0
    Uscore = 0
    Cmsg = ""
    chkFGrd.Value = False
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
    cmdChk.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    db1.Close
End Sub
Private Sub FldEnab()
    RCode.Enabled = True
    RDesc.Enabled = True
    Lscore.Enabled = True
    Uscore.Enabled = True
    cmbSlev.Enabled = True
    chkFGrd.Enabled = True
End Sub

Private Sub FldDisab()
    RCode.Enabled = False
    RDesc.Enabled = False
    Lscore.Enabled = False
    Uscore.Enabled = False
    cmbSlev.Enabled = False
    chkFGrd.Enabled = False
End Sub


Private Sub Lscore_gotfocus()
    Lscore.SelStart = 0
    Lscore.SelLength = Len(Lscore)
End Sub

Private Sub Uscore_gotfocus()
    Uscore.SelStart = 0
    Uscore.SelLength = Len(Uscore)
End Sub
