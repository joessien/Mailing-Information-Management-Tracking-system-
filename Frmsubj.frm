VERSION 5.00
Begin VB.Form FrmSubjG 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Tracking System"
   ClientHeight    =   4695
   ClientLeft      =   1755
   ClientTop       =   1395
   ClientWidth     =   8685
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   8685
   Begin VB.CheckBox ChkMustpass 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3120
      TabIndex        =   24
      ToolTipText     =   "Tick this box if students must pass this subject to get promoted to next class"
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox ChkExt 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3120
      TabIndex        =   21
      Top             =   3000
      Width           =   255
   End
   Begin VB.Data datgenTab 
      Caption         =   "Gentab"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox cmbSClass 
      Height          =   315
      Left            =   2160
      TabIndex        =   18
      Text            =   "cmbSClass"
      Top             =   2520
      Width           =   4215
   End
   Begin VB.PictureBox SSPanel1 
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   6795
      TabIndex        =   11
      Top             =   3840
      Width           =   6855
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
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2760
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmddel 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3840
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   360
         TabIndex        =   12
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
      Height          =   345
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   1
      Top             =   2040
      Width           =   4215
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
      Left            =   2520
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1560
      Width           =   795
   End
   Begin VB.PictureBox SSPanel5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   8595
      TabIndex        =   2
      Top             =   0
      Width           =   8655
   End
   Begin VB.PictureBox SSPanel3 
      Height          =   3855
      Left            =   6960
      ScaleHeight     =   3795
      ScaleWidth      =   1635
      TabIndex        =   5
      Top             =   840
      Width           =   1695
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H00C0C0FF&
         Caption         =   "E&xit"
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdbrow 
         Caption         =   "Bro&wse"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdbott 
         Caption         =   "&Bottom"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdtop 
         Caption         =   "&Top"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.PictureBox FormDets 
      BackColor       =   &H0080C0FF&
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
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   6795
      TabIndex        =   25
      Top             =   840
      Width           =   6855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Must Pass Subject:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   22
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "External Exam Subject:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   20
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Subject Class:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   17
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
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
      Left            =   2160
      TabIndex        =   16
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label DeptName 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label StaffNo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Subject Code:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "FrmSubjG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EditFlag As Boolean, MaxErrors As Integer
Dim rstDefTab As Recordset, AddFlag As Boolean
Dim db1 As Database, wrktemp As Workspace
Dim AType As String, ValBuff As String
Dim FirstTime As Integer, FirstPass As Integer
Dim GenClass As New DProLog
Dim rstSClass As Recordset, strpos As Integer

Private Sub CmdAdd_Click()
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
    mvSql = "Select code, desc, mustpass From defsubj ;"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No subjects defined", vbExclamation, "Subject Definition "
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "List of Subjects"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         frmBrowse.grdGeneral.Columns(0).Width = 1400
         frmBrowse.grdGeneral.Columns(1).Width = 2800
         frmBrowse.grdGeneral.Columns(2).Width = 1800
         Col0.Caption = "Subject Code"
         Col1.Caption = "Subject Description"
         Col2.Caption = "Mustpass Subject"
   End If
   frmBrowse.Show

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
   FldEnab
   RDesc.SetFocus
   cmdedit.Caption = "&Cancel"
   RCode.Enabled = False
   CmdDisab
Else
   cmdedit.Caption = "&Edit"
   RCode.Enabled = True
   CmdEnab
   EditFlag = False
   cmdsave.Enabled = False
   showdata
   FldDisab
End If
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
  Dim sql As String, digitflag As Integer, length As Integer
   Dim buff As String, AN As String

   On Error GoTo FindError1

   AN = InputBox$("Enter the Subject Code to find:")
   length = Len(AN)
   AN = UCase(AN)
   If length = 0 Then
      ErrorMessages (NOINPUT)
      Exit Sub
   End If
   If length = 1 Then
      buff = "code LIKE " & Chr$(34) & AN & "*" & Chr$(34)
   Else
      buff = "code LIKE " & Chr$(34) & AN & "*" & Chr$(34)
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
    If EditFlag = True Then
        rstDefTab.Edit
    End If
    If AddFlag = True Then
        rstDefTab.AddNew
        rstDefTab("Code") = "B" & UCase(RCode)
    End If
    rstDefTab("Desc") = UCase(RDesc)
    strpos = InStr(1, cmbSClass, ",", 1)
    rstDefTab("sclass") = Left(cmbSClass, strpos - 1)
    rstDefTab("extint") = ChkExt
    rstDefTab("mustpass") = ChkMustpass
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
    Dim flag As Integer, mtable As String
    GenClass.fleLogin mvUserid, "Accessed Subject Definition", Date, Time
    Call FormCentreMDI(Me)
    Set wrktemp = DBEngine.Workspaces(0)
    Set db1 = wrktemp.OpenDatabase(DBpath, True)
    Set rstDefTab = db1.OpenRecordset("DefsubJ", dbOpenDynaset)
    Set rstSClass = db1.OpenRecordset("DefSubclass", dbOpenDynaset)
    Me.Caption = MVCoyname
    If Not rstDefTab.EOF Then
       rstDefTab.MoveFirst
    End If
    EditFlag = False
    AddFlag = False
    cmddel.Visible = False
    cmdsave.Enabled = False
    ChkExt.Value = 0
    Me.Caption = MVCoyname
             ''__________________________________
    datGenTab.DatabaseName = DBpath
    datGenTab.RecordSource = "DefSubclass"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       cmbSClass.AddItem datGenTab.Recordset("code") + ", " + datGenTab.Recordset("Desc")
       datGenTab.Recordset.MoveNext
    Loop
    cmbSClass.ListIndex = 0
    showdata
    FldDisab

End Sub

Public Sub showdata()
    'On Error GoTo perror
     If rstDefTab.RecordCount >= 1 Then
        RCode = Right(rstDefTab("Code"), 3)
        RDesc = rstDefTab("Desc")
        If IsNull(rstDefTab!SClass) = True Then
           cmbSClass.ListIndex = 0
        Else
           GTBuff = "code = '" & rstDefTab("sclass") & "'"
           rstSClass.FindFirst GTBuff
           cmbSClass = rstSClass!code & "," & rstSClass!Desc
           FormDets.Caption = rstDefTab!code & " - " & rstDefTab!Desc
           If rstDefTab("extint") = True Then
              ChkExt.Value = 1
           Else
              ChkExt.Value = 0
           End If
           If rstDefTab("mustpass") = True Then
              ChkMustpass.Value = 1
           Else
              ChkMustpass.Value = 0
           End If
        End If
    End If
PError:
    Exit Sub
End Sub

Public Sub ClearData()
    RCode = ""
    RDesc = ""
    Sseq = 0
    ChkMustpass.Value = False
    ChkExt.Value = False
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
    db1.Close
End Sub
Private Sub FldEnab()
    RCode.Enabled = True
    RDesc.Enabled = True
    cmbSClass.Enabled = True
    ChkExt.Enabled = True
    ChkMustpass.Enabled = True
End Sub

Private Sub FldDisab()
    RCode.Enabled = False
    RDesc.Enabled = False
    cmbSClass.Enabled = False
    ChkExt.Enabled = False
    ChkMustpass.Enabled = False
End Sub


