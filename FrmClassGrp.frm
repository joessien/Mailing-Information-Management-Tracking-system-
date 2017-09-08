VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmClasGrp 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounting System"
   ClientHeight    =   5175
   ClientLeft      =   1755
   ClientTop       =   1395
   ClientWidth     =   9795
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   9795
   Begin VB.CheckBox GradClass 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   3840
      TabIndex        =   25
      Top             =   3960
      Width           =   255
   End
   Begin MSMask.MaskEdBox AvgPassMark 
      Height          =   375
      Left            =   3840
      TabIndex        =   24
      Top             =   3000
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   255
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "##0.00;(##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Data datgenTab 
      Caption         =   "Gentab"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox cmbSlev 
      Height          =   315
      Left            =   2520
      TabIndex        =   20
      Top             =   2520
      Width           =   4335
   End
   Begin VB.TextBox Sseq 
      Height          =   360
      Left            =   3840
      MaxLength       =   1
      TabIndex        =   19
      Top             =   3480
      Width           =   495
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   0
      TabIndex        =   11
      Top             =   4440
      Width           =   9735
      _Version        =   65536
      _ExtentX        =   17171
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
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
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
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmddel 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3600
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
      Left            =   2520
      MaxLength       =   40
      TabIndex        =   1
      Top             =   2040
      Width           =   4335
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
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1560
      Width           =   795
   End
   Begin Threed.SSPanel SSPanel5 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9735
      _Version        =   65536
      _ExtentX        =   17171
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "Class Group Definition "
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
      Height          =   2655
      Left            =   8040
      TabIndex        =   5
      Top             =   840
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   4683
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
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdbrow 
         Caption         =   "Bro&wse"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdbott 
         Caption         =   "&Bottom"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdtop 
         Caption         =   "&Top"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin Threed.SSPanel FormDets 
      Height          =   615
      Left            =   0
      TabIndex        =   27
      Top             =   840
      Width           =   8055
      _Version        =   65536
      _ExtentX        =   14208
      _ExtentY        =   1085
      _StockProps     =   15
      ForeColor       =   255
      BackColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelInner      =   1
      Font3D          =   2
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Graduation Class:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   26
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Minimum Promotion Avg. Pass Mark:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   23
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Class Promotion  Sequence:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Class Level:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   2640
      Width           =   1335
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
      Left            =   2520
      TabIndex        =   16
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label DeptName 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label StaffNo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Class Code:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
End
Attribute VB_Name = "FrmClasGrp"
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
Dim rstSLev As Recordset, strpos As Integer

Private Sub CmdAdd_Click()
On Error GoTo PError
 If cmdadd.Caption = "&Add" Then
    rstDefTab.AddNew
    AddFlag = True
    FldEnab
    ClearData
    RCode.SetFocus
    cmdSave.Enabled = True
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
   cmdSave.Enabled = False
   showdata
   FldDisab
End If
PError:
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
    mvSql = "Select * From Defclass"
    mvSql = mvSql + " order by sseq;"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No class group is defined", vbExclamation, "Class Group Definition"
         Exit Sub
     Else
         Set frmBrwse.datGeneral.Recordset = rst
         frmBrwse.grdGeneral.Caption = "LIST OF CLASS GROUPS"
         frmBrwse.Caption = Me.Caption
         Set Col0 = frmBrwse.grdGeneral.Columns(0)
         Set Col1 = frmBrwse.grdGeneral.Columns(1)
         Set Col2 = frmBrwse.grdGeneral.Columns(2)
         Set Col3 = frmBrwse.grdGeneral.Columns(3)
         Set Col4 = frmBrwse.grdGeneral.Columns(4)
         Set Col5 = frmBrwse.grdGeneral.Columns(5)
         frmBrwse.grdGeneral.Columns(0).Width = 1000
         frmBrwse.grdGeneral.Columns(1).Width = 2800
         frmBrwse.grdGeneral.Columns(2).Width = 1800
         frmBrwse.grdGeneral.Columns(3).Width = 2000
         frmBrwse.grdGeneral.Columns(4).Width = 3000
         frmBrwse.grdGeneral.Columns(5).Width = 3000
         Col0.Caption = "Group Code"
         Col1.Caption = "Group Description"
         Col2.Caption = "Group Level"
         Col3.Caption = "Progression Sequence"
         Col4.Caption = "Mim Promotion Avg Pass Mark"
         Col5.Caption = "Graduating Class"
   End If
   frmBrwse.Show
PError:
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
   cmdSave.Enabled = True
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
   cmdSave.Enabled = False
   showdata
   FldDisab
End If
PError:
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
  Dim sql As String, digitflag As Integer, length As Integer
   Dim buff As String, AN As String

   On Error GoTo FindError1

   AN = InputBox$("Enter the Class Group Code to find:")
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
        rstDefTab("Code") = UCase(RCode)
    End If
    rstDefTab("Desc") = UCase(RDesc)
    rstDefTab("sseq") = Sseq
    rstDefTab("gradclass") = GradClass
    strpos = InStr(1, cmbSlev, ",", 1)
    rstDefTab("slev") = Left(cmbSlev, strpos - 1)
    rstDefTab("avgpassmark") = AvgPassMark
    rstDefTab.Update
    AddFlag = False
    cmdadd.Caption = "&Add"
    CmdEnab
    cmdedit.Enabled = True
    cmdSave.Enabled = False
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
    Dim flag As Integer, mtable As String
    GenClass.fleLogin mvUserid, "Accessed Class Group Definition", Date, Time
    Call FormCentreMDI(Me)
    Set wrktemp = DBEngine.Workspaces(0)
    Set db1 = wrktemp.OpenDatabase(DBpath, True)
    Set rstDefTab = db1.OpenRecordset("Defclass", dbOpenDynaset)
    Set rstSLev = db1.OpenRecordset("DefSlev", dbOpenDynaset)
    Me.Caption = MVCoyname
    If Not rstDefTab.EOF Then
       rstDefTab.MoveFirst
    End If
    EditFlag = False
    AddFlag = False
    cmddel.Visible = False
    cmdSave.Enabled = False
    Me.Caption = MVCoyname
             ''__________________________________
    Do While Not rstSLev.EOF
       cmbSlev.AddItem rstSLev("code") + ", " + rstSLev("Desc")
       rstSLev.MoveNext
    Loop
    If Not (rstSLev.BOF And rstSLev.EOF) Then
       cmbSlev.ListIndex = 0
    End If
    showdata
    FldDisab
PError:
End Sub

Public Sub showdata()
On Error GoTo PError
     If rstDefTab.RecordCount >= 1 Then
        RCode = rstDefTab("Code")
        RDesc = rstDefTab("Desc")
        Sseq = rstDefTab("sseq")
        FormDets.Caption = rstDefTab!code & " - " & rstDefTab!Desc
        If IsNull(rstDefTab!sleV) = True Then
           cmbSlev.ListIndex = 0
        Else
           GTBuff = "code = '" & rstDefTab("slev") & "'"
           rstSLev.FindFirst GTBuff
           cmbSlev = rstSLev!code & "," & rstSLev!Desc
        End If
        AvgPassMark = rstDefTab("avgpassmark")
        If rstDefTab!GradClass = True Then
           GradClass.Value = 1
        Else
           GradClass.Value = 0
        End If
    End If
PError:
End Sub

Public Sub ClearData()
    RCode = ""
    RDesc = ""
    Sseq = 0
    AvgPassMark = 0
    GradClass.Value = 0
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
    Sseq.Enabled = True
    cmbSlev.Enabled = True
    AvgPassMark.Enabled = True
    GradClass.Enabled = True
End Sub

Private Sub FldDisab()
    RCode.Enabled = False
    RDesc.Enabled = False
    Sseq.Enabled = False
    cmbSlev.Enabled = False
    AvgPassMark.Enabled = False
    GradClass.Enabled = False
End Sub


