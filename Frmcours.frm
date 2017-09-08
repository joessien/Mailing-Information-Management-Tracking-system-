VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmDPCrsDef 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datapro"
   ClientHeight    =   4335
   ClientLeft      =   945
   ClientTop       =   1530
   ClientWidth     =   7350
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7350
   Begin Threed.SSPanel SSPanel2 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   7335
      _Version        =   65536
      _ExtentX        =   12938
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
         Left            =   720
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmddel 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2880
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   3960
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   5880
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   3255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
      _ExtentY        =   5741
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   3015
         Left            =   5760
         TabIndex        =   10
         Top             =   120
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   5318
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
         Begin VB.CommandButton cmdfind 
            Caption         =   "&Find"
            Height          =   435
            Left            =   240
            TabIndex        =   17
            Top             =   2040
            Width           =   975
         End
         Begin VB.CommandButton cmdbrow 
            Caption         =   "Bro&wse"
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   2520
            Width           =   975
         End
         Begin VB.CommandButton cmdtop 
            Caption         =   "&Top"
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Top             =   120
            Width           =   975
         End
         Begin VB.CommandButton cmdbott 
            Caption         =   "&Bottom"
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   1560
            Width           =   975
         End
         Begin VB.CommandButton cmdnext 
            Caption         =   "&Next"
            Height          =   375
            Left            =   240
            TabIndex        =   13
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmdprev 
            Caption         =   "&Previous"
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   1080
            Width           =   975
         End
      End
      Begin VB.TextBox CrsOrg 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   2040
         Width           =   3975
      End
      Begin MSMask.MaskEdBox CrsTenor 
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   327680
         MaxLength       =   4
         PromptChar      =   "_"
      End
      Begin VB.TextBox CrsCode 
         Height          =   285
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox CrsDesc 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   2400
         Width           =   3975
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   735
         Left            =   360
         TabIndex        =   23
         Top             =   240
         Width           =   5175
         _Version        =   65536
         _ExtentX        =   9128
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "Course Definition"
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
      Begin VB.Label Label3 
         Caption         =   "Days"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
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
         Left            =   1560
         TabIndex        =   11
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808080&
         Caption         =   "Duration:"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         Caption         =   "Organizers:"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label label1 
         BackColor       =   &H00808080&
         Caption         =   "Course Code:"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label label2 
         BackColor       =   &H00808080&
         Caption         =   "Description:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   2400
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmDPCrsDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rstCrs As Recordset, cmbval As String
Dim db6 As Database, wrktemp As Workspace
Dim AddFlag As Boolean, rstDept As Recordset
Dim TmpBuff As String, strpos As Byte
Dim RegTab As Recordset
Dim EditFlag As Boolean, FError As Boolean
Dim GenClass As New DProLog

Private Sub CmdAdd_Click()
 If cmdadd.Caption = "&Add" Then
    rstCrs.AddNew
    AddFlag = True
    ClearData
    CrsCode.SetFocus
    CmdDisab
    cmdedit.Enabled = False
    cmdsave.Enabled = True
    cmdadd.Enabled = True
    cmdadd.Caption = "&Cancel"
Else
   If Not (rstCrs.BOF And rstCrs.EOF) Then
        rstCrs.MoveFirst
   End If
   cmdadd.Caption = "&Add"
   CmdEnab
   AddFlag = False
   cmdedit.Enabled = True
   cmdsave.Enabled = False
End If
End Sub

Private Sub CmdBott_Click()
    Dim Count As Long
    If rstCrs.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstCrs.MoveLast
    showdata
End Sub

Private Sub CmdBrow_Click()
    Dim GenView As New frmBrowse
    Set GenView.datGeneral.Recordset = rstCrs
    GenView.Caption = Me.Caption & " - Records View"
    GenView.Show

End Sub

Private Sub CmdDel_Click()
   Dim i As Integer
   If rstCrs.RecordCount < 1 Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
    i = DeleteCheck()
    If i = vbYes Then
        rstCrs.Delete
        If Not rstCrs.BOF Then
          rstCrs.MovePrevious
          showdata
        Else
          ClearData
        End If
        showdata
    End If
End Sub

Private Sub CmdEdit_Click()
If cmdedit.Caption = "&Edit" Then
  If rstCrs.EOF Then
       MsgBox ("Empty Table")
       Exit Sub
   End If
   EditFlag = True
   cmdsave.Enabled = True
   rstCrs.Edit
   CrsTenor.SetFocus
   cmdedit.Caption = "&Cancel"
   CrsCode.Enabled = False
   CmdDisab
Else
   cmdedit.Caption = "&Edit"
   CrsCode.Enabled = True
   CmdEnab
   EditFlag = False
   cmdsave.Enabled = False
End If
 
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
  Dim sql As String, digitflag As Integer, length As Integer
   Dim buff As String, AN As String

   On Error GoTo FindError1

   AN = InputBox$("Enter the Course Code to find:")
   length = Len(AN)
   If length = 0 Then
      ErrorMessages (NOINPUT)
      Exit Sub
   End If
   buff = "Code LIKE '" & "C" & AN & "'"
   rstCrs.FindFirst buff
   If rstCrs.NOMATCH Then
      ErrorMessages (NOMATCH)
   Else
      showdata
   End If

FindError0:
   Exit Sub

FindError1:
   MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
   On Error GoTo 0
   Resume FindError0

End Sub

Private Sub CmdNext_Click()
  Dim flag As Integer
    
    On Error GoTo NextError
    If rstCrs.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstCrs.EOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEPAST)
        rstCrs.MoveLast
    Else
        rstCrs.MoveNext
        If rstCrs.EOF Then
            rstCrs.MoveLast
        End If
    End If
NextClear:
    showdata
    Exit Sub
NextError:
    ErrorMessages (NOMOVEPAST)
    rstCrs.Requery
    rstCrs.MoveLast
    On Error GoTo 0
    Resume NextClear
End Sub

Private Sub CmdPrev_Click()
 Dim flag As Integer
    
    On Error GoTo PrevError
    If rstCrs.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstCrs.BOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEFRONT)
        rstCrs.MoveFirst
    Else
        rstCrs.MovePrevious
        If rstCrs.BOF Then
            rstCrs.MoveFirst
        End If
    End If
PrevClear:
    showdata
    Exit Sub
PrevError:
    ErrorMessages (NOMOVEFRONT)
    rstCrs.Requery
    rstCrs.MoveFirst
    On Error GoTo 0
    Resume PrevClear
End Sub

Private Sub CmdSave_Click()
    Dim i As Integer
    On Error GoTo PError
    FError = False
    FldChk
    If FError = True Then
       Exit Sub
    End If
    If EditFlag = True Then
        rstCrs.Edit
    End If
    If AddFlag = True Then
        rstCrs.AddNew
        rstCrs("Code") = "C" & CrsCode
    End If
    rstCrs("Tenor") = CrsTenor
    rstCrs("Org") = CrsOrg
    rstCrs("Desc") = CrsDesc
    rstCrs.Update
    If Not (rstCrs.EOF And rstCrs.BOF) Then
        rstCrs.MoveLast
    End If
    cmdadd.Caption = "&Add"
    cmdedit.Enabled = True
    cmdsave.Enabled = False
    cmdedit.Caption = "&Edit"
    CrsCode.Enabled = True
    CmdEnab
    showdata
    EditFlag = False
    AddFlag = False
PError0:
    Exit Sub
PError:
    MsgBox "Invalid or Duplicate Input made, Check Entries", vbCritical, Me.Caption
    On Error GoTo 0
    Resume PError0

End Sub

Private Sub CmdTop_Click()
    If rstCrs.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstCrs.MoveFirst
    showdata
        
End Sub

Private Sub Form_Load()
    GenClass.fleLogin mvUserid, "Accessed Course Scheduling Definition", Date, Time
    Call FormCentreMDI(Me)
    Set wrktemp = DBEngine.Workspaces(0)
    Set db6 = wrktemp.OpenDatabase(DproDBpath, True)
    Set rstCrs = db6.OpenRecordset("CrsTab", dbOpenDynaset)
    Dim flag As Integer, mtable As String
    If Not rstCrs.EOF Then
       rstCrs.MoveFirst
    End If
    showdata
    EditFlag = False
    AddFlag = False
    cmdsave.Enabled = False
    cmddel.Visible = False

End Sub

Public Sub showdata()
     If rstCrs.RecordCount >= 1 Then
        CrsCode = Right(rstCrs("Code"), 3)
        CrsTenor = rstCrs("Tenor")
        CrsOrg = rstCrs("Org")
        CrsDesc = rstCrs("desc")
    End If
End Sub

Public Sub ClearData()
    strpos = 0
    CrsCode = ""
    CrsTenor = ""
    CrsOrg = ""
    CrsDesc = ""
    
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
    db6.Close
End Sub

Public Sub FldChk()
    If Len(CrsOrg) = 0 Then
        MsgBox "Null Organizers", vbExclamation, "Course Definition"
        CrsOrg.SetFocus
        FError = True
        Exit Sub
    End If
    If Len(CrsDesc) = 0 Then
        MsgBox "Invalid Course Description", vbExclamation, "Course Definition"
        CrsDesc.SetFocus
        FError = True
        Exit Sub
    End If
End Sub
