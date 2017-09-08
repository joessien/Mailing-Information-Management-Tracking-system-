VERSION 5.00
Begin VB.Form FrmQual 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Tracking System"
   ClientHeight    =   4110
   ClientLeft      =   1755
   ClientTop       =   1395
   ClientWidth     =   7830
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7830
   Begin VB.PictureBox SSPanel1 
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   7755
      TabIndex        =   12
      Top             =   3360
      Width           =   7815
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H00C0C0FF&
         Caption         =   "E&xit"
         Height          =   375
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2760
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmddel 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3840
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
         Size            =   9.75
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
      Top             =   2280
      Width           =   3615
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
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1800
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
      Left            =   -120
      ScaleHeight     =   675
      ScaleWidth      =   7755
      TabIndex        =   2
      Top             =   0
      Width           =   7815
   End
   Begin VB.PictureBox SSPanel3 
      Height          =   2415
      Left            =   6120
      ScaleHeight     =   2355
      ScaleWidth      =   1635
      TabIndex        =   6
      Top             =   840
      Width           =   1695
      Begin VB.CommandButton cmdbrow 
         Caption         =   "Bro&wse"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdbott 
         Caption         =   "&Bottom"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdtop 
         Caption         =   "&Top"
         Height          =   375
         Left            =   360
         TabIndex        =   8
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
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   5955
      TabIndex        =   18
      Top             =   840
      Width           =   6015
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
      Left            =   960
      TabIndex        =   5
      Top             =   2400
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
      Left            =   960
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Q"
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
      Top             =   1800
      Width           =   375
   End
End
Attribute VB_Name = "FrmQual"
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
        rstDefTab("Code") = "Q" & UCase(RCode)
    End If
    rstDefTab("Desc") = RDesc
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
    GenClass.fleLogin mvUserid, "Accessed Staff Qualifications Definition", Date, Time
    Call FormCentreMDI(Me)
    Set wrktemp = DBEngine.Workspaces(0)
    Set db1 = wrktemp.OpenDatabase(DBpath, True)
    Set rstDefTab = db1.OpenRecordset("DefQual", dbOpenDynaset)
    Me.Caption = MVCoyname
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
End Sub

Public Sub showdata()
     If rstDefTab.RecordCount >= 1 Then
        RCode = Right(rstDefTab("Code"), 3)
        RDesc = rstDefTab("Desc")
        FormDets.Caption = rstDefTab!code & " - " & rstDefTab!Desc
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
End Sub

Private Sub FldDisab()
    RCode.Enabled = False
    RDesc.Enabled = False
End Sub


