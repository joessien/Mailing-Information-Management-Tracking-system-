VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form FrmJob 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datapro"
   ClientHeight    =   4365
   ClientLeft      =   1890
   ClientTop       =   1260
   ClientWidth     =   6720
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6720
   Begin Threed.SSPanel SSPanel2 
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   3240
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
      Height          =   3255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6735
      _Version        =   65536
      _ExtentX        =   11880
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
         Left            =   5040
         TabIndex        =   6
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
         Begin VB.CommandButton cmdbrow 
            Caption         =   "Bro&wse"
            Height          =   375
            Left            =   360
            TabIndex        =   12
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdtop 
            Caption         =   "&Top"
            Height          =   375
            Left            =   360
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdbott 
            Caption         =   "&Bottom"
            Height          =   375
            Left            =   360
            TabIndex        =   10
            Top             =   1680
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
         Begin VB.CommandButton cmdprev 
            Caption         =   "&Previous"
            Height          =   375
            Left            =   360
            TabIndex        =   8
            Top             =   1200
            Width           =   975
         End
      End
      Begin VB.TextBox xDesc 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox xCode 
         Height          =   285
         Left            =   2160
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
         Width           =   4575
         _Version        =   65536
         _ExtentX        =   8070
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "Job Title Definition"
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
      Begin VB.Label Label8 
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "J"
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
         Left            =   1920
         TabIndex        =   7
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   "Job Description"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label StaffNo 
         BackColor       =   &H00808080&
         Caption         =   "Job Title"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EditFlag As Boolean, MaxErrors As Integer
Dim AType As String, ValBuff As String
Dim rstjob As Recordset, AddFlag As Boolean
Dim db6 As Database, wrktemp As Workspace
Dim FirstTime As Integer, FirstPass As Integer
Dim GenClass As New DProLog

Private Sub CmdAdd_Click()
 If cmdAdd.Caption = "&Add" Then
    rstjob.AddNew
    AddFlag = True
    ClearData
    xCode.SetFocus
    cmdSave.Enabled = True
    AddFlag = True
    CmdDisab
    cmdEdit.Enabled = False
    cmdAdd.Enabled = True
    cmdAdd.Caption = "&Cancel"
Else
    If Not rstjob.EOF Then
       rstjob.MoveFirst
    End If
   AddFlag = False
   cmdAdd.Caption = "&Add"
   CmdEnab
   cmdEdit.Enabled = True
   cmdSave.Enabled = False
   AddFlag = False
   showdata
End If
End Sub

Private Sub CmdBott_Click()
   Dim Count As Long
    If rstjob.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstjob.MoveLast
    showdata
End Sub

Private Sub CmdBrow_Click()
    Dim GenView As New frmBrowse
    Set GenView.datGeneral.Recordset = rstjob
    GenView.Caption = Me.Caption & " - Records View"
    GenView.Show

End Sub

Private Sub CmdDel_Click()
   Dim i As Integer
   If rstjob.RecordCount < 1 Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
    i = DeleteCheck()
    If i = vbYes Then
        rstjob.Delete
        If Not rstjob.BOF Then
          rstjob.MovePrevious
          showdata
        Else
          ClearData
        End If
    End If
End Sub

Private Sub CmdEdit_Click()
 If cmdEdit.Caption = "&Edit" Then
   If rstjob.EOF Then
       MsgBox ("Empty Table")
       Exit Sub
   End If
   EditFlag = True
   cmdSave.Enabled = True
   rstjob.Edit
   xDesc.SetFocus
   cmdEdit.Caption = "&Cancel"
   xCode.Enabled = False
   CmdDisab
Else
   cmdEdit.Caption = "&Edit"
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
    If rstjob.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstjob.EOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEPAST)
        rstjob.MoveLast
    Else
        rstjob.MoveNext
        If rstjob.EOF Then
            rstjob.MoveLast
        End If
    End If
NextClear:
    showdata
    Exit Sub
NextError:
    ErrorMessages (NOMOVEPAST)
    ''rstjob.Requery
    rstjob.MoveLast
    On Error GoTo 0
    Resume NextClear

End Sub

Private Sub CmdPrev_Click()
 Dim flag As Integer
    
    On Error GoTo PrevError
    If rstjob.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstjob.BOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEFRONT)
        rstjob.MoveFirst
    Else
        rstjob.MovePrevious
        If rstjob.BOF Then
            rstjob.MoveFirst
        End If
    End If
PrevClear:
    showdata
    Exit Sub
PrevError:
    ErrorMessages (NOMOVEFRONT)
    ''rstjob.Requery
    rstjob.MoveFirst
    On Error GoTo 0
    Resume PrevClear
End Sub

Private Sub CmdSave_Click()
  Dim i As Integer
    On Error GoTo PError
    
    If EditFlag = True Then
        rstjob.Edit
    End If
    If AddFlag = True Then
        rstjob.AddNew
        rstjob("code") = "J" & xCode
    End If
    rstjob("desc") = xDesc
    rstjob.Update
    AddFlag = False
    cmdAdd.Caption = "&Add"
    CmdEnab
    cmdEdit.Enabled = True
    cmdSave.Enabled = False
    cmdEdit.Caption = "&Edit"
    xCode.Enabled = True
    rstjob.MoveLast
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

Private Sub CmdTop_Click()
    If rstjob.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstjob.MoveFirst
    showdata
End Sub



Private Sub Form_Load()
    GenClass.fleLogin mvUserid, "Accessed Job Title", Date, Time
    Call FormCentreMDI(Me)
    Set wrktemp = DBEngine.Workspaces(0)
    Set db6 = wrktemp.OpenDatabase(DproDBpath, True)
    Set rstjob = db6.OpenRecordset("jobtab", dbOpenDynaset)
    Dim flag As Integer, mtable As String
    If Not rstjob.EOF Then
       rstjob.MoveFirst
    End If
    showdata
    EditFlag = False
    AddFlag = False
    cmdSave.Enabled = False
    cmdDel.Visible = False
End Sub

Public Sub showdata()
     If Not (rstjob.EOF And rstjob.BOF) Then
        xCode = Right(rstjob("code"), 3)
        xDesc = rstjob("desc")
    End If
End Sub

Public Sub ClearData()
    xCode = ""
    xDesc = ""

End Sub
Public Sub CmdEnab()
    cmdAdd.Enabled = True
    cmdDel.Enabled = True
    cmdBrow.Enabled = True
    cmdtop.Enabled = True
    cmdbott.Enabled = True
    cmdnext.Enabled = True
    cmdprev.Enabled = True
End Sub

Public Sub CmdDisab()
    cmdAdd.Enabled = False
    cmdDel.Enabled = False
    cmdBrow.Enabled = False
    cmdtop.Enabled = False
    cmdbott.Enabled = False
    cmdnext.Enabled = False
    cmdprev.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    db6.Close
End Sub
