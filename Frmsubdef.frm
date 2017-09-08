VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form FrmSubDef 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datapro"
   ClientHeight    =   4335
   ClientLeft      =   1890
   ClientTop       =   1260
   ClientWidth     =   6765
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6765
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
         Left            =   5280
         TabIndex        =   6
         Top             =   120
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
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
         Left            =   2280
         TabIndex        =   1
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox xCode 
         Height          =   285
         Left            =   2520
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
         Caption         =   "Qualifications Definition"
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
         Caption         =   "Q"
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
         Left            =   2280
         TabIndex        =   7
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   "Qualification Description:"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label StaffNo 
         BackColor       =   &H00808080&
         Caption         =   "Qualification Code:"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmSubDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EditFlag As Boolean, MaxErrors As Integer
Dim AType As String, ValBuff As String
Dim rstQual As Recordset, AddFlag As Boolean
Dim db6 As Database, wrktemp As Workspace
Dim FirstTime As Integer, FirstPass As Integer
Dim GenClass As New DProLog

Private Sub CmdAdd_Click()
 If cmdadd.Caption = "&Add" Then
    rstQual.AddNew
    AddFlag = True
    ClearData
    xCode.SetFocus
    cmdsave.Enabled = True
    AddFlag = True
    CmdDisab
    cmdedit.Enabled = False
    cmdadd.Enabled = True
    cmdadd.Caption = "&Cancel"
Else
    If Not rstQual.EOF Then
       rstQual.MoveFirst
    End If
   AddFlag = False
   cmdadd.Caption = "&Add"
   CmdEnab
   cmdedit.Enabled = True
   cmdsave.Enabled = False
   AddFlag = False
   showdata
End If
End Sub

Private Sub CmdBott_Click()
   Dim Count As Long
    If rstQual.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstQual.MoveLast
    showdata
End Sub

Private Sub CmdBrow_Click()
    Dim GenView As New frmBrowse
    Set GenView.datGeneral.Recordset = rstQual
    GenView.Caption = Me.Caption & " - Records View"
    GenView.Show

End Sub

Private Sub CmdDel_Click()
   Dim i As Integer
   If (rstQual.BOF And rstQual.EOF) Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
    i = DeleteCheck()
    If i = vbYes Then
        rstQual.Delete
        If Not rstQual.BOF Then
          rstQual.MovePrevious
          showdata
        Else
          ClearData
        End If
    End If
End Sub

Private Sub CmdEdit_Click()
 If cmdedit.Caption = "&Edit" Then
   If rstQual.EOF Then
       MsgBox ("Empty Table")
       Exit Sub
   End If
   EditFlag = True
   cmdsave.Enabled = True
   rstQual.Edit
   xDesc.SetFocus
   cmdedit.Caption = "&Cancel"
   xCode.Enabled = False
   CmdDisab
Else
   cmdedit.Caption = "&Edit"
   xCode.Enabled = True
   CmdEnab
   EditFlag = False
   cmdsave.Enabled = False
   showdata
End If
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdNext_Click()
    Dim flag As Integer
    On Error GoTo NextError
    If rstQual.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstQual.EOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEPAST)
        rstQual.MoveLast
    Else
        rstQual.MoveNext
        If rstQual.EOF Then
            rstQual.MoveLast
        End If
    End If
NextClear:
    showdata
    Exit Sub
NextError:
    ErrorMessages (NOMOVEPAST)
    ''rstQual.Requery
    rstQual.MoveLast
    On Error GoTo 0
    Resume NextClear

End Sub

Private Sub CmdPrev_Click()
 Dim flag As Integer
    On Error GoTo PrevError
    If rstQual.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstQual.BOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEFRONT)
        rstQual.MoveFirst
    Else
        rstQual.MovePrevious
        If rstQual.BOF Then
            rstQual.MoveFirst
        End If
    End If
PrevClear:
    showdata
    Exit Sub
PrevError:
    ErrorMessages (NOMOVEFRONT)
    ''rstQual.Requery
    rstQual.MoveFirst
    On Error GoTo 0
    Resume PrevClear
End Sub

Private Sub CmdSave_Click()
  Dim i As Integer
    On Error GoTo PError
    
    If EditFlag = True Then
        rstQual.Edit
    End If
    If AddFlag = True Then
        rstQual.AddNew
        rstQual("code") = "Q" & xCode
    End If
    rstQual("desc") = xDesc
    rstQual.Update
    AddFlag = False
    cmdadd.Caption = "&Add"
    CmdEnab
    cmdedit.Enabled = True
    cmdsave.Enabled = False
    cmdedit.Caption = "&Edit"
    xCode.Enabled = True
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
  
    If rstQual.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstQual.MoveFirst
    showdata
End Sub



Private Sub Form_Load()
    GenClass.fleLogin mvUserid, "Accessed Job Title", Date, Time
    Call FormCentreMDI(Me)
    Set wrktemp = DBEngine.Workspaces(0)
    Set db6 = wrktemp.OpenDatabase(DproDBpath, True)
    Set rstQual = db6.OpenRecordset("QualTab", dbOpenDynaset)
    Dim flag As Integer, mtable As String
    If Not rstQual.EOF Then
       rstQual.MoveFirst
    End If
    showdata
    EditFlag = False
    AddFlag = False
    cmdsave.Enabled = False
    cmddel.Visible = False
End Sub

Public Sub showdata()
     If Not (rstQual.BOF And rstQual.EOF) Then
        xCode = Right(rstQual("code"), 3)
        xDesc = rstQual("desc")
    End If
End Sub

Public Sub ClearData()
    xCode = ""
    xDesc = ""

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
