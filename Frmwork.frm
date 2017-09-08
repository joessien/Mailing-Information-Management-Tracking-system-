VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmDPwork 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datapro"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6135
   Begin VB.CommandButton cmdadd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox xDesc 
      Height          =   285
      Left            =   1920
      MaxLength       =   40
      TabIndex        =   1
      Text            =   " "
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox xCode 
      Height          =   285
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   0
      Text            =   " "
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdbrow 
      Caption         =   "Bro&wse"
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdtop 
      Caption         =   "&Top"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdbott 
      Caption         =   "&Bottom"
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdprev 
      Caption         =   "&Previous"
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   2160
      Width           =   975
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   360
      TabIndex        =   15
      Top             =   360
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "Work Definition"
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
      Caption         =   "W"
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
      TabIndex        =   14
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "Work Description:"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Work Code:"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "FrmDPwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EditFlag As Boolean, MaxErrors As Integer
Dim GenClass As New DProLog
Dim ValBuff As String, AddFlag As Boolean, TablPtr As Integer
Dim db6 As Database, wrktemp As Workspace, rstWork As Recordset
Dim FirstTime As Integer, FirstPass As Integer

Private Sub CmdAdd_Click()
 If cmdAdd.Caption = "&Add" Then
    rstWork.AddNew
    CmdDisab
    FldEnab
    AddFlag = True
    ClearData
    xCode.SetFocus
    AddFlag = True
    cmdEdit.Enabled = False
    cmdSave.Enabled = True
    cmdAdd.Enabled = True
    cmdAdd.Caption = "&Cancel"
Else
    If Not rstWork.EOF Then
       rstWork.MoveFirst
    End If
   AddFlag = False
   cmdAdd.Caption = "&Add"
   CmdEnab
   FldDisab
   AddFlag = False
   cmdEdit.Enabled = True
   cmdSave.Enabled = False
   showdata
End If
End Sub

Private Sub CmdBott_Click()
    Dim Count As Long
    If rstWork.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstWork.MoveLast
    showdata
End Sub

Private Sub CmdBrow_Click()
    Dim GenView As New frmBrowse
    Set GenView.datGeneral.Recordset = rstWork
    GenView.Caption = Me.Caption & " - Records View"
    GenView.Show
    End Sub

Private Sub CmdDel_Click()
   Dim i As Integer
   On Error GoTo DelError
   If rstWork.RecordCount < 1 Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
    i = DeleteCheck()
    If i = vbYes Then
        rstWork.Delete
        rstWork.MoveNext
        If rstWork.EOF Then rstWork.MovePrevious
    End If
DelClear:
    showdata
    Exit Sub
DelError:
    ErrorMessages (NODELETE)
    On Error GoTo 0
    Resume DelClear
End Sub

Private Sub CmdEdit_Click()
If cmdEdit.Caption = "&Edit" Then
    If rstWork.EOF Then
       MsgBox ("Empty Table")
       Exit Sub
   End If
   CmdDisab
   FldEnab
   EditFlag = True
   cmdSave.Enabled = True
   rstWork.Edit
   xDesc.SetFocus
   cmdEdit.Caption = "&Cancel"
   xCode.Enabled = False
Else
   cmdEdit.Caption = "&Edit"
   xCode.Enabled = True
   CmdEnab
   EditFlag = False
   cmdSave.Enabled = False
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
    If rstWork.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstWork.EOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEPAST)
        rstWork.MoveLast
    Else
        rstWork.MoveNext
        If rstWork.EOF Then
            rstWork.MoveLast
        End If
    End If
NextClear:
    showdata
    Exit Sub
NextError:
    ErrorMessages (NOMOVEPAST)
    rstWork.Requery
    rstWork.MoveLast
    On Error GoTo 0
    Resume NextClear

End Sub

Private Sub CmdPrev_Click()
  Dim flag As Integer
    
    On Error GoTo PrevError
    If rstWork.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstWork.BOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEFRONT)
        rstWork.MoveFirst
    Else
        rstWork.MovePrevious
        If rstWork.BOF Then
            rstWork.MoveFirst
        End If
    End If
PrevClear:
    showdata
    Exit Sub
PrevError:
    ErrorMessages (NOMOVEFRONT)
    On Error GoTo 0
    Resume PrevClear
End Sub

Private Sub CmdSave_Click()
  Dim i As Integer
    On Error GoTo PError
    
    If EditFlag = True Then
        rstWork.Edit
    End If
    If AddFlag = True Then
        rstWork.AddNew
        rstWork("Code") = "W" & xCode
    End If
    rstWork("desc") = xDesc
    rstWork.Update
    CmdEnab
    AddFlag = False
    cmdAdd.Caption = "&Add"
    cmdEdit.Enabled = True
    cmdSave.Enabled = False
    cmdEdit.Caption = "&Edit"
    xCode.Enabled = True
    rstWork.MoveLast
    EditFlag = False
    AddFlag = False
    showdata
    FldDisab
PError0:
    Exit Sub
    
PError:
    MsgBox "The error: " + Error$ + " has occurred", vbExclamation, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0
End Sub

Private Sub CmdTop_Click()

    If rstWork.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstWork.MoveFirst
    showdata
End Sub

Private Sub Form_Load()
    Call FormCentreMDI(Me)
    Set wrktemp = DBEngine.Workspaces(0)
    Set db6 = wrktemp.OpenDatabase(DproDBpath, True)
    Set rstWork = db6.OpenRecordset("WorkTab", dbOpenDynaset)
    Dim flag As Integer, mtable As String
    If Not rstWork.EOF Then
       rstWork.MoveFirst
    End If
    showdata
    EditFlag = False
    AddFlag = False
    cmdSave.Enabled = False
    ''cmdDel.Visible = False
    FldDisab
    GenClass.fleLogin mvUserid, "Accessed State Defination", Date, Time
End Sub

Public Sub showdata()
     If Not (rstWork.BOF And rstWork.EOF) Then
        xCode = Right(rstWork("Code"), 3)
        xDesc = rstWork("desc")
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
    cmdTop.Enabled = True
    cmdBott.Enabled = True
    cmdNext.Enabled = True
    cmdPrev.Enabled = True
End Sub

Public Sub CmdDisab()
    cmdAdd.Enabled = False
    cmdDel.Enabled = False
    cmdBrow.Enabled = False
    cmdTop.Enabled = False
    cmdBott.Enabled = False
    cmdNext.Enabled = False
    cmdPrev.Enabled = False

End Sub
Public Sub FldEnab()
    xCode.Enabled = True
    xDesc.Enabled = True
End Sub

Public Sub FldDisab()
    xCode.Enabled = False
    xDesc.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    db6.Close
End Sub
