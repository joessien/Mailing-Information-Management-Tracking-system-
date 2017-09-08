VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form FrmsubAlloc 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datapro"
   ClientHeight    =   3315
   ClientLeft      =   1410
   ClientTop       =   1965
   ClientWidth     =   6150
   DrawWidth       =   3
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3315
   ScaleWidth      =   6150
   Begin Threed.SSPanel SSPanel2 
      Height          =   975
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   1720
      _StockProps     =   15
      ForeColor       =   16576
      BackColor       =   11639171
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Font3D          =   2
      Begin VB.Image Image2 
         Height          =   720
         Left            =   120
         Picture         =   "FrmsubAlloc.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
   End
   Begin Threed.SSPanel SSPanel5 
      Height          =   975
      Left            =   720
      TabIndex        =   14
      Top             =   0
      Width           =   5415
      _Version        =   65536
      _ExtentX        =   9551
      _ExtentY        =   1720
      _StockProps     =   15
      Caption         =   "  Sports and Recreational Activities "
      ForeColor       =   16576
      BackColor       =   11639171
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Font3D          =   2
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   1215
      Left            =   0
      TabIndex        =   11
      Top             =   960
      Width           =   4695
      _Version        =   65536
      _ExtentX        =   8281
      _ExtentY        =   2143
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cmbReac 
         Height          =   315
         Left            =   1920
         TabIndex        =   15
         Top             =   720
         Width           =   2655
      End
      Begin VB.ComboBox cmbAppid 
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recreation Event:"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Personal Id:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   2175
      Left            =   4800
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   3836
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdbrow 
         Caption         =   "Bro&wse"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdbott 
         Caption         =   "&Bottom"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdtop 
         Caption         =   "&Top"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   4695
      _Version        =   65536
      _ExtentX        =   8281
      _ExtentY        =   1720
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Data datGentab 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1920
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmsubAlloc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstSports As Recordset, AddFlag As Boolean
Dim db6 As Database, wrktemp As Workspace
Dim EditFlag As Boolean, MaxErrors As Integer
Dim GenClass As New DProLog, strpos As Variant
Dim rstGenRecs As Recordset, tempBuff As Variant
Dim rstRecTab As Recordset

Private Sub CmdAdd_Click()
 If cmdadd.Caption = "&Add" Then
    rstSports.AddNew
    cmbAppid.SetFocus
    cmdsave.Enabled = True
    AddFlag = True
    CmdDisab
    cmdadd.Enabled = True
    cmdadd.Caption = "&Cancel"
Else
    If Not rstSports.EOF Then
       rstSports.MoveFirst
    End If
   AddFlag = False
   cmdadd.Caption = "&Add"
   CmdEnab
   cmdsave.Enabled = False
   showdata
End If
End Sub

Private Sub CmdBott_Click()
   Dim Count As Long
    If rstSports.BOF And rstSports.EOF Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstSports.MoveLast
    showdata
End Sub

Private Sub CmdBrow_Click()
    Dim GenView As New frmBrowse
    Set GenView.datGeneral.Recordset = rstSports
    GenView.Caption = Me.Caption & " - Records View"
    GenView.Show
End Sub

Private Sub CmdDel_Click()
  Dim i As Integer
   If rstSports.RecordCount < 1 Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
    i = DeleteCheck()
    If i = vbYes Then
        rstSports.Delete
        If Not (rstSports.BOF And rstSports.EOF) Then
          rstSports.MoveFirst
          showdata
        End If
    End If
End Sub
Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdNext_Click()
    Dim flag As Integer
    On Error GoTo NextError
    If rstSports.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstSports.EOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEPAST)
        rstSports.MoveLast
    Else
        rstSports.MoveNext
        If rstSports.EOF Then
            rstSports.MoveLast
        End If
    End If
NextClear:
    showdata
    Exit Sub
NextError:
    ErrorMessages (NOMOVEPAST)
    rstSports.MoveLast
    On Error GoTo 0
    Resume NextClear

End Sub

Private Sub CmdPrev_Click()
 Dim flag As Integer
    On Error GoTo PrevError
    If rstSports.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstSports.BOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEFRONT)
        rstSports.MoveFirst
    Else
        rstSports.MovePrevious
        If rstSports.BOF Then
            rstSports.MoveFirst
        End If
    End If
PrevClear:
    showdata
    Exit Sub
PrevError:
    ErrorMessages (NOMOVEFRONT)
    ''rstSports.Requery
    rstSports.MoveFirst
    On Error GoTo 0
    Resume PrevClear

End Sub

Private Sub CmdSave_Click()
  Dim i As Integer
    On Error GoTo PError
    If Len(Trim(cmbAppid)) = 0 Or Len(Trim(cmbAppid)) = 0 Then
        Exit Sub
    End If
    If AddFlag = True Then
        rstSports.AddNew
        strpos = InStr(1, cmbAppid, ",", 1)
        rstSports("appid") = Left(cmbAppid, strpos - 1)
    End If
    If EditFlag = True Then
        rstSports.Edit
    End If
    strpos = InStr(1, cmbReac, ",", 1)
    rstSports("scode") = Left(cmbReac, strpos - 1)
    rstSports.Update
    strpos = InStr(1, cmbAppid, ",", 1)
    tempBuff = Left(cmbAppid, strpos - 1)
    tempBuff = "appid = '" & tempBuff & "'"
    rstGenRecs.FindFirst tempBuff
    rstGenRecs.Edit
    rstGenRecs("chk7") = 1
    rstGenRecs.Update
    AddFlag = False
    EditFlag = False
    cmdadd.Caption = "&Add"
    CmdEnab
    cmdsave.Enabled = False
PError0:
    Exit Sub
PError:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0

End Sub

Private Sub CmdTop_Click()
    If rstSports.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstSports.MoveFirst
    showdata
End Sub

Private Sub Form_Load()
    GenClass.fleLogin mvUserid, "Accessed Organizational Grading", Date, Time
    Call FormCentreMDI(Me)
    Set wrktemp = DBEngine.Workspaces(0)
    Set db6 = wrktemp.OpenDatabase(DproDBpath, True)
    Set rstSports = db6.OpenRecordset("sports", dbOpenDynaset)
    Set rstGenRecs = db6.OpenRecordset("genrecs", dbOpenDynaset)
    Set rstRecTab = db6.OpenRecordset("rectab", dbOpenDynaset)
    Dim flag As Integer, mtable As String
    If Not rstSports.EOF Then
       rstSports.MoveFirst
    End If
    ''__________________________________
    datGentab.DatabaseName = DproDBpath
    datGentab.RecordSource = "GenRecs"
    datGentab.ReadOnly = False
    datGentab.Exclusive = False
    datGentab.Refresh
    Do While Not datGentab.Recordset.EOF
       cmbAppid.AddItem datGentab.Recordset!Appid + ", " + datGentab.Recordset!FName + " " + datGentab.Recordset!Surname
       datGentab.Recordset.MoveNext
    Loop
    'cmbAppid.ListIndex = 0
    ''__________________________________
    datGentab.DatabaseName = DproDBpath
    datGentab.RecordSource = "rectab"
    datGentab.ReadOnly = False
    datGentab.Exclusive = False
    datGentab.Refresh
    Do While Not datGentab.Recordset.EOF
       cmbReac.AddItem datGentab.Recordset.code + ", " + datGentab.Recordset.Desc
       datGentab.Recordset.MoveNext
    Loop
    'cmbreac.ListIndex = 0
    showdata
    cmdsave.Enabled = False
End Sub

Public Sub showdata()
     If Not (rstSports.BOF And rstSports.EOF) Then
        If Not (rstGenRecs.BOF And rstGenRecs.EOF) Then
            tempBuff = "appid = '" & rstSports("appid") & "'"
            rstGenRecs.FindFirst tempBuff
            cmbAppid = rstGenRecs!Appid & "," & rstGenRecs!fullname
        End If
        If Not (rstRecTab.BOF And rstRecTab.EOF) Then
            tempBuff = "Code = '" & rstSports("scode") & "'"
            rstRecTab.FindFirst tempBuff
            cmbReac = rstRecTab!code & "," & rstRecTab!Desc
        End If
    End If
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

