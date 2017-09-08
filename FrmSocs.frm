VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmDPSocs 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datapro"
   ClientHeight    =   3420
   ClientLeft      =   1500
   ClientTop       =   1485
   ClientWidth     =   7890
   DrawWidth       =   3
   ForeColor       =   &H000040C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3420
   ScaleWidth      =   7890
   Begin Threed.SSPanel SSPanel3 
      Height          =   975
      Left            =   720
      TabIndex        =   19
      Top             =   0
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8070
      _ExtentY        =   1720
      _StockProps     =   15
      Caption         =   "Club Subscription"
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
   Begin VB.Data datGenTab 
      Caption         =   "General table"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cmbAppid 
      Height          =   315
      Left            =   2280
      TabIndex        =   18
      Top             =   1200
      Width           =   3015
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   3375
      Left            =   5400
      TabIndex        =   6
      Top             =   0
      Width           =   2415
      _Version        =   65536
      _ExtentX        =   4260
      _ExtentY        =   5953
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
      BevelOuter      =   1
      BevelInner      =   2
      Begin VB.CommandButton cmdprev 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton cmdbott 
         Caption         =   "&Bottom"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton cmdtop 
         Caption         =   "&Top"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   855
      End
      Begin VB.Image StaffPic 
         BorderStyle     =   1  'Fixed Single
         Height          =   1815
         Left            =   240
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1935
      End
      Begin VB.Label fullname 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Width           =   1935
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   2520
      Width           =   5415
      _Version        =   65536
      _ExtentX        =   9551
      _ExtentY        =   1508
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
      Begin VB.CommandButton cmdbrow 
         Caption         =   "Bro&wse"
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   4440
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   3600
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmddel 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   960
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
   End
   Begin MSMask.MaskEdBox socDate 
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   1680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   327680
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin VB.TextBox SocName 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   2040
      Width           =   3735
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   975
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   1720
      _StockProps     =   15
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
         Picture         =   "FrmSocs.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Club Names:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Staff Identification"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
End
Attribute VB_Name = "frmDPSocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db6 As Database, wrktemp As Workspace
Dim rstSocData As Recordset
Dim rstGenrecs As Recordset
Dim PicName As String, mFlag As Boolean
Dim FldChk As Boolean
Dim tempBuff As String, strpos As Integer
Dim GenClass As New DProLog

Private Sub CmdAdd_Click()
 If cmdadd.Caption = "&Add" Then
   ClearData
   mFlag = True
   cmbAppid.SetFocus
   FldEnab
   CmdDisab
   cmdsave.Enabled = True
   cmdadd.Enabled = True
   cmdadd.Caption = "&Cancel"
Else
   CmdEnab
   cmdadd.Caption = "&Add"
   cmdsave.Enabled = False
   mFlag = False
End If
End Sub
Private Sub CmdDel_Click()
   Dim i As Integer
   If rstSocData.BOF And rstSocData.EOF Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
    i = DeleteCheck()
    If i = vbYes Then
        rstSocData.Delete
        rstSocData.Requery
        ClearData
    End If
End Sub

Private Sub CmdSave_Click()
    Dim i As Integer
    On Error GoTo PError
      FldChk = True
      FldVal
      If FldChk = False Then
        Exit Sub
      End If
      If mFlag = True Then
           rstSocData.AddNew
           rstSocData("appid") = Left(cmbAppid, strpos - 1)
           rstSocData("rind") = "X"
        Else
           rstSocData.Edit
        End If
        rstSocData("SoCdate") = socDate
        rstSocData("SoCName") = SocName
        rstSocData.Update
        tempBuff = Left(cmbAppid, strpos - 1)
        tempBuff = "appid = '" & tempBuff & "'"
        rstGenrecs.FindFirst tempBuff
        rstGenrecs.Edit
        rstGenrecs("chk8") = 1
        rstGenrecs.Update
        cmdedit.Caption = "&Edit"
        cmdadd.Caption = "&Add"
        cmdsave.Enabled = False
        cmbAppid.Enabled = True
        CmdEnab
        mFlag = False
        FldDisab
PError0:
    Exit Sub
PError:
    MsgBox "Invalid Input", vbCritical, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0
End Sub

Private Sub CmdEdit_Click()
 If cmdedit.Caption = "&Edit" Then
    If rstSocData.EOF And rstSocData.BOF Then
       MsgBox ("Empty Table")
       Exit Sub
   End If
   FldEnab
   mFlag = False
   cmdsave.Enabled = True
   socDate.SetFocus
   cmdedit.Caption = "&Cancel"
   cmbAppid.Enabled = False
   CmdDisab
   cmdedit.Enabled = True
Else
   cmdedit.Caption = "&Edit"
   cmbAppid.Enabled = True
   cmdsave.Enabled = False
   CmdEnab
   FldDisab
End If
End Sub

Private Sub Form_Load()
     GenClass.fleLogin mvUserid, "Accessed Club and Subscriptions", Date, Time
     Call FormCentreMDI(Me)
     Set wrktemp = DBEngine.Workspaces(0)
     Set db6 = wrktemp.OpenDatabase(DproDBpath, True)
     Set rstSocData = db6.OpenRecordset("SosData", dbOpenDynaset)
     Set rstGenrecs = db6.OpenRecordset("Genrecs", dbOpenDynaset)
     If Not rstGenrecs.EOF And rstGenrecs.BOF Then
        rstGenrecs.MoveFirst
     End If
      ''__________________________________
     datGenTab.DatabaseName = DproDBpath
     datGenTab.RecordSource = "genrecs"
     datGenTab.ReadOnly = False
     datGenTab.Exclusive = False
     datGenTab.Refresh
     Do While Not datGenTab.Recordset.EOF
       cmbAppid.AddItem datGenTab.Recordset("appid") + ", " + datGenTab.Recordset("fullName")
       datGenTab.Recordset.MoveNext
     Loop
     'cmbAppid.ListIndex = 0
     mFlag = False
     showdata
     FldDisab
     cmdsave.Enabled = False
     cmbAppid.Enabled = True
End Sub

Private Sub CmdBrow_Click()
    Dim GenView As New frmBrowse
    Set GenView.datGeneral.Recordset = rstSocData
    GenView.Caption = Me.Caption & " - Records View"
    GenView.Show

End Sub
Private Sub cmdexit_Click()
    Unload Me
End Sub

Public Sub showdata()
On Error GoTo PError
   If Not (rstSocData.EOF And rstSocData.BOF) Then
        tempBuff = "appid = '" & rstSocData("appid") & "'"
        rstGenrecs.FindFirst tempBuff
        cmbAppid = rstSocData!Appid & ", " & rstGenrecs!fullname
        PicName = rstGenrecs("imgpic")
        fullname = rstGenrecs("fullname")
        StaffPic = LoadPicture(PicName)
        socDate = rstSocData("SoCdate")
        SocName = rstSocData("SoCname")
    End If
PError0:
    Exit Sub
PError:
    MsgBox "Invalid Input", vbCritical, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0
End Sub

Public Sub ClearData()
        socDate = Date
        SocName = "None"
End Sub
Private Sub cmbAppid_Click()
On Error GoTo perror1
    strpos = InStr(1, cmbAppid, ",", 1)
    tempBuff = "appid = '" & Left(cmbAppid, strpos - 1) & "'"
    rstSocData.FindFirst tempBuff
    If rstSocData.NOMATCH Then
        ClearData
        rstGenrecs.FindFirst tempBuff
        fullname = rstGenrecs("fullname")
        PicName = rstGenrecs("imgpic")
        StaffPic = LoadPicture(PicName)
    Else
        showdata
    End If
perror1:
End Sub
Public Sub CmdEnab()
    cmdbrow.Enabled = True
    cmdedit.Enabled = True
    cmdadd.Enabled = True
    cmddel.Enabled = True
    cmdtop.Enabled = True
    cmdbott.Enabled = True
    cmdnext.Enabled = True
    cmdprev.Enabled = True
End Sub

Public Sub CmdDisab()
    cmdedit.Enabled = False
    cmdbrow.Enabled = False
    cmdadd.Enabled = False
    cmddel.Enabled = False
    cmdtop.Enabled = False
    cmdbott.Enabled = False
    cmdnext.Enabled = False
    cmdprev.Enabled = False
End Sub

Public Sub FldEnab()
    socDate.Enabled = True
    SocName.Enabled = True
End Sub

Public Sub FldDisab()
    socDate.Enabled = False
    SocName.Enabled = False
End Sub

Private Sub CmdBott_Click()
    Dim Count As Long
    If rstSocData.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstSocData.MoveLast
    showdata
End Sub
Private Sub CmdNext_Click()
        Dim flag As Integer
    
    On Error GoTo NextError
    If rstSocData.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstSocData.EOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEPAST)
        rstSocData.MoveLast
    Else
        rstSocData.MoveNext
        If rstSocData.EOF Then
            rstSocData.MoveLast
        End If
    End If
NextClear:
    showdata
    Exit Sub
NextError:
    ErrorMessages (NOMOVEPAST)
    rstSocData.Requery
    rstSocData.MoveLast
    On Error GoTo 0
    Resume NextClear

End Sub

Private Sub CmdPrev_Click()
  Dim flag As Integer
    
    On Error GoTo PrevError
    If rstSocData.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstSocData.BOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEFRONT)
        rstSocData.MoveFirst
    Else
        rstSocData.MovePrevious
        If rstSocData.BOF Then
            rstSocData.MoveFirst
        End If
    End If
PrevClear:
    showdata
    Exit Sub
PrevError:
    ErrorMessages (NOMOVEFRONT)
    rstSocData.Requery
    rstSocData.MoveFirst
    On Error GoTo 0
    Resume PrevClear
End Sub
Private Sub CmdTop_Click()
    If rstSocData.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstSocData.MoveFirst
    showdata
End Sub

Public Sub FldVal()
    If Len(socDate) = 0 Then
       Beep
       MsgBox "Invalid date", vbExclamation, "Social Records"
       socDate.SetFocus
       FldChk = False
       Exit Sub
    End If
    If Len(SocName) = 0 Then
       Beep
       MsgBox "Invalid Name", vbExclamation, "Social Records"
       SocName.SetFocus
       FldChk = False
       Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    db6.Close
End Sub
