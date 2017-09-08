VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmDPexpe 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datapro"
   ClientHeight    =   6345
   ClientLeft      =   1500
   ClientTop       =   1005
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7380
   Begin VB.Data datGentab 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Supervisor 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   1680
      Width           =   3615
   End
   Begin VB.ComboBox cmbAppid 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   3495
      Left            =   -60
      TabIndex        =   6
      Top             =   2220
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
      _ExtentY        =   6165
      _StockProps     =   15
      Caption         =   "Details of Experience"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   1
      BevelInner      =   2
      Font3D          =   2
      Alignment       =   6
      Begin VB.CommandButton cmdbott 
         Caption         =   "&Bottom"
         Height          =   495
         Left            =   6240
         TabIndex        =   16
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "&Previous"
         Height          =   495
         Left            =   6240
         TabIndex        =   17
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "&Next"
         Height          =   495
         Left            =   6240
         TabIndex        =   18
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdtop 
         Caption         =   "&Top"
         Height          =   495
         Left            =   6240
         TabIndex        =   19
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox details 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   240
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   600
         Width           =   5895
      End
   End
   Begin MSMask.MaskEdBox eDate 
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   327680
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox sdate 
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   327680
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   5640
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
      _ExtentY        =   1296
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
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   2820
         TabIndex        =   12
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1980
         TabIndex        =   13
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmddel 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   1140
         TabIndex        =   14
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   300
         TabIndex        =   15
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   6300
         TabIndex        =   11
         Top             =   180
         Width           =   975
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Supervisor:"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "Personal ID:"
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
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Date Employed:"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Date Left:"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   7560
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "FrmDPexpe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db6 As Database, wrktemp As Workspace
Dim rstJobExpe As Recordset
Dim rstGenRecs As Recordset
Dim mFlag As Boolean
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
   If rstJobExpe.BOF And rstJobExpe.EOF Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
    i = DeleteCheck()
    If i = vbYes Then
        rstJobExpe.Delete
        rstJobExpe.Requery
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
           rstJobExpe.AddNew
           strpos = InStr(1, cmbAppid, ",", 1)
           rstJobExpe("appid") = Left(cmbAppid, strpos - 1)
        Else
           rstJobExpe.Edit
        End If
        rstJobExpe("sdate") = CDate(SDate)
        rstJobExpe("edate") = CDate(Edate)
        rstJobExpe("supervisor") = Supervisor
        rstJobExpe("details") = details
        rstJobExpe.Update
        rstJobExpe.MoveLast
        tempBuff = Left(cmbAppid, strpos - 1)
        tempBuff = "appid = '" & tempBuff & "'"
        rstGenRecs.FindFirst tempBuff
        rstGenRecs.Edit
        rstGenRecs("chk6") = 1
        rstGenRecs.Update
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
    If rstJobExpe.EOF And rstJobExpe.BOF Then
       MsgBox ("Empty Table")
       Exit Sub
   End If
   FldEnab
   mFlag = False
   cmdsave.Enabled = True
   SDate.SetFocus
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

Private Sub CmdTop_Click()
    If (rstJobExpe.BOF And rstJobExpe.EOF) Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstJobExpe.MoveFirst
    tempBuff = "appid = '" & rstJobExpe("appid") & "'"
    rstGenRecs.FindFirst tempBuff
    If rstGenRecs.NOMATCH Then
        ClearData
    Else
       cmbAppid = rstGenRecs("appid") + ", " + rstGenRecs("fullName")
       showdata
    End If
End Sub

Private Sub Form_Load()
     GenClass.fleLogin mvUserid, "Accessed Club and Subscriptions", Date, Time
     Call FormCentreMDI(Me)
     Set wrktemp = DBEngine.Workspaces(0)
     Set db6 = wrktemp.OpenDatabase(DproDBpath, True)
     Set rstJobExpe = db6.OpenRecordset("jobexpe", dbOpenDynaset)
     Set rstGenRecs = db6.OpenRecordset("Genrecs", dbOpenDynaset)
     If Not rstGenRecs.EOF And rstGenRecs.BOF Then
        rstGenRecs.MoveFirst
     End If
      ''__________________________________
     datGentab.DatabaseName = DproDBpath
     datGentab.RecordSource = "genrecs"
     datGentab.ReadOnly = False
     datGentab.Exclusive = False
     datGentab.Refresh
     Do While Not datGentab.Recordset.EOF
       cmbAppid.AddItem datGentab.Recordset("appid") + ", " + datGentab.Recordset("fullName")
       datGentab.Recordset.MoveNext
     Loop
     'cmbAppid.ListIndex = 0
     mFlag = False
     showdata
     FldDisab
     cmdsave.Enabled = False
     cmbAppid.Enabled = True
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Public Sub showdata()
On Error GoTo PError
   If Not (rstJobExpe.EOF And rstJobExpe.BOF) Then
        Edate = rstJobExpe("edate")
        SDate = rstJobExpe("sdate")
        Supervisor = rstJobExpe("supervisor")
        details = rstJobExpe("details")
    End If
PError0:
    Exit Sub
PError:
    MsgBox "Invalid Input", vbCritical, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0
End Sub

Public Sub ClearData()
        Edate = Date
        SDate = SDate
        Supervisor = ""
        details = ""
End Sub
Private Sub cmbAppid_Click()
On Error GoTo perror1
    strpos = InStr(1, cmbAppid, ",", 1)
    tempBuff = "appid = '" & Left(cmbAppid, strpos - 1) & "'"
    rstJobExpe.FindFirst tempBuff
    If rstJobExpe.NOMATCH Then
        ClearData
    Else
        showdata
    End If
perror1:
End Sub
Public Sub CmdEnab()
    cmdedit.Enabled = True
    cmdadd.Enabled = True
    cmddel.Enabled = True
End Sub

Public Sub CmdDisab()
    cmdedit.Enabled = False
    cmdadd.Enabled = False
    cmddel.Enabled = False
End Sub

Public Sub FldEnab()
    Edate.Enabled = True
    SDate.Enabled = True
    Supervisor.Enabled = True
    details.Enabled = True
End Sub

Public Sub FldDisab()
    Edate.Enabled = False
    SDate.Enabled = False
    Supervisor.Enabled = False
    details.Enabled = False
End Sub

Public Sub FldVal()
    If Len(Trim(details)) = 0 Then
       Beep
       MsgBox "Invalid Details", vbExclamation, "Social Records"
       details.SetFocus
       FldChk = False
       Exit Sub
    End If
    If Len(Trim(Supervisor)) = 0 Then
       Beep
       MsgBox "Invalid Supervisor", vbExclamation, "Social Records"
       details.SetFocus
       FldChk = False
       Exit Sub
    End If
    If Len(Trim(Edate)) = 0 Then
       Beep
       MsgBox "Invalid end date", vbExclamation, "Social Records"
       details.SetFocus
       FldChk = False
       Exit Sub
    End If
    If Len(Trim(SDate)) = 0 Then
       Beep
       MsgBox "Invalid start date", vbExclamation, "Social Records"
       details.SetFocus
       FldChk = False
       Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    db6.Close
End Sub
Private Sub CmdNext_Click()
    Dim flag As Integer
    On Error GoTo NextError
    If (rstJobExpe.BOF And rstJobExpe.EOF) Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstJobExpe.EOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEPAST)
        rstJobExpe.MoveLast
    Else
        rstJobExpe.MoveNext
        If rstJobExpe.EOF Then
            rstJobExpe.MoveLast
        End If
    End If
NextClear:
    tempBuff = "appid = '" & rstJobExpe("appid") & "'"
    rstGenRecs.FindFirst tempBuff
    If rstGenRecs.NOMATCH Then
        ClearData
    Else
       cmbAppid = rstGenRecs("appid") + ", " + rstGenRecs("fullName")
       showdata
    End If
    Exit Sub
NextError:
    ErrorMessages (NOMOVEPAST)
    ''rstJobExpe.Requery
    rstJobExpe.MoveLast
    On Error GoTo 0
    Resume NextClear

End Sub

Private Sub CmdPrev_Click()
 Dim flag As Integer
    On Error GoTo PrevError
    If (rstJobExpe.BOF And rstJobExpe.EOF) Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstJobExpe.BOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEFRONT)
        rstJobExpe.MoveFirst
    Else
        rstJobExpe.MovePrevious
        If rstJobExpe.BOF Then
            rstJobExpe.MoveFirst
        End If
    End If
PrevClear:
    tempBuff = "appid = '" & rstJobExpe("appid") & "'"
    rstGenRecs.FindFirst tempBuff
    If rstGenRecs.NOMATCH Then
        ClearData
    Else
       cmbAppid = rstGenRecs("appid") + ", " + rstGenRecs("fullName")
       showdata
    End If
    Exit Sub
PrevError:
    ErrorMessages (NOMOVEFRONT)
    rstJobExpe.MoveFirst
    On Error GoTo 0
    Resume PrevClear
End Sub
Private Sub CmdBott_Click()
   Dim Count As Long
    If (rstJobExpe.BOF And rstJobExpe.EOF) Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstJobExpe.MoveLast
    tempBuff = "appid = '" & rstJobExpe("appid") & "'"
    rstGenRecs.FindFirst tempBuff
    If rstGenRecs.NOMATCH Then
        ClearData
    Else
       cmbAppid = rstGenRecs("appid") + ", " + rstGenRecs("fullName")
       showdata
    End If
End Sub
