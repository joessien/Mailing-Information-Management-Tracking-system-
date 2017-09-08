VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmTerms 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DataPro"
   ClientHeight    =   7080
   ClientLeft      =   1755
   ClientTop       =   1395
   ClientWidth     =   8460
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   8460
   Begin VB.CheckBox chkPromTerm 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2640
      TabIndex        =   31
      Top             =   2880
      Width           =   255
   End
   Begin VB.ComboBox cmbTerm 
      Height          =   315
      ItemData        =   "FrmTerms.frx":0000
      Left            =   2640
      List            =   "FrmTerms.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CheckBox ChkTerm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Check1"
      Height          =   255
      Left            =   2640
      TabIndex        =   28
      Top             =   2520
      Width           =   255
   End
   Begin Threed.SSPanel Cmsg 
      Height          =   375
      Left            =   720
      TabIndex        =   24
      Top             =   5760
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   661
      _StockProps     =   15
      BackColor       =   65280
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   11.99
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1935
      Left            =   720
      TabIndex        =   18
      Top             =   3480
      Width           =   5220
      _Version        =   65536
      _ExtentX        =   9208
      _ExtentY        =   3413
      _StockProps     =   15
      Caption         =   "Term Duration"
      ForeColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
      Alignment       =   6
      Begin Threed.SSCommand cmdChk 
         Height          =   855
         Left            =   4080
         TabIndex        =   23
         Top             =   480
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "Check"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
      End
      Begin MSMask.MaskEdBox Stdate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   21
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Endate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   22
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label duratioN 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2160
         TabIndex        =   26
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Duration:"
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
         Left            =   840
         TabIndex        =   25
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Start of Term:"
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
         Left            =   480
         TabIndex        =   20
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "End of Term:"
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
         Left            =   480
         TabIndex        =   19
         Top             =   1080
         Width           =   1455
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   6240
      Width           =   8415
      _Version        =   65536
      _ExtentX        =   14843
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
         Left            =   7080
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
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2640
      MaxLength       =   40
      TabIndex        =   1
      Top             =   1440
      Width           =   3615
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
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   0
      ToolTipText     =   "Maximum 4 Characters"
      Top             =   960
      Width           =   1035
   End
   Begin Threed.SSPanel SSPanel5 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8415
      _Version        =   65536
      _ExtentX        =   14843
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "Terms Definition "
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
      Height          =   5295
      Left            =   6720
      TabIndex        =   6
      Top             =   840
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   9340
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
      Begin VB.CommandButton cmdbrow 
         Caption         =   "Bro&wse"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdbott 
         Caption         =   "&Bottom"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdtop 
         Caption         =   "&Top"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Promotion Term:"
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
      Left            =   840
      TabIndex        =   32
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CummulativeTerm:"
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
      Left            =   600
      TabIndex        =   30
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enter Current Term:"
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
      Left            =   600
      TabIndex        =   27
      Top             =   2520
      Width           =   2055
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
      Left            =   1320
      TabIndex        =   5
      Top             =   1560
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
      Left            =   1920
      TabIndex        =   4
      Top             =   1080
      Width           =   615
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
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
End
Attribute VB_Name = "FrmTerms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EditFlag As Boolean, MaxErrors As Integer
Dim rstDefTab As Recordset, AddFlag As Boolean
Dim db1 As Database, wrktemp As Workspace
Dim AType As String, ValBuff As String
Dim firstTime As Integer, FirstPass As Integer
Dim GenClass As New DProLog, EFlag As Boolean
Dim RSTGen, rstReg As Recordset


Private Sub CmdAdd_Click()
 If cmdadd.Caption = "&Add" Then
    CmdDisab
    AddFlag = True
    FldEnab
    ClearData
    RCode.SetFocus
    cmdsave.Enabled = False
    cmdedit.Enabled = False
    cmdadd.Enabled = True
    cmdChk.Enabled = True
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
   cmdChk.Enabled = False
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
       Screen.MousePointer = vbHourglass
       mvBuff = "Select * From defterms Order By stdate;"
       Set RSTGen = db1.OpenRecordset(mvBuff, dbOpenSnapshot)
       Set GenView.datGeneral.Recordset = RSTGen
       GenView.Caption = Me.Caption
       GenView.Show
       Screen.MousePointer = vbDefault
    End Sub

Private Sub cmdChk_Click()
'On Error GoTo perror
   EFlag = False
   cmdsave.Enabled = False
   rstDefTab.Requery
   If Stdate = "" Or Endate = "" Then
      MsgBox "Invalid Dates: Dates cannoot be blank", vbCritical, "School Terms"
      Cmsg.BackColor = &HC0C0FF
      Cmsg.Caption = "Date Range is Invalid, Re-Enter"
      EFlag = True
       Exit Sub
   End If
   If RCode = "" Or RDesc = "" Then
      MsgBox "Invalid code or desscription: Cannot be blank", vbCritical, "School Terms"
      Cmsg.BackColor = &HC0C0FF
      Cmsg.Caption = "Date Range is Invalid, Re-Enter"
      EFlag = True
       Exit Sub
   End If
   If (rstDefTab.EOF And rstDefTab.EOF) Then
      EFlag = False
      cmdsave.Enabled = True
      duratioN.Caption = CDate(Endate) - CDate(Stdate)
      Cmsg.BackColor = &H80FF80
      Cmsg.Caption = "Date Range is accepted, Proceed to Save"
   Else
      If CDate(Stdate) >= CDate(Endate) Then
         MsgBox "Invalid Date Range: Endate cannot be less than or equal to Start Date", vbCritical, "School Terms"
         Cmsg.BackColor = &HC0C0FF
         Cmsg.Caption = "Date Range is Invalid, Re-Enter"
         EFlag = True
         Exit Sub
      End If
      'If Year(CDate(Stdate)) <> Year(CDate(Endate)) Then
      '   MsgBox "Invalid Dates: The academic session must be within the same year", vbCritical, "School Terms"
      '   Cmsg.BackColor = &HC0C0FF
      '   Cmsg.Caption = "Date is Invalid, Re-Enter"
      '   EFlag = True
      '   Exit Sub
      'End If
      rstDefTab.MoveFirst
      Do While Not rstDefTab.EOF
         'If EditFlag = False And rstDefTab!code <> RCode Then
            If CDate(Stdate) >= rstDefTab!Stdate And CDate(Stdate) <= rstDefTab!Endate Then
               MsgBox "Invalid Start Date: Date is already nested within " & rstDefTab!Desc, vbCritical, "School Terms"
               Cmsg.BackColor = &HC0C0FF
               Cmsg.Caption = "Date is Invalid, Re-Enter"
               EFlag = True
               Exit Sub
            End If
         'End If
         If CDate(Endate) >= rstDefTab!Stdate And CDate(Endate) <= rstDefTab!Endate Then
            MsgBox "Invalid End Date: date is already nested within " & rstDefTab!Desc, vbCritical, "School Terms"
            Cmsg.BackColor = &HC0C0FF
            Cmsg.Caption = "Date is Invalid, Re-Enter"
            EFlag = True
            Exit Sub
         End If
         If UCase(RCode) = rstDefTab!code And AddFlag = True Then
            MsgBox "Invalid term: Term is already defined", vbCritical, "School Terms"
            Cmsg.BackColor = &HC0C0FF
            Cmsg.Caption = "Invalid Grade, Re-Enter"
            EFlag = True
            Exit Sub
         End If
         If UCase(RDesc) = rstDefTab!Desc And AddFlag = True Then
            MsgBox "Invalid Description: Description already used for " & rstDefTab!Desc, vbCritical, "School Terms"
            Cmsg.BackColor = &HC0C0FF
            Cmsg.Caption = "Invalid Description, Re-Enter"
            EFlag = True
            Exit Sub
         End If
         rstDefTab.MoveNext
      Loop
      If EFlag = True Then
         Cmsg.BackColor = &HC0C0FF
         Cmsg.Caption = "Range is Invalid, Re-Enter"
      Else
         Cmsg.BackColor = &H80FF80
         Cmsg.Caption = "Range is accepted, Proceed to Save"
         duratioN = CDate(Endate) - CDate(Stdate)
         cmdsave.Enabled = True
      End If
   End If
   Exit Sub
PError:
   Cmsg.BackColor = &HC0C0FF
   Cmsg.Caption = "Range is Invalid, Re-Enter"
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
          ClearData
        End If
    End If

End Sub

Private Sub CmdEdit_Click()
 If cmdedit.Caption = "&Edit" Then
     If (rstDefTab.EOF And rstDefTab.EOF) Then
         MsgBox ("Empty Table")
         Exit Sub
     Else
         If rstDefTab.EOF Then
            rstDefTab.MovePrevious
         End If
     End If
   CmdDisab
   FldEnab
   EditFlag = True
   cmdChk.Enabled = True
   cmdsave.Enabled = True
   RCode.Enabled = False
   RDesc.SetFocus
   cmdedit.Caption = "&Cancel"
Else
   cmdedit.Caption = "&Edit"
   RCode.Enabled = True
   CmdEnab
   EditFlag = False
   cmdsave.Enabled = False
   cmdChk.Enabled = False
   showdata
   FldDisab
End If
End Sub

Private Sub cmdexit_Click()
    If IsNull(rstReg!cTerm) Or rstReg("Cterm") = "" Then
       MsgBox "No current Term is specified, Please select one term and specify as current term", "Critical error"
    Else
       Unload Me
    End If
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
    'On Error GoTo PError
    If EditFlag = True Then
        GTBuff = "code = '" & RCode & "'"
        rstDefTab.FindFirst GTBuff
        rstDefTab.Edit
    End If
    If AddFlag = True Then
        rstDefTab.AddNew
        rstDefTab("Code") = UCase(RCode)
    End If
    rstDefTab("Desc") = UCase(RDesc)
    rstDefTab("cumterm") = UCase(cmbTerm)
    rstDefTab("StDate") = UCase(Stdate)
    rstDefTab("EnDate") = UCase(Endate)
    rstDefTab("duration") = CDate(Endate) - CDate(Stdate)
    rstDefTab.Update
    If ChkTerm = 1 Then
       rstReg.Edit
       rstReg("Cterm") = UCase(RCode)
       rstReg.Update
    End If
    If chkPromTerm = 1 Then
       rstReg.Edit
       rstReg("promterm") = UCase(RCode)
       rstReg.Update
    End If
    cmdadd.Caption = "&Add"
    CmdEnab
    cmdedit.Enabled = True
    cmdsave.Enabled = False
    cmdChk.Enabled = False
    cmdedit.Caption = "&Edit"
    RCode.Enabled = True
    If AddFlag = True Then rstDefTab.MoveLast
    AddFlag = False
    showdata
    Cmsg = ""
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
    GenClass.fleLogin mvUserid, "Accessed Terms Definition", Date, Time
    Call FormCentreMDI(Me)
    Set wrktemp = DBEngine.Workspaces(0)
    Set db1 = wrktemp.OpenDatabase(DBpath, True)
    Set rstDefTab = db1.OpenRecordset("Defterms", dbOpenDynaset)
    Set rstReg = db1.OpenRecordset("DProreg", dbOpenDynaset)
    Me.Caption = MVCoyname
    If Not rstDefTab.EOF Then
       rstDefTab.MoveFirst
    End If
    EditFlag = False
    AddFlag = False
    ChkTerm.Value = False
    chkPromTerm.Value = False
    showdata
    FldDisab
    cmdChk.Enabled = False
    cmddel.Visible = True
    cmdsave.Enabled = False
    EFlag = False
    Me.Caption = MVCoyname
End Sub

Public Sub showdata()
     
     If rstDefTab.RecordCount = 1 Then
        rstDefTab.MoveFirst
     End If
     If rstDefTab.BOF And rstDefTab.EOF Then
        DoEvents
     Else
        RCode = rstDefTab("Code")
        RDesc = rstDefTab("Desc")
        Stdate = rstDefTab("StDate")
        Endate = rstDefTab("EnDate")
        cmbTerm = rstDefTab("cumterm")
        duratioN = rstDefTab("duration")
        If rstReg("Cterm") = rstDefTab("Code") Then
           ChkTerm.Value = 1
        Else
            ChkTerm.Value = 0
        End If
        If rstReg("promterm") = rstDefTab("Code") Then
           chkPromTerm.Value = 1
        Else
            chkPromTerm.Value = 0
        End If
    End If

End Sub

Public Sub ClearData()
    RCode = ""
    RDesc = ""
    Stdate = Date
    Endate = Date + 90
    Cmsg = ""
    duratioN = 0
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
    Stdate.Enabled = True
    Endate.Enabled = True
    ChkTerm.Enabled = True
    cmbTerm.Enabled = True
    chkPromTerm.Enabled = True
End Sub

Private Sub FldDisab()
    RCode.Enabled = False
    RDesc.Enabled = False
    Stdate.Enabled = False
    Endate.Enabled = False
    ChkTerm.Enabled = False
    chkPromTerm.Enabled = False
    cmbTerm.Enabled = False
    chkPromTerm.Enabled = False
End Sub



