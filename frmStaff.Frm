VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form FrmStaff 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Human Resource"
   ClientHeight    =   6630
   ClientLeft      =   1305
   ClientTop       =   1380
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   8715
   Begin VB.TextBox HmAddr1 
      Height          =   1245
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   36
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox PtelCon 
      Height          =   285
      Left            =   1320
      TabIndex        =   33
      Top             =   3120
      Width           =   1935
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   975
      Left            =   -240
      TabIndex        =   29
      Top             =   4920
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   1720
      _StockProps     =   15
      Caption         =   "Records References"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
      Alignment       =   6
      Begin VB.CommandButton cmdReltns 
         Caption         =   "Re&Lations"
         Height          =   375
         Left            =   3000
         TabIndex        =   32
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdEduc 
         Caption         =   "&Education"
         Height          =   375
         Left            =   1920
         TabIndex        =   31
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdref 
         Caption         =   "&Referees"
         Height          =   375
         Left            =   840
         TabIndex        =   30
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.ComboBox cmbStaffNo 
      Height          =   315
      Left            =   2160
      TabIndex        =   28
      Top             =   840
      Width           =   2775
   End
   Begin VB.ComboBox cmbState 
      Height          =   315
      Left            =   1320
      TabIndex        =   26
      Top             =   2760
      Width           =   1935
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   5175
      Left            =   5040
      TabIndex        =   10
      Top             =   720
      Width           =   3615
      _Version        =   65536
      _ExtentX        =   6376
      _ExtentY        =   9128
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
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sex:"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Age:"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Date of Birth:"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Place of Birth:"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nationality:"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Sex 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label ApplAge 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Birth 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Pbirth 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Nation 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label StaffNames 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   2400
         Width           =   3135
      End
      Begin VB.Image StaffPic 
         BorderStyle     =   1  'Fixed Single
         Height          =   1935
         Left            =   840
         Picture         =   "frmStaff.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.TextBox Surname 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox FName 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox OthNames 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.ComboBox mstatus 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00004000&
      Height          =   315
      ItemData        =   "frmStaff.frx":0542
      Left            =   1320
      List            =   "frmStaff.frx":054F
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Data datGenTab 
      Caption         =   "General Data"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   705
      Index           =   0
      Left            =   -240
      TabIndex        =   4
      Top             =   5880
      Width           =   8895
      _Version        =   65536
      _ExtentX        =   15690
      _ExtentY        =   1244
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
      BorderWidth     =   1
      BevelOuter      =   0
      BevelInner      =   2
      Begin VB.CommandButton cmdexit 
         Caption         =   "E&xit"
         Height          =   435
         Left            =   6480
         TabIndex        =   25
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdbrow 
         Caption         =   "&Browse"
         Height          =   435
         Left            =   3240
         TabIndex        =   24
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&save"
         Height          =   435
         Left            =   2160
         TabIndex        =   23
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdupd 
         Caption         =   "&Update"
         Height          =   435
         Left            =   1080
         TabIndex        =   22
         Top             =   120
         Width           =   1095
      End
   End
   Begin Threed.SSPanel SSPanel5 
      Height          =   615
      Left            =   600
      TabIndex        =   37
      Top             =   0
      Width           =   8055
      _Version        =   65536
      _ExtentX        =   14208
      _ExtentY        =   1085
      _StockProps     =   15
      Caption         =   "Staff Records"
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
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   615
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
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
   End
   Begin VB.Label Label6 
      Caption         =   "Home Address:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Telephone:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   34
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "State of Origin:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "Staff Identification:"
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
      TabIndex        =   9
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Surname:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "First Name:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Other Names:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Marital Status:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "FrmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim db1 As Database, wrktemp As Workspace
Dim GenData As Recordset, StTab As Recordset
Dim StidTab As Recordset, FVal As Boolean
Dim PicName As String, MAppid As String
Dim EditFlag As Boolean, MaxErrors As Integer
Dim PayTval As Integer, staffBuff As String
Dim strpos As Integer
Dim FirstTime As Integer, FirstPass As Integer
Dim GenClass As New QLog

Private Sub cmbStaffNo_LostFocus()
On Error GoTo pError
   strpos = InStr(1, cmbStaffNo, ",", 1)
   If strpos = 0 Then
     staffBuff = "StaffNo = '" & UCase(cmbStaffNo) & "'"
   Else
     staffBuff = "StaffNo = '" & Left(cmbStaffNo, strpos - 1) & "'"
   End If
    StidTab.FindFirst staffBuff
    If StidTab.NOMATCH Then
       Exit Sub
    Else
        cmbStaffNo = StidTab!StaffNo & ", " & StidTab!StaffNames
        showdata
    End If
pError:
End Sub

Private Sub cmdEduc_Click()
    FrmQual.Show
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdRef_Click()
On Error GoTo pError
    frmRefre.AppLbl.Caption = "Staff Id:"
    frmRefre.Show
    frmRefre.AppLbl.Caption = "File Number:"
pError:
End Sub

Private Sub cmdReltns_Click()
    FrmRelTn.Show
End Sub

Private Sub CmdSave_Click()
    Dim i As Integer
    On Error GoTo pError
    FVal = True
    FldVal
    If FVal = False Then
        Exit Sub
    End If
    If EditFlag = True Then
        GenData.Edit
        GenData("Surname") = Surname
        GenData("FName") = FName
        GenData("OthNames") = OthNames
        GenData("Ptelcon") = PtelCon
        GenData("HmAddr1") = HmAddr1
        GenData("MStatus") = mstatus
        strpos = InStr(1, cmbState, ",", 1)
        GenData("State") = Left(cmbState, strpos - 1)
        GenData.Update
        strpos = InStr(1, cmbStaffNo, ",", 1)
        staffBuff = "StaffNo = '" & Left(cmbStaffNo, strpos - 1) & "'"
        StidTab.FindFirst staffBuff
        StidTab.Edit
        StidTab("staffnames") = Surname & " " & FName
        StidTab.Update
        cmdupd.Caption = "&Update"
        showdata
        FldDisab
        CmdEnab
        EditFlag = False
        cmdsave.Enabled = False
     End If
        
PError0:
    Exit Sub
    
pError:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0

End Sub

Private Sub CmdUpd_Click()
 On Error GoTo pError
 If GenData("Appstat") = "Terminated" Then
    MsgBox "Staff Appointment has been terminated", vbInformation, "Staff Records"
    Exit Sub
 End If
 If cmdupd.Caption = "&Update" Then
    If GenData.EOF Then
       MsgBox ("Empty Table")
       Exit Sub
    End If
   FldEnab
   EditFlag = True
   cmdsave.Enabled = True
   GenData.Edit
   StidTab.Edit
   Surname.SetFocus
   cmdupd.Caption = "&Cancel"
   CmdDisab
Else
   FldDisab
   cmdupd.Caption = "&Update"
   CmdEnab
   EditFlag = False
   cmdsave.Enabled = False
End If
pError:
End Sub

Private Sub Form_Load()
On Error GoTo pError
     GenClass.fleLogin mvUserid, "Accessed Staff Records", Date, Time
     Call FormCentreMDI(Me)
     Set wrktemp = DBEngine.Workspaces(0)
     Set db1 = wrktemp.OpenDatabase(Dbpath, True)
     Set GenData = db1.OpenRecordset("Genrecs", dbOpenDynaset)
     Set StidTab = db1.OpenRecordset("CarManP", dbOpenDynaset)
     Set StTab = db1.OpenRecordset("StateTab", dbOpenDynaset)
     Me.Caption = mVCoyName
    ''__________________________________
    datGenTab.DatabaseName = Dbpath
    datGenTab.RecordSource = "StateTab"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       cmbState.AddItem datGenTab.Recordset("Statecode") + ", " + datGenTab.Recordset("StateName")
       datGenTab.Recordset.MoveNext
    Loop
    cmbState.ListIndex = 0
    ''__________________________________
    datGenTab.DatabaseName = Dbpath
    datGenTab.RecordSource = "carmanp"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       cmbStaffNo.AddItem datGenTab.Recordset("Staffno") + ", " + datGenTab.Recordset("StaffNames")
       datGenTab.Recordset.MoveNext
    Loop
    cmbStaffNo.ListIndex = 0
    '______________________________________
       Dim flag As Integer, mtable As String
         If Not GenData.EOF Then
            GenData.MoveFirst
         End If
         If Not StidTab.EOF Then
            StidTab.MoveFirst
         End If
    EditFlag = False
    cmdsave.Enabled = False
    showdata
    FldDisab
pError:
End Sub

Private Sub CmdBrow_Click()
On Error GoTo pError
    mvSql = "Select AppId, Appldate, posappl, surname, fname, othnames,sex,mstatus,birth, pbirth,age,ptelcon From Genrecs"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    Set frmBrowse.datGeneral.Recordset = rst
    frmBrowse.Caption = Me.Caption & " - Records View"
    frmBrowse.Show
pError:
End Sub


Public Sub showdata()
On Error GoTo perror1
     Dim StBuff As String
     If Not (GenData.BOF And GenData.EOF) Then
        strpos = InStr(1, cmbStaffNo, ",", 1)
        staffBuff = "StaffNo = '" & Left(cmbStaffNo, strpos - 1) & "'"
        StidTab.FindFirst staffBuff
        PicName = StidTab("staffpic")
        staffBuff = "AppId = '" & StidTab("AppId") & "'"
        GenData.FindFirst staffBuff
        Surname = GenData("SurName")
        FName = GenData("FName")
        OthNames = GenData("OthNames")
        Birth = GenData("Birth")
        Pbirth = GenData("PBirth")
        StBuff = "StateCode = '" & GenData("State") & "'"
        StTab.FindFirst StBuff
        cmbState = StTab!StateCode & "," & StTab!StateName
        Nation = GenData("Nation")
        PtelCon = GenData("PTelCon")
        HmAddr1 = GenData("HmAddr1")
        Sex = GenData("Sex")
        mstatus = GenData("MStatus")
        ApplAge = GenData("Age")
        StaffNames = StidTab("staffnames")
        PicName = StidTab("staffpic")
        StaffPic = LoadPicture(PicName)
    End If
    If GenData("Appstat") = "Terminated" Then
       MsgBox "Staff Appointment has been terminated", vbInformation, "Staff Records"
       FldDisab
    End If
perror1:
End Sub

Private Sub cmbStaffNo_Click()
    showdata
End Sub
Public Sub CmdEnab()
    cmdref.Enabled = True
    cmdEduc.Enabled = True
    cmdReltns.Enabled = True
    cmdbrow.Enabled = True
End Sub

Public Sub CmdDisab()
    cmdref.Enabled = False
    cmdbrow.Enabled = False
    cmdEduc.Enabled = False
    cmdReltns.Enabled = False
End Sub
Public Sub FldEnab()
        Surname.Enabled = True
        FName.Enabled = True
        OthNames.Enabled = True
        PtelCon.Enabled = True
        HmAddr1.Enabled = True
        mstatus.Enabled = True
End Sub

Public Sub FldDisab()
        Surname.Enabled = False
        FName.Enabled = False
        OthNames.Enabled = False
        PtelCon.Enabled = False
        HmAddr1.Enabled = False
        mstatus.Enabled = False
End Sub

Public Sub FldVal()
On Error GoTo pError
     If Len(Surname) = 0 Then
        MsgBox "Enter Surname ", vbExclamation, "Staff Records"
        Surname.SetFocus
        FVal = False
        Exit Sub
     End If
     If Len(FName) = 0 Then
        MsgBox "Enter First Name ", vbExclamation, "Staff Records"
        FVal = False
        FName.SetFocus
        Exit Sub
     End If
     If Len(HmAddr1) = 0 Then
        MsgBox "Enter Home Address ", vbExclamation, "Staff Records"
        FVal = False
        HmAddr1.SetFocus
        Exit Sub
     End If
pError:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    db1.Close
End Sub

