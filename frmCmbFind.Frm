VERSION 5.00
Begin VB.Form FrmListings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Tracking System"
   ClientHeight    =   6225
   ClientLeft      =   1305
   ClientTop       =   1380
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8280
   Begin VB.CommandButton Command20 
      Caption         =   "Out Going Mails by Subject"
      Height          =   495
      Left            =   4320
      TabIndex        =   20
      Top             =   4920
      Width           =   3375
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Received Mails by Reference Number"
      Height          =   495
      Left            =   4320
      TabIndex        =   19
      Top             =   4440
      Width           =   3375
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Received Mails by Faculty/Division"
      Height          =   495
      Left            =   4320
      TabIndex        =   18
      Top             =   5400
      Width           =   3375
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Removed Mails by Date"
      Height          =   495
      Left            =   480
      TabIndex        =   17
      Top             =   4920
      Width           =   3375
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Removed Mails by Reference"
      Height          =   495
      Left            =   480
      TabIndex        =   16
      Top             =   4440
      Width           =   3375
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Mails on Hold"
      Height          =   495
      Left            =   480
      TabIndex        =   15
      Top             =   5400
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Received Mails by Faculty/Division"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   3720
      Width           =   3375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Sent Mails by Faculty/Division"
      Height          =   495
      Left            =   4320
      TabIndex        =   8
      Top             =   3720
      Width           =   3375
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Sent Mails by Subject"
      Height          =   495
      Left            =   4320
      TabIndex        =   14
      Top             =   3240
      Width           =   3375
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Sent Mails by Reference Number"
      Height          =   495
      Left            =   4320
      TabIndex        =   13
      Top             =   2760
      Width           =   3375
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Sent Mails by OutDate"
      Height          =   495
      Left            =   4320
      TabIndex        =   12
      Top             =   2280
      Width           =   3375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Sent Mails by Originator"
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   1800
      Width           =   3375
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Sent Mails by Originating Department"
      Height          =   495
      Left            =   4320
      TabIndex        =   11
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Sent Mails by Beneficiary"
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   840
      Width           =   3375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Out Going Mails by Subject"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   3240
      Width           =   3375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Received Mails by Reference Number"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   2760
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Received Mails by In Date"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   2280
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Received Mails by Originator"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   3375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Received Mails by Originating Department"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Received Mails by Beneficiary"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label fhdr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mail Listings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "FrmListings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim db1 As Database, wrktemp As Workspace
Dim FVal As Boolean
Dim PicName As String, MAppid As String
Dim EditFlag As Boolean, MaxErrors As Integer
Dim PayTval As Integer, studBuff As String
Dim strpos As Integer

Dim rstStudmast As Recordset
Dim rstCNames As Recordset
Dim rstSports As Recordset
Dim rstSkills As Recordset
Dim rstTalent As Recordset
Dim rstcSup As Recordset

Dim GenClass As New mProLog
Private Sub cmbStudno_LostFocus()
    StudFind
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'On Error GoTo PError
     GenClass.fleLogin mvUserid, "Accessed Combination Search Enquiry", Date, Time
     Call FormCentreMDI(Me)
     Set wrktemp = DBEngine.Workspaces(0)
     Set db1 = wrktemp.OpenDatabase(DBpath, True)
     Set rstStudmast = db1.OpenRecordset("Studmast", dbOpenDynaset)
     Set rstCNames = db1.OpenRecordset("DefClassNames", dbOpenDynaset)
     Set rstSports = db1.OpenRecordset("DefSports", dbOpenDynaset)
     Set rstSkills = db1.OpenRecordset("Defskill", dbOpenDynaset)
     Set rstTalent = db1.OpenRecordset("Deftalents", dbOpenDynaset)
     Set rstcSup = db1.OpenRecordset("qryclasssup", dbOpenDynaset)
     Me.Caption = MVCoyname
    ''__________________________________
    datgenTab.DatabaseName = DBpath
    datgenTab.RecordSource = "studmast"
    datgenTab.ReadOnly = False
    datgenTab.Exclusive = False
    datgenTab.Refresh
    Do While Not datgenTab.Recordset.EOF
       cmbStudNo.AddItem datgenTab.Recordset("studno") + ", " + datgenTab.Recordset("StudNames")
       datgenTab.Recordset.MoveNext
    Loop
    cmbStudNo.ListIndex = 0
    '______________________________________
    ClearData
PError:
End Sub

Public Sub showdata()
'On Error GoTo perror1
     Dim StBuff As String
     If Not (rstStudmast.BOF And rstStudmast.EOF) Then
        PicName = rstStudmast("studpic")
        StudPic = LoadPicture(PicName)
        StudNo = rstStudmast("studno")
        StudNames = rstStudmast("studnames")
        stFileNo = rstStudmast("stfileno")
        regDate = rstStudmast("regdate")
        Sex = rstStudmast("Sex")
        DOB = rstStudmast("dob")
        Age = rstStudmast("age")
        HMAddr1 = rstStudmast("hmaddr1")
        Guardian = rstStudmast("guardian")
        GDPhone = rstStudmast("gdphone")
        stemail = rstStudmast("stemail")
        
        GTBuff = "Code = '" & rstStudmast("cclass") & "'"
        rstCNames.FindFirst GTBuff
        If rstCNames.NOMATCH Then
           cclass = ""
        Else
           If IsNull(rstCNames!NickName) Then
              cclass = rstCNames!Desc & ", " & rstCNames!CGrp
           Else
              cclass = rstCNames!Desc & ", " & rstCNames!NickName & " - " & rstCNames!CGrp
           End If
        End If
        
        GTBuff = "Code = '" & rstStudmast("skills") & "'"
        rstSkills.FindFirst GTBuff
        If rstSkills.NOMATCH Then
           skills = ""
        Else
           skills = rstSkills!Desc
        End If
        
        
        GTBuff = "Code = '" & rstStudmast("talent") & "'"
        rstTalent.FindFirst GTBuff
        If rstTalent.NOMATCH Then
           talent = ""
        Else
           talent = rstTalent!Desc
        End If
        
        GTBuff = "Code = '" & rstStudmast("sports") & "'"
        rstSports.FindFirst GTBuff
        If rstSports.NOMATCH Then
           sports = ""
        Else
           sports = rstSports!Desc
        End If
        
        GTBuff = "rCode = '" & rstStudmast("cclass") & "'"
        rstcSup.FindFirst GTBuff
        If rstcSup.NOMATCH Then
           CMaster = ""
        Else
           CMaster = rstcSup!StudNames
        End If
        
        
    End If
perror1:
End Sub

Private Sub cmbStudno_Click()
    StudFind
End Sub

Private Sub Form_Unload(Cancel As Integer)
    db1.Close
End Sub
Public Sub StudFind()
On Error GoTo PError
   strpos = InStr(1, cmbStudNo, ",", 1)
   If strpos = 0 Then
     studBuff = "studno = '" & UCase(cmbStudNo) & "'"
   Else
     studBuff = "studno = '" & Left(cmbStudNo, strpos - 1) & "'"
   End If
    rstStudmast.FindFirst studBuff
    If rstStudmast.NOMATCH Then
       MsgBox "Mail with the Registration Number entered is not found.", vbCritical, "Enquiry"
    Else
        cmbStudNo = rstStudmast!StudNo & ", " & rstStudmast!StudNames
        showdata
    End If
PError:
End Sub

Public Sub ClearData()
On Error GoTo PError
        PicName = PicsPath & "picinit.bmp"
        StudPic = LoadPicture(PicName)
        StudNo.Caption = ""
        StudNames = ""
        stFileNo = ""
        regDate = ""
        Sex = ""
        DOB = ""
        Age = ""
        HMAddr1 = ""
        Guardian = ""
        GDPhone = ""
        stemail = ""
        cclass = ""
        skills = ""
        talent = ""
        sports = ""
        CMaster = ""
PError:
End Sub
