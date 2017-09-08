VERSION 5.00
Begin VB.Form FrmStudEnq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Tracking System"
   ClientHeight    =   6765
   ClientLeft      =   1305
   ClientTop       =   1380
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   9645
   Begin VB.PictureBox SSPanel3 
      BackColor       =   &H00C0C0C0&
      Height          =   6015
      Left            =   3840
      ScaleHeight     =   5955
      ScaleWidth      =   5715
      TabIndex        =   10
      Top             =   720
      Width           =   5775
      Begin VB.Data datGenTab 
         Caption         =   "General Data"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cmbStudNo 
         Height          =   315
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   3855
      End
      Begin VB.PictureBox StudNames 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   360
         ScaleHeight     =   555
         ScaleWidth      =   5235
         TabIndex        =   29
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label stemail 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   35
         Top             =   5400
         Width           =   3375
      End
      Begin VB.Label sports 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   34
         Top             =   5040
         Width           =   3375
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sports:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Talents"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   31
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label talent 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   30
         Top             =   4680
         Width           =   3375
      End
      Begin VB.Label HMAddr1 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   2160
         TabIndex        =   26
         Top             =   2760
         Width           =   3375
         WordWrap        =   -1  'True
      End
      Begin VB.Label CMaster 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   25
         Top             =   3960
         Width           =   3375
      End
      Begin VB.Label skills 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   4320
         Width           =   3375
      End
      Begin VB.Label GDPhone 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   23
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label Guardian 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label cclass 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Skills:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Class Master:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Guadian's Name"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Class:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00400000&
         Caption         =   "Student No:"
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
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Phone Number:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Home Address:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   2760
         Width           =   1695
      End
   End
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H00C0C0C0&
      Height          =   6015
      Left            =   0
      ScaleHeight     =   5955
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   720
      Width           =   3735
      Begin VB.CommandButton cmdexit 
         Caption         =   "E&xit"
         Height          =   555
         Left            =   2160
         TabIndex        =   28
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label stFileNo 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Date of Registration:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sex:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Age:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Date of Birth:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "File Number:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Sex 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Age 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label DOB 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label regDate 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label StudNo 
         Alignment       =   2  'Center
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
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   600
         TabIndex        =   1
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Image StudPic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2055
         Left            =   600
         Picture         =   "frmStudEnq.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.PictureBox SSPanel5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   9555
      TabIndex        =   27
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "FrmStudEnq"
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

Dim GenClass As New DProLog
Private Sub cmbStudno_LostFocus()
    StudFind
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'On Error GoTo PError
     GenClass.fleLogin mvUserid, "Accessed Student Enquiry", Date, Time
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
    datGenTab.DatabaseName = DBpath
    datGenTab.RecordSource = "studmast"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       cmbStudNo.AddItem datGenTab.Recordset("studno") + ", " + datGenTab.Recordset("StudNames")
       datGenTab.Recordset.MoveNext
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
       MsgBox "Student with the Registration Number entered is not found.", vbCritical, "Enquiry"
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
