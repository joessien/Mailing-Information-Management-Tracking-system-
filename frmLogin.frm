VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4935
   ClientLeft      =   2700
   ClientTop       =   1950
   ClientWidth     =   6210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmLogin.frx":030A
   ScaleHeight     =   4935
   ScaleWidth      =   6210
   Begin VB.TextBox txtUserId 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4080
      MaxLength       =   15
      TabIndex        =   0
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtNPswd 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   315
      HideSelection   =   0   'False
      IMEMode         =   3  'DISABLE
      Left            =   3960
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtUserPswd 
      Appearance      =   0  'Flat
      Height          =   315
      HideSelection   =   0   'False
      IMEMode         =   3  'DISABLE
      Left            =   4080
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdAccept 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Accept"
      Height          =   375
      Left            =   3360
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Quit"
      Height          =   375
      Left            =   4680
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtConfm 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   315
      HideSelection   =   0   'False
      IMEMode         =   3  'DISABLE
      Left            =   3960
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label SSP3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   5775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "User ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
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
      Left            =   1920
      TabIndex        =   10
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter New Password:"
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
      Left            =   1920
      TabIndex        =   11
      Top             =   3240
      Width           =   1935
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GenClass As New mProLog, mvErrFlg As Integer
Dim mvCtr As Integer, mvReply As String
Dim mvPswd As String, nvGrs As Integer
Dim DB3 As Database, PathCtrl As Recordset
Dim rstUserFle As Recordset, NodCtr As Integer
Dim rstPswHst As Recordset, FAmvBuff As String
Dim rstReg, rstTerms As Recordset, Mctr As Integer
Dim v As String

Private Sub clsAll()
    Dim i As Long
    For i = Forms.Count - 1 To 0 Step -1
        Unload Forms(i)
        Unload Me
    Next
End Sub

Private Sub cmdExit_Click()
    GenClass.fleLogin mvUserid, "Quiting Application", Date, Time
    clsAll
    Unload FrmMtrack
    Unload frmsplash
    Unload Me
End Sub

Private Sub cmdquit_Click()
    GenClass.fleLogin mvUserid, "Quiting from PassWord Change", Date, Time
    clsAll
    Unload FrmMtrack
    Unload frmsplash
    Unload Me
End Sub

Private Sub txtUserID_LostFocus()
On Error GoTo PError
    If txtUserID = "" Then
        Beep
        mvReply = MsgBox("User Identification not specified, do you want to exit ?", vbYesNo, "Login")
        If mvReply = vbYes Then
            GenClass.fleLogin mvUserid, "Quiting Application", Date, Time
            'clsAll
            Unload FrmMtrack
            Unload frmsplash
            Unload Me
            DB3.Close
        Else
            txtUserID = ""
            txtUserID.SetFocus
        End If
    Else
        FAmvBuff = "Select * from UserFle Where UserID = '" _
            & txtUserID & "'"
        Set rstUserFle = DB3.OpenRecordset(FAmvBuff, dbOpenDynaset)
        If rstUserFle.EOF And rstUserFle.BOF Then
           Beep
           MsgBox "Invalid User Identification, Please Re-enter", vbExclamation
           txtUserID = ""
           txtUserID.SetFocus
           mnu1 = False
           mnu2 = False
           mnu3 = False
           mnu4 = False
           mnu5 = False
           mnu6 = False
           mnu7 = False
           mnu8 = False
           mnu9 = False
           mnu10 = False
           mnu11 = False
           mnu12 = False
           mnu13 = False
           mnu14 = False
           mnu15 = False
           
        Else
           If rstUserFle!alock = True Then
              GenClass.fleLogin mvUserid, "Attemted access to a locked account", Date, Time
              MsgBox "Your account has been locked, Contact the Administrator", vbCritical
              Unload FrmMtrack
              Unload frmsplash
              Unload Me
              DB3.Close
              End
           End If
           mvPswd = rstUserFle!UserPswd
           mvAcad = rstUserFle!acad
           GCurrTerm = rstReg!cTerm
           Crest = rstReg!ClassSize
           GautoMs = rstReg!automs
           GLogo = rstReg!logo
           GmedLim = rstReg!medlim
           GSControl = rstReg!sControl
           mvUserName = rstUserFle!UserName
           mnu1 = rstUserFle!mnu1
           mnu2 = rstUserFle!mnu2
           mnu3 = rstUserFle!mnu3
           mnu4 = rstUserFle!mnu4
           mnu5 = rstUserFle!mnu5
           mnu6 = rstUserFle!mnu6
           mnu7 = rstUserFle!mnu7
           mnu8 = rstUserFle!mnu8
           mnu9 = rstUserFle!mnu9
           mnu10 = rstUserFle!mnu10
           mnu11 = rstUserFle!mnu11
           mnu12 = rstUserFle!mnu12
           mnu13 = rstUserFle!mnu13
           mnu14 = rstUserFle!mnu14
           mnu15 = rstUserFle!mnu15
        End If
        mvCtr = mvCtr + 1
    End If
    If mvCtr > 3 Then
        GenClass.fleLogin txtUserID, "Attempt Logging In", Date, Time
        MsgBox "Illegal Access Atempted, Application will Quit", vbCritical
        Unload frmsplash
        Unload Me
        DB3.Close
    End If
PError:
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    frmLogin.Top = 1700
    frmLogin.Left = 3000

    
    '______________________________________________
    ' opens a text file to read path of database
    mvDbFile = ""
    Open "C:\windows\system32\mtpro.txt" For Input As #1
    Do While Not EOF(1)
       Input #1, mvDbFile
    Loop
    Close #1
    '___________________________________________
    Set DB3 = OpenDatabase(mvDbFile)
    Set rstUserFle = DB3.OpenRecordset("UserFle", dbOpenDynaset)
    Set rstReg = DB3.OpenRecordset("Dproreg", dbOpenDynaset)
    Set rstPswHst = DB3.OpenRecordset("PswHst", dbOpenDynaset)
    If rstReg!NodYear - Date <= 0 Then
       MsgBox "Your application has expired due to invalid Licence. Contact the vendor.", vbCritical, "Licence Information"
       End
    End If
    If rstReg!NodYear - Date < 7 Then
       NodYear = rstReg!NodYear - Date
       MsgBox "This application will stop working in " & Str(NodYear) & " days from today. Contact the vendor to validate your license.", vbCritical, "Licence Information"
    End If
    If rstReg!cVal <> 666 Then
       If rstReg!NodYear - Date > 8 Then
          NodCtr = rstReg!NodYear - Date
          rstReg.Edit
          rstReg!NodYear = rstReg!NodYear - NodCtr + 8
          rstReg.Update
       Else
          MsgBox "Your days of free usage is over. Buy this software if you wish to continue using it.", vbCritical, "Purchase Information"
       End If
    End If
    DBpath = mvDbFile
    Rptpath = rstReg!Rptpath
    DocsPath = rstReg!DocsPath
    PicsPath = rstReg!PicsPath
    CInit = rstReg!ClassInit
    GAutoProm = rstReg!Autoprom
    Label3.Visible = True
    Label6.Visible = True
    Label21.Visible = False
    Label22.Visible = False
    txtNPswd.Visible = False
    txtConfm.Visible = False
    cmdAccept.Visible = False
    cmdQuit.Visible = False
    coyname = rstReg!RegName
    MVCoyname = rstReg!RegName
    GlbState = rstReg!gstate
    gBunitD = rstReg!gbud
    SSP3.Caption = coyname
    mvUserid = ""
    txtUserPswd = ""
    txtUserID = ""
    mvCtr = 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdOk_Click()
On Error GoTo PError
    If txtUserPswd <> "" Then
        If txtUserPswd <> mvPswd Then
            MsgBox "Invalid Password, Please Re-enter", vbExclamation
            txtUserPswd = ""
            mvCtr = mvCtr + 1
            If mvCtr > 3 Then
                If mnu15 = False Then
                   rstUserFle.Edit
                   rstUserFle("alock") = True
                   rstUserFle.Update
                   GenClass.fleLogin txtUserID, "Account locked due to 3 attempted login failures", Date, Time
                   MsgBox "Your account has been locked, Contact the Administrator", vbCritical
                End If
                Unload frmsplash
                Unload Me
                DB3.Close
            Else
                txtUserPswd.SetFocus
            End If
        Else
            mvUserid = txtUserID
            GenClass.fleLogin mvUserid, "Logged In", Date, Time
            If Date > rstUserFle!PassExp Then
                If rstUserFle!PassGrs > 0 Then
                    If MsgBox("Your Password Has Expired," & Chr(13) & _
                        "You Have " + _
                        Str(rstUserFle!PassGrs) & " Grace Left" & Chr(13) & _
                        "Want To Change Password?", _
                        vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                        UpdPas
                    Else
                        rstUserFle.Edit
                        rstUserFle("PassGrs") = rstUserFle!PassGrs - 1
                        rstUserFle.Update
                        RunNPrg
                    End If
                Else
                    MsgBox "Your Password Has Expired" & Chr(13) & _
                        "You Have " + _
                        Str(rstUserFle!PassGrs) & " Grace Left" & Chr(13) & _
                        "You Must Change Your Password?", vbCritical
                        UpdPas
                End If
            Else
                RunNPrg
            End If
        End If
    End If
PError:
End Sub

Private Sub RunNPrg()
    menuLoad
    FrmMtrack.Show
    DB3.Close
    Unload Me
End Sub

Private Sub UpdPas()
    cmdOk.Visible = False
    cmdExit.Visible = False
    Label3.Visible = False
    Label6.Visible = False
    txtUserPswd.Visible = False
    txtUserID.Visible = False
    Label21.Visible = True
    Label22.Visible = True
    txtNPswd.Visible = True
    txtConfm.Visible = True
    cmdAccept.Visible = True
    cmdQuit.Visible = True
    txtNPswd.SetFocus
End Sub

Private Sub cmdAccept_Click()
On Error GoTo PError
    mvErrFlg = 0
    ValFlds
    If mvErrFlg = 0 Then
        ' Update Password History Table
        rstPswHst.AddNew
        rstPswHst("UserID") = txtUserID
        rstPswHst("UserPswd") = rstUserFle!UserPswd
        rstPswHst.Update
        ' Update Password Table With New Password
        rstUserFle.Edit
        rstUserFle("UserPswd") = txtNPswd
        rstUserFle("PassExp") = Date + rstReg!PassExp
        'rstUserFle("PassExp") = Date +  rstReg!passCtrl
        rstUserFle("PassGrs") = rstReg!PassGrs
        rstUserFle.Update
        RunNPrg
    End If
PError:
End Sub

Private Sub ValFlds()
On Error GoTo PError:
    If txtNPswd = "" Or txtNPswd <> txtConfm Then
        MsgBox "Invalid Password", vbExclamation
        mvErrFlg = 1
        txtNPswd.SetFocus
        Exit Sub
    End If
    FAmvBuff = "Select * from PswHst Where UserID = '" _
        & txtUserID & "' And UserPswd = '" _
        & txtNPswd & "'"
    Set rstPswHst = DB3.OpenRecordset(FAmvBuff, dbOpenDynaset)
    If Not (rstPswHst.EOF And rstPswHst.BOF) Or _
            mvPswd = txtNPswd Then
        MsgBox "Password Previously Used, Please Re-enter", vbExclamation
        txtNPswd.SetFocus
        mvErrFlg = 1
    End If
PError:
End Sub

Private Sub txtConfm_GotFocus()
    txtConfm.SelStart = 0
    txtConfm.SelLength = Len(txtConfm)
End Sub

Private Sub txtNPswd_GotFocus()
    txtNPswd.SelStart = 0
    txtNPswd.SelLength = Len(txtNPswd)
End Sub
Public Sub menuLoad()
    With FrmMtrack
        '________________________clear access rights
        ' ----------User setup
        '.mnu10.Enabled = False
        .mnu11.Enabled = False
        .mnu12.Enabled = False
        .mnu13.Enabled = False
        ' ----------company setup
        '.mnu20.Enabled = False
        .mnu21.Enabled = False
        '.mnu22.Enabled = False
        .mnu23.Enabled = False
        '.mnu24.Enabled = False
        .mnu25.Enabled = False
        ' ----------Incoming Mail
        '.mnu30.Enabled = False
        .mnu31.Enabled = False
        .mnu32.Enabled = False
        '.mnu33.Enabled = False
        .mnu34.Enabled = False
        ' ----------Out going Mail
        '.mnu40.Enabled = False
        .mnu41.Enabled = False
        .mnu42.Enabled = False
        '.mnu43.Enabled = False
        .mnu44.Enabled = False
        ' ----------Reports Mail
        '.mnu50.Enabled = False
        ' ----------Searching criteria
        .mnu60.Enabled = False
        .mnu61.Enabled = False
        .mnu62.Enabled = False
        .mnu65.Enabled = False
        
        ' ----------Referencing
        '.mnu70.Enabled = False
        .mnu71.Enabled = False
        .mnu72.Enabled = False
        .mnu73.Enabled = False
        .mnu74.Enabled = False
        .mnu75.Enabled = False
        .mnu76.Enabled = False
        .mnu77.Enabled = False
        .mnu78.Enabled = False
        .mnu79.Enabled = False
        .mnu7a.Enabled = False
        '__________________________________________________________load access rights
        
        '________________________Assign access rights
        ' ----------User Logon
        .mnu10.Enabled = True
        .mnu11.Enabled = mnu1
        .mnu12.Enabled = True
        .mnu13.Enabled = True
        ' ----------company  and user setup
        .mnu20.Enabled = True
        .mnu21.Enabled = mnu2
        .mnu22.Enabled = True
        .mnu23.Enabled = mnu3
        .mnu24.Enabled = True
        .mnu25.Enabled = mnu3
        ' ----------Incoming Mail
        .mnu30.Enabled = True
        .mnu31.Enabled = mnu4
        .mnu32.Enabled = mnu5
        .mnu33.Enabled = True
        .mnu34.Enabled = mnu6
        ' ----------Out going Mail
        .mnu40.Enabled = True
        .mnu41.Enabled = mnu7
        .mnu42.Enabled = mnu8
        .mnu43.Enabled = True
        .mnu44.Enabled = mnu9
        ' ----------Reports
        .mnu50.Enabled = True
        .mnu51.Enabled = mnu7
        .mnu54.Enabled = mnu4
        .mnu55.Enabled = mnu7
        .mnu55a.Enabled = mnu7
        .mnu56.Enabled = mnu6
        .mnu57.Enabled = mnu9
        .mnu58.Enabled = mnu4
        
        ' ----------Search Mails
        .mnu60.Enabled = mnu10
        .mnu61.Enabled = mnu10
        .mnu65.Enabled = mnu10
        ' ----------Killed Mails Searching
        .mnu60.Enabled = True
        .mnu62.Enabled = mnu11
        ' ----------Referencing
        .mnu70.Enabled = True
        .mnu71.Enabled = mnu12
        .mnu72.Enabled = mnu12
        .mnu73.Enabled = mnu12
        .mnu74.Enabled = mnu12
        .mnu75.Enabled = mnu12
        .mnu76.Enabled = mnu12
        .mnu77.Enabled = mnu12
        .mnu78.Enabled = mnu12
        .mnu79.Enabled = mnu12
        .mnu7a.Enabled = mnu12
    End With
    FrmMtrack.Show
    Unload frmLogin
    Unload frmsplash
End Sub
