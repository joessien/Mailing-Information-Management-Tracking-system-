VERSION 5.00
Begin VB.Form frmCPasswd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Tracking System"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5910
   Begin VB.PictureBox SSPanel5 
      BackColor       =   &H00C0C0C0&
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2715
      ScaleWidth      =   5835
      TabIndex        =   4
      Top             =   480
      Width           =   5895
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   4680
         TabIndex        =   11
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4680
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox CurrPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2400
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtConfm 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2400
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtUserPswd 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2400
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Current Password"
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
         Height          =   225
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Name:"
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
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Confirm New Password:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "New Password:"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1470
         Width           =   1455
      End
   End
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   5835
      TabIndex        =   3
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmCPasswd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db1 As Database, wrktemp As Workspace
Dim mvBuff As String, mvErrFlg As Integer, mcurrpass As String
Dim rstUserFle As Recordset
Dim GenClass As New mProLog
Dim rstPswHst, rstReg As Recordset, mvPswd As String, i As Integer

Private Sub FindAcc()
On Error GoTo PError
    mvBuff = "Select * from userfle Where UserID = '" _
        & mvUserid & "'"
    Set rstUserFle = db1.OpenRecordset(mvBuff, dbOpenDynaset)
    If rstUserFle.EOF And rstUserFle.BOF Then
        Set rstUserFle = db1.OpenRecordset("hrUsers", dbOpenDynaset)
        mvBuff = "Select * from hrUsers Where UserID = '" _
        & mvUserid & "'"
        Set rstUserFle = db1.OpenRecordset(mvBuff, dbOpenDynaset)
        If rstUserFle.EOF And rstUserFle.BOF Then
           MsgBox "User Not Found", vbInformation, Me.Caption
           Unload Me
        Else
           Label2.Caption = rstUserFle!UserName
           Label2.Refresh
           mvPswd = rstUserFle!UserPswd
         End If
    Else
        Label2.Caption = rstUserFle!UserName
        Label2.Refresh
        mvPswd = rstUserFle!UserPswd
    End If
PError:
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
On Error GoTo PError
    mvErrFlg = 0
    ValFlds
    If mvErrFlg = 0 Then
        ' Update Password History Table
        rstPswHst.AddNew
        rstPswHst("UserID") = mvUserid
        rstPswHst("UserPswd") = rstUserFle!UserPswd
        rstPswHst.Update
        ' Update Password Table With New Password
        rstUserFle.Edit
        rstUserFle("UserPswd") = txtUserPswd
        rstUserFle("PassExp") = Date + rstReg!PassExp
        'rstUserFle("PassGrs") = rstGlCtrl!PassGrs
        rstUserFle.Update
        Me.Hide
        MsgBox "Password Successfully Changed", vbOKOnly, Me.Caption
        GenClass.fleLogin mvUserid, "Changed Password ", Date, Time
        Unload Me
    End If
PError:
End Sub

Private Sub Currpass_GotFocus()
    CurrPass.SelStart = 0
    CurrPass.SelLength = Len(CurrPass)
End Sub

Private Sub Currpass_LostFocus()
On Error GoTo PError
    If Trim(CurrPass) <> mvPswd Then
       i = MsgBox("Invalid Current Password, Do you want to quit?", vbYesNo, "Password Change")
       If i = vbYes Then
          Unload Me
       Else
          CurrPass.SetFocus
       End If
    End If
    Exit Sub
PError:
End Sub

Private Sub Form_Load()
    On Error GoTo PError
    Call FormCentreMDI(Me)
    Set wrktemp = DBEngine.Workspaces(0)
    Set db1 = wrktemp.OpenDatabase(DBpath, True)
    Set rstReg = db1.OpenRecordset("DProReg", dbOpenDynaset)
    Set rstPswHst = db1.OpenRecordset("pswhst", dbOpenDynaset)
    Set rstUserFle = db1.OpenRecordset("Userfle", dbOpenDynaset)
    FindAcc
    Screen.MousePointer = vbDefault
    CurrPass.TabIndex = 0
    Me.Caption = MVCoyname
    GenClass.fleLogin mvUserid, "Accessed Change Password ", Date, Time
PError:
End Sub

Private Sub ValFlds()
    On Error GoTo PError
    If Trim(txtUserPswd) = "" Or txtUserPswd <> txtConfm Then
        MsgBox "Invalid Password", vbExclamation
        mvErrFlg = 1
        txtUserPswd.SetFocus
        Exit Sub
    End If
    mvBuff = "Select * from PswHst Where UserID = '" _
    & mvUserid & "' And UserPswd = '" _
    & txtUserPswd & "'"
    Set rstPswHst = db1.OpenRecordset(mvBuff, dbOpenDynaset)
    If Not (rstPswHst.EOF And rstPswHst.BOF) _
        Or mvPswd = txtUserPswd Then
        MsgBox "Password Previously Used, Please Re-enter", vbExclamation
        txtUserPswd.SetFocus
        mvErrFlg = 1
    End If
PError:
End Sub

Private Sub txtUserPswd_gotFocus()
    txtUserPswd.SelStart = 0
    txtUserPswd.SelLength = Len(txtUserPswd)
End Sub

Private Sub txtConfm_GotFocus()
    txtConfm.SelStart = 0
    txtConfm.SelLength = Len(txtConfm)
End Sub
