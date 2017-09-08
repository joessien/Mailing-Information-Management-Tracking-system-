VERSION 5.00
Begin VB.Form frmuser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Tracking System"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7290
   Begin VB.CommandButton cmdCanc 
      BackColor       =   &H00FF8080&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4800
      Width           =   1065
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FF8080&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4800
      Width           =   1065
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FF8080&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4800
      Width           =   1065
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FF8080&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4800
      Width           =   1065
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00FF8080&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4800
      Width           =   1065
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FF8080&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4800
      Width           =   1065
   End
   Begin VB.CheckBox chk12 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Global Reference Setup "
      Height          =   255
      Left            =   2760
      TabIndex        =   20
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CheckBox chk11 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Del/Dom Mails Search"
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CheckBox chk10 
      BackColor       =   &H00FFC0C0&
      Caption         =   "In/Out Mails Search "
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CheckBox chk9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Mail Domiciliation"
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CheckBox chk8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Outgoing Mails Update"
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CheckBox chk7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Outgoing Mails Checkout"
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CheckBox chk6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Incoming Mails Delete"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CheckBox chk5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Incoming Mails Update"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CheckBox chk4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Incoming Mails Registration"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CheckBox chk3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "User Access Setup"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CheckBox chk2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "System Configuration"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CheckBox chk1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Password Change"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox txtUserID 
      Height          =   285
      Left            =   1080
      MaxLength       =   9
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1200
      Width           =   4335
   End
   Begin VB.TextBox txtUserPswd 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtConfm 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3840
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CheckBox ChkLock 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Lock"
      Height          =   255
      Left            =   6360
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdBrow 
      BackColor       =   &H00FF8080&
      Caption         =   "Bro&wse"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3360
      Width           =   870
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00FF8080&
      Caption         =   "&Find"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdBott 
      BackColor       =   &H00FF8080&
      Caption         =   "&Bottom"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2760
      Width           =   870
   End
   Begin VB.CommandButton cmdPrev 
      BackColor       =   &H00FF8080&
      Caption         =   "&Previous"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2400
      Width           =   870
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FF8080&
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2760
      Width           =   870
   End
   Begin VB.CommandButton cmdTop 
      BackColor       =   &H00FF8080&
      Caption         =   "&Top"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2400
      Width           =   870
   End
   Begin VB.Label fhdr 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Access Configuration"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   0
      TabIndex        =   33
      Top             =   120
      Width           =   7215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   7080
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
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
      Left            =   120
      TabIndex        =   8
      Top             =   1590
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
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
      Left            =   120
      TabIndex        =   7
      Top             =   855
      Width           =   855
   End
   Begin VB.Label Label2 
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
      Left            =   120
      TabIndex        =   6
      Top             =   1215
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Confirm:"
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
      Left            =   3000
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mvOpt As Integer, mvReply As Integer
Dim db6 As Database, mvErrFlg As Integer
Dim wrktemp As Workspace, MAcada As Boolean
Dim rstUserFle As Recordset, rst, rstReg  As Recordset
Dim GenClass As New mProLog
Dim mvBuff As String, mvLRec As String

Private Sub FindAcc()
    mvBuff = "Select * from UserFle Where UserID = '" _
        & txtUserID & "'"
    Set rstUserFle = db6.OpenRecordset(mvBuff, dbOpenDynaset)
     Select Case mvOpt
        Case 1
            If Not (rstUserFle.EOF And rstUserFle.BOF) Then
                MsgBox "User ID Already Defined", vbInformation, Me.Caption
                txtUserID.SetFocus
            End If
        Case Is > 1
            If rstUserFle.EOF And rstUserFle.BOF Then
                MsgBox "User ID Does Not Exits", vbInformation, Me.Caption
                txtUserID.SetFocus
            Else
                showdata
            End If
    End Select
End Sub

Private Sub CmdAdd_Click()
    FldEnab
    ClearFlds
    CmdDisab
    cmdEdit.Enabled = False
    cmdDel.Enabled = False
    EnabSavCan
    mvOpt = 1
    cmdAdd.Enabled = False
    txtUserID.SetFocus
End Sub

Private Sub CmdBrow_Click()
    Screen.MousePointer = vbHourglass
    mvSql = "Select UserID, UserName From UserFle"
    Set rst = db6.OpenRecordset(mvSql, dbOpenSnapshot)
     If (rst.BOF And rst.EOF) Then
         MsgBox "No users are defined", vbExclamation, "User Administration"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "List of System Users"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         frmBrowse.grdGeneral.Columns(0).Width = 1400
         frmBrowse.grdGeneral.Columns(1).Width = 2800
         Col0.Caption = "User Id"
         Col1.Caption = "User Names"
   End If
   Screen.MousePointer = vbDefault
   frmBrowse.Show
End Sub

Private Sub cmdCanc_Click()
    ClearFlds
    disSavCan
    EnabCmd
    CmdEnab
    If (rstUserFle.EOF And rstUserFle.BOF) Then
        cmdEdit.Enabled = False
        cmdDel.Enabled = False
    End If
    cmdAdd.Enabled = True
    FldDisab
End Sub

Private Sub CmdDel_Click()
    mvOpt = 3
    If txtUserID = "" Then
       MsgBox "User ID Not Specified", vbInformation, Me.Caption
       Exit Sub
    End If
    FindAcc
    mvReply = DeleteCheck()
    If mvReply = vbYes Then
        rstUserFle.Delete
        MsgBox "User Details Deleted", vbInformation, Me.Caption
        FldEnab
        ClearFlds
        FldDisab
    Else
        MsgBox "User Details Not Deleted", vbExclamation, Me.Caption
    End If
End Sub

Private Sub CmdEdit_Click()
   mvOpt = 2
   If txtUserID = "" Then
       MsgBox "User ID Not Specified", vbInformation, Me.Caption
       Exit Sub
   End If
   FindAcc
   FldEnab
   txtUserID.Enabled = False
   CmdDisab
   cmdAdd.Enabled = False
   cmdDel.Enabled = False
   EnabSavCan
   txtUserName.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdTop_Click()
    mvBuff = "Select * from UserFle Order By UserID;"
    Set rstUserFle = db6.OpenRecordset(mvBuff, dbOpenDynaset)
    If Not rstUserFle.BOF Then
       rstUserFle.MoveFirst
       showdata
       FldDisab
    End If
End Sub

Private Sub CmdNext_Click()
    Dim flag As Integer
    On Error GoTo NextError
    If rstUserFle.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstUserFle.EOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEPAST)
        rstUserFle.MoveLast
    Else
        rstUserFle.MoveNext
        If rstUserFle.EOF Then
            rstUserFle.MoveLast
        End If
    End If
NextClear:
    showdata
    Exit Sub
NextError:
    ErrorMessages (NOMOVEPAST)
    rstUserFle.Requery
    rstUserFle.MoveLast
    On Error GoTo 0
    Resume NextClear
End Sub

Private Sub CmdPrev_Click()
  Dim flag As Integer
    
    On Error GoTo PrevError
    If rstUserFle.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstUserFle.BOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEFRONT)
        rstUserFle.MoveFirst
    Else
        rstUserFle.MovePrevious
        If rstUserFle.BOF Then
            rstUserFle.MoveFirst
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

Private Sub CmdBott_Click()
    mvBuff = "Select * from UserFle Order By UserID;"
    Set rstUserFle = db6.OpenRecordset(mvBuff, dbOpenDynaset)
    If Not (rstUserFle.BOF And rstUserFle.EOF) Then
       rstUserFle.MoveLast
       showdata
       FldDisab
    End If
End Sub

Private Sub cmdFind_Click()
    Dim mvLength As Integer
    Dim mvBuff As String, mvAns As String
    mvLength = 0
    mvAns = InputBox$("Enter User ID:")
    mvLength = Len(mvAns)
    If mvLength = 0 Then
       MsgBox "User Selected Cancel or Nothing To Find", vbInformation, Me.Caption
       Exit Sub
    End If
    mvBuff = "Select * from UserFle Where UserID = '" _
        & mvAns & "'"
    Set rstUserFle = db6.OpenRecordset(mvBuff, dbOpenDynaset)
    If rstUserFle.EOF And rstUserFle.BOF Then
        MsgBox "User ID Not Found", vbInformation, Me.Caption
        Exit Sub
    Else
        showdata
    End If
End Sub

Private Sub CmdSave_Click()
    mvErrFlg = 0
    ValFlds
    If mvErrFlg = 0 Then
        Select Case mvOpt
            Case 1
                rstUserFle.AddNew
                rstUserFle("UserID") = UCase(txtUserID)
                rstUserFle("acad") = False
                UpdTab
            Case 2
                rstUserFle.Edit
                rstUserFle("acad") = MAcada
                UpdTab
        End Select
        cmdAdd.Enabled = True
        CmdEnab
        disSavCan
        EnabCmd
        FldDisab
        mvBuff = "Select * from UserFle Order By UserID;"
        Set rstUserFle = db6.OpenRecordset(mvBuff, dbOpenDynaset)
        GTBuff = "userid = '" & UCase(txtUserID) & "'"
        rstUserFle.FindFirst GTBuff
    End If
End Sub

Private Sub showdata()
    If Not (rstUserFle.BOF And rstUserFle.EOF) Then
        mvLRec = txtUserID
        txtUserID = rstUserFle!userid
        txtUserPswd = rstUserFle!UserPswd
        txtConfm = txtUserPswd
        txtUserName = rstUserFle!UserName
        If rstUserFle!mnu1 = True Then
            chk1.Value = 1
        Else
            chk1.Value = 0
        End If
        If rstUserFle!mnu2 = True Then
            chk2.Value = 1
        Else
            chk2.Value = 0
        End If
        If rstUserFle!mnu3 = True Then
            chk3.Value = 1
        Else
            chk3.Value = 0
        End If
        If rstUserFle!mnu4 = True Then
            chk4.Value = 1
        Else
            chk4.Value = 0
        End If
        If rstUserFle!mnu5 = True Then
            chk5.Value = 1
        Else
            chk5.Value = 0
        End If
        If rstUserFle!mnu6 = True Then
            chk6.Value = 1
        Else
            chk6.Value = 0
        End If
        If rstUserFle!mnu7 = True Then
            chk7.Value = 1
        Else
            chk7.Value = 0
        End If
        If rstUserFle!mnu8 = True Then
            chk8.Value = 1
        Else
            chk8.Value = 0
        End If
        If rstUserFle!mnu9 = True Then
            chk9.Value = 1
        Else
            chk9.Value = 0
        End If
        If rstUserFle!mnu10 = True Then
            chk10.Value = 1
        Else
            chk10.Value = 0
        End If
        If rstUserFle!mnu11 = True Then
            chk11.Value = 1
        Else
            chk11.Value = 0
        End If
        If rstUserFle!mnu12 = True Then
            chk12.Value = 1
        Else
            chk12.Value = 0
        End If
        If rstUserFle!alock = True Then
           ChkLock.Value = 1
        Else
           ChkLock.Value = 0
        End If
        MAcada = rstUserFle!acad
        FldDisab
    End If
End Sub

Private Sub CmdEnab()
    cmdFind.Enabled = True
    cmdbrow.Enabled = True
    cmdtop.Enabled = True
    cmdbott.Enabled = True
    cmdnext.Enabled = True
    cmdprev.Enabled = True
End Sub

Private Sub EnabCmd()
    cmdAdd.Enabled = True
    cmdEdit.Enabled = True
    cmdDel.Enabled = True
End Sub

Private Sub CmdDisab()
    cmdFind.Enabled = False
    cmdbrow.Enabled = False
    cmdtop.Enabled = False
    cmdbott.Enabled = False
    cmdnext.Enabled = False
    cmdprev.Enabled = False
    If rstUserFle.EOF And rstUserFle.BOF Then
        cmdEdit.Enabled = False
        cmdDel.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Call FormCentreMDI(Me)
    ClearFlds
    disSavCan
    Screen.MousePointer = vbHourglass
    Set wrktemp = DBEngine.Workspaces(0)
    Set db6 = wrktemp.OpenDatabase(DBpath, True)
    Set rstUserFle = db6.OpenRecordset("userfle", dbOpenDynaset)
    Set rstReg = db6.OpenRecordset("DProReg", dbOpenDynaset)
    If rstUserFle.EOF And rstUserFle.BOF Then
        CmdDisab
        cmdEdit.Enabled = False
        cmdDel.Enabled = False
    Else
        rstUserFle.MoveFirst
    End If
    FldDisab
    GenClass.fleLogin mvUserid, "Accessed User Administration", Date, Time
    Screen.MousePointer = vbDefault
End Sub

Private Sub ClearFlds()
    txtUserID = ""
    txtUserName = ""
    txtUserPswd = ""
    txtConfm = ""
    chk1.Value = False
    chk2.Value = False
    chk3.Value = False
    chk4.Value = False
    chk5.Value = False
    chk6.Value = False
    chk7.Value = False
    chk8.Value = False
    chk9.Value = False
    chk10.Value = False
    chk11.Value = False
    chk12.Value = False
    ChkLock.Value = False
   mvOpt = 0
End Sub

Private Sub disSavCan()
    cmdSave.Enabled = False
    cmdCanc.Enabled = False
End Sub

Private Sub EnabSavCan()
    cmdSave.Enabled = True
    cmdCanc.Enabled = True
End Sub

Private Sub UpdTab()
On Error GoTo PError
    rstUserFle("UserName") = txtUserName
    rstUserFle("UserPswd") = txtUserPswd
    rstUserFle("PassExp") = Date + rstReg!PassExp
    rstUserFle("mnu1") = chk1.Value
    rstUserFle("mnu2") = chk2.Value
    rstUserFle("mnu3") = chk3.Value
    rstUserFle("mnu4") = chk4.Value
    rstUserFle("mnu5") = chk5.Value
    rstUserFle("mnu6") = chk6.Value
    rstUserFle("mnu7") = chk7.Value
    rstUserFle("mnu8") = chk8.Value
    rstUserFle("mnu9") = chk9.Value
    rstUserFle("mnu10") = chk10.Value
    rstUserFle("mnu11") = chk11.Value
    rstUserFle("mnu12") = chk12.Value
    rstUserFle("alock") = ChkLock.Value
    rstUserFle.Update
    Exit Sub
PError:
    MsgBox "Record already exist", vbExclamation
End Sub

Private Sub ValFlds()
    If txtUserID = "" Then
        MsgBox "Invalid User ID", vbExclamation
        mvErrFlg = 1
        txtUserID.SetFocus
        Exit Sub
    End If
    If txtUserName = "" Then
        MsgBox "Invalid User Name", vbExclamation
        mvErrFlg = 1
        txtUserName.SetFocus
        Exit Sub
    End If
    If txtUserPswd = "" Or txtUserPswd <> txtConfm Then
        MsgBox "Invalid Password", vbExclamation
        mvErrFlg = 1
        txtUserPswd.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    db6.Close
End Sub

Private Sub txtUserID_LostFocus()
    FindAcc
End Sub

Private Sub txtUserName_GotFocus()
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName)
End Sub

Private Sub txtUserPswd_gotFocus()
    txtUserPswd.SelStart = 0
    txtUserPswd.SelLength = Len(txtUserPswd)
End Sub

Private Sub txtConfm_GotFocus()
    txtConfm.SelStart = 0
    txtConfm.SelLength = Len(txtConfm)
End Sub

Private Sub FldEnab()
    txtUserID.Enabled = True
    txtUserName.Enabled = True
    txtUserPswd.Enabled = True
    txtConfm.Enabled = True
    chk1.Enabled = True
    chk2.Enabled = True
    chk3.Enabled = True
    chk4.Enabled = True
    chk5.Enabled = True
    chk6.Enabled = True
    chk7.Enabled = True
    chk8.Enabled = True
    chk9.Enabled = True
    chk10.Enabled = True
    chk11.Enabled = True
    chk12.Enabled = True
    ChkLock.Enabled = True
    
    
End Sub

Private Sub FldDisab()
    txtUserID.Enabled = False
    txtUserName.Enabled = False
    txtUserPswd.Enabled = False
    txtConfm.Enabled = False
    chk1.Enabled = False
    chk2.Enabled = False
    chk3.Enabled = False
    chk4.Enabled = False
    chk5.Enabled = False
    chk6.Enabled = False
    chk7.Enabled = False
    chk8.Enabled = False
    chk9.Enabled = False
    chk10.Enabled = False
    chk11.Enabled = False
    chk12.Enabled = False
    ChkLock.Enabled = False
End Sub
