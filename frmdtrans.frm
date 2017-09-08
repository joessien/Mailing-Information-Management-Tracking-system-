VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmtransf 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retail Plus"
   ClientHeight    =   5145
   ClientLeft      =   1410
   ClientTop       =   2250
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5145
   ScaleWidth      =   7200
   Begin VB.TextBox reasoN 
      Height          =   855
      Left            =   1680
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   3360
      Width           =   5175
   End
   Begin VB.ComboBox cmbTTo 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2340
      Width           =   3855
   End
   Begin MSMask.MaskEdBox tDate 
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   2880
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.ComboBox cmbStudNo 
      Height          =   315
      Left            =   2280
      TabIndex        =   8
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Data datGenTab 
      Caption         =   "Gen Table"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   795
      Left            =   0
      TabIndex        =   2
      Top             =   4320
      Width           =   7155
      _Version        =   65536
      _ExtentX        =   12621
      _ExtentY        =   1402
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
         Caption         =   "E&xit"
         Height          =   375
         Left            =   5880
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4200
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "&Transfer"
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   1215
      Left            =   5640
      TabIndex        =   4
      Top             =   1800
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   2143
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
         Caption         =   "List"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdfind 
         Caption         =   "&Find"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   975
      End
   End
   Begin Threed.SSPanel sstrans 
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   6975
      _Version        =   65536
      _ExtentX        =   12303
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "Student Class Transfer"
      BackColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   4
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label TFrom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   1680
      TabIndex        =   15
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   360
      TabIndex        =   13
      Top             =   2355
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   7560
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer from:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   435
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "Student Number:"
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
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "frmtransf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MaxErrors As Integer, strpos As String
Dim db1 As Database, wrktemp As Workspace
Dim rstTrans As Recordset
Dim rstStudmast As Recordset
Dim mvCCode, MVstNames As String
Dim rstCNames As Recordset
Dim tempBuff As String, NameBuff As String
Dim StudBuff As String, EditFlag As Boolean
Dim AddFlag As Boolean, GenClass As New DProLog

Private Sub cmbStudNo_Click()
'On Error GoTo PError
    If AddFlag = False Then
       ClearData
       strpos = InStr(1, cmbStudNo, ",", 1)
       StudBuff = "StudNo = '" & Left(cmbStudNo, strpos - 1) & "'"
       rstStudmast.FindFirst StudBuff
       If rstStudmast.NOMATCH Then
           MsgBox "Student has no previous class assigned", vbCritical, "Critical Error"
           Exit Sub
        Else
           mvCCode = Trim(rstStudmast!cclass)
           GTBuff = "Code = '" & UCase(mvCCode) & "'"
           rstCNames.FindFirst GTBuff
           If rstCNames.NOMATCH Then
               MsgBox "Critical Error: Class code not defined but specified as ClassRoom", vbCritical, "Critical Error"
           Else
              TFrom.Caption = rstCNames!Code & ",  " & rstCNames!Desc & " - " & rstCNames!CGrp
              MVstNames = rstCNames!Code & ",  " & rstCNames!Desc & " - " & rstCNames!CGrp
           End If
        End If
    Else
        ClearData
    End If
perror:
End Sub

Private Sub cmbStudNo_LostFocus()
'On Error GoTo PError
   strpos = InStr(1, cmbStudNo, ",", 1)
   If strpos = 0 Then
     StudBuff = "StudNo = '" & UCase(cmbStudNo) & "'"
   Else
     StudBuff = "StudNo = '" & Left(cmbStudNo, strpos - 1) & "'"
   End If
    rstStudmast.FindFirst StudBuff
    If rstStudmast.NOMATCH Then
       Exit Sub
    Else
       mvCCode = Trim(rstStudmast!cclass)
       GTBuff = "Code = '" & UCase(mvCCode) & "'"
       rstCNames.FindFirst GTBuff
       If rstCNames.NOMATCH Then
           MsgBox "Critical Error: Class code not defined but specified as ClassRoom", vbCritical, "Critical Error"
       Else
           TFrom.Caption = rstCNames!Code & ",  " & rstCNames!Desc & " - " & rstCNames!CGrp
           MVstNames = rstCNames!Code & ",  " & rstCNames!Desc & " - " & rstCNames!CGrp
       End If
    End If
perror:
End Sub

Private Sub CmdAdd_Click()
On Error GoTo perror
If cmdadd.Caption = "&Transfer" Then
    FldEnab
   rstTrans.AddNew
   AddFlag = True
   cmbStudNo.SetFocus
   cmdSave.Enabled = True
   CmdDisab
   cmdSave.Enabled = True
   cmdadd.Enabled = True
   cmdadd.Caption = "&Cancel"
Else
    FldDisab
   cmdadd.Caption = "&Transfer"
   CmdEnab
   cmdSave.Enabled = False
   AddFlag = False
End If
perror:
End Sub

Private Sub Cmdfind_Click()
  Dim sql As String, digitflag As Integer, length As Integer
   Dim AN As String
   'On Error GoTo FindError1
   AN = InputBox$("Enter the Student number to find:")
   length = Len(AN)
   If length = 0 Then
      ErrorMessages (NOINPUT)
      Exit Sub
   End If
   If length = 1 Then
      StudBuff = "StudNo LIKE " & Chr$(34) & AN & "*" & Chr$(34)
   Else
      StudBuff = "StudNo LIKE " & Chr$(34) & AN & "*" & Chr$(34)
   End If
    rstStudmast.FindFirst StudBuff
    If rstStudmast.NOMATCH Then
       MsgBox "Student number not found", vbCritical, "Critical Error"
       Exit Sub
    Else
       cmbStudNo = rstStudmast("StudNo") + ", " + rstStudmast("StudNames")
       mvCCode = Trim(rstStudmast!cclass)
       GTBuff = "Code = '" & UCase(mvCCode) & "'"
       rstCNames.FindFirst GTBuff
       If rstCNames.NOMATCH Then
           MsgBox "Critical Error: Class code not defined but specified as ClassRoom", vbCritical, "Critical Error"
       Else
           TFrom.Caption = rstCNames!Code & ",  " & rstCNames!Desc & " - " & rstCNames!CGrp
           MVstNames = rstCNames!Code & ",  " & rstCNames!Desc & " - " & rstCNames!CGrp
       End If
    End If

FindError0:
   Exit Sub

FindError1:
   MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
   On Error GoTo 0
   Resume FindError0

End Sub

Private Sub CmdSave_Click()
    Dim i As Integer
    'On Error GoTo PError
    If Len(Trim(reasoN)) = 0 Then
       Beep
       MsgBox "Reason for transfer cannot be blank", vbExclamation, "Student Records"
       reasoN.SetFocus
       Exit Sub
    End If
    If Len(tDate) = 0 Then
       Beep
       MsgBox "Transfer date cannot be blank", vbExclamation, "Student Records"
       tDate.SetFocus
       Exit Sub
    End If
    If Left(cmbTTo, strpos - 1) = mvCCode Then
       Beep
       MsgBox "Both classes are same. Transfer cannot be done to same class", vbExclamation, "Student Records"
       Exit Sub
    End If
     If EditFlag = True Then
        rstTrans.Edit
        EditFlag = False
    End If
    If AddFlag = True Then
        rstTrans.AddNew
        strpos = InStr(1, cmbStudNo, ",", 1)
        rstTrans("StudNo") = Left(cmbStudNo, strpos - 1)
        StudBuff = Trim(Left(cmbStudNo, strpos - 1))
        AddFlag = False
    End If
    rstTrans("TFrom") = mvCCode
    rstTrans("TFromD") = rstCNames!Desc
    strpos = InStr(1, cmbTTo, ",", 1)
    rstTrans("TTo") = Left(cmbTTo, strpos - 1)
    rstTrans("TToD") = Trim(Mid(cmbTTo, strpos + 2))
    rstTrans("reason") = reasoN
    rstTrans("tdate") = tDate
    rstTrans("staffno") = mvUserid
    rstTrans.Update
    strpos = InStr(1, cmbStudNo, ",", 1)
    GTBuff = "studno = '" & UCase(Left(cmbStudNo, strpos - 1)) & "'"
     rstStudmast.FindFirst GTBuff
     If rstStudmast.NOMATCH Then
         MsgBox "Student Record is not found in Students master file", vbCritical, Me.Caption
     Else
         rstStudmast.Edit
         strpos = InStr(1, cmbTTo, ",", 1)
         rstStudmast("cclass") = Left(cmbTTo, strpos - 1)
         rstStudmast.Update
   End If
    cmdadd.Caption = "&Transfer"
    cmbStudNo.Enabled = True
    CmdEnab
    AddFlag = False
    cmdSave.Enabled = False
    rstTrans.MoveLast
    FldDisab
PError0:
    Exit Sub
    
perror:
    MsgBox "Invalid data or Duplicated Entry, Check.", vbCritical, Me.Caption
    On Error GoTo 0
    Resume PError0

End Sub

Private Sub Form_Load()
On Error GoTo perror
     GenClass.fleLogin mvUserid, "Accessed Student Departmental transfer", Date, Time
     Call FormCentreMDI(Me)
     Set wrktemp = DBEngine.Workspaces(0)
     Set db1 = wrktemp.OpenDatabase(DBpath, True)
     Set rstCNames = db1.OpenRecordset("defclassnames", dbOpenDynaset)
     Set rstStudmast = db1.OpenRecordset("studmast", dbOpenDynaset)
     Set rstTrans = db1.OpenRecordset("studtrans", dbOpenDynaset)
     Me.Caption = MVCoyname

        ''__________________________________
        datGenTab.DatabaseName = DBpath
        datGenTab.RecordSource = "defclassnames"
        datGenTab.ReadOnly = False
        datGenTab.Exclusive = False
        datGenTab.Refresh
        Do While Not datGenTab.Recordset.EOF
           cmbTTo.AddItem datGenTab.Recordset("code") + ", " + datGenTab.Recordset("desc") + " - " + datGenTab.Recordset("cgrp")
           datGenTab.Recordset.MoveNext
        Loop
        If datGenTab.Recordset.RecordCount > 1 Then
           cmbTTo.ListIndex = 0
        End If
        
        
    ''__________________________________
    datGenTab.DatabaseName = DBpath
    datGenTab.RecordSource = "studmast"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       cmbStudNo.AddItem datGenTab.Recordset("StudNo") + ", " + datGenTab.Recordset("StudNames")
       datGenTab.Recordset.MoveNext
    Loop
        If datGenTab.Recordset.RecordCount > 1 Then
           cmbStudNo.ListIndex = 0
        End If
    ''__________________________________
    FldDisab
    ClearData
    AddFlag = False
    EditFlag = False
    cmdSave.Enabled = False
perror:
End Sub

Private Sub CmdBrow_Click()
'On Error GoTo PError
    cmdSave.Enabled = False
    AddFlag = False
    mvSql = "Select * From studtrans"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No students transfer historyd ", vbExclamation, "Class Transfer "
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "Students Transfer History"
         frmBrowse.Caption = Me.Caption & " - Records View"
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         frmBrowse.grdGeneral.Columns(0).Width = 1400
         frmBrowse.grdGeneral.Columns(1).Width = 2200
         frmBrowse.grdGeneral.Columns(2).Width = 800
         frmBrowse.grdGeneral.Columns(3).Width = 2200
         frmBrowse.grdGeneral.Columns(4).Width = 800
         frmBrowse.grdGeneral.Columns(5).Width = 2200
         frmBrowse.grdGeneral.Columns(6).Width = 2800
         frmBrowse.grdGeneral.Columns(7).Width = 1000
         frmBrowse.grdGeneral.Columns(7).Width = 1400
         Col0.Caption = "Student Number"
         Col1.Caption = "Student Name"
         Col2.Caption = "CF Code"
         Col3.Caption = "Trf From Class Description"
         Col4.Caption = "Trf To Code"
         Col5.Caption = "Trf To Class Description"
         Col6.Caption = "Reason"
         Col7.Caption = "Transfer Date"
         Col8.Caption = "Transfered By"
   End If
   frmBrowse.Show
perror:
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub


Public Sub ClearData()
        tDate = Date
        cmbTTo = ""
        TFrom.Caption = ""
        reasoN = ""
End Sub

Public Sub CmdEnab()
    cmdadd.Enabled = True
    cmdbrow.Enabled = True
    cmdFind.Enabled = True
End Sub

Public Sub CmdDisab()
    cmdadd.Enabled = False
    cmdbrow.Enabled = False
    cmdFind.Enabled = False
End Sub


Public Sub FldEnab()
    cmbStudNo.Enabled = True
    cmbTTo.Enabled = True
    tDate.Enabled = True
End Sub

Public Sub FldDisab()
    cmbStudNo.Enabled = True
    cmbTTo.Enabled = False
    tDate.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    db1.Close
End Sub
