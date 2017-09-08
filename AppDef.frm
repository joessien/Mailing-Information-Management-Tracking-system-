VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmAppDef 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datapro"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7410
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   0
      Top             =   3600
      Width           =   1095
   End
   Begin Threed.SSPanel SSPanel5 
      Height          =   2775
      Left            =   6240
      TabIndex        =   24
      Top             =   720
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   4895
      _StockProps     =   15
      Caption         =   "Scroll"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Alignment       =   6
      Begin VB.CommandButton cmdBrow 
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
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton cmdTop 
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
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdBott 
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
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdNext 
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
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdPrev 
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
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdFind 
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
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   855
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   615
      Left            =   0
      TabIndex        =   19
      Top             =   3600
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   1085
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdCanc 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   120
         Width           =   855
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   2775
      Left            =   0
      TabIndex        =   17
      Top             =   720
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   4895
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   5895
         Begin MSComDlg.CommonDialog GetFile 
            Left            =   3960
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox txtImpDesc 
            Height          =   285
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   2
            Top             =   720
            Width           =   2535
         End
         Begin MSMask.MaskEdBox IntBatNo 
            Height          =   255
            Left            =   1920
            TabIndex        =   1
            ToolTipText     =   "Enter value between  900000 to 999998"
            Top             =   360
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   615
            Left            =   4800
            TabIndex        =   20
            Top             =   1560
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   1085
            _StockProps     =   15
            BackColor       =   8421376
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
            BevelInner      =   2
            Begin VB.CommandButton cmdOk 
               BackColor       =   &H0000FF00&
               Caption         =   "OK"
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
               Left            =   120
               MaskColor       =   &H000000FF&
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.TextBox txtTable 
            Height          =   285
            Left            =   1920
            MaxLength       =   15
            TabIndex        =   5
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CommandButton cmdSelPath 
            BackColor       =   &H0000FF00&
            Caption         =   "Select Database"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Select the Database of the Source Table"
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label txtImpPath 
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1920
            TabIndex        =   4
            Top             =   1200
            Width           =   3615
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Selected Table:"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Batch Number:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Application Description:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   1815
         End
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7335
      _Version        =   65536
      _ExtentX        =   12938
      _ExtentY        =   1085
      _StockProps     =   15
      Caption         =   "Interface Definition"
      BackColor       =   14737632
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
   End
End
Attribute VB_Name = "frmAppDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GenClass As New DProLog
Dim db2 As Database, db1 As Database
Dim rstImpTab As Recordset, rstImpTrn As Recordset
Dim mvBuff As String, rst As Recordset
Dim mvReply As String, mvErrFlg As Integer
Dim mvLRec As String

Private Sub CmdDisab()
    cmdtop.Enabled = False
    cmdnext.Enabled = False
    cmdprev.Enabled = False
    cmdbott.Enabled = False
    cmdFind.Enabled = False
    cmdbrow.Enabled = False
End Sub

Private Sub CmdEnab()
    cmdtop.Enabled = True
    cmdnext.Enabled = True
    cmdprev.Enabled = True
    cmdbott.Enabled = True
    cmdFind.Enabled = True
    cmdbrow.Enabled = True
End Sub

Private Sub CmdAdd_Click()
    FldEnab
    ClearData
    IntBatNo.SetFocus
    cmdAdd.Enabled = False
    cmdCanc.Enabled = True
End Sub

Private Sub cmdCanc_Click()
    ClearData
    FldDisab
    cmdAdd.Enabled = True
    cmdCanc.Enabled = False
End Sub

Private Sub CmdDel_Click()
On Error GoTo DelError
    If IntBatNo = 0 Then
        MsgBox "Select Application Batch Number To Delete", vbInformation, Me.Caption
        Exit Sub
    End If
    If rstImpTab.EOF And rstImpTab.BOF Then
        MsgBox "Table is empty", vbInformation, Me.Caption
        Exit Sub
    End If
    mvReply = DeleteCheck()
    If mvReply = vbYes Then
        rstImpTab.Delete
        ClearData
        MsgBox "Application Link Removed", vbInformation, Me.Caption
    Else
        MsgBox "Application Link Not Removed", vbExclamation, Me.Caption
    End If
'   If (rstImpTab.BOF And rstImpTab.EOF) Then
'       Cleardata
'   End If
DelClear:
    Exit Sub
DelError:
    ErrorMessages (NODELETE)
    On Error GoTo 0
    Resume DelClear
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
On Error GoTo Perror
    mvErrFlg = 0
    ValFlds
    If mvErrFlg = 0 Then
        rstImpTab.AddNew
        rstImpTab("impBatno") = IntBatNo
        rstImpTab("impPath") = txtImpPath
        rstImpTab("impTab") = txtTable
        rstImpTab("impDesc") = txtImpDesc
        rstImpTab.Update
        FldDisab
        CmdEnab
    End If
Perror:
End Sub

Private Sub Form_Load()
On Error GoTo Perror
     Me.Caption = coyname
     GenClass.fleLogin mvUserid, "Access Application Link", Date, Time
    Call FormCentreMDI(Me)
     Screen.MousePointer = vbHourglass
     Set db2 = OpenDatabase(DBpath)
     mvBuff = "Select * from ImpTab Order By ImpBatno;"
     Set rstImpTab = db2.OpenRecordset(mvBuff, dbOpenDynaset)
     FldDisab
     If rstImpTab.EOF And rstImpTab.BOF Then
        cmdAdd.Enabled = True
        cmdDel.Enabled = False
        cmdCanc.Enabled = False
        CmdDisab
     End If
Perror:
     Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSelPath_Click()
On Error GoTo Perror
    GetFile.FileName = ""
    GetFile.Filter = "All Files (*.*)|*.*|Database Files (*.mdb)|*.mdb"
    GetFile.FilterIndex = 2
    GetFile.DefaultExt = "MDB"
    GetFile.Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
    GetFile.ShowOpen
    'Getfile.Action = 1
    txtImpPath = GetFile.FileName
    If Len(txtImpPath) = 0 Then
        MsgBox "No file was selected", vbInformation, Me.Caption
        Exit Sub
    Else
        txtTable.Enabled = True
        cmdOk.Enabled = True
        txtTable.SetFocus
    End If
Perror:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    GenClass.fleLogin mvUserid, "Exit Application Link", Date, Time
    db2.Close
End Sub
Private Sub IntBatNo_GotFocus()
    IntBatNo.MaxLength = 6
    IntBatNo.SelStart = 0
    IntBatNo.SelLength = Len(IntBatNo)
End Sub

Private Sub IntBatNo_LostFocus()
On Error GoTo Perror
    If IntBatNo = 0 Then
        ClearData
        FldDisab
        cmdAdd.Enabled = True
        cmdCanc.Enabled = False
        Exit Sub
    End If
    If IntBatNo < 999900 Or IntBatNo > 999989 Then
        MsgBox "Batch Number Out of Range [999900 to 999989]", vbCritical, Me.Caption
        IntBatNo.SetFocus
    Else
       mvBuff = "Select * from ImpTab Where ImpBatno = '" _
        & IntBatNo & "'"
       Set rstImpTab = db2.OpenRecordset(mvBuff, dbOpenDynaset)
       If Not (rstImpTab.EOF And rstImpTab.BOF) Then
          MsgBox "Batch Already Defined", vbCritical, Me.Caption
          IntBatNo = 0
          IntBatNo.SetFocus
       End If
    End If
Perror:
End Sub

Private Sub txtTable_gotfocus()
    txtTable.SelStart = 0
    txtTable.SelLength = Len(txtTable)
End Sub

Private Sub FldDisab()
On Error GoTo Perror
    IntBatNo.Enabled = False
    txtTable.Enabled = False
    txtImpDesc.Enabled = False
    cmdOk.Enabled = False
    cmdSelPath.Enabled = False
    If rstImpTab.EOF And rstImpTab.BOF Then
        cmdDel.Enabled = False
    Else
        cmdDel.Enabled = True
    End If
    cmdAdd.Enabled = True
Perror:
End Sub

Private Sub FldEnab()
    IntBatNo.Enabled = True
    txtImpDesc.Enabled = True
    cmdOk.Enabled = False
    cmdSelPath.Enabled = True
    cmdDel.Enabled = False
End Sub

Private Sub ValFlds()
On Error GoTo ImpClr
    If IntBatNo < 999900 Or IntBatNo > 999989 Then
        MsgBox "Batch Number Out of Range [999900 to 999989]", vbCritical, Me.Caption
        IntBatNo.SetFocus
        mvErrFlg = 1
        Exit Sub
    End If
    If Trim(txtImpDesc) = "" Then
        MsgBox "Invalid Application Description", vbCritical, Me.Caption
        txtImpDesc.SetFocus
        mvErrFlg = 1
        Exit Sub
    End If
    If Trim(txtImpPath) = "" Then
        MsgBox "Invalid Application Path", vbCritical, Me.Caption
        txtImpPath.TabIndex = 0
        cmdSelPath.SetFocus
        mvErrFlg = 1
        Exit Sub
    End If
    If Trim(UCase(txtImpPath)) = Trim(UCase(DBpath)) Then
        MsgBox "Selected Path Cannot Be Same With Application", vbCritical, Me.Caption
        txtImpPath.TabIndex = 0
        cmdSelPath.SetFocus
        mvErrFlg = 1
        Exit Sub
    End If
    If Trim(txtTable) = "" Then
        MsgBox "Invalid Table Name", vbCritical, Me.Caption
        txtTable.SetFocus
        mvErrFlg = 1
        Exit Sub
    End If
    On Error GoTo ImpErr
    Set db1 = OpenDatabase(txtImpPath, True)
    Set rstImpTrn = db1.OpenRecordset(txtTable, dbOpenSnapshot)
ImpClr:
    db1.Close
    Exit Sub
ImpErr:
    If Err.Number = 3078 Then
        MsgBox "Table Specified Does Not Exist", vbCritical, Me.Caption
        txtTable.SetFocus
        mvErrFlg = 1
        db1.Close
        Exit Sub
    End If
    On Error Resume Next
    Resume ImpClr
End Sub

Private Sub CmdBrow_Click()
On Error GoTo Perror
    Screen.MousePointer = vbHourglass
    mvBuff = "Select * from ImpTab Order By ImpDesc;"
    Set rst = db2.OpenRecordset(mvBuff, dbOpenSnapshot)
    Set frmBrowse.datGeneral.Recordset = rst
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "Interface Definition"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 2800
         frmBrowse.grdGeneral.Columns(2).Width = 2800
         frmBrowse.grdGeneral.Columns(3).Width = 1200
         Col0.Caption = "Reference"
         Col1.Caption = "Application Description"
         Col2.Caption = "Application's Database Path"
         Col3.Caption = "Import Data Table"
    frmBrowse.Show
Perror:
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdTop_Click()
    mvBuff = "Select * from ImpTab Order By ImpBatno;"
    Set rstImpTab = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    If Not rstImpTab.BOF Then
       rstImpTab.MoveFirst
       showdata
       FldDisab
    End If
End Sub

Private Sub CmdNext_Click()
On Error GoTo Perror
    mvBuff = ""
    If rstImpTab.EOF Or rstImpTab.BOF Then
        mvBuff = ""
    Else
        If Not IsNull(rstImpTab!ImpBatNo) Then
            mvBuff = rstImpTab!ImpBatNo
        End If
    End If
    mvBuff = "Select * from ImpTab Where ImpBatno > '" _
        & mvBuff & "' Order By ImpBatno;"
    Set rstImpTab = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    If Not (rstImpTab.EOF And rstImpTab.BOF) Then
        showdata
        FldDisab
    End If
Perror:
End Sub

Private Sub CmdPrev_Click()
    mvBuff = "Select * from ImpTab Where ImpBatno = '" _
        & mvLRec & "'"
    Set rstImpTab = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    If Not (rstImpTab.EOF And rstImpTab.BOF) Then
        rstImpTab.MoveLast
        showdata
        FldDisab
    End If
End Sub

Private Sub CmdBott_Click()
On Error GoTo Perror
    mvBuff = "Select * from ImpTab Order By ImpBatno;"
    Set rstImpTab = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    If Not (rstImpTab.BOF And rstImpTab.EOF) Then
       rstImpTab.MoveLast
       showdata
       FldDisab
    End If
Perror:
End Sub

Private Sub cmdFind_Click()
On Error GoTo Perror
    Dim mvLength As Integer
    Dim mvBuff As String, mvAns As String
    mvLength = 0
    mvAns = InputBox$("Enter Batch Number:")
    mvLength = Len(mvAns)
    If mvLength = 0 Then
       MsgBox "User Select Cancel or Nothing To Find", vbInformation, Me.Caption
       Exit Sub
    End If
    mvBuff = "Select * from ImpTab Where ImpBatNo = '" _
        & mvAns & "'"
    Set rstImpTab = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    If rstImpTab.EOF And rstImpTab.BOF Then
        MsgBox "Batch Number Not Found", vbInformation, Me.Caption
        Exit Sub
    Else
        showdata
    End If
Perror:
End Sub

Private Sub showdata()
On Error GoTo Perror
     If Not (rstImpTab.EOF And rstImpTab.BOF) Then
        mvLRec = IntBatNo
        IntBatNo = rstImpTab!ImpBatNo
        txtImpDesc = rstImpTab!ImpDesc
        txtImpPath = rstImpTab!ImpPath
        txtTable = rstImpTab!ImpTab
    End If
Perror:
End Sub

Private Sub ClearData()
    IntBatNo = 0
    txtImpDesc = ""
    txtImpPath = ""
    txtTable = ""
End Sub
