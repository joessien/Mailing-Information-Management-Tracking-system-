VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form FrmDPProc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Ledger "
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5925
   Begin Threed.SSPanel SSPanel2 
      Height          =   3975
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   5895
      _Version        =   65536
      _ExtentX        =   10398
      _ExtentY        =   7011
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
      Begin VB.ComboBox cmbRAccNum 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   10
         Top             =   120
         Width           =   3375
      End
      Begin Threed.SSPanel SSPer 
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   5415
         _Version        =   65536
         _ExtentX        =   9551
         _ExtentY        =   873
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
         Begin VB.Label lblProg 
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   135
         End
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00808080&
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Click To Exit Process"
         Top             =   3480
         Width           =   1095
      End
      Begin Threed.SSPanel SSPYEnd 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Select the Retain Earning Account for End of Year Closing"
         Top             =   120
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Retain Earning Account:"
         BackColor       =   8421376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   1800
         Width           =   5655
         _Version        =   65536
         _ExtentX        =   9975
         _ExtentY        =   2778
         _StockProps     =   15
         Caption         =   "Processing Details"
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
         Outline         =   -1  'True
         Font3D          =   1
         Alignment       =   6
         Begin VB.CommandButton cmdCont 
            Caption         =   "&Continue"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   "Click To Commence Processing"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "&Process"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4320
            TabIndex        =   6
            ToolTipText     =   "Click To Commence Processing"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Type of Processing:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label txtProType 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2040
            TabIndex        =   12
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label txtProDate 
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2040
            TabIndex        =   4
            Top             =   720
            Width           =   3375
         End
         Begin VB.Label Label1 
            Caption         =   "Date:"
            Height          =   255
            Left            =   1080
            TabIndex        =   3
            Top             =   720
            Width           =   375
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Retain Earning Account Desc."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label LblActivity 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   5415
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   885
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _Version        =   65536
      _ExtentX        =   10398
      _ExtentY        =   1561
      _StockProps     =   15
      Caption         =   "PROCESSING"
      BackColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   2
      Font3D          =   2
      Begin Threed.SSPanel SSPanel4 
         Height          =   855
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1508
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
         Begin VB.Image Image2 
            Height          =   720
            Left            =   120
            Picture         =   "frmDPproc.frx":0000
            Stretch         =   -1  'True
            Top             =   120
            Width           =   495
         End
      End
   End
End
Attribute VB_Name = "FrmDPProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GenClass As New GLClass, rstGLPL As Recordset
Dim db2 As Database, mvErrFlg As Integer
Dim dbs1 As Database, mvRTyp As Integer
Dim rstGLMast As Recordset, rstGLTrn As Recordset
Dim rstGLRecr As Recordset, rstGLCtrl As Recordset
Dim rstGLHist As Recordset, rstAccTab As Recordset
Dim rstAutoTrn As Recordset, rstAutoCtrl As Recordset
Dim mvperinc As Single, mvMsg As String, mvDat As Date
Dim mvBuff As String, mvperctr As Single
Dim mvPL As Currency, mvGLHDR As String, mvGCode As String
Dim mvStrPos As String, mvTyp As String, mvCode As String
Dim lblPer As Byte, mvRecAll As Long, mvCls As Integer
Dim rstTBalance As Recordset, mvALMBal As Currency, mvAcurBal As Currency
Dim mvBLMBal As Currency, mvBcurBal As Currency
Dim mvCLMBal As Currency, mvCcurBal As Currency
Dim mvDLMBal As Currency, mvDcurBal As Currency
Dim mvELMBal As Currency, mvEcurBal As Currency
Dim mvFLMBal As Currency, mvFcurBal As Currency
Dim mvGLMBal As Currency, mvGcurBal As Currency
Dim IntBatNo As Long, rstCurRecr As Recordset

Private Sub OpnBalUpd()
    mvBuff = "Select * from GLMast Order By GLNumber;"
    Set rstGLMast = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    mvBuff = "Select * from GLTrnHst Where GLNumber = '';"
    Set rstGLHist = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    LblActivity.Caption = "Generating A/C Opening Balance ....."
    LblActivity.Refresh
    lblProg.Width = 0
    mvperctr = 1
    If Not (rstGLMast.EOF And rstGLMast.BOF) Then
        rstGLMast.MoveLast
        mvperinc = 100 / rstGLMast.RecordCount
        rstGLMast.MoveFirst
    End If
    While Not rstGLMast.EOF
        rstGLHist.AddNew
        rstGLHist("GLNumber") = UCase(rstGLMast!GLNumber)
        rstGLHist("GLTrnDat") = Format(mvDat, "dd-mm-yyyy")
        If IsNull(rstGLMast!GLBal) Then
            rstGLHist("GLTrnAmt") = 0
        Else
            rstGLHist("GLTrnAmt") = rstGLMast!GLBal
        End If
        rstGLHist("GLPartlars") = " "
        rstGLHist("GLInstno") = " "
        rstGLHist("GLAcctDat") = rstGLHist!GLtrnDat
        rstGLHist.Update
        rstGLMast.MoveNext
        UpdatePercentDialog (mvperctr)
        mvperctr = mvperctr + 1
    Wend
    rstGLHist.Close
    rstGLMast.Close
    lblProg.Caption = "100%"
End Sub

Private Sub cmdexit_Click()
    mvCls = 0
    GenClass.fleLogin mvUserid, "Quit General Ledger Processing", Date, Time
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mvCls <> 0 Then
        MsgBox "You Cannot Terminate Application Prematurely", vbCritical
        Cancel = True
        'Resume
    Else
        Cancel = mvCls
    End If
End Sub

Private Sub cmdok_Click()
    cmdexit.Enabled = False
    cmdok.Visible = False
    If rstGLCtrl!EODCtrl <> "2" Then
        MsgBox "Recurrent Entries Yet To Be Activated", vbCritical
        mvCls = 0
        Unload Me
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    LblActivity.Visible = True
    SSPer.Visible = True
    txtProDate.Caption = Format(rstGLCtrl!CAcctDate, "dd/mm/yyyy")
    txtProType.Caption = "End of Day Processing"
    mvCls = 1
    'Import Subsidiary Ledger Entries
    If Trim(rstGLCtrl!SubPath) <> "" Then
        Set dbs1 = OpenDatabase(Trim(rstGLCtrl("SubPath")))
        mvBuff = "Select * from " _
            & Trim(rstGLCtrl!SubCtrl)
        Set rstAutoCtrl = dbs1.OpenRecordset(mvBuff, dbOpenDynaset)
        mvBuff = "Select * from " _
            & Trim(rstGLCtrl!SubTrn)
        Set rstAutoTrn = dbs1.OpenRecordset(mvBuff, dbOpenSnapshot)
        If rstAutoCtrl!ProcCtrl = 1 Then
            LblActivity.Caption = "Extracting Subsidiary Ledger Entries ......"
            IntBatNo = 999996
            ImpSubTrn
            rstAutoCtrl.Edit
            rstAutoCtrl("ProcCtrl") = 0
            rstAutoCtrl("EODCtrl") = 0
            rstAutoCtrl.Update
        End If
'       rstAutoTrn.Close
'       rstAutoCtrl.Close
        dbs1.Close
    End If
    'Import Auxiliary Entries
    If Trim(rstGLCtrl!AutoPath) <> "" Then
        Set dbs1 = OpenDatabase(Trim(rstGLCtrl!AutoPath))
        mvBuff = "Select * from " _
            & Trim(rstGLCtrl!AutoCtrl)
        Set rstAutoCtrl = dbs1.OpenRecordset(mvBuff, dbOpenDynaset)
        mvBuff = "Select * from " _
            & Trim(rstGLCtrl!AutoTrn)
        Set rstAutoTrn = dbs1.OpenRecordset(mvBuff, dbOpenSnapshot)
        If rstAutoCtrl!ProcCtrl = 1 Then
            LblActivity.Caption = "Extracting Auxiliary Entries ......"
            IntBatNo = 999997
            ImpSubTrn
            rstAutoCtrl.Edit
            rstAutoCtrl("ProcCtrl") = 0
            rstAutoCtrl.Update
        End If
'       rstAutoTrn.Close
'       rstAutoCtrl.Close
        dbs1.Close
    End If
    'Set Update Control
    rstGLCtrl.Edit
    rstGLCtrl("EODCtrl") = "1"
    rstGLCtrl.Update
    mvDat = rstGLCtrl!NAcctDate
    LblActivity.Caption = "Extracting Due Recurrent Entries ........"
    LblActivity.Refresh
    mvBuff = "Delete * from CurRecr;"
    db2.Execute mvBuff
    mvBuff = Format(rstGLCtrl!NAcctDate, "m-d-yy")
    mvBuff = "Select * from GLRecr Where " _
        & "GLDueDat < #" & mvBuff & "# And " _
        & "GLExpDat >= #" & mvBuff & "#"
    Set rstGLRecr = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    mvBuff = "Select * from CurRecr;"
    Set rstCurRecr = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    mvBuff = "Select * from GLTrn;"
    Set rstGLTrn = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    lblProg.Width = 0
    mvperctr = 1
    If Not (rstGLRecr.EOF And rstGLRecr.BOF) Then
        rstGLRecr.MoveLast
        mvperinc = 100 / rstGLRecr.RecordCount
        rstGLRecr.MoveFirst
    End If
    While Not rstGLRecr.EOF
        If rstGLRecr!PostInd Then
            ' Update CurRecr with transaction
            UpdCur
            ' Create Debit Leg
            rstGLTrn.AddNew
            rstGLTrn("GLNumber") = rstGLRecr!GLDrNumber
            rstGLTrn("GLCentCode") = rstGLRecr!GLDRCentCode
            rstGLTrn("GLTrnAmt") = -1 * rstGLRecr!GLTrnamt
            RecFld
            ' Create Credit Leg
            rstGLTrn.AddNew
            rstGLTrn("GLNumber") = rstGLRecr!GLCrNumber
            rstGLTrn("GLCentCode") = rstGLRecr!GLCRCentCode
            rstGLTrn("GLTrnAmt") = rstGLRecr!GLTrnamt
            RecFld
            ' Reset Entry Next Due Date
            rstGLRecr.Edit
            rstGLRecr("PostInd") = False
            rstGLRecr("GLDueDat") = rstGLRecr!GLDueDat _
                                    + rstGLRecr!GLFreq
            rstGLRecr.Update
            UpdatePercentDialog (mvperctr)
            mvperctr = mvperctr + 1
        End If
        rstGLRecr.MoveNext
    Wend
    rstGLRecr.Close
    mvRTyp = 1
    lblProg.Caption = "100%"
'   Set rstGLMast = db2.OpenRecordset("GLMast", dbOpenDynaset)
    PostTrn
    ' Check for Monthend & Process if Monthend
    If Month(rstGLCtrl!NAcctDate) <> _
                    Month(rstGLCtrl!CAcctDate) Then
        mvRTyp = 2
        txtProType.Caption = "End of Month Processing"
        txtProType.Refresh
        ' Set Opening Balance To G/L Balance
        Screen.MousePointer = vbHourglass
        LblActivity.Caption = "Reseting Account Opening Balance ....."
        LblActivity.Refresh
        lblProg.Width = 0
        mvperctr = 1
        mvBuff = "Select * from GLMast Order By GLNumber;"
        Set rstGLMast = db2.OpenRecordset(mvBuff, dbOpenDynaset)
        If Not (rstGLMast.EOF And rstGLMast.BOF) Then
            rstGLMast.MoveLast
            mvperinc = 100 / rstGLMast.RecordCount
            rstGLMast.MoveFirst
        End If
        While Not rstGLMast.EOF
            rstGLMast.Edit
            rstGLMast("GLOpnBbal") = rstGLMast!GLBal
            rstGLMast.Update
            UpdatePercentDialog (mvperctr)
            mvperctr = mvperctr + 1
            rstGLMast.MoveNext
        Wend
        rstGLMast.Close
        lblProg.Caption = "100%"
        
        '' Set Opening Balance into History
        OpnBalUpd
        rstGLCtrl.Edit
        rstGLCtrl("LBatno") = 0
        rstGLCtrl.Update
        Screen.MousePointer = vbDefault
        
        ' Check for Year End & Process Year End Activities
        If Month(rstGLCtrl!CAcctDate) = _
                            rstGLCtrl!EoyMth Then
            mvRTyp = 3
            txtProType.Caption = "End of Year Processing"
            SSPYEnd.Visible = True
            SSPer.Visible = True
            LblActivity.Visible = False
            lblProg.Visible = False
            Label3.Caption = ""
            Label3.Visible = True
            cmbRAccNum.Visible = True
            cmdCont.Visible = True
            cmbRAccNum.Clear
            mvBuff = "Select * from GLMast Where " _
                & "GLTyp = 'C' And GLLev <> 'H' " _
                & "Order By GLNar;"
            Set rstGLMast = db2.OpenRecordset(mvBuff, dbOpenDynaset)
            If Not (rstGLMast.EOF And _
                            rstGLMast.BOF) Then
                rstGLMast.MoveFirst
            End If
            While Not rstGLMast.EOF
                If rstGLMast!GLTyp = "C" And rstGLMast!GLLev <> "H" Then
                    cmbRAccNum.AddItem _
                        rstGLMast!GLNar + _
                        " | " + rstGLMast!GLNumber
                End If
                rstGLMast.MoveNext
            Wend
            LblActivity.Caption = ""
            lblProg.Caption = ""
            LblActivity.Refresh
            lblProg.Width = 0
            LblActivity.Visible = True
            lblProg.Visible = True
            cmbRAccNum.TabIndex = 0
            cmdCont.TabIndex = 1
            cmbRAccNum.SetFocus
            Exit Sub
        End If
    End If
    rstGLCtrl.Edit
    rstGLCtrl("LAcctDate") = rstGLCtrl!CAcctDate
    rstGLCtrl("CAcctDate") = rstGLCtrl!NAcctDate
    rstGLCtrl.Update
    mvCls = 0
    Unload Me
End Sub

Private Sub UpdCur()
    rstCurRecr.AddNew
    rstCurRecr("GLCRCentCode") = rstGLRecr!GLCRCentCode
    rstCurRecr("GLCRNumber") = rstGLRecr!GLCrNumber
    rstCurRecr("GLDRCentCode") = rstGLRecr!GLDRCentCode
    rstCurRecr("GLDRNumber") = rstGLRecr!GLDrNumber
    rstCurRecr("GLPartlars") = rstGLRecr!GLPartLars
    rstCurRecr("GLRecNo") = rstGLRecr!GLRecNo
    rstCurRecr("GLDueDat") = rstGLRecr!GLDueDat
    rstCurRecr("GLExpDat") = rstGLRecr!GLExpDat
    rstCurRecr("GLTrnAmt") = rstGLRecr!GLTrnamt
    rstCurRecr("GLCreDat") = rstGLRecr!GLCreDat
    rstCurRecr("GLFreq") = rstGLRecr!GLFreq
    rstCurRecr.Update
End Sub

Private Sub ImpSubTrn()
    LblActivity.Refresh
    mvBuff = "Select * from GLTrn;"
    Set rstGLTrn = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    lblProg.Width = 0
    mvperctr = 1
    If Not (rstAutoTrn.EOF And rstAutoTrn.BOF) Then
        rstAutoTrn.MoveLast
        mvperinc = 100 / rstAutoTrn.RecordCount
        rstAutoTrn.MoveFirst
    End If
    While Not rstAutoTrn.EOF
        rstGLTrn.AddNew
        rstGLTrn("GLNumber") = rstAutoTrn!GLNumber
        rstGLTrn("GLCentCode") = rstAutoTrn!GLCentCode
        rstGLTrn("GLTrnAmt") = rstAutoTrn!GLTrnamt
        rstGLTrn("GLPartlars") = rstAutoTrn!GLPartLars
        rstGLTrn("GLInstno") = rstAutoTrn!GLInstno
        rstGLTrn("GLTrnDat") = rstAutoTrn!GLtrnDat
        rstGLTrn("GLBatno") = IntBatNo
        rstGLTrn.Update
        UpdatePercentDialog (mvperctr)
        mvperctr = mvperctr + 1
        rstAutoTrn.MoveNext
    Wend
    rstGLTrn.Close
    lblProg.Caption = "100%"
End Sub

Private Sub cmbRAccNum_GotFocus()
    cmbRAccNum.SelStart = 0
    cmbRAccNum.SelLength = Len(cmbRAccNum)
End Sub

Private Sub cmbRAccNum_LostFocus()
    mvErrFlg = 0
    ChkAccNo
    If mvErrFlg = 1 Then
        MsgBox mvMsg, vbInformation
        cmbRAccNum.SetFocus
        Exit Sub
    End If
    If MsgBox("Okay To Commence End Of Year Reversal?", _
        vbQuestion + vbYesNo, Me.Caption) = vbYes Then
        cmdCont.Visible = False
        Screen.MousePointer = vbHourglass
        mvBuff = "Select * from GLTrn;"
        Set rstGLTrn = db2.OpenRecordset(mvBuff, dbOpenDynaset)
        LblActivity.Caption = "Generating Reversal Transactions ....."
        LblActivity.Refresh
        GenRev
        PostTrn
        rstGLCtrl.Edit
        rstGLCtrl("LAcctDate") = rstGLCtrl!CAcctDate
        rstGLCtrl("CAcctDate") = rstGLCtrl!NAcctDate
        rstGLCtrl.Update
        mvCls = 0
        Screen.MousePointer = vbDefault
        Unload Me
    Else
        cmbRAccNum.SetFocus
        mvCls = 1
    End If
End Sub

Public Sub ChkAccNo()
    cmbRAccNum = Trim(cmbRAccNum)
    If cmbRAccNum = "" Then
        mvMsg = "Account Number Cannot Be Blank"
        mvErrFlg = 1
        Exit Sub
    End If
    mvStrPos = InStr(cmbRAccNum, ",") + 1
    If mvStrPos <= 0 Then
       mvStrPos = 1
    End If
    cmbRAccNum = Trim(Mid(cmbRAccNum, mvStrPos))
    mvBuff = "Select * from GLMast Where GLNumber = '" _
        & cmbRAccNum & "'"
    Set rstGLMast = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    If rstGLMast.EOF And rstGLMast.BOF Then
        mvMsg = "Account Number Does Not Exist"
        mvErrFlg = 1
     Else
        Label3.Caption = rstGLMast!GLNar
    '_____ Check if account is control account
        If rstGLMast!GLLev = "H" Or _
            cmbRAccNum = rstGLCtrl!PlaccNum Or _
            rstGLMast!GLTyp <> "C" Then
            mvMsg = "Retain Earning Must Be A Capital Account"
            mvErrFlg = 1
        End If
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo LDErr
    Me.Caption = coyName
    Screen.MousePointer = vbHourglass
    Set db2 = OpenDatabase(mvDBpath, True)
    GenClass.fleLogin mvUserid, "Access General Ledger Processing", Date, Time
    GenClass.frmCentre FrmGLProc
    mvBuff = "Select * from GLCtrl;"
    Set rstGLCtrl = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    mvBuff = "Select * from GLRecr;"
    Set rstGLRecr = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    txtProDate.Caption = Format(rstGLCtrl!CAcctDate, "dd/mm/yyyy")
    txtProType.Caption = "End of Day Processing"
    ''_________________________________
    SSPYEnd.Visible = False
    SSPer.Visible = False
    lblPer = lblProg.Width
    Label3.Visible = False
    cmbRAccNum.Visible = False
    cmdCont.Visible = False
    mvBuff = "Delete * from Tbalance;"
    db2.Execute mvBuff
    mvBuff = "Delete * from GLPL;"
    db2.Execute mvBuff
    cmdexit.TabIndex = 0
LDClr:
    Screen.MousePointer = vbDefault
    Exit Sub
LDErr:
    If Err.Number = 3356 Then
        Screen.MousePointer = vbDefault
        MsgBox "Database Is In Use, Please Try Later", vbCritical
        SSPYEnd.Visible = False
        SSPer.Visible = False
        Label3.Visible = False
        cmbRAccNum.Visible = False
        cmdCont.Visible = False
        cmdok.Enabled = False
        cmdexit.TabIndex = 0
        Resume LDClr
    Else
        MsgBox "Error: " + Error$ + " Has Occurred", vbExclamation, Me.Caption
    End If
    On Error Resume Next
    Resume LDClr
End Sub

Private Sub RecFld()
    rstGLTrn("GLPartlars") = rstGLRecr!GLPartLars
    rstGLTrn("GLInstno") = rstGLRecr!GLRecNo
    rstGLTrn("GLTrnDat") = rstGLRecr!GLDueDat
    rstGLTrn("GLBatno") = 999999
    rstGLTrn.Update
End Sub

Private Sub PostTrn()
    ' Update Trial Balance
    mvBuff = "Select * from TBalance;"
    Set rstTBalance = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    LblActivity.Caption = "Setting Trial Balance ....."
    LblActivity.Refresh
    lblProg.Width = 0
    mvperctr = 1
    mvGLHDR = " "
    mvBuff = "Select * from GLMast Order By GLNumber;"
    Set rstGLMast = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    If Not (rstGLMast.EOF And rstGLMast.BOF) Then
        rstGLMast.MoveLast
        mvperinc = 100 / rstGLMast.RecordCount
        rstGLMast.MoveFirst
    End If
    While Not rstGLMast.EOF
        If rstGLMast!GLLev = "H" Then
            mvGLHDR = rstGLMast!GLNumber
        End If
        rstTBalance.AddNew
        rstTBalance("RTyp") = mvRTyp
        rstTBalance("GLHDR") = mvGLHDR
        rstTBalance("GLNumber") = rstGLMast!GLNumber
        rstTBalance("GLNar") = rstGLMast!GLNar
        rstTBalance("GLOpnBal") = rstGLMast!GLBal
        rstTBalance("GLOpnBbal") = rstGLMast!GLOpnBbal
        rstTBalance("DebitAmt") = 0
        rstTBalance("CreditAmt") = 0
        rstTBalance("GLBal") = 0
        rstTBalance.Update
        rstGLMast.MoveNext
        UpdatePercentDialog (mvperctr)
        mvperctr = mvperctr + 1
    Wend
    lblProg.Caption = "100%"
    
    ' Update G/L Masterfile with Transaction
    LblActivity.Caption = "Posting Transactions ....."
    LblActivity.Refresh
    lblProg.Width = 0
    mvperctr = 1
    mvBuff = "Select * from GLTrn;"
    Set rstGLTrn = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    If Not (rstGLTrn.EOF And rstGLTrn.BOF) Then
        rstGLTrn.MoveLast
        mvperinc = 100 / rstGLTrn.RecordCount
        rstGLTrn.MoveFirst
    End If
    While Not rstGLTrn.EOF
        mvBuff = "Select * from GLMast Where GLNumber = '" _
            & rstGLTrn!GLNumber & "'"
        Set rstGLMast = db2.OpenRecordset(mvBuff, dbOpenDynaset)
        If Not (rstGLMast.EOF And rstGLMast.BOF) Then
            rstGLMast.Edit
            If IsNull(rstGLMast!GLBal) Then
                rstGLMast("GLBal") = rstGLTrn!GLTrnamt
            Else
                rstGLMast("GLBal") = rstGLMast("GLBal") + _
                                        rstGLTrn!GLTrnamt
            End If
            rstGLMast("GLLMvtDat") = rstGLCtrl!CAcctDate
            rstGLMast.Update
        Else
            MsgBox "Please Note, G/L Account " & _
                    rstGLTrn!GLNumber & " Not Found", vbCritical
        End If
        ' Update TBalance with Entry
        mvBuff = "Select * from TBalance Where GLNumber = '" _
            & rstGLTrn!GLNumber & "'"
        Set rstTBalance = db2.OpenRecordset(mvBuff, dbOpenDynaset)
        If Not (rstTBalance.EOF And rstTBalance.BOF) Then
            rstTBalance.Edit
            If rstGLTrn!GLTrnamt < 0 Then
                rstTBalance("DebitAmt") = rstTBalance!DebitAmt + _
                                        Abs(rstGLTrn!GLTrnamt)
            Else
                rstTBalance("CreditAmt") = rstTBalance!CreditAmt + _
                                        rstGLTrn!GLTrnamt
            End If
            rstTBalance("GLLMvtDat") = rstGLCtrl!CAcctDate
            rstTBalance.Update
        Else
            MsgBox "Please Note, G/L Account " & _
                    rstGLTrn!GLNumber & " Not Found", vbCritical
        End If
        rstGLTrn.MoveNext
        UpdatePercentDialog (mvperctr)
        mvperctr = mvperctr + 1
    Wend
    lblProg.Caption = "100%"
    
    ' Set Shadow Balance To G/L Balance
    mvPL = 0
    LblActivity.Caption = "Set Shadow Balance ....."
    LblActivity.Refresh
    lblProg.Width = 0
    mvperctr = 1
    mvBuff = "Select * from GLMast Order By GLNumber;"
    Set rstGLMast = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    If Not (rstGLMast.EOF And rstGLMast.BOF) Then
        rstGLMast.MoveLast
        mvperinc = 100 / rstGLMast.RecordCount
        rstGLMast.MoveFirst
'       rstTBalance.MoveFirst
    End If
    While Not rstGLMast.EOF
        mvBuff = "Select * from TBalance Where GLNumber = '" _
            & rstGLMast!GLNumber & "'"
        Set rstTBalance = db2.OpenRecordset(mvBuff, dbOpenDynaset)
        If Not (rstTBalance.EOF And rstTBalance.BOF) Then
            rstTBalance.Edit
            rstTBalance("GLBal") = rstGLMast!GLBal
            rstTBalance.Update
        Else
            MsgBox "Please Note, G/L Account " & _
                    rstGLMast!GLNumber & " Not Found", vbCritical
        End If
        rstGLMast.Edit
        rstGLMast("GLSBal") = rstGLMast!GLBal
        If InStr("EXITP", rstGLMast!GLTyp) <> 0 Then
            mvPL = mvPL + rstGLMast!GLBal
        End If
        rstGLMast.Update
        rstGLMast.MoveNext
        UpdatePercentDialog (mvperctr)
        mvperctr = mvperctr + 1
    Wend
    mvBuff = "Select * from GLMast Where GLNumber = '" _
        & rstGLCtrl!PlaccNum & "'"
    Set rstGLMast = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    If Not (rstGLMast.EOF And rstGLMast.BOF) Then
        rstGLMast.Edit
        rstGLMast("GLBal") = mvPL
        rstGLMast("GLLMvtDat") = rstGLCtrl!CAcctDate
        rstGLMast.Update
    End If
    lblProg.Caption = "100%"
    
    ' Update History File with Transactions
    LblActivity.Caption = "Update History with Transactions ....."
    LblActivity.Refresh
    lblProg.Width = 0
    mvperctr = 1
    mvBuff = "Select * from GLTrnHst;"
    Set rstGLHist = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    mvBuff = "Select * from GLTrn;"
    Set rstGLTrn = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    If Not (rstGLTrn.EOF And rstGLTrn.BOF) Then
        rstGLTrn.MoveLast
        mvperinc = 100 / rstGLTrn.RecordCount
        rstGLTrn.MoveFirst
    End If
    While Not rstGLTrn.EOF
        rstGLHist.AddNew
        rstGLHist("GLNumber") = rstGLTrn!GLNumber
        rstGLHist("GLPartlars") = rstGLTrn!GLPartLars
        rstGLHist("GLInstno") = rstGLTrn!GLInstno
        rstGLHist("GLTrnAmt") = rstGLTrn!GLTrnamt
        rstGLHist("GLTrnDat") = rstGLTrn!GLtrnDat
        rstGLHist("trtime") = rstGLTrn("trtime")
        rstGLHist("GLBatno") = rstGLTrn!GLBatNo
        rstGLHist("GLAcctDat") = rstGLCtrl!CAcctDate
        rstGLHist.Update
        UpdatePercentDialog (mvperctr)
        mvperctr = mvperctr + 1
        rstGLTrn.MoveNext
    Wend
    rstGLTrn.Close
    mvBuff = "Delete * from GLTrn;"
    db2.Execute mvBuff
    lblProg.Caption = "100%"
    
    ' Update P & L Table
    mvBuff = "Select * from GLPL;"
    Set rstGLPL = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    LblActivity.Caption = "Generating Income Statement ....."
    LblActivity.Refresh
    lblProg.Width = 0
    mvperctr = 1
    mvBuff = "Select * from GLMast Order By GLNumber;"
    Set rstGLMast = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    If Not (rstGLMast.EOF And rstGLMast.BOF) Then
        rstGLMast.MoveLast
        mvperinc = 100 / rstGLMast.RecordCount
        mvperinc = mvperinc / 5
    End If
    ' Update for Interest Income
    mvTyp = "T"
    mvCode = "A"
    mvGCode = "B"
    If UCase(rstGLCtrl!GLTyp) = "S" Then
        mvMsg = "** Total Sales Income **"
    Else
        mvMsg = "** Total Interest Income **"
    End If
    UpdPL
    mvALMBal = mvGLMBal
    mvAcurBal = mvGcurBal
    ' Update for Interest Expense
    mvTyp = "P"
    mvCode = "C"
    mvGCode = "D"
    mvMsg = "** Total Cost of Sales **"
    If UCase(rstGLCtrl!GLTyp) = "S" Then
        mvMsg = "** Total Cost Of Sales **"
    Else
        mvMsg = "** Total Interest Expense **"
    End If
    UpdPL
    mvBLMBal = mvGLMBal
    mvBcurBal = mvGcurBal
    mvCLMBal = mvALMBal + mvBLMBal
    mvCcurBal = mvAcurBal + mvBcurBal
    rstGLPL.AddNew
    rstGLPL("RTyp") = mvRTyp
    rstGLPL("GLTyp") = "E"
    rstGLPL("GLNar") = "** Gross Profit **"
    rstGLPL("GLOpnBbal") = mvCLMBal
    rstGLPL("GLBal") = mvCcurBal
    rstGLPL.Update
    'Update for Other Income
    mvTyp = "I"
    mvCode = "F"
    mvGCode = "G"
    mvMsg = "** Total Other Income **"
    UpdPL
    mvDLMBal = mvGLMBal
    mvDcurBal = mvGcurBal
    ' Update for Operating Expenses
    mvTyp = "E"
    mvCode = "H"
    mvGCode = "I"
    mvMsg = "** Total Operating Expenses **"
    UpdPL
    mvELMBal = mvGLMBal
    mvEcurBal = mvGcurBal
    ' Update for PBT
    mvFLMBal = mvCLMBal + mvDLMBal + mvELMBal
    mvFcurBal = mvCcurBal + mvDcurBal + mvEcurBal
    rstGLPL.AddNew
    rstGLPL("RTyp") = mvRTyp
    rstGLPL("GLTyp") = "J"
    rstGLPL("GLNar") = "** Profit Before Tax **"
    rstGLPL("GLOpnBbal") = mvFLMBal
    rstGLPL("GLBal") = mvFcurBal
    rstGLPL.Update
    ' Update for Tax
    mvTyp = "X"
    mvCode = "K"
    mvGCode = "L"
    mvMsg = "** Taxation **"
    UpdPL
    rstGLPL.AddNew
    rstGLPL("RTyp") = mvRTyp
    rstGLPL("GLTyp") = "M"
    rstGLPL("GLNar") = "** Profit After Tax **"
    rstGLPL("GLOpnBbal") = mvFLMBal + mvGLMBal
    rstGLPL("GLBal") = mvFcurBal + mvGcurBal
    rstGLPL.Update
    ' End of P&L
    lblProg.Caption = "100%"
    Screen.MousePointer = vbDefault
End Sub

Private Sub GenRev()
    LblActivity.Caption = "Generating Reversal Entries ......"
    LblActivity.Refresh
    lblProg.Width = 0
    mvperctr = 1
    mvBuff = "Select * from GLMast Order By GLNumber;"
    Set rstGLMast = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    If Not (rstGLMast.EOF And rstGLMast.BOF) Then
        rstGLMast.MoveLast
        mvperinc = 100 / rstGLMast.RecordCount
        rstGLMast.MoveFirst
    End If
    While Not rstGLMast.EOF
        If (rstGLMast!GLTyp = "P" Or _
           rstGLMast!GLTyp = "I" Or _
           rstGLMast!GLTyp = "E" Or _
           rstGLMast!GLTyp = "X" Or _
           rstGLMast!GLTyp = "T") And _
           rstGLMast!GLBal <> 0 Then
            ' Generate Contra Leg
            rstGLTrn.AddNew
            rstGLTrn("GLNumber") = rstGLMast!GLNumber
            rstGLTrn("GLTrnAmt") = -1 * rstGLMast!GLBal
            UpdRev
            ' Generate Retain Earning Leg
            rstGLTrn.AddNew
            rstGLTrn("GLNumber") = cmbRAccNum
            rstGLTrn("GLTrnAmt") = rstGLMast!GLBal
            UpdRev
        End If
        UpdatePercentDialog (mvperctr)
        mvperctr = mvperctr + 1
        rstGLMast.MoveNext
    Wend
    lblProg.Caption = "100%"
End Sub

Private Sub UpdRev()
    rstGLTrn("GLPartlars") = "End of Year Reversal"
    rstGLTrn("GLInstno") = "EOYREV"
    rstGLTrn("GLTrnDat") = rstGLCtrl!CAcctDate
    rstGLTrn("GLBatNo") = 0
    rstGLTrn.Update
End Sub

Private Sub UpdPL()
    mvGLMBal = 0
    mvGcurBal = 0
    mvBuff = "Select * from GLMast Order By GLNumber;"
    Set rstGLMast = db2.OpenRecordset(mvBuff, dbOpenDynaset)
    If Not (rstGLMast.EOF And rstGLMast.BOF) Then
        rstGLMast.MoveFirst
    End If
    While Not rstGLMast.EOF
        If rstGLMast!GLTyp = mvTyp Then
            If rstGLMast!GLLev = "H" Then
                mvGLHDR = rstGLMast!GLNumber
            End If
            rstGLPL.AddNew
            rstGLPL("RTyp") = mvRTyp
            rstGLPL("GLTyp") = mvCode
            rstGLPL("GLHDR") = mvGLHDR
            rstGLPL("GLNumber") = rstGLMast!GLNumber
            rstGLPL("GLNar") = rstGLMast!GLNar
            rstGLPL("GLOpnBbal") = rstGLMast!GLOpnBbal
            rstGLPL("GLBal") = rstGLMast!GLBal
            rstGLPL("GLLMVTDat") = rstGLMast!GLLMVTDat
            rstGLPL.Update
            mvGLMBal = mvGLMBal + rstGLMast!GLOpnBbal
            mvGcurBal = mvGcurBal + rstGLMast!GLBal
        End If
        UpdatePercentDialog (mvperctr)
        mvperctr = mvperctr + 1
        rstGLMast.MoveNext
    Wend
    rstGLPL.AddNew
    rstGLPL("RTyp") = mvRTyp
    rstGLPL("GLTyp") = mvGCode
    rstGLPL("GLNar") = mvMsg
    rstGLPL("GLOpnBbal") = mvGLMBal
    rstGLPL("GLBal") = mvGcurBal
    rstGLPL.Update
End Sub

Private Sub UpdatePercentDialog(newval As Single)
    lblProg.Caption = Str$(Int(newval * mvperinc)) & "%"
    lblProg.Width = (Int(newval * mvperinc) * (LblActivity.Width - 160)) / 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ExtErr
    GenClass.fleLogin mvUserid, "Exit from General Ledger Processing", Date, Time
    db2.Close
ExtClr:
    Exit Sub
ExtErr:
    If Err.Number = 91 Then
        Resume ExtClr
    Else
        MsgBox "Error: " + Error$ + " Has Occurred", vbExclamation, Me.Caption
    End If
    On Error Resume Next
    Resume ExtClr
End Sub

