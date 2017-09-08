VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMailDom 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Tracking System"
   ClientHeight    =   7095
   ClientLeft      =   1305
   ClientTop       =   1080
   ClientWidth     =   11070
   DrawWidth       =   3
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7095
   ScaleWidth      =   11070
   Begin VB.PictureBox SSPanel5 
      BackColor       =   &H00B19983&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11490
      TabIndex        =   22
      Top             =   6360
      Width           =   11550
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdDom 
         BackColor       =   &H0000FF00&
         Caption         =   "&Domicile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00B19983&
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
         Height          =   375
         Left            =   5640
         TabIndex        =   48
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.PictureBox SSPanel3 
      BackColor       =   &H00B19983&
      Height          =   6375
      Left            =   0
      ScaleHeight     =   6315
      ScaleWidth      =   10995
      TabIndex        =   19
      Top             =   0
      Width           =   11055
      Begin VB.TextBox recBy 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2520
         MaxLength       =   25
         TabIndex        =   46
         Top             =   5760
         Width           =   2775
      End
      Begin VB.CommandButton cmdLDom 
         Caption         =   "Dom List"
         Height          =   375
         Left            =   8640
         TabIndex        =   45
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox INdate 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   42
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox remarks 
         BackColor       =   &H00C0FFFF&
         Height          =   855
         Left            =   6720
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         ToolTipText     =   "Enter the reason for retaining or domiciling the mail in your unit."
         Top             =   5280
         Width           =   3975
      End
      Begin VB.Frame OptMailAct 
         BackColor       =   &H00B19983&
         Caption         =   "Mail Action"
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
         Height          =   1335
         Left            =   5520
         TabIndex        =   36
         Top             =   2760
         Width           =   1575
         Begin VB.OptionButton mAct 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Free Mail"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   38
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton mAct 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Hold Mail"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   37
            Top             =   840
            Width           =   1095
         End
      End
      Begin VB.TextBox mailDesc 
         Appearance      =   0  'Flat
         Height          =   1335
         Left            =   7200
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   2400
         Width           =   3495
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Find Mail"
         Height          =   375
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H8000000B&
         Caption         =   "Retrieve Mail"
         Height          =   375
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox CmbBy 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   10
         Top             =   4200
         Width           =   2895
      End
      Begin VB.TextBox BenNames 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   4
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox CPhone 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7200
         MaxLength       =   15
         TabIndex        =   12
         Top             =   2040
         Width           =   3255
      End
      Begin VB.ComboBox cmbPrior 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   11
         Top             =   4560
         Width           =   2895
      End
      Begin VB.ComboBox cmbFac 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   5
         Top             =   2400
         Width           =   2895
      End
      Begin VB.ComboBox cmbRBU 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   9
         Top             =   3840
         Width           =   2895
      End
      Begin VB.ComboBox cmbSubj 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   7
         Top             =   3120
         Width           =   2895
      End
      Begin VB.ComboBox cmbsClass 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   8
         Top             =   3480
         Width           =   2895
      End
      Begin VB.ComboBox cmbDept 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   6
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox MailRef 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   390
         Left            =   2640
         MaxLength       =   15
         TabIndex        =   0
         Top             =   960
         Width           =   2295
      End
      Begin VB.Data DatTMP 
         Caption         =   "Tmptab"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox OrigNames 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Data datgenTab 
         Caption         =   "Gentab"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   8760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog GetFile 
         Left            =   2880
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "Bmp"
         FilterIndex     =   2
      End
      Begin VB.CommandButton cmdList 
         Caption         =   "In &List"
         Height          =   375
         Left            =   9720
         TabIndex        =   17
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdbott 
         Caption         =   "Botto&m"
         Height          =   375
         Left            =   9720
         TabIndex        =   16
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   8640
         TabIndex        =   15
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   9720
         TabIndex        =   13
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdtop 
         Caption         =   "&Top"
         Height          =   375
         Left            =   8640
         TabIndex        =   14
         Top             =   3840
         Width           =   975
      End
      Begin MSMask.MaskEdBox OutDate 
         Height          =   375
         Left            =   2520
         TabIndex        =   43
         ToolTipText     =   "Enter the date of mail despatch."
         Top             =   5280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   12648447
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
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Dom Authorized By:"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   47
         Top             =   5880
         Width           =   1815
      End
      Begin VB.Label Label16 
         BackColor       =   &H00B19983&
         Caption         =   "Date Domiciled:"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   44
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00B19983&
         Caption         =   "Reason :"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5760
         TabIndex        =   40
         Top             =   5280
         Width           =   975
      End
      Begin VB.Label Label18 
         BackColor       =   &H00B19983&
         Caption         =   "Beneficiary's Names:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   39
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Recieved By::"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   34
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Domicle Incoming Mail "
         BeginProperty Font 
            Name            =   "Bodoni MT Condensed"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   840
         TabIndex        =   33
         Top             =   240
         Width           =   10095
      End
      Begin VB.Label Label26 
         BackColor       =   &H00B19983&
         Caption         =   "Date Recieved:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5880
         TabIndex        =   32
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackColor       =   &H00B19983&
         Caption         =   "Mail Description:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5880
         TabIndex        =   31
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label15 
         BackColor       =   &H00B19983&
         Caption         =   "Contact Phone:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5880
         TabIndex        =   30
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Mail Priority:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   29
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Faculty / Division:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   28
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Recieving Business Unit:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   27
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Mail Subject:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   26
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Classification:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label House 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Academic Department:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11520
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label Label8 
         BackColor       =   &H00B19983&
         Caption         =   "Originator'sName:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11400
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Image Image2 
         Height          =   855
         Left            =   120
         Picture         =   "frmMailDom.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00B19983&
         Caption         =   "Mail Reference:"
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
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   1080
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmMailDom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim db1 As Database, wrktemp As Workspace
Dim PicName As String, DateBuff, AgeBuff As String

Dim rstState, RegTab, rstTitle As Recordset
Dim rstsClass, rstBUnitD As Recordset
Dim rstBUnitR, rstFac, rstDept As Recordset
Dim rstPrior, rstSubj As Recordset
Dim rstMMast, rstSelData As Recordset

Dim AddFlag As Boolean, PBuff As Variant
Dim EditFlag As Boolean, MaxErrors As Integer
Dim FldChk As Boolean, Insql As String
Dim strpos As Integer, tempBuff As String
Dim FirstTime As Integer, FirstPass As Integer
Dim GenClass As New mProLog, buff As String
Dim BType As String

Private Sub cmbsubj_click()
  SubjClass
End Sub

Private Sub SubjClass()
    
    strpos = InStr(1, cmbSubj, ",", 1)
    GTBuff = Trim(Left(cmbSubj, strpos - 1))
    '_____________________________find subject classification
    'rstsClass.Requery
    mvSql = "Select * from defsubclass Where subjg = '" & GTBuff & "'"
    Set rstSelData = db1.OpenRecordset(mvSql, dbOpenDynaset)
    cmbsClass.Clear
    If (rstSelData.BOF And rstSelData.EOF) Then
       cmbsClass = " "
    Else
       rstSelData.Requery
       Do While Not rstSelData.EOF
          cmbsClass.AddItem rstSelData!RCode + ", " + rstSelData!RDesc
          rstSelData.MoveNext
       Loop
       If Not (rstSelData.BOF And rstSelData.EOF) Then
         cmbsClass.ListIndex = 0
       End If
   End If
End Sub

Private Sub cmdLDom_Click()
On Error GoTo PError
    mvSql = "Select * From dommailquery order by outdate"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF DOMICILED MAILS"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         frmBrowse.grdGeneral.Columns(0).Width = 1000
         frmBrowse.grdGeneral.Columns(1).Width = 1000
         frmBrowse.grdGeneral.Columns(2).Width = 1000
         frmBrowse.grdGeneral.Columns(3).Width = 3000
         frmBrowse.grdGeneral.Columns(4).Width = 2000
         frmBrowse.grdGeneral.Columns(5).Width = 2000
         frmBrowse.grdGeneral.Columns(6).Width = 2000
         frmBrowse.grdGeneral.Columns(7).Width = 2800
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2000
         Col0.Caption = "Mail Ref."
         Col1.Caption = "Date In       "
         Col2.Caption = "Date Dom."
         Col3.Caption = "Reason for Domiciling"
         Col4.Caption = "Mail Details"
         Col5.Caption = "Deparment"
         Col6.Caption = "Subjects"
         Col7.Caption = "Mail Sub Classification"
         Col8.Caption = "Beneficiary"
         Col9.Caption = "Authorized By"
   End If
   frmBrowse.Show
PError: End Sub

Private Sub cmdFind_Click()
  Dim sql As String, digitflag As Integer
   Dim buff As String, ans As String

   On Error GoTo FindError1

   ans = InputBox$("Enter the Mail's Reference to find:")
   ans = UCase(ans)
   If Len(ans) = 0 Then
      ErrorMessages (NOINPUT)
      Exit Sub
   Else
      buff = "mailref LIKE " & Chr$(34) & ans & "*" & Chr$(34)
   End If
   rstMMast.FindFirst buff
   If rstMMast.NOMATCH Then
      ErrorMessages (NOMATCH)
      Exit Sub
   Else
      showdata
   End If
FindError0:
   Exit Sub

FindError1:
   MsgBox "ERROR:   " + Error$ + " has occurred", vbCritical, Me.Caption
   On Error GoTo 0
End Sub

Private Sub cmdList_Click()
On Error GoTo PError
    mvSql = "Select * From inmailquery order by indate"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No Mails found", vbExclamation, "Mail Records Review"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF INCOMING MAILS"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         Set Col8 = frmBrowse.grdGeneral.Columns(8)
         Set Col9 = frmBrowse.grdGeneral.Columns(9)
         Set col10 = frmBrowse.grdGeneral.Columns(10)
         frmBrowse.grdGeneral.Columns(0).Width = 1200
         frmBrowse.grdGeneral.Columns(1).Width = 800
         frmBrowse.grdGeneral.Columns(2).Width = 1400
         frmBrowse.grdGeneral.Columns(3).Width = 2000
         frmBrowse.grdGeneral.Columns(4).Width = 2200
         frmBrowse.grdGeneral.Columns(5).Width = 1000
         frmBrowse.grdGeneral.Columns(6).Width = 1000
         frmBrowse.grdGeneral.Columns(7).Width = 2800
         frmBrowse.grdGeneral.Columns(8).Width = 2800
         frmBrowse.grdGeneral.Columns(9).Width = 2200
         frmBrowse.grdGeneral.Columns(10).Width = 3000
         Col0.Caption = "Mail Ref."
         Col1.Caption = "     "
         Col2.Caption = "Originated By"
         Col3.Caption = "Department"
         Col4.Caption = "Beneficiary"
         Col5.Caption = "Indate"
         Col6.Caption = "Outdate"
         Col7.Caption = "Subject"
         Col8.Caption = "Classification"
         Col9.Caption = "Received By"
         col10.Caption = "Mail Description"
   End If
   frmBrowse.Show
PError:
End Sub



Private Sub cmdOk_Click()
  FindMail
End Sub
Private Sub FindMail()
  Dim sql As String, digitflag As Integer
   Dim buff As String, ans As String
   On Error GoTo FindError1
   ans = UCase(Mailref)
   If Len(ans) = 0 Then
      MsgBox "The mail reference given is invalid, Retry", vbInformation, "Mail Records"
      ClearData
      Exit Sub
   Else
      buff = "mailref LIKE " & Chr$(34) & ans & "*" & Chr$(34)
   End If
   rstMMast.FindFirst buff
   If rstMMast.NOMATCH Then
      ErrorMessages (NOMATCH)
      ClearData
      Exit Sub
   Else
      showdata
   End If
FindError0:
   Exit Sub

FindError1:
   MsgBox "ERROR:   " + Error$ + " has occurred", vbCritical, Me.Caption
   On Error GoTo 0
End Sub

Private Sub mailref_gotfocus()
    Mailref.SelStart = 0
    Mailref.SelLength = Len(Mailref)
End Sub

Private Sub MailRef_LostFocus()
   FindMail
End Sub

Private Sub cphone_gotfocus()
    CPhone.SelStart = 0
    CPhone.SelLength = Len(CPhone)
End Sub


Private Sub maildesc_gotfocus()
    mailDesc.SelStart = 0
    mailDesc.SelLength = Len(mailDesc)
End Sub


Private Sub CmdBott_Click()
On Error GoTo PError
    Dim Count As Long
    If rstMMast.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstMMast.MoveLast
    showdata
PError:
End Sub


Private Sub CmdDom_Click()
   Dim i As Integer
 On Error GoTo PError
   lbl = ""
   If rstMMast.RecordCount < 1 Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
   If rstMMast!FMail = False Then
      MsgBox "This mail is on Hold thus cannot be domiciled. Free the mail and try again.", vbInformation, "Mail Domiciliation"
      Exit Sub
   End If
   If Len(Trim(recBy)) = 0 Then
       Beep
       MsgBox "You must enter the name of the person that authorized the domicilation of this mail.", vbExclamation, "Mail Domiciliation"
       recBy.SetFocus
       Exit Sub
   End If
   If Len(Trim(remarks)) = 0 Then
       Beep
       MsgBox "You must enter your reason for domiciling this mail.", vbExclamation, "Mail Domiciliation"
       remarks.SetFocus
       Exit Sub
   End If
    If Len(Trim(OutDate)) = 0 Then
       MsgBox "Mail Outdate in cannot be blank", vbCritical, "Mail Forwarding"
       OutDate.SetFocus
       Exit Sub
    End If
    If CDate(OutDate) > Date Then
       MsgBox "The date you sent this mail out cannot be greater than today's date, Re-Enter.", vbCritical, "Mail Forwarding"
       OutDate.SetFocus
       Exit Sub
    End If
    If CDate(INdate) > CDate(OutDate) Then
       MsgBox "The date you sent this mail out cannot be less than the day mail was recieved, Re-Enter.", vbCritical, "Mail Forwarding"
       OutDate.SetFocus
       Exit Sub
    End If
   
   
   i = DomCheck()
   If i = vbYes Then
        If rstMMast("outmail") = False Then
            GTBuff = "mailref = '" & rstMMast("mailref") & "'"
            rstMMast.FindFirst GTBuff
            If rstMMast.NOMATCH Then
                MsgBox "Mail record not found", vbInformation, "Mail Records"
            Else
                rstMMast.Edit
                rstMMast("remarks") = remarks
                rstMMast("outdate") = OutDate
                rstMMast("outrecby") = recBy
                strpos = InStr(1, cmbRBU, ",", 1)
                rstMMast("cmbrbu") = Trim(Left(cmbRBU, strpos - 1))
                rstMMast.Update
                mvSql = " INSERT INTO dommails SELECT * FROM [mailmast] where mailref = '" & Mailref & "'"
                db1.Execute mvSql
                rstMMast.Delete
                lbl = "This Mail has been domiciled."
                ClearData
             End If
        Else
            MsgBox "Mail record has been forwarded outside your business unit and cannnot be domiciled", vbInformation, "Mail Delete"
            Exit Sub
        End If
        CmdEnab
    End If
PError:
End Sub

Private Sub CmdNext_Click()
    Dim flag As Integer
    On Error GoTo NextError
    If rstMMast.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstMMast.EOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEPAST)
        rstMMast.MoveLast
    Else
        rstMMast.MoveNext
        If rstMMast.EOF Then
            rstMMast.MoveLast
        End If
    End If
NextClear:
    showdata
    Exit Sub
NextError:
    ErrorMessages (NOMOVEPAST)
    rstMMast.MoveLast
    On Error GoTo 0
Resume NextClear
End Sub

Private Sub CmdPrev_Click()
    Dim flag As Integer
    On Error GoTo PrevError
    If rstMMast.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstMMast.BOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEFRONT)
        rstMMast.MoveFirst
    Else
        rstMMast.MovePrevious
        If rstMMast.BOF Then
            rstMMast.MoveFirst
        End If
    End If
PrevClear:
    showdata
    Exit Sub
PrevError:
    ErrorMessages (NOMOVEFRONT)
    rstMMast.MoveFirst
    On Error GoTo 0
    Resume PrevClear
End Sub

Private Sub CmdTop_Click()
On Error GoTo PError
    If rstMMast.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstMMast.MoveFirst
    showdata
PError:
End Sub

Private Sub Form_Load()
On Error GoTo PError
     GenClass.fleLogin mvUserid, "Deleted an Incoming Mail", Date, Time
     Dim AppStatBuff As String
     Call FormCentreMDI(Me)
     Set wrktemp = DBEngine.Workspaces(0)
     Set db1 = wrktemp.OpenDatabase(DBpath, True)
     
     Set rstState = db1.OpenRecordset("defstate", dbOpenDynaset)
     Set RegTab = db1.OpenRecordset("dproreg", dbOpenDynaset)
     Set rstsClass = db1.OpenRecordset("defsubclass", dbOpenDynaset)
     Set rstDept = db1.OpenRecordset("Defdept", dbOpenDynaset)
     Set rstBUnitD = db1.OpenRecordset("defbunitd", dbOpenDynaset)
     Insql = "Select * from mailmast Where outmail = false"
     Set rstMMast = db1.OpenRecordset(Insql, dbOpenDynaset)
     Set rstBUnitR = db1.OpenRecordset("defbunitr", dbOpenDynaset)
     Set rstFac = db1.OpenRecordset("deffac", dbOpenDynaset)
     Set rstPrior = db1.OpenRecordset("defprior", dbOpenDynaset)
     Set rstSubj = db1.OpenRecordset("defsubj", dbOpenDynaset)
     Set rstTitle = db1.OpenRecordset("deftitle", dbOpenDynaset)
     
     Dim flag As Integer, mtable As String
     Me.Caption = MVCoyname
     OutDate = Date
     OptMailAct.Enabled = False
     'cLogo = RegTab!coyid
     'cTerm.Caption = GDescTerm
     
    ''__________________________________
    'If Not (rstTitle.BOF And rstTitle.EOF) Then
    '   rstTitle.Requery
    '   Do While Not rstTitle.EOF
    '      OcmbTit.AddItem rstTitle!RCode
    '      BCmbTit.AddItem rstTitle!RCode
    '      rstTitle.MoveNext
    '   Loop
    '   BCmbTit.ListIndex = 5
    '   OcmbTit.ListIndex = 5
    'End If
    ''__________________________________
    If Not (rstBUnitD.BOF And rstBUnitD.EOF) Then
       rstBUnitD.Requery
       Do While Not rstBUnitD.EOF
          cmbRBU.AddItem rstBUnitD!RCode + ", " + rstBUnitD!RDesc
          rstBUnitD.MoveNext
       Loop
       mvBuff = "rCode = '" & gBunitD & "'"
       rstBUnitD.FindFirst mvBuff
       cmbRBU = rstBUnitD!RCode & "," & rstBUnitD!RDesc
       
       ' select only the staff that belongs to the business unit
       GTBuff = gBunitD
       mvSql = "Select * from defbunitr Where bunit = '" & GTBuff & "'"
       Set rstSelData = db1.OpenRecordset(mvSql, dbOpenDynaset)
       CmbBy.Clear
       If (rstSelData.BOF And rstSelData.EOF) Then
          CmbBy = " "
       Else
          rstSelData.MoveFirst
          Do While Not rstSelData.EOF
             CmbBy.AddItem rstSelData!RCode + ", " + rstSelData!RDesc
             rstSelData.MoveNext
          Loop
          CmbBy.ListIndex = 0
       End If
       
    End If
    ''__________________________________
    If Not (rstPrior.BOF And rstPrior.EOF) Then
       rstPrior.Requery
       Do While Not rstPrior.EOF
          cmbPrior.AddItem rstPrior!RCode + ", " + rstPrior!RDesc
          rstPrior.MoveNext
       Loop
       cmbPrior.ListIndex = 0
    End If
    ''__________________________________
    If Not (rstSubj.BOF And rstSubj.EOF) Then
       rstSubj.Requery
       Do While Not rstSubj.EOF
          cmbSubj.AddItem rstSubj!RCode + ", " + rstSubj!RDesc
          rstSubj.MoveNext
       Loop
       cmbSubj.ListIndex = 0
       ' select only the sub classification that belongs to the subject
          rstSubj.MoveFirst
          GTBuff = Trim(rstSubj!RCode)
          mvSql = "Select * from defsubclass Where subjg = '" & GTBuff & "'"
          Set rstSelData = db1.OpenRecordset(mvSql, dbOpenDynaset)
          cmbsClass.Clear
          If (rstSelData.BOF And rstSelData.EOF) Then
             cmbsClass = " "
          Else
             rstSelData.MoveFirst
             Do While Not rstSelData.EOF
                cmbsClass.AddItem rstSelData!RCode + ", " + rstSelData!RDesc
                rstSelData.MoveNext
             Loop
             cmbsClass.ListIndex = 0
          End If
    End If
    ''__________________________________
    If Not (rstFac.BOF And rstFac.EOF) Then
       rstFac.Requery
       Do While Not rstFac.EOF
          cmbFac.AddItem rstFac!RCode + ", " + rstFac!RDesc
          rstFac.MoveNext
       Loop
       cmbFac.ListIndex = 0
       ' select only the department that belongs to the faculty
       rstFac.MoveFirst
       GTBuff = Trim(rstFac!RCode)
       mvSql = "Select * from defdept Where facc = '" & GTBuff & "'"
       Set rstSelData = db1.OpenRecordset(mvSql, dbOpenDynaset)
       cmbDept.Clear
       If (rstSelData.BOF And rstSelData.EOF) Then
          cmbDept = " "
       Else
          rstSelData.MoveFirst
          Do While Not rstSelData.EOF
             cmbDept.AddItem rstSelData!RCode + ", " + rstSelData!RDesc
             rstSelData.MoveNext
          Loop
          cmbDept.ListIndex = 0
       End If
    End If
    ''__________________________________
    ''__________________________________
    ''CmdAdd.SetFocus
    AddFlag = False
    EditFlag = False
    ClearData
    If Not (rstMMast.BOF And rstMMast.EOF) Then
        rstMMast.Requery
        rstMMast.MoveFirst
        Mailref = rstMMast("mailref")
        showdata
    End If
    FldDisab
    Mailref.Enabled = True
PError:
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub
Public Sub showdata()
  On Error GoTo PError
  lbl = ""
  If rstMMast.RecordCount = 1 Then
     rstMMast.MoveFirst
  End If
  If Not (rstMMast.BOF And rstMMast.EOF) Then
        Mailref = rstMMast!Mailref
        INdate = rstMMast!INdate
        OutDate = Date
        remarks = ""
        recBy = ""
        Select Case rstMMast!FMail
        Case True
           mAct(0).Value = True
        Case False
           mAct(1).Value = True
        End Select
        
        If IsNull(rstMMast("ocmbtit")) Then
           OrigNames = " "
        Else
           OrigNames = rstMMast("OcmbTit") + " "
        End If
        If IsNull(rstMMast("ofname")) Then
           OrigNames = OrigNames + " "
        Else
           OrigNames = OrigNames + rstMMast("ofname") + " "
        End If
        If IsNull(rstMMast("oothnames")) Then
           OrigNames = OrigNames + " "
        Else
           OrigNames = OrigNames + rstMMast("oothnames") + " "
        End If
        If IsNull(rstMMast("osurname")) Then
           OrigNames = OrigNames + " "
        Else
           OrigNames = OrigNames + rstMMast("osurname") + " "
        End If
        
        If IsNull(rstMMast("bcmbtit")) Then
           BenNames = " "
        Else
           BenNames = rstMMast("bcmbTit") + " "
        End If
        
        If IsNull(rstMMast("bfname")) Then
           BenNames = BenNames + " "
        Else
           BenNames = BenNames + rstMMast("bfname") + " "
        End If
        If IsNull(rstMMast("bothnames")) Then
           BenNames = BenNames + " "
        Else
           BenNames = BenNames + rstMMast("bothnames") + " "
        End If
        If IsNull(rstMMast("bsurname")) Then
           BenNames = BenNames + " "
        Else
           BenNames = BenNames + rstMMast("bsurname") + " "
        End If
        If IsNull(rstMMast("cphone")) Then
           CPhone = " "
        Else
           CPhone = rstMMast("cphone")
        End If
        If IsNull(rstMMast("maildesc")) Then
           mailDesc = " "
        Else
           mailDesc = rstMMast("maildesc")
        End If
        '________________________________________
        
        GTBuff = "rCode = '" & rstMMast("cmbdept") & "'"
        rstDept.FindFirst GTBuff
        cmbDept = rstDept!RCode & "," & rstDept!RDesc
        
        GTBuff = "rCode = '" & rstMMast("cmbsubj") & "'"
        rstSubj.FindFirst GTBuff
        cmbSubj = rstSubj!RCode & "," & rstSubj!RDesc
        
        GTBuff = "rCode = '" & rstMMast("cmbsclass") & "'"
        rstsClass.FindFirst GTBuff
        cmbsClass = rstsClass!RCode & "," & rstsClass!RDesc
        
        GTBuff = "rCode = '" & rstMMast("cmbrbu") & "'"
        rstBUnitD.FindFirst GTBuff
        cmbRBU = rstBUnitD!RCode & "," & rstBUnitD!RDesc
        
        GTBuff = "rCode = '" & rstMMast("cmbby") & "'"
        rstBUnitR.FindFirst GTBuff
        CmbBy = rstBUnitR!RCode & "," & rstBUnitR!RDesc
        
        
        GTBuff = "rCode = '" & rstMMast("cmbfac") & "'"
        rstFac.FindFirst GTBuff
        cmbFac = rstFac!RCode & "," & rstFac!RDesc
        
        GTBuff = "rCode = '" & rstMMast("cmbprior") & "'"
        rstPrior.FindFirst GTBuff
        cmbPrior = rstPrior!RCode & "," & rstPrior!RDesc
        
   Else
        MsgBox "Mail Records not found", vbCritical, "mail Records"
   End If
PError0:
     Exit Sub
PError:
    MsgBox "Error has occured, check data and retry", vbCritical, "Mail Records"
    On Error GoTo 0
    Resume PError0
End Sub

Public Sub ClearData()
On Error GoTo PError
        Mailref = ""
        mailDesc = ""
        INdate = Date
        CPhone = ""
PError:
End Sub
Private Sub Form_Unload(Cancel As Integer)
    db1.Close
End Sub


Public Sub CmdEnab()
    cmdLDom.Enabled = True
    cmdtop.Enabled = True
    cmdbott.Enabled = True
    cmdnext.Enabled = True
    cmdprev.Enabled = True
    cmdList.Enabled = True
End Sub

Public Sub CmdDisab()
    cmdLDom.Enabled = False
    cmdtop.Enabled = False
    cmdbott.Enabled = False
    cmdnext.Enabled = False
    cmdprev.Enabled = False
    cmdList.Enabled = False
End Sub

Public Sub FldEnab()
On Error GoTo PError
        Mailref.Enabled = True
        INdate.Enabled = True
        mailDesc.Enabled = True
        CPhone.Enabled = True
        cmbDept.Enabled = True
        cmbSubj.Enabled = True
        cmbsClass.Enabled = True
        cmbRBU.Enabled = True
        cmbFac.Enabled = True
        cmbPrior.Enabled = True
        CmbBy.Enabled = True
PError:
End Sub

Public Sub FldDisab()
On Error GoTo PError
        Mailref.Enabled = False
        INdate.Locked = True
        mailDesc.Locked = True
        CPhone.Locked = True
        cmbDept.Enabled = False
        cmbSubj.Enabled = False
        cmbsClass.Enabled = False
        cmbRBU.Enabled = False
        cmbFac.Enabled = False
        cmbPrior.Enabled = False
        CmbBy.Enabled = False
PError:
End Sub
Public Sub FldVal()
  On Error GoTo PError
    If EditFlag = False Then
        GTBuff = "mailref = '" & Mailref & "'"
        rstMMast.FindFirst GTBuff
        If rstMMast.NOMATCH Then
          DoEvents
        Else
          Beep
          MsgBox "Mail reference is already allocated", vbExclamation, "Mail Records"
          Mailref.SetFocus
          FldChk = False
          Exit Sub
        End If
    End If
    'If Len(MailRef) > 12 Then
    '   Beep
    '   MsgBox "Mail number must not be more than 12 characters. Enter a valid Mail Number", vbExclamation, "Mail Records"
    '   MailRef.SetFocus
    '   FldChk = False
    '   Exit Sub
    'End If
    If Len(Trim(cmbRBU)) = 0 Then
       Beep
       MsgBox "Business Unit's Name cannot be blank. Select a valid Business unit", vbExclamation, "Mail Records"
       cmbRBU.SetFocus
       FldChk = False
       Exit Sub
    End If
    If Len(Trim(CmbBy)) = 0 Then
       Beep
       MsgBox "Mail recipients's Name cannot be blank. Select a valid Recipient", vbExclamation, "Mail Records"
       CmbBy.SetFocus
       FldChk = False
       Exit Sub
    End If
    
    If Len(Trim(cmbDept)) = 0 Then
       Beep
       MsgBox "Department's Name cannot be blank. Select a valid Department", vbExclamation, "Mail Records"
       cmbDept.SetFocus
       FldChk = False
       Exit Sub
    End If
    If Len(Trim(cmbsClass)) = 0 Then
       Beep
       MsgBox "Sub Classification cannot be blank. Select a valid subclassification", vbExclamation, "Mail Records"
       cmbsClass.SetFocus
       FldChk = False
       Exit Sub
    End If
    
    
    If Len(mailDesc) < 10 Then
       Beep
       MsgBox "You must describe the mail in at least ten characters", vbExclamation, "Mail Records"
       mailDesc.SetFocus
       FldChk = False
       Exit Sub
    End If
    
    If (CDate(INdate) > Date) Then
       Beep
       MsgBox "Invalid Registration Date; Mail Registration date cannot be greater than today's date", vbExclamation, "Mail Records"
       INdate.SetFocus
       FldChk = False
       Exit Sub
    End If
PError:
End Sub
Private Sub cmbfac_click()
  FiltDept
End Sub
Private Sub FiltDept()
    
    strpos = InStr(1, cmbFac, ",", 1)
    GTBuff = Trim(Left(cmbFac, strpos - 1))
    '_____________________________find departments for the faculty
    'rstsClass.Requery
    mvSql = "Select * from defdept Where facc = '" & GTBuff & "'"
    Set rstSelData = db1.OpenRecordset(mvSql, dbOpenDynaset)
    cmbDept.Clear
    If (rstSelData.BOF And rstSelData.EOF) Then
       cmbDept = " "
    Else
       rstSelData.Requery
       Do While Not rstSelData.EOF
          cmbDept.AddItem rstSelData!RCode + ", " + rstSelData!RDesc
          rstSelData.MoveNext
       Loop
       If Not (rstSelData.BOF And rstSelData.EOF) Then
         cmbDept.ListIndex = 0
       End If
   End If
End Sub
Private Sub cmbrbu_click()
  Filtstaff
End Sub
Private Sub Filtstaff()
    
    strpos = InStr(1, cmbRBU, ",", 1)
    GTBuff = Trim(Left(cmbRBU, strpos - 1))
    '_____________________________find staff for the Selected Business Unit
    'rstsClass.Requery
    mvSql = "Select * from defbunitr Where bunit = '" & GTBuff & "'"
    Set rstSelData = db1.OpenRecordset(mvSql, dbOpenDynaset)
    CmbBy.Clear
    If (rstSelData.BOF And rstSelData.EOF) Then
       CmbBy = " "
    Else
       rstSelData.Requery
       Do While Not rstSelData.EOF
          CmbBy.AddItem rstSelData!RCode + ", " + rstSelData!RDesc
          rstSelData.MoveNext
       Loop
       If Not (rstSelData.BOF And rstSelData.EOF) Then
         CmbBy.ListIndex = 0
       End If
   End If
End Sub

Private Sub recBy_gotfocus()
    recBy.SelStart = 0
    recBy.SelLength = Len(recBy)
End Sub

Private Sub remarks_gotfocus()
    remarks.SelStart = 0
    remarks.SelLength = Len(remarks)
End Sub
