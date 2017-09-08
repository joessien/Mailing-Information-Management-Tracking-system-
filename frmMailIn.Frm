VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMailIn 
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
      TabIndex        =   31
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
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   495
         Left            =   2640
         TabIndex        =   19
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H0000FF00&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   5520
         TabIndex        =   52
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
      TabIndex        =   9
      Top             =   0
      Width           =   11055
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
         Left            =   5640
         TabIndex        =   49
         Top             =   4800
         Width           =   1695
         Begin VB.OptionButton mAct 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Hold Mail"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   51
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton mAct 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Free Mail"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   50
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.TextBox mailDesc 
         Appearance      =   0  'Flat
         Height          =   1815
         Left            =   7560
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   48
         ToolTipText     =   "Describe the mail content here."
         Top             =   4320
         Width           =   3255
      End
      Begin VB.CommandButton cmdList 
         Caption         =   "In &List"
         Height          =   375
         Left            =   9885
         TabIndex        =   25
         Top             =   3000
         Width           =   975
      End
      Begin VB.ComboBox CmbBy 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2640
         TabIndex        =   15
         Top             =   5520
         Width           =   2775
      End
      Begin VB.ComboBox BCmbTit 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7320
         TabIndex        =   5
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox BOthNames 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7320
         MaxLength       =   30
         TabIndex        =   7
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox BFname 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7320
         MaxLength       =   30
         TabIndex        =   6
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox BSurname 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   7320
         MaxLength       =   30
         TabIndex        =   8
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox CPhone 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7560
         MaxLength       =   15
         TabIndex        =   18
         Top             =   3960
         Width           =   3015
      End
      Begin VB.ComboBox OcmbTit 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2640
         TabIndex        =   1
         Top             =   1800
         Width           =   2295
      End
      Begin VB.ComboBox cmbPrior 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2640
         TabIndex        =   16
         Top             =   5880
         Width           =   2775
      End
      Begin VB.ComboBox cmbFac 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2640
         TabIndex        =   10
         Top             =   3720
         Width           =   2775
      End
      Begin VB.ComboBox cmbRBU 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2640
         TabIndex        =   14
         Top             =   5160
         Width           =   2775
      End
      Begin VB.ComboBox cmbSubj 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2640
         TabIndex        =   12
         Top             =   4440
         Width           =   2775
      End
      Begin VB.ComboBox cmbsClass 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2640
         TabIndex        =   13
         Top             =   4800
         Width           =   2775
      End
      Begin VB.ComboBox cmbDept 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2640
         TabIndex        =   11
         Top             =   4080
         Width           =   2775
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
         Width           =   2055
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
      Begin VB.TextBox OOthNames 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   3
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox OFName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   2
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox OSurname 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   4
         Top             =   2880
         Width           =   2295
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
      Begin MSMask.MaskEdBox InDate 
         Height          =   315
         Left            =   7560
         TabIndex        =   17
         Top             =   3600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdbott 
         Caption         =   "&Bottom"
         Height          =   375
         Left            =   9885
         TabIndex        =   24
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   9885
         TabIndex        =   23
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   9885
         TabIndex        =   22
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdtop 
         Caption         =   "&Top"
         Height          =   375
         Left            =   9885
         TabIndex        =   21
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Recieved By::"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   720
         TabIndex        =   47
         Top             =   5520
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Incoming Mails Registration"
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
         Left            =   960
         TabIndex        =   46
         Top             =   120
         Width           =   10095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00B19983&
         Caption         =   "Beneficiary's Title:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5280
         TabIndex        =   45
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackColor       =   &H00B19983&
         Caption         =   "Beneficiary's Surname:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5280
         TabIndex        =   44
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Label13 
         BackColor       =   &H00B19983&
         Caption         =   "Beneficiary's First Name:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5280
         TabIndex        =   43
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label18 
         BackColor       =   &H00B19983&
         Caption         =   "Beneficiary's Other Names:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5280
         TabIndex        =   42
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label26 
         BackColor       =   &H00B19983&
         Caption         =   "Date Recieved:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6240
         TabIndex        =   41
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackColor       =   &H00B19983&
         Caption         =   "Mail Description:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6240
         TabIndex        =   40
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label15 
         BackColor       =   &H00B19983&
         Caption         =   "Contact Phone:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6240
         TabIndex        =   39
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Mail Priority:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   720
         TabIndex        =   38
         Top             =   5880
         Width           =   1815
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Faculty / Division:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   720
         TabIndex        =   37
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Recieving Business Unit:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   720
         TabIndex        =   36
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Mail Subject:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   720
         TabIndex        =   35
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Classification:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   720
         TabIndex        =   34
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label House 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Academic Department:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   720
         TabIndex        =   33
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   11400
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label8 
         BackColor       =   &H00B19983&
         Caption         =   "Originator'sTitle:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   720
         TabIndex        =   32
         Top             =   1800
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
         Picture         =   "frmMailIn.frx":0000
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
         TabIndex        =   30
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00B19983&
         Caption         =   "Originator's Surname:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   720
         TabIndex        =   29
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00B19983&
         Caption         =   "Originator's First Name:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   720
         TabIndex        =   28
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00B19983&
         Caption         =   "Originators Other Names:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   720
         TabIndex        =   27
         Top             =   2520
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmMailIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim db1 As Database, wrktemp As Workspace
Dim PicName As String, DateBuff, AgeBuff As String
Dim FMail As Boolean

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

Private Sub CmbBy_lostfocus()
On Error GoTo PError:
   cmbVal
PError:
End Sub

Private Sub cmbDept_lostfocus()
On Error GoTo PError:
  cmbVal
PError:
End Sub

Private Sub cmbFac_LostFocus()
On Error GoTo PError:
   cmbVal
PError:
End Sub


Private Sub cmbPrior_lostfocus()
On Error GoTo PError:
   cmbVal
PError:
End Sub

Private Sub cmbRBU_LostFocus()
On Error GoTo PError:
   cmbVal
PError:
End Sub

Private Sub cmbsClass_LostFocus()
On Error GoTo PError:
   cmbVal
PError:
End Sub

Private Sub cmbsubj_click()
  cmbsClass.Enabled = True
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

Private Sub cmbSubj_LostFocus()
On Error GoTo PError:
  cmbVal
PError:
End Sub

Private Sub cmdList_Click()
On Error GoTo PError
    mvSql = "Select * From inmailquery order by mailref"
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
         Col0.Caption = "Reference"
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



Private Sub mAct_Click(Index As Integer)
On Error GoTo PError
    Select Case Index
        Case 0
            mAct(0).Value = True
        Case 1
            mAct(1).Value = True
    End Select
PError:
End Sub

Private Sub mailref_gotfocus()
    Mailref.SelStart = 0
    Mailref.SelLength = Len(Mailref)
End Sub

Private Sub OcmbTit_lostfocus()
On Error GoTo PError:
  cmbVal
PError:
End Sub
Private Sub BCmbTit_lostfocus()
On Error GoTo PError:
  cmbVal
PError:
End Sub


Private Sub oSurname_gotfocus()
    osurname.SelStart = 0
    osurname.SelLength = Len(osurname)
End Sub

Private Sub cphone_gotfocus()
    CPhone.SelStart = 0
    CPhone.SelLength = Len(CPhone)
End Sub

Private Sub oOthNames_gotfocus()
    OOthNames.SelStart = 0
    OOthNames.SelLength = Len(OOthNames)
End Sub

Private Sub bSurname_gotfocus()
    BSurname.SelStart = 0
    BSurname.SelLength = Len(BSurname)
End Sub
Private Sub oFName_gotfocus()
    BFname.SelStart = 0
    BFname.SelLength = Len(BFname)
End Sub

Private Sub bOthNames_gotfocus()
    BOthNames.SelStart = 0
    BOthNames.SelLength = Len(BOthNames)
End Sub

Private Sub maildesc_gotfocus()
    mailDesc.SelStart = 0
    mailDesc.SelLength = Len(mailDesc)
End Sub

Private Sub CmdAdd_Click()
On Error GoTo PError
lbl = ""
If cmdAdd.Caption = "&Add" Then
   FldEnab
   mAct(0).Value = True
   cmbPrior.ListIndex = 2
   rstMMast.AddNew
   AddFlag = True
   CmdDisab
   ClearData
   Mailref = ""
   Mailref.SetFocus
   'cmdEdit.Enabled = False
   cmdSave.Enabled = True
   cmdAdd.Enabled = True
   cmdAdd.Caption = "&Cancel"
   'SubjClass
Else
   ClearData
   FldDisab
   AddFlag = False
   cmdAdd.Caption = "&Add"
   CmdEnab
   'cmdEdit.Enabled = True
   cmdSave.Enabled = False
End If
PError:
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

Private Sub CmdDel_Click()
   Dim i As Integer
 On Error GoTo PError
   If rstMMast.RecordCount < 1 Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
    i = DeleteCheck()
    If i = vbYes Then
        If rstMMast("rind") = "X" Then
            GTBuff = "Fileno = '" & rstMMast("Fileno") & "'"
            rstMMast.FindFirst GTBuff
            If rstMMast.NOMATCH Then
                MsgBox "Mail record not found", vbInformation, "Mail Records"
            Else
                rstMMast.Delete
            End If
        Else
            MsgBox "Mail record has entries and cannnot be deleted", vbInformation, "Mail Records"
            Exit Sub
        End If
        showdata
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

Private Sub CmdSave_Click()
    Dim i As Integer
    On Error GoTo PError
    FldChk = True
    FldVal
    If FldChk = False Then
        Exit Sub
    End If
    FldChk = True
    cmbVal
    If FldChk = False Then
        Exit Sub
    End If
    If AddFlag = True Then
        rstMMast.AddNew
           'If Len(mailref) < 12 Then
           '   Beep
           '   MsgBox "Mail number must not be less than 12 characters. Enter a valid Mail Number", vbExclamation, "Mail Records"
           '   mailref.SetFocus
           '   Exit Sub
           'End If
           rstMMast("mailref") = Mailref
    Else
        RegTab.Edit
    End If
    rstMMast("ocmbtit") = UCase(OcmbTit)
    rstMMast("oSurname") = UCase(osurname)
    rstMMast("oFName") = UCase(OFName)
    rstMMast("oOthNames") = UCase(OOthNames)
    rstMMast("bcmbtit") = UCase(BCmbTit)
    rstMMast("bSurname") = UCase(BSurname)
    rstMMast("bFName") = UCase(BFname)
    rstMMast("bOthNames") = UCase(BOthNames)
    rstMMast("fmail") = FMail
    rstMMast("benNames") = UCase(BSurname) & ", " & UCase(BFname) & " " & UCase(BOthNames)
    If CDate(INdate) > Date Then
       i = MsgBox("Recieved date cannot be greater than today's date. Do you want to correct it?", vbYesNo, "Incoming Mails ")
       If i = vbYes Then
          INdate.SetFocus
       Else
          DoEvents
       End If
       Exit Sub
    End If
    rstMMast("indate") = INdate
    rstMMast("cphone") = CPhone
    rstMMast("inmail") = True
    rstMMast("maildesc") = mailDesc
    
    strpos = InStr(1, cmbSubj, ",", 1)
    rstMMast("cmbsubj") = Trim(Left(cmbSubj, strpos - 1))
    
    strpos = InStr(1, cmbDept, ",", 1)
    rstMMast("cmbdept") = Trim(Left(cmbDept, strpos - 1))
    
    strpos = InStr(1, cmbsClass, ",", 1)
    rstMMast("cmbsClass") = Trim(Left(cmbsClass, strpos - 1))
    
    strpos = InStr(1, cmbRBU, ",", 1)
    rstMMast("cmbRBU") = Trim(Left(cmbRBU, strpos - 1))
    
    strpos = InStr(1, cmbFac, ",", 1)
    rstMMast("cmbFac") = Trim(Left(cmbFac, strpos - 1))
    
    strpos = InStr(1, cmbPrior, ",", 1)
    rstMMast("cmbPrior") = Trim(Left(cmbPrior, strpos - 1))
    
    strpos = InStr(1, CmbBy, ",", 1)
    rstMMast("CmbBy") = Trim(Left(CmbBy, strpos - 1))
    
    rstMMast.Update
    cmdAdd.Caption = "&Add"
    'cmdEdit.Caption = "&Edit"
    'cmdEdit.Enabled = True
    CmdEnab
    FldDisab
    If AddFlag = True Then
       AddFlag = False
       rstMMast.MoveLast
    Else
       EditFlag = False
    End If
    lbl = "This Mail has been registered."
    cmdSave.Enabled = False
PError0:
    Exit Sub
    
PError:
    MsgBox "Check the fields for invalid data.", vbCritical, "Mail Records"
    On Error GoTo 0
    Resume PError0

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
     GenClass.fleLogin mvUserid, "Registered a new mail", Date, Time
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
     FMail = True
     mAct(0).Value = True
     cmbDept.Enabled = False
     cmbsClass.Enabled = False
     CmbBy.Enabled = False
     'cLogo = RegTab!coyid
     'cTerm.Caption = GDescTerm
     
    ''__________________________________
    If Not (rstTitle.BOF And rstTitle.EOF) Then
       rstTitle.Requery
       Do While Not rstTitle.EOF
          OcmbTit.AddItem rstTitle!RCode
          BCmbTit.AddItem rstTitle!RCode
          rstTitle.MoveNext
       Loop
       BCmbTit.ListIndex = 5
       OcmbTit.ListIndex = 5
    End If
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
       cmbPrior.ListIndex = 3
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
    cmdSave.Enabled = False
    ClearData
    If Not (rstMMast.BOF And rstMMast.EOF) Then
        rstMMast.Requery
        rstMMast.MoveFirst
        Mailref = rstMMast("mailref")
        showdata
    End If
    FldDisab
PError:
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub
Public Sub showdata()
   lbl = ""
  On Error GoTo PError
  If rstMMast.RecordCount = 1 Then
     rstMMast.MoveFirst
  End If
  If Not (rstMMast.BOF And rstMMast.EOF) Then
        Mailref = rstMMast!Mailref
        INdate = rstMMast!INdate
        Select Case rstMMast!FMail
        Case True
           mAct(0).Value = True
        Case False
           mAct(1).Value = True
        End Select
        
        If IsNull(rstMMast("ocmbtit")) Then
           OcmbTit = " "
        Else
           OcmbTit = rstMMast("OcmbTit")
        End If
        If IsNull(rstMMast("bcmbtit")) Then
           BCmbTit = " "
        Else
           BCmbTit = rstMMast("bcmbTit")
        End If
        If IsNull(rstMMast("osurname")) Then
           osurname = " "
        Else
           osurname = rstMMast("osurname")
        End If
        If IsNull(rstMMast("bsurname")) Then
           BSurname = " "
        Else
           BSurname = rstMMast("bsurname")
        End If
        If IsNull(rstMMast("oothnames")) Then
           OOthNames = " "
        Else
           OOthNames = rstMMast("oothnames")
        End If
        If IsNull(rstMMast("bothnames")) Then
           BOthNames = " "
        Else
           BOthNames = rstMMast("bothnames")
        End If
        If IsNull(rstMMast("ofname")) Then
           OFName = " "
        Else
           OFName = rstMMast("ofname")
        End If
        If IsNull(rstMMast("bfname")) Then
           BFname = " "
        Else
           BFname = rstMMast("bfname")
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
        osurname = ""
        OFName = ""
        OOthNames = ""
        BSurname = ""
        BFname = ""
        BOthNames = ""
        mailDesc = ""
        INdate = Date
        CPhone = ""
PError:
End Sub
Private Sub Form_Unload(Cancel As Integer)
    db1.Close
End Sub


Public Sub CmdEnab()
    cmdAdd.Enabled = True
    cmdtop.Enabled = True
    cmdbott.Enabled = True
    cmdnext.Enabled = True
    cmdprev.Enabled = True
    cmdList.Enabled = True
End Sub

Public Sub CmdDisab()
    cmdAdd.Enabled = False
    cmdtop.Enabled = False
    cmdbott.Enabled = False
    cmdnext.Enabled = False
    cmdprev.Enabled = False
    cmdList.Enabled = False

End Sub

Public Sub FldEnab()
On Error GoTo PError
        Mailref.Enabled = True
        osurname.Enabled = True
        OFName.Enabled = True
        OOthNames.Enabled = True
        BSurname.Enabled = True
        BFname.Enabled = True
        BOthNames.Enabled = True
        OcmbTit.Enabled = True
        BCmbTit.Enabled = True
        INdate.Enabled = True
        mailDesc.Enabled = True
        CPhone.Enabled = True
        cmbSubj.Enabled = True
        cmbRBU.Enabled = True
        cmbFac.Enabled = True
        cmbPrior.Enabled = True
        OptMailAct.Enabled = True
PError:
End Sub
Public Sub CFldEnab()
        cmbDept.Enabled = True
        cmbsClass.Enabled = True
        CmbBy.Enabled = True
End Sub
Public Sub FldDisab()
On Error GoTo PError
        Mailref.Enabled = False
        osurname.Enabled = False
        OFName.Enabled = False
        OOthNames.Enabled = False
        BSurname.Enabled = False
        BFname.Enabled = False
        BOthNames.Enabled = False
        OcmbTit.Enabled = False
        BCmbTit.Enabled = False
        INdate.Enabled = False
        mailDesc.Enabled = False
        CPhone.Enabled = False
        cmbDept.Enabled = False
        cmbSubj.Enabled = False
        cmbsClass.Enabled = False
        cmbRBU.Enabled = False
        cmbFac.Enabled = False
        cmbPrior.Enabled = False
        CmbBy.Enabled = False
        OptMailAct.Enabled = False
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
    If Len(Trim(osurname)) = 0 Then
       Beep
       MsgBox "Originator's Surname cannot be blank. Enter a valid Surname", vbExclamation, "Mail Records"
       osurname.SetFocus
       FldChk = False
       Exit Sub
    End If
    If Len(Trim(BSurname)) = 0 Then
       Beep
       MsgBox "Beneficiary's Surname cannot be blank. Enter a valid Surname", vbExclamation, "Mail Records"
       BSurname.SetFocus
       FldChk = False
       Exit Sub
    End If
    If Len(Trim(OFName)) = 0 Then
       Beep
       MsgBox "Originator's First Name cannot be blank. Enter a valid First Name", vbExclamation, "Mail Records"
       OFName.SetFocus
       FldChk = False
       Exit Sub
    End If
    If Len(Trim(BFname)) = 0 Then
       Beep
       MsgBox "Beneficiay's First Name cannot be blank. Enter a valid First Name", vbExclamation, "Mail Records"
       BFname.SetFocus
       FldChk = False
       Exit Sub
    End If
    If Len(Trim(Mailref)) = 0 Then
       Beep
       MsgBox "Mail reference cannot be blank. Enter a valid Reference", vbExclamation, "Mail Records"
       BFname.SetFocus
       FldChk = False
       Exit Sub
    End If
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
  cmbDept.Enabled = True
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
  CmbBy.Enabled = True
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

Public Sub cmbVal()
    strpos = InStr(1, OcmbTit, ",", 1)
    If strpos = 0 Then
       mvBuff = OcmbTit
    Else
      mvBuff = Left(OcmbTit, strpos - 1)
    End If
    If Len(Trim(OcmbTit)) = 0 Then
       Mailref.SetFocus
       FldChk = False
       Exit Sub
    End If
    mvSql = "rcode = '" & Trim(mvBuff) & "'"
    rstTitle.FindFirst mvSql
    If rstTitle.NOMATCH Then
        MsgBox "Invalid reference entered.", vbInformation, Me.Caption
        OcmbTit.SetFocus
        FldChk = False
        Exit Sub
    Else
        OcmbTit = rstTitle!RCode
    End If
'___________________________________________
    strpos = InStr(1, BCmbTit, ",", 1)
    If strpos = 0 Then
       mvBuff = BCmbTit
    Else
      mvBuff = Left(BCmbTit, strpos - 1)
    End If
    If Len(Trim(BCmbTit)) = 0 Then
       Mailref.SetFocus
       FldChk = False
       Exit Sub
    End If
    mvSql = "rcode = '" & Trim(mvBuff) & "'"
    rstTitle.FindFirst mvSql
    If rstTitle.NOMATCH Then
        MsgBox "Invalid reference entered.", vbInformation, Me.Caption
        BCmbTit.SetFocus
        FldChk = False
        Exit Sub
    Else
        BCmbTit = rstTitle!RCode
    End If
'_________________________________________
    strpos = InStr(1, cmbFac, ",", 1)
    If strpos = 0 Then
       mvBuff = cmbFac
    Else
      mvBuff = Left(cmbFac, strpos - 1)
    End If
    If Len(Trim(cmbFac)) = 0 Then
       FldChk = False
       Mailref.SetFocus
       Exit Sub
    End If
    mvSql = "rcode = '" & Trim(mvBuff) & "'"
    rstFac.FindFirst mvSql
    If rstFac.NOMATCH Then
        MsgBox "Invalid reference entered.", vbInformation, Me.Caption
        FldChk = False
        cmbFac.SetFocus
        Exit Sub
    Else
        cmbFac = rstFac!RCode + ", " + rstFac!RDesc
        'FiltDept
    End If
'_________________________________________
    strpos = InStr(1, cmbDept, ",", 1)
    If strpos = 0 Then
       mvBuff = cmbDept
    Else
      mvBuff = Left(cmbDept, strpos - 1)
    End If
    If Len(Trim(cmbDept)) = 0 Then
       FldChk = False
       cmbFac.SetFocus
       Exit Sub
    End If
    mvSql = "rcode = '" & Trim(mvBuff) & "'"
    rstDept.FindFirst mvSql
    If rstDept.NOMATCH Then
        MsgBox "Invalid reference entered.", vbInformation, Me.Caption
        FldChk = False
        cmbDept.SetFocus
        Exit Sub
    Else
        cmbDept = rstDept!RCode + ", " + rstDept!RDesc
    End If

'_________________________________________
    strpos = InStr(1, cmbSubj, ",", 1)
    If strpos = 0 Then
       mvBuff = cmbSubj
    Else
      mvBuff = Left(cmbSubj, strpos - 1)
    End If
    If Len(Trim(cmbSubj)) = 0 Then
        FldChk = False
       Mailref.SetFocus
       Exit Sub
    End If
    mvSql = "rcode = '" & Trim(mvBuff) & "'"
    rstSubj.FindFirst mvSql
    If rstSubj.NOMATCH Then
        MsgBox "Invalid reference entered.", vbInformation, Me.Caption
        cmbSubj.SetFocus
        FldChk = False
        Exit Sub
    Else
        cmbSubj = rstSubj!RCode + ", " + rstSubj!RDesc
    End If

'_________________________________________
    strpos = InStr(1, cmbsClass, ",", 1)
    If strpos = 0 Then
       mvBuff = cmbsClass
    Else
      mvBuff = Left(cmbsClass, strpos - 1)
    End If
    If Len(Trim(cmbsClass)) = 0 Then
       cmbFac.SetFocus
        FldChk = False
       Exit Sub
    End If
    mvSql = "rcode = '" & Trim(mvBuff) & "'"
    rstsClass.FindFirst mvSql
    If rstsClass.NOMATCH Then
        MsgBox "Invalid reference entered.", vbInformation, Me.Caption
        cmbsClass.SetFocus
        FldChk = False
        Exit Sub
    Else
        cmbsClass = rstsClass!RCode + ", " + rstsClass!RDesc
    End If

'_________________________________________
    strpos = InStr(1, cmbRBU, ",", 1)
    If strpos = 0 Then
       mvBuff = cmbRBU
    Else
      mvBuff = Left(cmbRBU, strpos - 1)
    End If
    If Len(Trim(cmbRBU)) = 0 Then
       Mailref.SetFocus
        FldChk = False
       Exit Sub
    End If
    mvSql = "rcode = '" & Trim(mvBuff) & "'"
    rstBUnitD.FindFirst mvSql
    If rstBUnitD.NOMATCH Then
        MsgBox "Invalid reference entered.", vbInformation, Me.Caption
        cmbRBU.SetFocus
        FldChk = False
        Exit Sub
    Else
        cmbRBU = rstBUnitD!RCode + ", " + rstBUnitD!RDesc
    End If

'_________________________________________
    strpos = InStr(1, CmbBy, ",", 1)
    If strpos = 0 Then
       mvBuff = CmbBy
    Else
      mvBuff = Left(CmbBy, strpos - 1)
    End If
    If Len(Trim(CmbBy)) = 0 Then
       cmbRBU.SetFocus
        FldChk = False
       Exit Sub
    End If
    mvSql = "rcode = '" & Trim(mvBuff) & "'"
    rstBUnitR.FindFirst mvSql
    If rstBUnitR.NOMATCH Then
        MsgBox "Invalid reference entered.", vbInformation, Me.Caption
        CmbBy.SetFocus
        FldChk = False
        Exit Sub
    Else
        CmbBy = rstBUnitR!RCode + ", " + rstBUnitR!RDesc
    End If

'_________________________________________
    strpos = InStr(1, cmbPrior, ",", 1)
    If strpos = 0 Then
       mvBuff = cmbPrior
    Else
      mvBuff = Left(cmbPrior, strpos - 1)
    End If
    If Len(Trim(cmbPrior)) = 0 Then
       Mailref.SetFocus
        FldChk = False
       Exit Sub
    End If
    mvSql = "rcode = '" & Trim(mvBuff) & "'"
    rstPrior.FindFirst mvSql
    If rstPrior.NOMATCH Then
        MsgBox "Invalid reference entered.", vbInformation, Me.Caption
        cmbPrior.SetFocus
        FldChk = False
        Exit Sub
    Else
        cmbPrior = rstPrior!RCode + ", " + rstPrior!RDesc
    End If

'_________________________________________



End Sub
