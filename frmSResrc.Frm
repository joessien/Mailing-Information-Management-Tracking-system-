VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmSResrc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datapro"
   ClientHeight    =   9090
   ClientLeft      =   555
   ClientTop       =   1125
   ClientWidth     =   17745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   17745
   Begin MSDBGrid.DBGrid grdGeneral 
      Bindings        =   "frmSResrc.frx":0000
      Height          =   7455
      Left            =   9000
      OleObjectBlob   =   "frmSResrc.frx":0019
      TabIndex        =   15
      Top             =   600
      Width           =   8715
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   8895
      _Version        =   65536
      _ExtentX        =   15690
      _ExtentY        =   6588
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
      Begin VB.Data datGenTab 
         Caption         =   "General Data"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox cmbStaffNo 
         Height          =   315
         Left            =   2880
         TabIndex        =   2
         Top             =   2640
         Width           =   4455
      End
      Begin VB.Label telno 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4200
         TabIndex        =   33
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Category:"
         Height          =   255
         Left            =   2880
         TabIndex        =   32
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label cat 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4200
         TabIndex        =   14
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label qual 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4200
         TabIndex        =   13
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label job 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4200
         TabIndex        =   12
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label dept 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4200
         TabIndex        =   11
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label disc 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4200
         TabIndex        =   10
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Qualification:"
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contact No :"
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Job Title:"
         Height          =   255
         Left            =   2880
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Department:"
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Discipline:"
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.Image StaffPic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2235
         Left            =   120
         Picture         =   "frmSResrc.frx":09EF
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2475
      End
      Begin VB.Label StaffNames 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   3120
         Width           =   6795
      End
      Begin VB.Label Label1 
         BackColor       =   &H00400000&
         Caption         =   "Staff Identification:"
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
         Left            =   120
         TabIndex        =   3
         Top             =   2640
         Width           =   2415
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   3675
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   4440
      Width           =   8895
      _Version        =   65536
      _ExtentX        =   15690
      _ExtentY        =   6482
      _StockProps     =   15
      Caption         =   " Resource Allocation"
      ForeColor       =   12582912
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   0
      BevelInner      =   2
      Font3D          =   3
      Alignment       =   6
      Begin VB.Data datGeneral 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox cmbPClass 
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1320
         Width           =   3975
      End
      Begin VB.ComboBox cmbSports 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1320
         Width           =   3375
      End
      Begin VB.ComboBox cmbMast 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2280
         Width           =   3375
      End
      Begin VB.ComboBox cmbSClass 
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   840
         Width           =   3975
      End
      Begin VB.ComboBox cmbSkills 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1800
         Width           =   3375
      End
      Begin VB.ComboBox cmbsubJ 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   840
         Width           =   3375
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   735
         Left            =   120
         TabIndex        =   36
         Top             =   2760
         Width           =   8655
         _Version        =   65536
         _ExtentX        =   15266
         _ExtentY        =   1296
         _StockProps     =   15
         BackColor       =   16744703
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox cmbRemAll 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   240
            Width           =   6495
         End
         Begin VB.CommandButton cmdDel 
            Caption         =   "&Remove Resource"
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sports:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Class:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   5160
         TabIndex        =   23
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Skills:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subject:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Class Master:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2280
         Width           =   1095
      End
   End
   Begin Threed.SSPanel SSPanel5 
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   17775
      _Version        =   65536
      _ExtentX        =   31353
      _ExtentY        =   1085
      _StockProps     =   15
      Caption         =   "Staff Responsibility Allocations "
      ForeColor       =   16576
      BackColor       =   11639171
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Font3D          =   2
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   855
      Left            =   0
      TabIndex        =   27
      Top             =   8160
      Width           =   17775
      _Version        =   65536
      _ExtentX        =   31353
      _ExtentY        =   1508
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
      Begin VB.CommandButton cmdAddG 
         Caption         =   "Add &Group"
         Height          =   435
         Left            =   3960
         TabIndex        =   43
         ToolTipText     =   "Will Add selected Subject to Group of Class the Selected Class Belongs "
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddL 
         Caption         =   "Add &Level"
         Height          =   435
         Left            =   5280
         TabIndex        =   42
         ToolTipText     =   "Will Add selected Subject to School Level where the Selected Class Belongs "
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H000000FF&
         Caption         =   "&Reset"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Will Clear all Resource Alllocations"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdTeach 
         Caption         =   "List &Teachers"
         Height          =   435
         Left            =   13200
         TabIndex        =   40
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdList 
         Caption         =   "List &Subjects"
         Height          =   435
         Left            =   11880
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H00C0FFFF&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   16440
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   435
         Left            =   14640
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddC 
         Caption         =   "Add &Class"
         Height          =   435
         Left            =   2640
         TabIndex        =   28
         ToolTipText     =   "Will Add selected Subject to Selected Class"
         Top             =   240
         Width           =   1215
      End
      Begin Threed.SSCommand CmdCUpd 
         Height          =   615
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Assigns Teachers to Classrooms"
         Top             =   120
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Update Classrooms"
         ForeColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
         Outline         =   0   'False
      End
      Begin VB.Label SubName 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   9000
         TabIndex        =   39
         Top             =   240
         Width           =   2775
      End
   End
End
Attribute VB_Name = "FrmSResrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim db1 As Database, wrktemp As Workspace
Dim rstLMast As Recordset, FVal As Boolean
Dim PicName As String, AddFlag As Boolean
Dim strVal As String, tempBuff As String

Dim rstClassZ As Recordset, i As Integer
Dim rstClass As Recordset, rstSelAlloc As Recordset
Dim rstSkills As Recordset, rstSubj As Recordset
Dim rstStAlloc As Recordset, rstSports As Recordset
Dim rstFilt As Recordset

Dim rstDept, rstCRoom, rstMCRoom As Recordset
Dim rstJob As Recordset, rstDisc As Recordset
Dim rstQual As Recordset, rstCat As Recordset

Dim MsubJCode, mSubjDesc As String
Dim mSptCode, mSptDesc As String
Dim mCMCode, mCMDESC As String, CRFlag As Boolean

Dim mRTaget, mRCode As String
Dim strpos, strpos1 As Integer, DelFlag As Boolean
Dim firstTime As Integer, FirstPass, AddType As Integer
Dim GenClass As New DProLog, BRStaffno As String



Private Sub cmbPClass_Click()
    'On Error GoTo PError
    Dim i As Integer
   If firstTime = 1 Then
      Exit Sub
   End If
    If AddFlag = True Then
        i = ConfirmP()
        If i = vbYes Then
            FiltRecP
            strpos = InStr(1, cmbStaffNo, ",", 1)
            tmpStr1 = Left(cmbStaffNo, strpos - 1)
            strpos = InStr(1, cmbSports, ",", 1)
            tmpStr2 = Left(cmbSports, strpos - 1)
            rstFilt.MoveFirst
            Do While Not rstFilt.EOF
               'strpos = InStr(1, cmbPClass, ",", 1)
               'tmpStr3 = Left(cmbPClass, strpos - 1)
               mvSql = "DELETE * FROM TabStaffResrc where staffno ='" & tmpStr1 & "' and rcode = '" & tmpStr2 & "' and rtaget = '" & rstFilt!code & "'"
               db1.Execute mvSql
               rstCRoom.Requery
               rstStAlloc.AddNew
               strpos = InStr(1, cmbSports, ",", 1)
               mSptCode = Left(cmbSports, strpos - 1)
               mSptDesc = UCase(Mid(cmbSports, strpos + 1))
               strpos = InStr(1, cmbStaffNo, ",", 1)
               tempBuff = "StaffNo = '" & Left(cmbStaffNo, strpos - 1) & "'"
               BRStaffno = Left(cmbStaffNo, strpos - 1)
               rstStAlloc("staffno") = Left(cmbStaffNo, strpos - 1)
               rstStAlloc("staffnames") = Mid(cmbStaffNo, strpos + 1)
               strpos = InStr(1, cmbSClass, ",", 1)
               rstStAlloc("rcode") = mSptCode
               rstStAlloc("rdesc") = mSptDesc
               rstStAlloc("rtaget") = rstFilt!code
               rstStAlloc("tdesc") = rstFilt!code & " " & rstFilt!Desc & " " & rstFilt!NickName
               rstStAlloc("rdate") = Date
               rstStAlloc("rapply") = "Sports to Train"
               rstStAlloc.Update
               rstFilt.MoveNext
            Loop
            Brefresh
            cmbPClass.Enabled = True
        Else
            MsgBox "Class to supervise Sports not allocated", vbExclamation, "Resource Allocation"
            cmbPClass.Enabled = False
        End If
    End If
    
PError0:
    Exit Sub
PError:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0
End Sub

Private Sub cmbsClass_Click()
    On Error GoTo PError
    Dim i As Integer
   If firstTime = 1 Then
      Exit Sub
   End If
    If AddFlag = True Then
        i = ConfirmS()
        If i = vbYes Then
            FiltRecS
            strpos = InStr(1, cmbsubJ, ",", 1)
            MsubJCode = Left(cmbsubJ, strpos - 1)
            mSubjDesc = Mid(cmbsubJ, strpos + 1)
            strpos = InStr(1, cmbStaffNo, ",", 1)
            tmpStr1 = Left(cmbStaffNo, strpos - 1)
            rstFilt.MoveFirst
            Do While Not rstFilt.EOF
               'strpos1 = InStr(1, cmbSClass, ",", 1)
               'tmpStr2 = Left(cmbSClass, strpos1 - 1)
               mvSql = "DELETE * FROM TabStaffResrc where staffno ='" & tmpStr1 & "' and rcode = '" & MsubJCode & "' and rtaget = '" & rstFilt!code & "'"
               db1.Execute mvSql
               rstCRoom.Requery
               rstStAlloc.AddNew
               strpos = InStr(1, cmbStaffNo, ",", 1)
               tempBuff = "StaffNo = '" & Left(cmbStaffNo, strpos - 1) & "'"
               BRStaffno = Left(cmbStaffNo, strpos - 1)
               rstStAlloc("staffno") = Left(cmbStaffNo, strpos - 1)
               rstStAlloc("staffnames") = Mid(cmbStaffNo, strpos + 1)
               strpos = InStr(1, cmbSClass, ",", 1)
               rstStAlloc("rcode") = MsubJCode
               rstStAlloc("rdesc") = mSubjDesc
               rstStAlloc("rtaget") = rstFilt!code
               rstStAlloc("tdesc") = rstFilt!code & " " & rstFilt!Desc & " " & rstFilt!NickName
               rstStAlloc("rdate") = Date
               rstStAlloc("rapply") = "Subject to Teach"
               rstStAlloc.Update
               rstFilt.MoveNext
            Loop
            CRFlag = True
            Brefresh
            cmbSClass.Enabled = True
        Else
            MsgBox "Class to lecture not Allocated", vbExclamation, "Resource Allocation"
            cmbSClass.Enabled = False
        End If
    End If
PError0:
    Exit Sub
PError:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0
End Sub


Private Sub cmbsubJ_click()
    'On Error GoTo PError
    Dim i As Integer
   If firstTime = 1 Then
      Exit Sub
   End If
    If AddFlag = True Then
            strpos = InStr(1, cmbsubJ, ",", 1)
            MsubJCode = Left(cmbsubJ, strpos - 1)
            mSubjDesc = Mid(cmbsubJ, strpos + 1)
    End If
cmbSClass.Enabled = True
PError0:
    Exit Sub
PError:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0
End Sub

Private Sub cmbSkills_Click()
    'On Error GoTo perror
    Dim i As Integer
   If firstTime = 1 Then
      Exit Sub
   End If
    If AddFlag = True Then
        i = Confirm()
        If i = vbYes Then
            rstStAlloc.AddNew
            strpos = InStr(1, cmbStaffNo, ",", 1)
            tempBuff = "StaffNo = '" & Left(cmbStaffNo, strpos - 1) & "'"
            BRStaffno = Left(cmbStaffNo, strpos - 1)
            rstStAlloc("staffnames") = Mid(cmbStaffNo, strpos + 1)
            rstStAlloc("staffno") = Left(cmbStaffNo, strpos - 1)
            strpos = InStr(1, cmbSkills, ",", 1)
            rstStAlloc("rcode") = Left(cmbSkills, strpos - 1)
            rstStAlloc("rdesc") = UCase(Mid(cmbSkills, strpos + 1))
            rstStAlloc("rtaget") = "9999"
            rstStAlloc("tdesc") = "STUDENT GROUP"
            rstStAlloc("rdate") = Date
            rstStAlloc("rapply") = "Special Skill to Harness"
            rstStAlloc.Update
            Brefresh
        Else
            MsgBox "Skills Resource not Added", vbExclamation, "Resource Allocation"
        End If
    End If
PError0:
    Exit Sub
PError:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0

End Sub

Private Sub cmbMast_Click()
    'On Error GoTo PError
    Dim i As Integer
   If firstTime = 1 Then
      Exit Sub
   End If
    If AddFlag = True Then
        i = ConfirmM()
        If i = vbYes Then
            FiltRecM
            rstFilt.MoveFirst
            Do While Not rstFilt.EOF
               rstStAlloc.AddNew
               strpos = InStr(1, cmbStaffNo, ",", 1)
               tempBuff = "StaffNo = '" & Left(cmbStaffNo, strpos - 1) & "'"
               BRStaffno = Left(cmbStaffNo, strpos - 1)
               rstStAlloc("staffno") = Left(cmbStaffNo, strpos - 1)
               rstStAlloc("staffnames") = Mid(cmbStaffNo, strpos + 1)
               strpos = InStr(1, cmbMast, ",", 1)
               rstStAlloc("rcode") = rstFilt!code
               rstStAlloc("rdesc") = "CLASS SUPERVISOR"
               'rstStAlloc("rdesc") = Mid(cmbMast, strpos + 1)
               rstStAlloc("rtaget") = rstFilt!code
               rstStAlloc("tdesc") = rstFilt!code & " " & rstFilt!Desc & " " & rstFilt!NickName
               rstStAlloc("rdate") = Date
               rstStAlloc("rapply") = "Class to Supervise"
               rstStAlloc.Update
               rstFilt.MoveNext
            Loop
            Brefresh
        Else
            MsgBox "Class Master Responsibility not Allocated", vbExclamation, "Resource Allocation"
        End If
    End If
PError0:
    Exit Sub
PError:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0
End Sub
Private Sub cmbsports_Click()
    'On Error GoTo PError
    Dim i As Integer
   If firstTime = 1 Then
      Exit Sub
   End If
    If AddFlag = True Then
       strpos = InStr(1, cmbSports, ",", 1)
       mSptCode = Left(cmbSports, strpos - 1)
       mSptDesc = UCase(Mid(cmbSports, strpos + 1))
    End If
PError0:
cmbPClass.Enabled = True
    Exit Sub
PError:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0
End Sub

Private Sub cmbRemAll_Click()
'On Error GoTo PError
   If firstTime = 1 Then
      Exit Sub
   End If
   Dim i As Integer
   Dim Buff1 As Variant, Buff2 As Variant
   If rstStAlloc.BOF And rstStAlloc.EOF Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
   If DelFlag = True Then
        i = DeleteCheck()
        If i = vbYes Then
            strpos = InStr(1, cmbStaffNo, ",", 1)
            Buff1 = Left(cmbStaffNo, strpos - 1)
            BRStaffno = Left(cmbStaffNo, strpos - 1)
            strpos = InStr(1, cmbRemAll, ",", 1)
            Buff2 = Left(cmbRemAll, strpos - 1)
            mvSql = "Staffno = '" & Buff1 & "'" & " and RCode ='" & Buff2 & "'"
            rstStAlloc.FindFirst mvSql
            rstStAlloc.Delete
            CRFlag = True
            rstStAlloc.Requery
            Brefresh
        End If
        ' referesh combo box
        datgenTab.DatabaseName = DBpath
        datgenTab.RecordSource = "TabStaffResrc"
        datgenTab.ReadOnly = False
        datgenTab.Exclusive = False
        datgenTab.Refresh
        mvSql = "Select * from TabStaffResrc Where Staffno = '" & BRStaffno & "'"
        Set rstSelAlloc = db1.OpenRecordset(mvSql, dbOpenDynaset)
        rstSelAlloc.Requery
        If rstSelAlloc.RecordCount <> 0 Then
           rstSelAlloc.MoveFirst
        End If
        cmbRemAll.Clear
        Do While Not rstSelAlloc.EOF
           cmbRemAll.AddItem rstSelAlloc("Rcode") + ", " + rstSelAlloc("RDesc") + " - " + rstSelAlloc("tDesc")
           rstSelAlloc.MoveNext
        Loop
   End If
PError:
End Sub

Private Sub cmbStaffNo_LostFocus()
 On Error GoTo PError
    datgenTab.DatabaseName = DBpath
    datgenTab.RecordSource = "TabStaffResrc"
    datgenTab.ReadOnly = False
    datgenTab.Exclusive = False
    datgenTab.Refresh
    strpos = InStr(1, cmbStaffNo, ",", 1)
    If strpos = 0 Then
        tempBuff = cmbStaffNo
        BRStaffno = tempBuff
    Else
        tempBuff = Left(cmbStaffNo, strpos - 1)
        BRStaffno = Left(cmbStaffNo, strpos - 1)
    End If
    tempBuff = "StaffNo = '" & tempBuff & "'"
    rstLMast.FindFirst tempBuff
    If rstLMast.NOMATCH Then
        Exit Sub
    Else
        cmbStaffNo = rstLMast!StaffNo & ", " & rstLMast!StaffNames
    End If
    
    mvSql = "Select * from TabStaffResrc Where Staffno = '" & BRStaffno & "'"
    Set rstSelAlloc = db1.OpenRecordset(mvSql, dbOpenDynaset)


    rstSelAlloc.Requery
    If rstSelAlloc.RecordCount <> 0 Then
      rstSelAlloc.MoveFirst
    End If
    cmbRemAll.Clear
    Do While Not rstSelAlloc.EOF
         cmbRemAll.AddItem rstSelAlloc("Rcode") + ", " + rstSelAlloc("RDesc") + " - " + rstSelAlloc("tDesc")
         rstSelAlloc.MoveNext
    Loop
    ClearData
    showdata
    Brefresh
PError:
End Sub

Private Sub cmdAddG_Click()
'On Error GoTo PError
AddType = 1
If cmdAddG.Caption = "Add &Group" Then
    'rstStAlloc.AddNew
    AddFlag = True
    cmdAddG.Caption = "&Finish"
    cmdAddC.Enabled = False
    cmdAddG.Enabled = True
    cmdAddL.Enabled = False
    cmdReset.Enabled = False
Else
    cmdAddG.Caption = "Add &Group"
    AddFlag = False
    ' referesh combo box
    datgenTab.DatabaseName = DBpath
    datgenTab.RecordSource = "TabStaffResrc"
    datgenTab.ReadOnly = False
    datgenTab.Exclusive = False
    datgenTab.Refresh
    mvSql = "Select * from TabStaffResrc Where Staffno = '" & BRStaffno & "'"
    Set rstSelAlloc = db1.OpenRecordset(mvSql, dbOpenDynaset)
    rstSelAlloc.Requery
    If rstSelAlloc.RecordCount <> 0 Then
       rstSelAlloc.MoveFirst
    End If
    cmbRemAll.Clear
    Do While Not rstSelAlloc.EOF
       cmbRemAll.AddItem rstSelAlloc("Rcode") + ", " + rstSelAlloc("RDesc") + " - " + rstSelAlloc("tDesc")
       rstSelAlloc.MoveNext
    Loop
    cmbSClass.Enabled = False
    cmbPClass.Enabled = False
    cmdAddC.Enabled = True
    cmdAddG.Enabled = True
    cmdAddL.Enabled = True
    cmdReset.Enabled = True
End If
PError:
End Sub

Private Sub cmdAddL_Click()
On Error GoTo PError
AddType = 2
If cmdAddL.Caption = "Add &Level" Then
    'rstStAlloc.AddNew
    AddFlag = True
    cmdAddL.Caption = "&Finish"
    cmdAddC.Enabled = False
    cmdAddG.Enabled = False
    cmdAddL.Enabled = True
    cmdReset.Enabled = False
Else
    cmdAddL.Caption = "Add &Level"
    AddFlag = False
    ' referesh combo box
    datgenTab.DatabaseName = DBpath
    datgenTab.RecordSource = "TabStaffResrc"
    datgenTab.ReadOnly = False
    datgenTab.Exclusive = False
    datgenTab.Refresh
    mvSql = "Select * from TabStaffResrc Where Staffno = '" & BRStaffno & "'"
    Set rstSelAlloc = db1.OpenRecordset(mvSql, dbOpenDynaset)
    rstSelAlloc.Requery
    If rstSelAlloc.RecordCount <> 0 Then
       rstSelAlloc.MoveFirst
    End If
    cmbRemAll.Clear
    Do While Not rstSelAlloc.EOF
       cmbRemAll.AddItem rstSelAlloc("Rcode") + ", " + rstSelAlloc("RDesc") + " - " + rstSelAlloc("tDesc")
       rstSelAlloc.MoveNext
    Loop
    cmbSClass.Enabled = False
    cmbPClass.Enabled = False
    cmdAddC.Enabled = True
    cmdAddG.Enabled = True
    cmdAddL.Enabled = True
    cmdReset.Enabled = True
End If
PError:
End Sub

Private Sub CmdCUpd_Click()
SubName.Caption = ""
If rstCRoom.BOF And rstCRoom.EOF Then
    MsgBox "Class rooms have not been created", vbExclamation, "Teachers Assignment"
Else
   mvSql = "Select * from qrysubjteach "
   Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
   rst.Requery
   rst.MoveFirst
   If rst.RecordCount > 0 Then
       i = classupD()
       If i = vbYes Then
          strVal = " "
          mvSql = "UPDATE mastclassroom SET staffno = '" & strVal & "'"
          db1.Execute mvSql
          mvSql = "UPDATE mastclassroom SET staffnames = '" & strVal & "'"
          db1.Execute mvSql
          rstMCRoom.Requery
          
          Do While Not rst.EOF
              mRTaget = rst!rtaget
              GTBuff = rst!RDesc
              mRCode = rst!RCode
              mvSql = "roomid = '" & mRTaget & "' and subjcode = '" & mRCode & "'"
              rstMCRoom.MoveFirst
              rstMCRoom.FindFirst mvSql
              If rstMCRoom.NOMATCH Then
                 PMsg = "The Classroom " & mRTaget & " does not offer the Subject " & mRCode & " - " & GTBuff & " for teaching by " & rst!StaffNames
                 PMsg = PMsg + ". Either Add the subject to the class or Remove the subject from the Teacher's Role."
                 MsgBox PMsg, vbExclamation, "Teachers Assignment"
              Else
                 rstMCRoom.Edit
                 rstMCRoom("staffno") = rst!StaffNo
                 rstMCRoom("staffnames") = rst!StaffNames
                 rstMCRoom.Update
              End If
              rst.MoveNext
              SubName.Caption = "Update in progress ..."
           Loop
           SubName.Caption = "Completed"
           CRFlag = False
       Else
           MsgBox "Class rooms subjects updated aborted", vbExclamation, "Teachers Assignment"
       End If
   Else
       MsgBox "No Teachers have been assigned the responsibility to teach", vbExclamation, "Teachers Assignment"
   End If
End If
End Sub


Private Sub CmdDel_Click()
'On Error GoTo PError
If cmdDel.Caption = "&Remove Resource" Then
    DelFlag = True
    cmdDel.Caption = "&End Remove"
    cmbPClass.Enabled = False
    cmbSClass.Enabled = False
    cmbSkills.Enabled = False
    cmbSports.Enabled = False
    cmbMast.Enabled = False
    cmbsubJ.Enabled = False
Else
    cmdDel.Caption = "&Remove Resource"
    DelFlag = False
    cmbPClass.Enabled = False
    cmbSClass.Enabled = False
    cmbSkills.Enabled = True
    cmbSports.Enabled = True
    cmbMast.Enabled = True
    cmbsubJ.Enabled = True
End If
PError:

End Sub

Private Sub cmdexit_Click()
   Dim i As Integer
   If CRFlag = True Then
      i = CRCheck()
      If i = vbYes Then
         CmdCUpd.DoClick
      Else
         MsgBox "Learning Resource not updated. You can run UPDATE CLASSROOM later to make the resource available.", vbExclamation, "Staff Responsibilities"
      End If
  End If
  Unload Me
End Sub


Private Sub cmdList_Click()
On Error GoTo PError
     mvSql = "Select staffno,rcode,rdesc,tdesc,rapply From TabStaffResrc"
     Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
     If (rst.BOF And rst.EOF) Then
         MsgBox "No Subjects assigned", vbExclamation, "Staff Responsibilities"
         cmbRemAll.Clear
         Set datGeneral.Recordset = rst
         grdGeneral.Caption = Me.Caption
     Else
         Set datGeneral.Recordset = rst
         grdGeneral.Caption = Me.Caption
         Set Col0 = grdGeneral.Columns(0)
         Set Col1 = grdGeneral.Columns(1)
         Set Col2 = grdGeneral.Columns(2)
         Set Col3 = grdGeneral.Columns(3)
         Set Col4 = grdGeneral.Columns(4)
         grdGeneral.Columns(0).Width = 1000
         grdGeneral.Columns(1).Width = 600
         grdGeneral.Columns(2).Width = 2400
         grdGeneral.Columns(3).Width = 2800
         grdGeneral.Columns(4).Width = 1800
         Col0.Caption = "Staff No."
         Col1.Caption = "Code"
         Col2.Caption = "Resource Allocation"
         Col3.Caption = "Class"
         Col4.Caption = "Responsibility "
         
   End If
PError:
End Sub

Private Sub cmdRefresh_Click()
    Brefresh
End Sub

Private Sub CmdAddC_Click()
On Error GoTo PError
AddType = 0
If cmdAddC.Caption = "Add &Class" Then
    'rstStAlloc.AddNew
    AddFlag = True
    cmdAddC.Caption = "&Finish"
    cmdAddC.Enabled = True
    cmdAddG.Enabled = False
    cmdAddL.Enabled = False
    cmdReset.Enabled = False
Else
    cmdAddC.Caption = "Add &Class"
    AddFlag = False
    ' referesh combo box
    datgenTab.DatabaseName = DBpath
    datgenTab.RecordSource = "TabStaffResrc"
    datgenTab.ReadOnly = False
    datgenTab.Exclusive = False
    datgenTab.Refresh
    mvSql = "Select * from TabStaffResrc Where Staffno = '" & BRStaffno & "'"
    Set rstSelAlloc = db1.OpenRecordset(mvSql, dbOpenDynaset)
    rstSelAlloc.Requery
    If rstSelAlloc.RecordCount <> 0 Then
       rstSelAlloc.MoveFirst
    End If
    cmbRemAll.Clear
    Do While Not rstSelAlloc.EOF
       cmbRemAll.AddItem rstSelAlloc("Rcode") + ", " + rstSelAlloc("RDesc") + " - " + rstSelAlloc("tDesc")
       rstSelAlloc.MoveNext
    Loop
    cmbSClass.Enabled = False
    cmbPClass.Enabled = False
    cmdAddC.Enabled = True
    cmdAddG.Enabled = True
    cmdAddL.Enabled = True
    cmdReset.Enabled = True
End If
PError:
End Sub

Private Sub cmdReset_Click()
On Error GoTo PError
   If mnu15 = False Then
      MsgBox "This Role can only be carried out by Administrator, Access denied. ", vbExclamation, "Staff Responsibilities"
   Else
      cmdAddC.Enabled = False
      cmdAddG.Enabled = False
      cmdAddL.Enabled = False
      cmdReset.Enabled = True
      i = cont()
      If i = vbYes Then
         db1.Execute " Delete * from TabStaffResrc"
         MsgBox "All Staff Resource Allocations have been Reset", vbExclamation, "Staff Responsibilities"
      Else
         MsgBox "Staff Resource Allocations Reset aborted", vbExclamation, "Staff Responsibilities"
      End If
      Brefresh
      cmdAddC.Enabled = True
      cmdAddG.Enabled = True
      cmdAddL.Enabled = True
      cmdReset.Enabled = True
  End If
PError:
End Sub

Private Sub CmdTeach_Click()
On Error GoTo PError
     mvSql = "Select staffno, staffnames,dept,qual,disc,job From lectmast"
     Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
     If (rst.BOF And rst.EOF) Then
         MsgBox "No Teachers Records Found", vbExclamation, "Staff Responsibilities"
         cmbRemAll.Clear
         Set datGeneral.Recordset = rst
         grdGeneral.Caption = Me.Caption
     Else
         Set datGeneral.Recordset = rst
         grdGeneral.Caption = Me.Caption
         Set Col0 = grdGeneral.Columns(0)
         Set Col1 = grdGeneral.Columns(1)
         Set Col2 = grdGeneral.Columns(2)
         Set Col3 = grdGeneral.Columns(3)
         Set Col4 = grdGeneral.Columns(4)
         Set Col5 = grdGeneral.Columns(5)
         grdGeneral.Columns(0).Width = 1000
         grdGeneral.Columns(1).Width = 2800
         grdGeneral.Columns(2).Width = 1000
         grdGeneral.Columns(3).Width = 1000
         grdGeneral.Columns(4).Width = 1000
         grdGeneral.Columns(5).Width = 1000
         Col0.Caption = "Staff No."
         Col1.Caption = "Staff Names"
         Col2.Caption = "Department"
         Col3.Caption = "Qualification"
         Col4.Caption = "Discipline"
         Col5.Caption = "Job Title "
         
   End If
PError:
End Sub

Private Sub Form_Load()
'On Error GoTo PError
     Dim flag As Integer, mtable As String
     GenClass.fleLogin mvUserid, "Accessed Staff Responsibility Allocation", Date, Time
     Call FormCentreMDI(Me)
     Set wrktemp = DBEngine.Workspaces(0)
     Set db1 = wrktemp.OpenDatabase(DBpath, True)
     Set rstSubj = db1.OpenRecordset("defsubj", dbOpenDynaset)
     Set rstClass = db1.OpenRecordset("defclass", dbOpenDynaset)
     Set rstStAlloc = db1.OpenRecordset("TabStaffResrc", dbOpenDynaset)
     Set rstSelAlloc = db1.OpenRecordset("TabStaffRtmp", dbOpenDynaset)
     Set rstSkills = db1.OpenRecordset("defskill", dbOpenDynaset)
     Set rstLMast = db1.OpenRecordset("lectmast", dbOpenDynaset)
     Set rstSports = db1.OpenRecordset("defsports", dbOpenDynaset)
     
     Set rstDept = db1.OpenRecordset("defdept", dbOpenDynaset)
     Set rstDisc = db1.OpenRecordset("defdisc", dbOpenDynaset)
     Set rstJob = db1.OpenRecordset("defjob", dbOpenDynaset)
     Set rstQual = db1.OpenRecordset("defqual", dbOpenDynaset)
     Set rstCat = db1.OpenRecordset("defstaffcat", dbOpenDynaset)
     
     Set rstCRoom = db1.OpenRecordset("defclassnames", dbOpenDynaset)
     Set rstMCRoom = db1.OpenRecordset("MastClassRoom", dbOpenDynaset)
   
     Me.Caption = MVCoyname
     firstTime = 1
     CRFlag = False
     cmbPClass.Enabled = False
     cmbSClass.Enabled = False

    ''__________________________________1
    datgenTab.DatabaseName = DBpath
    datgenTab.RecordSource = "lectmast"
    datgenTab.ReadOnly = False
    datgenTab.Exclusive = False
    datgenTab.Refresh
    Do While Not datgenTab.Recordset.EOF
       cmbStaffNo.AddItem datgenTab.Recordset("Staffno") + ", " + datgenTab.Recordset("StaffNames")
       datgenTab.Recordset.MoveNext
    Loop
    If datgenTab.Recordset.RecordCount <> 0 Then
       cmbStaffNo.ListIndex = 0
    End If
    ''_________________________________2_
    datgenTab.DatabaseName = DBpath
    datgenTab.RecordSource = "defclassnames"
    datgenTab.ReadOnly = False
    datgenTab.Exclusive = False
    datgenTab.Refresh
    datgenTab.Recordset.MoveFirst
    Do While Not datgenTab.Recordset.EOF
       If IsNull(datgenTab.Recordset("nickname")) Then
            cmbSClass.AddItem datgenTab.Recordset("code") + ", " + datgenTab.Recordset("Desc") + " - " + datgenTab.Recordset("cgrp")
            cmbPClass.AddItem datgenTab.Recordset("code") + ", " + datgenTab.Recordset("Desc") + " - " + datgenTab.Recordset("cgrp")
            cmbMast.AddItem datgenTab.Recordset("code") + ", " + datgenTab.Recordset("Desc") + " - " + datgenTab.Recordset("cgrp")
       Else
            cmbSClass.AddItem datgenTab.Recordset("code") + ", " + datgenTab.Recordset("Desc") + " " + datgenTab.Recordset("Nickname") + " - " + datgenTab.Recordset("cgrp")
            cmbPClass.AddItem datgenTab.Recordset("code") + ", " + datgenTab.Recordset("Desc") + " " + datgenTab.Recordset("Nickname") + " - " + datgenTab.Recordset("cgrp")
            cmbMast.AddItem datgenTab.Recordset("code") + ", " + datgenTab.Recordset("Desc") + " " + datgenTab.Recordset("Nickname") + " - " + datgenTab.Recordset("cgrp")
       End If
       datgenTab.Recordset.MoveNext
    Loop
    If datgenTab.Recordset.RecordCount <> 0 Then
       cmbSClass.ListIndex = 0
       cmbPClass.ListIndex = 0
       cmbMast.ListIndex = 0
    End If
    mvSql = "Select * from TabStaffResrc Where Staffno = '" & BRStaffno & "'"
    Set rstSelAlloc = db1.OpenRecordset(mvSql, dbOpenDynaset)
    rstSelAlloc.Requery
    If rstSelAlloc.RecordCount <> 0 Then
      rstSelAlloc.MoveFirst
    End If
    cmbRemAll.Clear
    Do While Not rstSelAlloc.EOF
       cmbRemAll.AddItem rstSelAlloc("Rcode") + ", " + rstSelAlloc("RDesc") + " - " + rstSelAlloc("tDesc")
       rstSelAlloc.MoveNext
    Loop
    If rstSelAlloc.RecordCount <> 0 Then
       cmbRemAll.ListIndex = 0
    End If
    
    ''__________________________________5
    datgenTab.DatabaseName = DBpath
    datgenTab.RecordSource = "defskill"
    datgenTab.ReadOnly = False
    datgenTab.Exclusive = False
    datgenTab.Refresh
    Do While Not datgenTab.Recordset.EOF
       cmbSkills.AddItem datgenTab.Recordset("code") + ", " + datgenTab.Recordset("Desc")
       datgenTab.Recordset.MoveNext
    Loop
    If datgenTab.Recordset.RecordCount <> 0 Then
       cmbSkills.ListIndex = 0
    End If
    ''__________________________________6
    datgenTab.DatabaseName = DBpath
    datgenTab.RecordSource = "defsubj"
    datgenTab.ReadOnly = False
    datgenTab.Exclusive = False
    datgenTab.Refresh
    Do While Not datgenTab.Recordset.EOF
       cmbsubJ.AddItem datgenTab.Recordset("code") + ", " + datgenTab.Recordset("Desc")
       datgenTab.Recordset.MoveNext
    Loop
    If datgenTab.Recordset.RecordCount <> 0 Then
       cmbsubJ.ListIndex = 0
    End If
    '____________________________________7
    
    datgenTab.DatabaseName = DBpath
    datgenTab.RecordSource = "defsports"
    datgenTab.ReadOnly = False
    datgenTab.Exclusive = False
    datgenTab.Refresh
    Do While Not datgenTab.Recordset.EOF
       cmbSports.AddItem datgenTab.Recordset("code") + ", " + datgenTab.Recordset("Desc")
       datgenTab.Recordset.MoveNext
    Loop
    If datgenTab.Recordset.RecordCount <> 0 Then
       cmbSports.ListIndex = 0
    End If
    '____________________________________
         If Not rstLMast.EOF Then
            rstLMast.MoveFirst
         End If
    showdata
    DelFlag = False
    firstTime = 0
    AddType = 0
    cmdAddC.Enabled = True
    cmdAddG.Enabled = True
    cmdAddL.Enabled = True
    cmdReset.Enabled = True
PError:
End Sub

Public Sub showdata()
On Error GoTo perror1
     Dim StBuff As String
     If Not (rstLMast.BOF And rstLMast.EOF) Then
        strpos = InStr(1, cmbStaffNo, ",", 1)
        tempBuff = "StaffNo = '" & Left(cmbStaffNo, strpos - 1) & "'"
        rstLMast.FindFirst tempBuff
        If rstLMast.NOMATCH Then
             MsgBox "Staff record not found", vbExclamation, "Staff Responsibilities"
        Else
           PicName = rstLMast("staffpic")
           telno = rstLMast("ptelcon")
         
           GTBuff = "code = '" & rstLMast("dept") & "'"
           rstDept.FindFirst GTBuff
           dept = rstDept!Desc
    
           GTBuff = "code = '" & rstLMast("job") & "'"
           rstJob.FindFirst GTBuff
           job = rstJob!Desc
    
           GTBuff = "code = '" & rstLMast("disc") & "'"
           rstDisc.FindFirst GTBuff
           disc = rstDisc!Desc
    
           GTBuff = "code = '" & rstLMast("qual") & "'"
           rstQual.FindFirst GTBuff
           qual = rstQual!Desc
    
           GTBuff = "code = '" & rstLMast("cat") & "'"
           rstCat.FindFirst GTBuff
           cat = rstCat!Desc
       
            'Sex = rstLMast("Sex")
            StaffNames = rstLMast("staffnames")
            PicName = rstLMast("staffpic")
            StaffPic = LoadPicture(PicName)
         End If
    End If
    Exit Sub
perror1:
    PicName = PicsPath & "picinit.bmp"
    StaffPic = LoadPicture(PicName)
End Sub

Private Sub cmbStaffNo_Click()
    strpos = InStr(1, cmbStaffNo, ",", 1)
    tempBuff = Left(cmbStaffNo, strpos - 1)
    BRStaffno = Left(cmbStaffNo, strpos - 1)
    
    mvSql = "Select * from TabStaffResrc Where Staffno = '" & BRStaffno & "'"
    Set rstSelAlloc = db1.OpenRecordset(mvSql, dbOpenDynaset)
    rstSelAlloc.Requery
    If rstSelAlloc.RecordCount <> 0 Then
      rstSelAlloc.MoveFirst
    End If
    cmbRemAll.Clear
    Do While Not rstSelAlloc.EOF
         cmbRemAll.AddItem rstSelAlloc("Rcode") + ", " + rstSelAlloc("RDesc") + " - " + rstSelAlloc("tDesc")
         rstSelAlloc.MoveNext
    Loop
    ClearData
    showdata
    Brefresh
PError:
End Sub
Private Sub ClearData()
    MsubJCode = ""
    mSubjDesc = ""
    mSptCode = ""
    mSptDesc = ""
    mCMCode = ""
    mCMDESC = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
    db1.Close
End Sub
Public Sub Brefresh()
On Error GoTo PError
     mvSql = "SELECT staffno,rcode,rdesc,tdesc,rapply From TabStaffResrc WHERE [Staffno]='" & BRStaffno & "'"
     Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
     
     If (rst.BOF And rst.EOF) Then
         'MsgBox "No Resources assigned", vbExclamation, "Staff Responsibilities"
         cmbRemAll.Clear
         Set datGeneral.Recordset = rst
         grdGeneral.Caption = Me.Caption
     Else
         Set datGeneral.Recordset = rst
         grdGeneral.Caption = Me.Caption
         Set Col0 = grdGeneral.Columns(0)
         Set Col1 = grdGeneral.Columns(1)
         Set Col2 = grdGeneral.Columns(2)
         Set Col3 = grdGeneral.Columns(3)
         Set Col4 = grdGeneral.Columns(4)
         grdGeneral.Columns(0).Width = 1000
         grdGeneral.Columns(1).Width = 600
         grdGeneral.Columns(2).Width = 2400
         grdGeneral.Columns(3).Width = 2800
         grdGeneral.Columns(4).Width = 1800
         Col0.Caption = "Staff No."
         Col1.Caption = "Code"
         Col2.Caption = "Resource Allocation"
         Col3.Caption = "Class"
         Col4.Caption = "Responsibility "
         'FGrid.Columns(4).NumberFormat = "##,###,###,###.00"
         'FGrid.Columns(5).NumberFormat = "##,###,###,###.00"
   End If
PError:
End Sub
Public Sub FiltRecS()
     Dim FltClass, FltGrp, FltLev As String
     Select Case AddType
        Case 0
           strpos = InStr(1, cmbSClass, ",", 1)
           FltClass = Trim(Left(cmbSClass, strpos - 1))
           mvSql = "Select * from defclassnames Where code = '" & FltClass & "'"
           Set rstFilt = db1.OpenRecordset(mvSql, dbOpenDynaset)
        Case 1
           strpos1 = InStr(1, cmbSClass, "-", 1)
           FltGrp = Trim(Mid(cmbSClass, strpos1 + 1))
           mvSql = "Select * from defclassnames Where cgrp = '" & FltGrp & "'"
           Set rstFilt = db1.OpenRecordset(mvSql, dbOpenDynaset)
        Case 2
          strpos1 = InStr(1, cmbSClass, "-", 1)
          FltGrp = Trim(Mid(cmbSClass, strpos1 + 1))
          GTBuff = "code = '" & FltGrp & "'"
          rstClass.FindFirst GTBuff
          FltLev = rstClass!sleV
          mvSql = "Select * from defclass Where slev = '" & FltLev & "'"
          Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
          rst.MoveFirst
          db1.Execute "DELETE * FROM defcnamestmp"
          Do While Not rst.EOF
              db1.Execute " INSERT INTO defcnamestmp SELECT * FROM [defclassnames] where cgrp = '" & rst!code & "';"
              rst.MoveNext
          Loop
          mvSql = "Select * from defcnamestmp"
          Set rstFilt = db1.OpenRecordset(mvSql, dbOpenDynaset)
    End Select
End Sub
Public Sub FiltRecP()
     Dim FltClass, FltGrp, FltLev As String
     Select Case AddType
        Case 0
           strpos = InStr(1, cmbPClass, ",", 1)
           FltClass = Trim(Left(cmbPClass, strpos - 1))
           mvSql = "Select * from defclassnames Where code = '" & FltClass & "'"
           Set rstFilt = db1.OpenRecordset(mvSql, dbOpenDynaset)
        Case 1
           strpos1 = InStr(1, cmbPClass, "-", 1)
           FltGrp = Trim(Mid(cmbPClass, strpos1 + 1))
           mvSql = "Select * from defclassnames Where cgrp = '" & FltGrp & "'"
           Set rstFilt = db1.OpenRecordset(mvSql, dbOpenDynaset)
        Case 2
          strpos1 = InStr(1, cmbPClass, "-", 1)
          FltGrp = Trim(Mid(cmbPClass, strpos1 + 1))
          GTBuff = "code = '" & FltGrp & "'"
          rstClass.FindFirst GTBuff
          FltLev = rstClass!sleV
          mvSql = "Select * from defclass Where slev = '" & FltLev & "'"
          Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
          rst.MoveFirst
          db1.Execute "DELETE * FROM defcnamestmp"
          Do While Not rst.EOF
              db1.Execute " INSERT INTO defcnamestmp SELECT * FROM [defclassnames] where cgrp = '" & rst!code & "';"
              rst.MoveNext
          Loop
          mvSql = "Select * from defcnamestmp"
          Set rstFilt = db1.OpenRecordset(mvSql, dbOpenDynaset)
    End Select
End Sub

Public Sub FiltRecM()
     Dim FltClass, FltGrp, FltLev As String
     Select Case AddType
        Case 0
           strpos = InStr(1, cmbMast, ",", 1)
           FltClass = Trim(Left(cmbMast, strpos - 1))
           mvSql = "Select * from defclassnames Where code = '" & FltClass & "'"
           Set rstFilt = db1.OpenRecordset(mvSql, dbOpenDynaset)
        Case 1
           strpos1 = InStr(1, cmbMast, "-", 1)
           FltGrp = Trim(Mid(cmbMast, strpos1 + 1))
           mvSql = "Select * from defclassnames Where cgrp = '" & FltGrp & "'"
           Set rstFilt = db1.OpenRecordset(mvSql, dbOpenDynaset)
        Case 2
          strpos1 = InStr(1, cmbMast, "-", 1)
          FltGrp = Trim(Mid(cmbMast, strpos1 + 1))
          GTBuff = "code = '" & FltGrp & "'"
          rstClass.FindFirst GTBuff
          FltLev = rstClass!sleV
          mvSql = "Select * from defclass Where slev = '" & FltLev & "'"
          Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
          rst.MoveFirst
          db1.Execute "DELETE * FROM defcnamestmp"
          Do While Not rst.EOF
              db1.Execute " INSERT INTO defcnamestmp SELECT * FROM [defclassnames] where cgrp = '" & rst!code & "';"
              rst.MoveNext
          Loop
          mvSql = "Select * from defcnamestmp"
          Set rstFilt = db1.OpenRecordset(mvSql, dbOpenDynaset)
    End Select
End Sub


