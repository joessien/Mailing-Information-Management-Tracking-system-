VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmClass 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DataPro"
   ClientHeight    =   9075
   ClientLeft      =   1755
   ClientTop       =   1395
   ClientWidth     =   12480
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   12480
   Begin Threed.SSPanel SSPelect 
      Height          =   2295
      Left            =   120
      TabIndex        =   35
      Top             =   5880
      Width           =   6615
      _Version        =   65536
      _ExtentX        =   11668
      _ExtentY        =   4048
      _StockProps     =   15
      Caption         =   "Elective Course Management"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
      Alignment       =   6
      Begin VB.ComboBox cmbRem 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   1800
         Width           =   3615
      End
      Begin Threed.SSCommand cmdFix 
         Height          =   855
         Left            =   120
         TabIndex        =   43
         Top             =   600
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "FI&X"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin VB.ComboBox cmbElect 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   600
         Width           =   3615
      End
      Begin VB.ComboBox cmbsubJ 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1080
         Width           =   3615
      End
      Begin Threed.SSCommand cmdRem 
         Height          =   495
         Left            =   120
         TabIndex        =   45
         Top             =   1680
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "RE&MOVE ELECTIVE"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Elective Group:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   1560
         TabIndex        =   41
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Class Subject:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   1680
         TabIndex        =   39
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.TextBox NickName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   33
      Top             =   2520
      Width           =   4455
   End
   Begin Threed.SSPanel LCName 
      Height          =   615
      Left            =   0
      TabIndex        =   32
      Top             =   840
      Width           =   12375
      _Version        =   65536
      _ExtentX        =   21828
      _ExtentY        =   1085
      _StockProps     =   15
      Caption         =   "Class Name "
      ForeColor       =   255
      BackColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelInner      =   1
      Font3D          =   2
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "E&xit"
      Height          =   735
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8280
      Width           =   1335
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   2895
      Left            =   9000
      TabIndex        =   21
      Top             =   1560
      Width           =   3375
      _Version        =   65536
      _ExtentX        =   5953
      _ExtentY        =   5106
      _StockProps     =   15
      Caption         =   "Lists"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
      Alignment       =   6
      Begin VB.CommandButton cmdAllStud 
         Caption         =   "All Students &By Class"
         Height          =   495
         Left            =   240
         TabIndex        =   31
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdSubDept 
         Caption         =   "All Subjects By &Dept"
         Height          =   495
         Left            =   1800
         TabIndex        =   29
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmDSubAS 
         Caption         =   "&All Subjects By Class"
         Height          =   495
         Left            =   1800
         TabIndex        =   25
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdClass 
         Caption         =   "All C&lasses"
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton CmdSubj 
         Caption         =   "Class Sub&jects"
         Height          =   495
         Left            =   1800
         TabIndex        =   24
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdStud 
         Caption         =   "Class &Students"
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.ComboBox cmbCSup 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2160
      TabIndex        =   20
      Text            =   "cmbCSup"
      Top             =   4440
      Width           =   4455
   End
   Begin MSMask.MaskEdBox CSize 
      Height          =   375
      Left            =   2160
      TabIndex        =   18
      Top             =   3480
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   3
      PromptChar      =   "_"
   End
   Begin VB.Data datgenTab 
      Caption         =   "Gentab"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox cmbCGrp 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2160
      TabIndex        =   16
      Text            =   "cmbCGrp"
      Top             =   3000
      Width           =   4455
   End
   Begin Threed.SSPanel SSAction 
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   8280
      Width           =   10935
      _Version        =   65536
      _ExtentX        =   19288
      _ExtentY        =   1296
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
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmddel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Remove"
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   495
         Left            =   1800
         TabIndex        =   30
         ToolTipText     =   "Assigns Subjects to Classrooms"
         Top             =   120
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Update Classroom"
         ForeColor       =   16384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   4
         Font3D          =   2
      End
      Begin Threed.SSCommand cmdStat 
         Height          =   495
         Left            =   6840
         TabIndex        =   46
         ToolTipText     =   "Update Class population statistics"
         Top             =   120
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Statistics"
         ForeColor       =   4194304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   4
         Font3D          =   2
      End
      Begin VB.Label lblStats 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   8280
         TabIndex        =   47
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.TextBox RDesc 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   1
      Top             =   2040
      Width           =   4455
   End
   Begin VB.TextBox RCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1560
      Width           =   795
   End
   Begin Threed.SSPanel SSPanel5 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12375
      _Version        =   65536
      _ExtentX        =   21828
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "Class Room Administration"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
      Begin VB.Data datGeneral 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   8400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   2895
      Left            =   6840
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   5106
      _StockProps     =   15
      Caption         =   " Browse"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FloodColor      =   255
      Font3D          =   2
      Alignment       =   6
      Begin VB.CommandButton Cmdfind 
         Caption         =   "&Find Class"
         Height          =   375
         Left            =   360
         TabIndex        =   28
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton cmdbott 
         Caption         =   "&Bottom"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdtop 
         Caption         =   "&Top"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
   End
   Begin MSMask.MaskEdBox SubjNo 
      Height          =   375
      Left            =   2160
      TabIndex        =   36
      ToolTipText     =   "Automatically  updated when you assign subjects to class "
      Top             =   3960
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   3
      PromptChar      =   "_"
   End
   Begin MSDBGrid.DBGrid grdGeneral 
      Bindings        =   "Frmclass.frx":0000
      Height          =   3615
      Left            =   6840
      OleObjectBlob   =   "Frmclass.frx":0019
      TabIndex        =   42
      Top             =   4560
      Width           =   5595
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Class Subjects:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   37
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nickname:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   34
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label cnote 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   120
      TabIndex        =   27
      Top             =   4920
      Width           =   6615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Class Prefect:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Class Size:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ClassGroup:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label DeptName 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Uniform Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label StaffNo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Class Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
End
Attribute VB_Name = "FrmClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EditFlag As Boolean, MaxErrors As Integer
Dim AddFlag, FixFlag, RemFlag As Boolean
Dim db1 As Database, wrktemp As Workspace
Dim AType As String, ValBuff As String
Dim firstTime As Integer, FirstPass As Integer
Dim GenClass As New DProLog

Dim rstCRoom, rstStats As Recordset
Dim rstClass, TabSelAlloc As Recordset
Dim rstcSup, rstStudmast As Recordset
Dim rstCNames, rstElect, rstDefEl As Recordset

Dim mvCgrp, mvClCode, mvsubjcode  As String
Dim i, strpos As Integer

Public Sub Brefresh()
On Error GoTo PError
     mvSql = "Select roomid,subjdesc,eldesc from tabelect"
     Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
     If (rst.BOF And rst.EOF) Then
         'MsgBox "No Resources assigned", vbExclamation, "Class Administration"
         Set datGeneral.Recordset = rst
         grdGeneral.Caption = "NO ELECTIVE SUBJECTS FOR SELECTED CLASS"
     Else
         Set datGeneral.Recordset = rst
         grdGeneral.Caption = "ELECTIVE SUBJECTS FOR SELECTED CLASS"
         Set Col0 = grdGeneral.Columns(0)
         Set Col1 = grdGeneral.Columns(1)
         Set Col2 = grdGeneral.Columns(2)
         grdGeneral.Columns(0).Width = 600
         grdGeneral.Columns(1).Width = 2600
         grdGeneral.Columns(2).Width = 2000
         Col0.Caption = "Class"
         Col1.Caption = "Subject Description"
         Col2.Caption = "Elective Group"
   End If
PError:
End Sub
Private Sub cmbRem_Click()
If IsEmpty(cmbRem) Or Trim(Len(cmbRem)) = 0 Then
   MsgBox "Invalid code", vbExclamation, "Class Administration"
   Exit Sub
Else
   cmbRem.ListIndex = 0
End If
If RemFlag = True Then
   strpos = InStr(1, cmbRem, ",", 1)
   strpos1 = InStr(1, cmbRem, "-", 1)
   tmpStr1 = Left(cmbRem, strpos - 1)
   i = strpos1 - strpos
   mSubjDesc = Trim(Mid(cmbRem, strpos + 2, i - 2))
   tmpStr2 = Trim(Mid(cmbRem, strpos1 + 1))
   db1.Execute "UPDATE mastclassroom SET elcode = 'CORE' WHERE roomid = '" & tmpStr1 & "' and subjdesc = '" & mSubjDesc & "' and elcode = '" & tmpStr2 & "'"
   CalcSubjNo
   CRRefresh
End If
End Sub
Public Sub CRRefresh()
   mvSql = "Select roomID,SubjDesc,ElCode from mastclassroom Where roomid = '" & rstCNames!code & "' and elcode <> 'CORE'"
   Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
   If (rst.BOF And rst.EOF) Then
       Set datGeneral.Recordset = rst
       grdGeneral.Caption = "NO ELECTIVE SUBJECTS FOR SELECTED CLASS"
    Else
       Set datGeneral.Recordset = rst
       grdGeneral.Caption = "ELECTIVE SUBJECTS FOR SELECTED CLASS"
       Set Col0 = grdGeneral.Columns(0)
       Set Col1 = grdGeneral.Columns(1)
       Set Col2 = grdGeneral.Columns(2)
       grdGeneral.Columns(0).Width = 800
       grdGeneral.Columns(1).Width = 2400
       grdGeneral.Columns(2).Width = 2000
       Col0.Caption = "Class"
       Col1.Caption = "Subject Description"
       Col2.Caption = "Elective Group"
       cmbRem.Clear
       rst.MoveFirst
       Do While Not rst.EOF
          cmbRem.AddItem rst("roomid") + ", " + rst("subjdesc") + " - " + rst("elcode")
          rst.MoveNext
       Loop
       ' load elective tab with current record from classroom
       mvSql = "Select * from mastclassroom Where roomid = '" & rstCNames!code & "' and elcode <> 'CORE'"
       Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
       db1.Execute "DELETE * FROM Tabelect"
       rst.MoveFirst
       Do While Not rst.EOF
          rstElect.AddNew
          rstElect("roomid") = rst!roomid
          rstElect("subjcode") = rst!subjcode
          rstElect("elcode") = rst!elcode
          rstElect("subjdesc") = rst!subjdesc
          GTBuff = "code = '" & rst!elcode & "'"
          rstDefEl.FindFirst GTBuff
          rstElect("eldesc") = rstDefEl!Desc
          rstElect.Update
          rst.MoveNext
       Loop
    End If
End Sub
Private Sub cmbsubJ_click()
On Error GoTo PError
Dim i As Integer
If FixFlag = True Then
   strpos = InStr(1, cmbsubJ, ",", 1)
   MsubJCode = Left(cmbsubJ, strpos - 1)
   mSubjDesc = Mid(cmbsubJ, strpos + 1)
   strpos = InStr(1, cmbElect, ",", 1)
   MsubJCode = Left(cmbsubJ, strpos - 1)
   tmpStr1 = Left(cmbElect, strpos - 1)
   mvSql = "DELETE * FROM Tabelect where subjcode ='" & MsubJCode & "' and elcode = '" & tmpStr1 & "' and roomid = 'N" & RCode & "'"
   db1.Execute mvSql
   rstElect.Requery
   rstElect.AddNew
   rstElect("roomid") = "N" + RCode
   rstElect("subjcode") = MsubJCode
   rstElect("elcode") = tmpStr1
   rstElect("subjdesc") = mSubjDesc
   rstElect("eldesc") = Trim(Mid(cmbElect, strpos + 2))
   rstElect.Update
   Brefresh
   Exit Sub
PError:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
End If
End Sub

Private Sub Cmdclear_Click()
   Dim i As Integer
   If rstElect.RecordCount < 1 Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
    i = DeleteCheck()
    If i = vbYes Then
        db1.Execute "DELETE * FROM tabelect"
        Brefresh
    End If

End Sub

Private Sub cmdFix_Click()
If cmdFix.Caption = "FI&X" Then
   FixFlag = True
   cmdFix.Caption = "&CLEAR"
   cmbElect.Enabled = True
   cmbsubJ.Enabled = True
Else
   cmdFix.Caption = "FI&X"
   cmbElect.Enabled = False
   cmbsubJ.Enabled = False
   If rstElect.RecordCount < 1 Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
    i = DeleteCheck()
    If i = vbYes Then
        db1.Execute "DELETE * FROM tabelect"
        Brefresh
    End If
End If
End Sub

Private Sub cmdRem_Click()
If cmdRem.Caption = "RE&MOVE ELECTIVE" Then
   RemFlag = True
   cmdRem.Caption = "&END REMOVE"
   cmdRem.ForeColor = &HFF00&
   cmbElect.Enabled = False
   cmbsubJ.Enabled = False
   cmbRem.Enabled = True
Else
   cmdRem.Caption = "RE&MOVE ELECTIVE"
   cmdRem.ForeColor = &HFF&
   cmbElect.Enabled = True
   cmbsubJ.Enabled = True
   cmbRem.Enabled = False
   RemFlag = False
End If
End Sub

Private Sub cmdStat_Click()
'On Error GoTo pERROR
Dim maLe, feMale, eTot As Integer
lblStats.Visible = True
lblStats.Caption = "Updating ..."
lblStats.Refresh
If Not (rstCNames.BOF And rstCNames.EOF) Then rstCNames.MoveFirst
db1.Execute "DELETE * FROM tabstats"
rstStats.Requery
Do While Not rstCNames.EOF
   maLe = 0
   feMale = 0
   eTot = 0
   mvSql = "Select * from studmast Where cclass = '" & rstCNames!code & "'"
   Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
   If Not (rst.BOF And rst.EOF) Then
      rst.MoveFirst
      Do While Not rst.EOF
        If rst!Sex = "Female" Then feMale = feMale + 1
        If rst!Sex = "Male" Then maLe = maLe + 1
        eTot = eTot + 1
        rst.MoveNext
      Loop
      rstStats.AddNew
      rstStats("classid") = rstCNames!code
      rstStats("male") = maLe
      rstStats("female") = feMale
      rstStats("total") = eTot
      rstStats.Update
   End If
   rstCNames.MoveNext
Loop
lblStats.Caption = "Completed"
PError:
End Sub

Private Sub cmdStat_LostFocus()
   lblStats.Visible = False
End Sub

Private Sub rcode_lostfocus()
   SubjLoad
End Sub
Public Sub SubjLoad()
On Error GoTo PError
   db1.Execute "DELETE * FROM tabelect"
   CRRefresh
   mvSql = "Select * from mastclassroom Where roomID = 'N" & RCode & "'"
   Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
   cmbsubJ.Clear
   Do While Not rst.EOF
      cmbsubJ.AddItem rst("subjcode") + ", " + rst("subjdesc")
      rst.MoveNext
   Loop
   If rst.RecordCount <> 0 Then
     cmbsubJ.ListIndex = 0
   End If
   Exit Sub
PError:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
End Sub

Private Sub cmdAllStud_Click()
On Error GoTo PError
    mvSql = "Select Cclass,StudNo,stFileno, studnames, regdate, dob, sex From studmast"
    mvSql = mvSql + " order by Cclass;"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No students have been registered in selected class", vbExclamation, "Class Administration "
         Exit Sub
     Else
         Set frmBrwse.datGeneral.Recordset = rst
         frmBrwse.grdGeneral.Caption = "LIST OF ALL STUDENTS"
         frmBrwse.Caption = Me.Caption
         Set Col0 = frmBrwse.grdGeneral.Columns(0)
         Set Col1 = frmBrwse.grdGeneral.Columns(1)
         Set Col2 = frmBrwse.grdGeneral.Columns(2)
         Set Col3 = frmBrwse.grdGeneral.Columns(3)
         Set Col4 = frmBrwse.grdGeneral.Columns(4)
         Set Col5 = frmBrwse.grdGeneral.Columns(5)
         Set Col6 = frmBrwse.grdGeneral.Columns(6)
         frmBrwse.grdGeneral.Columns(0).Width = 1000
         frmBrwse.grdGeneral.Columns(1).Width = 1400
         frmBrwse.grdGeneral.Columns(2).Width = 1800
         frmBrwse.grdGeneral.Columns(3).Width = 2800
         frmBrwse.grdGeneral.Columns(4).Width = 1000
         frmBrwse.grdGeneral.Columns(5).Width = 1000
         frmBrwse.grdGeneral.Columns(6).Width = 1000
         Col0.Caption = "Current Class"
         Col1.Caption = "Student No."
         Col2.Caption = "File Number"
         Col3.Caption = "Student Name"
         Col4.Caption = "Reg. date"
         Col5.Caption = "Date of Birth "
         Col6.Caption = "Sex"
   End If
   frmBrwse.Show
PError:
End Sub

Private Sub cmDSubAS_Click()
On Error GoTo PError
     mvSql = "Select roomid,roomdesc,classid,subjcode,subjdesc,elcode,staffnames From mastclassroom order by roomid;"
     Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
     If (rst.BOF And rst.EOF) Then
         MsgBox "No subjects defined", vbExclamation, "Class Subjects "
         Exit Sub
     Else
         Set frmBrwse.datGeneral.Recordset = rst
         frmBrwse.grdGeneral.Caption = "LIST OF ALL SUBJECTS SORTED BY CLASS"
         frmBrwse.Caption = Me.Caption
         Set Col0 = frmBrwse.grdGeneral.Columns(0)
         Set Col1 = frmBrwse.grdGeneral.Columns(1)
         Set Col2 = frmBrwse.grdGeneral.Columns(2)
         Set Col3 = frmBrwse.grdGeneral.Columns(3)
         Set Col4 = frmBrwse.grdGeneral.Columns(4)
         Set Col5 = frmBrwse.grdGeneral.Columns(5)
         Set Col6 = frmBrwse.grdGeneral.Columns(6)
         frmBrwse.grdGeneral.Columns(0).Width = 600
         frmBrwse.grdGeneral.Columns(1).Width = 1200
         frmBrwse.grdGeneral.Columns(2).Width = 1000
         frmBrwse.grdGeneral.Columns(3).Width = 1000
         frmBrwse.grdGeneral.Columns(4).Width = 2800
         frmBrwse.grdGeneral.Columns(5).Width = 1400
         frmBrwse.grdGeneral.Columns(6).Width = 2800
         Col0.Caption = "Class Code."
         Col1.Caption = "Class Name"
         Col2.Caption = "Class Group"
         Col3.Caption = "Subject Code"
         Col4.Caption = "Subject Description "
         Col5.Caption = "Offered As"
         Col6.Caption = "Teacher's Name"
   End If
   frmBrwse.Show
PError:
End Sub

Private Sub CmdAdd_Click()
 If cmdadd.Caption = "&Add" Then
    CmdDisab
    LCName.Caption = ""
    rstCNames.AddNew
    AddFlag = True
    FixFlag = False
    cmdFix.Enabled = True
    cmdSave.Enabled = True
    cmdadd.Enabled = True
    cnote.Visible = True
    cnote.Visible = True
    cmdedit.Enabled = False
    cmddel.Enabled = False
    cmdRem.Enabled = False
    cmbRem.Enabled = False
    cmbsubJ.Enabled = False
    cmbElect.Enabled = False
    FldEnab
    ClearData
    RCode.SetFocus
    cmdadd.Caption = "&Cancel"
    PMsg = "After a new Class is created, you will be able to assign subjects and students to the class. "
    PMsg = PMsg + "Then return here for Administration to select subjects in that class that are Electives. "
    PMsg = PMsg + "Also return here for administration to select one of the students as Class Prefect."
    cnote.Caption = PMsg
Else
    If Not rstCNames.EOF Then
       rstCNames.MoveFirst
    End If
    cmdFix.Caption = "FI&X"
    cmdadd.Caption = "&Add"
    CmdEnab
    EditFlag = False
    AddFlag = False
    FixFlag = False
    cmdRem.Enabled = True
    cmdedit.Enabled = True
    cmddel.Enabled = True
    cmdFix.Enabled = False
    cmdSave.Enabled = False
    cmbsubJ.Enabled = False
    cmbElect.Enabled = False
    cnote.Visible = False
    showdata
    FldDisab
End If
End Sub

Private Sub CmdBott_Click()
    FixFlag = False
    Dim Count As Long
    If rstCNames.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstCNames.MoveLast
    showdata
End Sub

Private Sub CmdClass_Click()
  On Error GoTo PError
     mvSql = "Select * From defclassnames"
     Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
     If (rst.BOF And rst.EOF) Then
         MsgBox "No class defined", vbExclamation, "Class definition "
         Exit Sub
     Else
         Set frmBrwse.datGeneral.Recordset = rst
         frmBrwse.grdGeneral.Caption = "LIST OF ALL CLASSES"
         frmBrwse.Caption = Me.Caption
         Set Col0 = frmBrwse.grdGeneral.Columns(0)
         Set Col1 = frmBrwse.grdGeneral.Columns(1)
         Set Col2 = frmBrwse.grdGeneral.Columns(2)
         Set Col3 = frmBrwse.grdGeneral.Columns(3)
         Set Col4 = frmBrwse.grdGeneral.Columns(4)
         Set Col5 = frmBrwse.grdGeneral.Columns(5)
         Set Col6 = frmBrwse.grdGeneral.Columns(6)
         Set Col7 = frmBrwse.grdGeneral.Columns(7)
         frmBrwse.grdGeneral.Columns(0).Width = 600
         frmBrwse.grdGeneral.Columns(1).Width = 1800
         frmBrwse.grdGeneral.Columns(2).Width = 1800
         frmBrwse.grdGeneral.Columns(3).Width = 800
         frmBrwse.grdGeneral.Columns(4).Width = 800
         frmBrwse.grdGeneral.Columns(5).Width = 1000
         frmBrwse.grdGeneral.Columns(6).Width = 1400
         frmBrwse.grdGeneral.Columns(7).Width = 2600
         Col0.Caption = "Class Code."
         Col1.Caption = "Class Name"
         Col2.Caption = "Nickname"
         Col3.Caption = "Class Group"
         Col4.Caption = "Class Size"
         Col5.Caption = "Subject No."
         Col6.Caption = "Prefect No. "
         Col7.Caption = "Class Prefect's Name"
   End If
   frmBrwse.Show
PError:
End Sub

Private Sub CmdDel_Click()
   Dim i As Integer
   If rstCNames.RecordCount < 1 Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
   PMsg = "While there may be a valid reason why you want to remove a Class, this is a very dangerous operation. "
   PMsg = PMsg + "Some consequences are: You will lose all information about students, subjects, and teachers in that Class. "
   PMsg = PMsg + "You will invalidate all processes, reports and enquiries that refer to that Class. "
   PMsg = PMsg + "If you just created this Class in error with no students, teachers assigned or not  processed one Term, you may remove it without consequences. "
   PMsg = PMsg + "When you Remove a Class, system will log the activity with your name and time that you did it. This action is IRREDEEMABLE. "
   PMsg = PMsg + "Are you sure you wish to continue ?"
   i = MsgBox(PMsg, vbYesNo, "Class Administration")
   
   If i = vbYes Then
       PMsg = "Deleted the Class: N" & RCode & " " & RDesc & " " & NickName & " from the system"
       GenClass.fleLogin mvUserid, PMsg, Date, Time
       rstCNames.Delete
       db1.Execute "DELETE * FROM mastclassroom Where roomid = 'N" & RCode & "'"
       If Not rstCNames.BOF Then
         rstCNames.MovePrevious
         showdata
       Else
         ClearData
       End If
   End If
End Sub

Private Sub CmdEdit_Click()
On Error GoTo PError
 If cmdedit.Caption = "&Edit" Then
     If IsEmpty("rstCNames") Then
         MsgBox ("Empty Table")
         Exit Sub
     Else
         If rstCNames.EOF Then
            rstCNames.MovePrevious
         End If
     End If
   FixFlag = False
   RCode.Enabled = False
   cmdRem.Enabled = False
   cmbRem.Enabled = False
   cmbsubJ.Enabled = False
   cmbElect.Enabled = False
   cmdFix.Enabled = True
   cmdSave.Enabled = True
   EditFlag = True
   rstCNames.Edit
   FldEnab
   RDesc.SetFocus
   cmdedit.Caption = "&Cancel"
   CmdDisab
   SubjLoad
Else
   cmdedit.Caption = "&Edit"
   cmdFix.Caption = "FI&X"
   FixFlag = False
   EditFlag = False
   cmdSave.Enabled = False
   cmbsubJ.Enabled = False
   cmbElect.Enabled = False
   cmdFix.Enabled = False
   RCode.Enabled = True
   cmdRem.Enabled = True
   CmdEnab
   showdata
   FldDisab
End If
PError:
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
On Error GoTo PError
    Dim mvLength As Integer
    Dim mvBuff As String, mvAns As String
    mvLength = 0
    mvAns = InputBox$("Enter Class ID:")
    mvLength = Len(mvAns)
    If mvLength = 0 Then
       MsgBox "No input provided, will close", vbInformation, Me.Caption
       Exit Sub
    End If
    mvBuff = "code = '" _
        & mvAns & "'"
    rstCNames.FindFirst mvBuff
    If rstCNames.NOMATCH Then
        MsgBox "Class ID Not Found", vbInformation, Me.Caption
        Exit Sub
    Else
        showdata
    End If
PError:
End Sub

Private Sub CmdNext_Click()
        Dim flag As Integer
    
    On Error GoTo NextError
    FixFlag = False
    If rstCNames.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstCNames.EOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEPAST)
        rstCNames.MoveLast
    Else
        rstCNames.MoveNext
        If rstCNames.EOF Then
            rstCNames.MoveLast
        End If
    End If
NextClear:
    showdata
    Exit Sub
NextError:
    ErrorMessages (NOMOVEPAST)
    ''rstCNames.Requery
    rstCNames.MoveLast
    On Error GoTo 0
    Resume NextClear

End Sub

Private Sub CmdPrev_Click()
  Dim flag As Integer
    On Error GoTo PrevError
    FixFlag = False
    If rstCNames.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstCNames.BOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEFRONT)
        rstCNames.MoveFirst
    Else
        rstCNames.MovePrevious
        If rstCNames.BOF Then
            rstCNames.MoveFirst
        End If
    End If
PrevClear:
    showdata
    Exit Sub
PrevError:
    ErrorMessages (NOMOVEFRONT)
    rstCNames.Requery
    rstCNames.MoveFirst
    On Error GoTo 0
    Resume PrevClear
End Sub

Private Sub CmdSave_Click()
  Dim i As Integer
    'On Error GoTo PError
    If Len(Trim(RCode)) = 0 And AddFlag = True Then
       MsgBox "Class code cannot be blank", vbCritical, "Class Administration"
       RCode.SetFocus
       Exit Sub
    End If
    ' Check that there is no elective subject selection that has only one subject
    If Not (rstDefEl.BOF And rstDefEl.EOF) Then
       rstDefEl.MoveFirst
       Do While Not rstDefEl.EOF
          i = 0
          mvSql = "Select * from tabelect Where elcode = '" & rstDefEl!code & "'"
          Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
          If Not (rst.BOF And rst.EOF) Then rst.MoveFirst
          Do While Not rst.EOF
             i = i + 1
             rst.MoveNext
          Loop
          If i = 1 Then
             If rst.EOF Then rst.MovePrevious
             MsgBox "The subject " & rst!subjdesc & " cannot be an elective onto itself. Electives should be atleast two subjects. Remove the subject or add another subject to the group", vbExclamation, "Class Administration "
             Exit Sub
          End If
          rstDefEl.MoveNext
       Loop
    End If
   'end of check
    If EditFlag = True Then
        rstCNames.Edit
    End If
    If AddFlag = True Then
        rstCNames.AddNew
        rstCNames("Code") = "N" & UCase(RCode)
    End If
    mvClCode = "N" & UCase(RCode)  ' used for updating the classroom
    rstCNames("Desc") = UCase(RDesc)
    rstCNames("nickname") = UCase(NickName)
    rstCNames("csize") = CSize
    strpos = InStr(1, cmbCGrp, ",", 1)
    If strpos = 0 Or Len(Trim(cmbCGrp)) = 0 Then
        rstCNames("cgrp") = " "
         mvCgrp = ""
    Else
        rstCNames("cgrp") = Left(cmbCGrp, strpos - 1)
        mvCgrp = Trim(UCase(Left(cmbCGrp, strpos - 1))) ' used for updating the classroom
    End If
    strpos = InStr(1, cmbCSup, ",", 1)
    If strpos = 0 Or Len(Trim(cmbCSup)) = 0 Then
        rstCNames("csupname") = " "
    Else
        rstCNames("csupno") = Left(cmbCSup, strpos - 1)
        rstCNames("csupname") = Mid(cmbCSup, strpos + 1)
    End If
    
    rstCNames.Update
    'update class room
    initCSubj
    'init classroom with elective subjects
    rstElect.Requery
    If Not (rstElect.BOF And rstElect.EOF) Then
       rstElect.MoveFirst
       Do While Not rstElect.EOF
          GTBuff = "roomid = '" & rstElect!roomid & "' and subjcode = '" & rstElect!subjcode & "'"
          rstCRoom.FindFirst GTBuff
          If rstCRoom.NOMATCH Then
             MsgBox "An elective subject " & rstElect!subjdesc & " for class  " & rstElect!roomid & " not found. Contact Administrator ", vbExclamation, "Class Administration "
          Else
             rstCRoom.Edit
             rstCRoom("elcode") = rstElect!elcode
             rstCRoom.Update
          End If
          rstElect.MoveNext
       Loop
    End If
    db1.Execute "DELETE * FROM tabelect"
    rstElect.Requery
    'find and update number of subjects in class
    rstCNames.MoveFirst
    Do While Not rstCNames.EOF
    i = 0
       mvSql = "Select * from mastclassroom Where roomid = '" & rstCNames!code & "' and elcode = 'CORE'"
       Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
       If Not (rst.BOF And rst.EOF) Then rst.MoveFirst
       Do While Not rst.EOF
         i = i + 1
         rst.MoveNext
       Loop
       If Not (rstDefEl.BOF And rstDefEl.EOF) Then
          rstDefEl.MoveFirst
          Do While Not rstDefEl.EOF
             mvSql = "Select * from mastclassroom Where roomid = '" & rstCNames!code & "' and elcode = '" & rstDefEl!code & "'"
             Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
             If Not (rst.BOF And rst.EOF) Then i = i + 1
             rstDefEl.MoveNext
          Loop
       End If
       rstCNames.Edit
       rstCNames("subjno") = i
       rstCNames.Update
       rstCNames.MoveNext
    Loop
    ' no of subject update completed
    cmdadd.Caption = "&Add"
    CmdEnab
    cmdFix.Caption = "FI&X"
    cmdRem.Enabled = True
    cmdedit.Enabled = True
    cmddel.Enabled = True
    RCode.Enabled = True
    cmbElect.Enabled = False
    cmbsubJ.Enabled = False
    RemFlag = False
    EditFlag = False
    AddFlag = False
    cnote.Visible = False
    AddFlag = False
    cmdSave.Enabled = False
    cmdFix.Enabled = False
    cmbRem.Enabled = False
    cmdedit.Caption = "&Edit"
    'rstCNames.MoveLast
    GTBuff = "code = 'N" & RCode & "'"
    rstCNames.FindFirst GTBuff
    showdata
    FldDisab
    cmdStat.DoClick
    MsgBox "Class rooms update has been completed", vbExclamation, "Class Administration"
PError0:
    Exit Sub
PError:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0
End Sub
Public Sub CalcSubjNo()
On Error GoTo PError
    rstCNames.MoveFirst
    i = 0
    mvSql = "Select * from mastclassroom Where roomid = '" & rstCNames!code & "' and elcode = 'CORE'"
    Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
    If Not (rst.BOF And rst.EOF) Then rst.MoveFirst
    Do While Not rst.EOF
       i = i + 1
       rst.MoveNext
    Loop
    If Not (rstDefEl.BOF And rstDefEl.EOF) Then
    rstDefEl.MoveFirst
    Do While Not rstDefEl.EOF
        mvSql = "Select * from mastclassroom Where roomid = '" & rstCNames!code & "' and elcode = '" & rstDefEl!code & "'"
        Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
        If Not (rst.BOF And rst.EOF) Then
           i = i + 1
        End If
        rstDefEl.MoveNext
    Loop
    End If
    rstCNames.Edit
    rstCNames("subjno") = i
    rstCNames.Update
PError:
End Sub
Private Sub cmdSup_Click()
    Dim GenView As New frmBrwse
    Set GenView.datGeneral.Recordset = rstClassZ
    GenView.Caption = Me.Caption
    GenView.Show

End Sub

Private Sub cmdStud_Click()
On Error GoTo PError
    GTBuff = "N" & RCode
    mvSql = "Select Cclass,StudNo,stFileno, studnames, regdate, dob, sex From studmast where Cclass = '"
    mvSql = mvSql + GTBuff & "' order by Cclass;"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No students have been registered in selected class", vbExclamation, "Class Administration "
         Exit Sub
     Else
         Set frmBrwse.datGeneral.Recordset = rst
         frmBrwse.grdGeneral.Caption = "LIST OF STUDENTS IN SELECTED CLASS"
         frmBrwse.Caption = Me.Caption
         Set Col0 = frmBrwse.grdGeneral.Columns(0)
         Set Col1 = frmBrwse.grdGeneral.Columns(1)
         Set Col2 = frmBrwse.grdGeneral.Columns(2)
         Set Col3 = frmBrwse.grdGeneral.Columns(3)
         Set Col4 = frmBrwse.grdGeneral.Columns(4)
         Set Col5 = frmBrwse.grdGeneral.Columns(5)
         Set Col6 = frmBrwse.grdGeneral.Columns(6)
         frmBrwse.grdGeneral.Columns(0).Width = 1000
         frmBrwse.grdGeneral.Columns(1).Width = 1400
         frmBrwse.grdGeneral.Columns(2).Width = 1800
         frmBrwse.grdGeneral.Columns(3).Width = 2800
         frmBrwse.grdGeneral.Columns(4).Width = 1200
         frmBrwse.grdGeneral.Columns(5).Width = 1200
         frmBrwse.grdGeneral.Columns(6).Width = 1000
         Col0.Caption = "Current Class"
         Col1.Caption = "Student No."
         Col2.Caption = "File Number"
         Col3.Caption = "Student Name"
         Col4.Caption = "Reg. date"
         Col5.Caption = "Date of Birth "
         Col6.Caption = "Sex"
   End If
   frmBrwse.Show
PError:
End Sub

Private Sub cmdSubDept_Click()
On Error GoTo PError

     mvSql = "Select roomid,roomdesc,classid,subjcode,subjdesc,elcode From mastclassroom order by subjcode;"
     Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
     If (rst.BOF And rst.EOF) Then
         MsgBox "No subjects defined", vbExclamation, "Class Subjects "
         Exit Sub
     Else
         Set frmBrwse.datGeneral.Recordset = rst
         frmBrwse.grdGeneral.Caption = "LIST OF ALL SUBJECTS SORTED BY CLASS"
         frmBrwse.Caption = Me.Caption
         Set Col0 = frmBrwse.grdGeneral.Columns(0)
         Set Col1 = frmBrwse.grdGeneral.Columns(1)
         Set Col2 = frmBrwse.grdGeneral.Columns(2)
         Set Col3 = frmBrwse.grdGeneral.Columns(3)
         Set Col4 = frmBrwse.grdGeneral.Columns(4)
         Set Col5 = frmBrwse.grdGeneral.Columns(5)
         'Set Col6 = frmBrwse.grdGeneral.Columns(6)
         frmBrwse.grdGeneral.Columns(0).Width = 600
         frmBrwse.grdGeneral.Columns(1).Width = 1200
         frmBrwse.grdGeneral.Columns(2).Width = 1000
         frmBrwse.grdGeneral.Columns(3).Width = 1000
         frmBrwse.grdGeneral.Columns(4).Width = 2800
         frmBrwse.grdGeneral.Columns(5).Width = 1400
         'frmBrwse.grdGeneral.Columns(6).Width = 2800
         Col0.Caption = "Class Code."
         Col1.Caption = "Class Name"
         Col2.Caption = "Class Group"
         Col3.Caption = "Subject Code"
         Col4.Caption = "Subject Description "
         Col5.Caption = "Offered As"
         'Col6.Caption = "Class Prefect's Name"
   End If
   frmBrwse.Show
PError:
End Sub

Private Sub cmdSubj_Click()
On Error GoTo PError
     GTBuff = "N" & RCode
     mvSql = "Select roomid,roomdesc,classid,subjcode,subjdesc,elcode,staffnames From mastclassroom where roomid = '" & GTBuff & "'"
     Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
     If (rst.BOF And rst.EOF) Then
         MsgBox "No subjects defined for selected class", vbExclamation, "Class Subjects "
         Exit Sub
     Else
         Set frmBrwse.datGeneral.Recordset = rst
         frmBrwse.grdGeneral.Caption = "LIST OF SUBJECTS IN SELECTED CLASS"
         frmBrwse.Caption = Me.Caption
         Set Col0 = frmBrwse.grdGeneral.Columns(0)
         Set Col1 = frmBrwse.grdGeneral.Columns(1)
         Set Col2 = frmBrwse.grdGeneral.Columns(2)
         Set Col3 = frmBrwse.grdGeneral.Columns(3)
         Set Col4 = frmBrwse.grdGeneral.Columns(4)
         Set Col5 = frmBrwse.grdGeneral.Columns(5)
         Set Col6 = frmBrwse.grdGeneral.Columns(6)
         frmBrwse.grdGeneral.Columns(0).Width = 600
         frmBrwse.grdGeneral.Columns(1).Width = 1200
         frmBrwse.grdGeneral.Columns(2).Width = 1000
         frmBrwse.grdGeneral.Columns(3).Width = 1000
         frmBrwse.grdGeneral.Columns(4).Width = 2800
         frmBrwse.grdGeneral.Columns(5).Width = 1400
         frmBrwse.grdGeneral.Columns(6).Width = 2800
         Col0.Caption = "Class Code."
         Col1.Caption = "Class Name"
         Col2.Caption = "Class Group"
         Col3.Caption = "Subject Code"
         Col4.Caption = "Subject Description "
         Col5.Caption = "Offered As"
         Col6.Caption = "Teacher's Name"
   End If
   frmBrwse.Show
PError:
End Sub

Private Sub CmdTop_Click()
    FixFlag = False
    If rstCNames.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstCNames.MoveFirst
    showdata
End Sub

Private Sub initCSubj()
On Error GoTo PError:
   mvSql = "DELETE * FROM mastclassroom Where roomid = '" & mvClCode & "'"
   db1.Execute mvSql
   rstCRoom.Requery
   mvSql = "Select * from TabCsAlloc Where classid = '" & mvCgrp & "'"
   Set TabSelAlloc = db1.OpenRecordset(mvSql, dbOpenDynaset)
   TabSelAlloc.Requery
   If Not (TabSelAlloc.BOF And TabSelAlloc.EOF) Then
      TabSelAlloc.MoveFirst
      Do While Not TabSelAlloc.EOF
         rstCRoom.AddNew
         rstCRoom("roomid") = mvClCode
         rstCRoom("roomdesc") = UCase(RDesc)
         rstCRoom("classid") = mvCgrp
         rstCRoom("subjcode") = TabSelAlloc!SubCode
         mvsubjcode = TabSelAlloc!SubCode
         rstCRoom("subjdesc") = TabSelAlloc!subdesc
         rstCRoom("subdate") = TabSelAlloc!subdate
         rstCRoom("subclass") = TabSelAlloc!subclass
         rstCRoom("subclassid") = TabSelAlloc!sublev
         rstCRoom("extsubj") = TabSelAlloc!extsubJ
         rstCRoom("elcode") = "CORE"
         'find teachers for the subjects in that class
         mvSql = "Select * from TabStaffResrc Where rtaget = '" & mvClCode & "' and rcode ='" & mvsubjcode & "'"
         Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
         If Not (rst.BOF And rst.EOF) Then
            If IsNull(rst!StaffNo) Then
               rstCRoom("staffno") = ""
               rstCRoom("staffnames") = ""
            Else
               rstCRoom("staffno") = rst!StaffNo
               rstCRoom("staffnames") = rst!StaffNames
            End If
         Else
            rstCRoom("staffno") = ""
            rstCRoom("staffnames") = ""
         End If
         'end find teacher
         rstCRoom.Update
         TabSelAlloc.MoveNext
      Loop
      cnote.Caption = "Update in progress " & Progind
   Else
      DoEvents
   End If
   rstCNames.MoveNext
   Progind = Progind & "."
   cnote.Caption = "Completed"
PError:
End Sub

Private Sub Form_Load()
    Dim flag As Integer, mtable As String
    GenClass.fleLogin mvUserid, "Accessed Class Room Administration", Date, Time
    Call FormCentreMDI(Me)
    Set wrktemp = DBEngine.Workspaces(0)
    Set db1 = wrktemp.OpenDatabase(DBpath, True)
    Set rstCNames = db1.OpenRecordset("DefclassNames", dbOpenDynaset)
    Set rstClass = db1.OpenRecordset("DefClass", dbOpenDynaset)
    Set rstcSup = db1.OpenRecordset("qryclasssup", dbOpenDynaset)
    Set rstStudmast = db1.OpenRecordset("studmast", dbOpenDynaset)
    Set rstCRoom = db1.OpenRecordset("mastclassroom", dbOpenDynaset)
    Set rstElect = db1.OpenRecordset("tabelect", dbOpenDynaset)
    Set rstDefEl = db1.OpenRecordset("defelect", dbOpenDynaset)
    Set rstStats = db1.OpenRecordset("tabstats", dbOpenDynaset)
    Me.Caption = MVCoyname
    If Not rstCNames.EOF Then
       rstCNames.MoveFirst
    End If
    EditFlag = False
    AddFlag = False
    FixFlag = False
    'cmddel.Visible = False
    cmdSave.Enabled = False
    cmdFix.Enabled = False
    Me.Caption = MVCoyname
             ''__________________________________
    If Not (rstClass.BOF And rstClass.EOF) Then rstClass.MoveFirst
    Do While Not rstClass.EOF
       cmbCGrp.AddItem rstClass("code") + ", " + rstClass("Desc")
       rstClass.MoveNext
    Loop
    If rstClass.RecordCount > 1 Then
       cmbCGrp.ListIndex = 0
    End If
    '_______________
    cmbElect.Clear
    If Not (rstDefEl.BOF And rstDefEl.EOF) Then
       rstDefEl.MoveFirst
       Do While Not rstDefEl.EOF
          cmbElect.AddItem rstDefEl("code") + ", " + rstDefEl("Desc")
          rstDefEl.MoveNext
       Loop
       cmbElect.ListIndex = 0
    End If
    '_______________
    
    If rstCNames.BOF And rstCNames.EOF Then
        cmbCSup = " "
    Else
        rstCNames.MoveFirst
        GTBuff = Trim(rstCNames!code)
        mvSql = "Select studno,studnames From studmast where Cclass = '" & GTBuff & "'"
        Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
        rst.Requery
        Do While Not rst.EOF
           cmbCSup.AddItem rst("Studno") + ", " + rst("studnames")
           rst.MoveNext
         Loop
         If rst.RecordCount > 1 Then
            cmbCSup.ListIndex = 0
         End If
     End If
    '__________
    showdata
    cnote.Visible = False
    cmbElect.Enabled = False
    cmbsubJ.Enabled = False
    cmbRem.Enabled = False
    RemFlag = False
    FldDisab

End Sub

Public Sub showdata()
'On Error GoTo PError
     If rstCNames.RecordCount = 1 Then
        rstCNames.MoveFirst
     End If
     If rstCNames.EOF = True Then
        rstCNames.MovePrevious
     End If
     If rstCNames.RecordCount >= 1 Then
        RCode = Right(rstCNames("Code"), 3)
        RDesc = rstCNames("Desc")
        SubjLoad
        If IsNull(rstCNames!NickName) Then
           NickName = ""
        Else
           NickName = rstCNames("nickname")
        End If
        'display number of subjects to be taken by each student in class
        If IsNull(rstCNames!SubjNo) Then
           SubjNo = 0
        Else
           SubjNo = rstCNames!SubjNo
           mvSql = "Select roomID,SubjDesc,ElCode from mastclassroom Where roomid = '" & rstCNames!code & "' and elcode <> 'CORE'"
           Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
           If (rst.BOF And rst.EOF) Then
              Set datGeneral.Recordset = rst
              grdGeneral.Caption = "NO ELECTIVE SUBJECTS FOR SELECTED CLASS"
           Else
              Set datGeneral.Recordset = rst
              grdGeneral.Caption = "ELECTIVE SUBJECTS FOR SELECTED CLASS"
              Set Col0 = grdGeneral.Columns(0)
              Set Col1 = grdGeneral.Columns(1)
              Set Col2 = grdGeneral.Columns(2)
              grdGeneral.Columns(0).Width = 800
              grdGeneral.Columns(1).Width = 2400
              grdGeneral.Columns(2).Width = 2000
              Col0.Caption = "Class"
              Col1.Caption = "Subject Description"
              Col2.Caption = "Elective Group"
              cmbRem.Clear
              rst.MoveFirst
              Do While Not rst.EOF
                 cmbRem.AddItem rst("roomid") + ", " + rst("subjdesc") + " - " + rst("elcode")
                 rst.MoveNext
              Loop
           End If
        End If
        CSize = rstCNames("cSize")
        GTBuff = "code = '" & rstCNames("cGrp") & "'"
        rstClass.FindFirst GTBuff
        cmbCGrp = rstClass!code & "," & rstClass!Desc
        LCName.Caption = rstCNames!code & " - " & RDesc & " - " & rstClass!Desc
        If IsNull(rstCNames!csupno) = True Then
           'no supervisor so show list of students in that class if any
           GTBuff = Trim(rstCNames!code)
           mvSql = "Select studno,studnames From studmast where Cclass = '" & GTBuff & "'"
           Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
           rst.Requery
           cmbCSup.Clear
           Do While Not rst.EOF
              cmbCSup.AddItem rst("Studno") + ", " + rst("studnames")
              rst.MoveNext
           Loop
           cmbCSup = ""
        Else
           ' load list of students in that class
           GTBuff = Trim(rstCNames!code)
           mvSql = "Select studno,studnames From studmast where Cclass = '" & GTBuff & "'"
           Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
           rst.Requery
           cmbCSup.Clear
           Do While Not rst.EOF
              cmbCSup.AddItem rst("Studno") + ", " + rst("studnames")
              rst.MoveNext
           Loop
           If rst.RecordCount > 1 Then
              GTBuff = "studno = '" & rstCNames("csupno") & "'"
              rstStudmast.FindFirst GTBuff
              cmbCSup = rstStudmast!StudNo & "," & rstStudmast!StudNames
           Else
              cmbCSup = " "
           End If
           cmbCSup.Enabled = False
           ' end listing
        End If
    End If
PError:
End Sub

Public Sub ClearData()
    RCode = ""
    RDesc = ""
    CSize = 40
    SubjNo = 0
    NickName = ""
    cmbCSup.Clear
End Sub
Public Sub CmdEnab()
    cmdadd.Enabled = True
    cmddel.Enabled = True
    cmdtop.Enabled = True
    cmdbott.Enabled = True
    cmdnext.Enabled = True
    cmdprev.Enabled = True
    cmdFind.Enabled = True
End Sub

Public Sub CmdDisab()
    cmdadd.Enabled = False
    cmddel.Enabled = False
    cmdtop.Enabled = False
    cmdbott.Enabled = False
    cmdnext.Enabled = False
    cmdprev.Enabled = False
    cmdFind.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    db1.Close
End Sub
Private Sub FldEnab()
    RCode.Enabled = True
    RDesc.Enabled = True
    NickName.Enabled = True
    CSize.Enabled = True
    cmbCGrp.Enabled = True
    cmbCSup.Enabled = True
End Sub

Private Sub FldDisab()
    RCode.Enabled = False
    RDesc.Enabled = False
    NickName.Enabled = False
    CSize.Enabled = False
    cmbCGrp.Enabled = False
    cmbCSup.Enabled = False

End Sub

