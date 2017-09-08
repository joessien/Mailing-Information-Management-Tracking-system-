VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmTest 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datapro"
   ClientHeight    =   9840
   ClientLeft      =   1890
   ClientTop       =   1260
   ClientWidth     =   15960
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   15960
   Begin Threed.SSPanel SSPanel2 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   9000
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   " "
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
      Begin Threed.SSCommand cmdSave 
         Height          =   495
         Left            =   720
         TabIndex        =   0
         Top             =   120
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "&Save Result"
         ForeColor       =   4194304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   5
         Font3D          =   2
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H0080FFFF&
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
         Height          =   375
         Left            =   3120
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   9855
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   16095
      _Version        =   65536
      _ExtentX        =   28390
      _ExtentY        =   17383
      _StockProps     =   15
      Caption         =   " "
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
      BevelWidth      =   3
      BevelInner      =   1
      Begin VB.ComboBox cmbSClass 
         Height          =   315
         Left            =   11760
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   960
         Width           =   4095
      End
      Begin VB.ComboBox cmbsubJ 
         Height          =   315
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   960
         Width           =   3615
      End
      Begin VB.Data datGeneral 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Data datGenTab 
         Caption         =   "General Data"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   13800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin Threed.SSPanel PanSSA 
         Height          =   3975
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Width           =   4335
         _Version        =   65536
         _ExtentX        =   7646
         _ExtentY        =   7011
         _StockProps     =   15
         Caption         =   "Student Score Adjustment"
         ForeColor       =   12583104
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
         FloodColor      =   12583104
         Font3D          =   2
         Alignment       =   6
         Begin VB.ComboBox cmbStud 
            Height          =   315
            Left            =   1080
            TabIndex        =   8
            Top             =   1800
            Width           =   3135
         End
         Begin VB.ComboBox cmbrefS 
            Height          =   315
            ItemData        =   "FrmTest.frx":0000
            Left            =   1080
            List            =   "FrmTest.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   600
            Width           =   3135
         End
         Begin VB.CommandButton cmdUpd 
            BackColor       =   &H0000FF00&
            Caption         =   "&Update"
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
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   3360
            Width           =   975
         End
         Begin VB.TextBox sComment 
            Appearance      =   0  'Flat
            Height          =   525
            Left            =   1080
            MaxLength       =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   2640
            Width           =   3135
         End
         Begin MSMask.MaskEdBox StudScore 
            Height          =   375
            Left            =   2040
            TabIndex        =   9
            Top             =   3360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Eval Ref:"
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Student NO:"
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter New  Score:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   3480
            Width           =   2055
         End
         Begin VB.Label StudNames 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   1080
            TabIndex        =   15
            Top             =   2160
            Width           =   3135
         End
         Begin VB.Label evalSubjS 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   1080
            TabIndex        =   14
            Top             =   960
            Width           =   3135
         End
         Begin VB.Label LClass 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   1080
            TabIndex        =   13
            Top             =   1320
            Width           =   3135
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Subject:"
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Class:"
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label label55 
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
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   2760
            Width           =   735
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   735
         Left            =   4440
         TabIndex        =   21
         Top             =   120
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "Test Results Administration"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "01"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin MSDBGrid.DBGrid grdGeneral 
         Bindings        =   "FrmTest.frx":0004
         Height          =   6375
         Left            =   4560
         OleObjectBlob   =   "FrmTest.frx":001D
         TabIndex        =   22
         ToolTipText     =   "When updating the grid, always use the down arrow to commit change into the cell"
         Top             =   2640
         Width           =   11355
      End
      Begin Threed.SSPanel PanLR 
         Height          =   2535
         Left            =   120
         TabIndex        =   23
         Top             =   6480
         Width           =   4335
         _Version        =   65536
         _ExtentX        =   7646
         _ExtentY        =   4471
         _StockProps     =   15
         Caption         =   "List Results"
         ForeColor       =   12583104
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
         Font3D          =   2
         Alignment       =   6
         Begin VB.ComboBox cmbRefL 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   600
            Width           =   3255
         End
         Begin VB.CommandButton cmdVoid 
            BackColor       =   &H008080FF&
            Caption         =   "&Void"
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
            Left            =   3000
            MaskColor       =   &H0080C0FF&
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   1920
            Width           =   975
         End
         Begin VB.CommandButton cmdView 
            BackColor       =   &H0000FF00&
            Caption         =   "&Approval"
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
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox VReason 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   960
            MaxLength       =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            Top             =   1440
            Width           =   3255
         End
         Begin VB.Label evalSubjL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   960
            Width           =   4095
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Eval. Ref:"
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label15 
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
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1560
            Width           =   735
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   1455
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   4335
         _Version        =   65536
         _ExtentX        =   7646
         _ExtentY        =   2566
         _StockProps     =   15
         Caption         =   "Evaluation Reference"
         ForeColor       =   12583104
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
         Font3D          =   2
         Alignment       =   6
         Begin VB.TextBox NExRef 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            MaxLength       =   12
            TabIndex        =   33
            Top             =   600
            Width           =   1335
         End
         Begin Threed.SSCommand cmdOk 
            Height          =   495
            Left            =   3240
            TabIndex        =   32
            Top             =   720
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   873
            _StockProps     =   78
            Caption         =   "&Accept"
            BevelWidth      =   3
            Font3D          =   4
         End
         Begin MSMask.MaskEdBox evDate 
            Height          =   255
            Left            =   1680
            TabIndex        =   34
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Subject Reference:"
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Evaluation Date:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   240
            TabIndex        =   35
            Top             =   960
            Width           =   1455
         End
      End
      Begin Threed.SSCommand cmdRoles 
         Height          =   375
         Left            =   5520
         TabIndex        =   37
         ToolTipText     =   "Teaching responsibilities defined for current user"
         Top             =   2160
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "&Classroom Roles"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
      End
      Begin Threed.SSPanel LCtr 
         Height          =   615
         Left            =   15120
         TabIndex        =   38
         ToolTipText     =   "Total number of students in selected class"
         Top             =   1800
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   15
         ForeColor       =   255
         BackColor       =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         BevelInner      =   1
         Font3D          =   2
      End
      Begin Threed.SSPanel ssPM 
         Height          =   615
         Left            =   9600
         TabIndex        =   39
         ToolTipText     =   "Maximum score obtainable in this assessment for this class"
         Top             =   1800
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   15
         ForeColor       =   8388608
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
         BevelWidth      =   3
         BevelInner      =   1
         Font3D          =   2
      End
      Begin Threed.SSPanel cTerm 
         Height          =   735
         Left            =   -240
         TabIndex        =   40
         Top             =   120
         Width           =   4335
         _Version        =   65536
         _ExtentX        =   7646
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "Term"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         BevelInner      =   2
         Font3D          =   4
      End
      Begin Threed.SSCommand cmdList 
         Height          =   375
         Left            =   7320
         TabIndex        =   41
         ToolTipText     =   "Full Responsibilities defined for current user"
         Top             =   2160
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "&Responsibilities"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   615
         Left            =   4560
         TabIndex        =   42
         Top             =   9120
         Width           =   11295
         _Version        =   65536
         _ExtentX        =   19923
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   " "
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
         Begin VB.TextBox teaCher 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   480
            MaxLength       =   12
            TabIndex        =   58
            ToolTipText     =   "If you are not the teacher of this subject, Enter the Teacher's ID number here"
            Top             =   120
            Width           =   1935
         End
         Begin VB.ComboBox cmbStudEval 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   5160
            TabIndex        =   43
            Top             =   120
            Width           =   3615
         End
         Begin Threed.SSCommand cmdvResult 
            Height          =   375
            Left            =   9360
            TabIndex        =   44
            Top             =   120
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "View All Results"
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Select Student Number:"
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
            Left            =   2880
            TabIndex        =   45
            Top             =   240
            Width           =   2175
         End
      End
      Begin Threed.SSCommand cmdCRooms 
         Height          =   375
         Left            =   11760
         TabIndex        =   57
         ToolTipText     =   "Teaching responsibilities defined for current user"
         Top             =   2160
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "&Class List"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   4680
         TabIndex        =   56
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Class:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   11040
         TabIndex        =   55
         Top             =   960
         Width           =   735
      End
      Begin VB.Label CName 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   11760
         TabIndex        =   54
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label CGrp 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   11760
         TabIndex        =   53
         Top             =   1800
         Width           =   3135
      End
      Begin VB.Label SubjnamE 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   5520
         TabIndex        =   52
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label subjmasT 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   5520
         TabIndex        =   51
         Top             =   1800
         Width           =   3615
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Class Group:"
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
         Left            =   10680
         TabIndex        =   50
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Class Name:"
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
         Left            =   10680
         TabIndex        =   49
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Teacher:"
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
         Left            =   4680
         TabIndex        =   48
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
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
         Left            =   4680
         TabIndex        =   47
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblAms 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Unrestricted"
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   9480
         TabIndex        =   46
         ToolTipText     =   $"FrmTest.frx":09F3
         Top             =   2400
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim db1 As Database, wrktemp As Workspace
Dim GenClass As New DProLog
Dim MAStudno(10000) As String

Dim firstTime  As Boolean, EFlag As Boolean
Dim FVal As Boolean, AddFlag As Boolean

Dim rstLMast As Recordset, rstStudmast As Recordset
Dim rstClass As Recordset, rstSelClass As Recordset
Dim rstSubj As Recordset, rstCNames As Recordset
Dim rstMCRoom As Recordset, rstEvals As Recordset
Dim rstInpTab As Recordset, rstSelRef As Recordset
Dim rstAutoB As Recordset, rstMarkS, rstEvalstest As Recordset

Dim MsubJCode, mSubjDesc, tempTaB As String
Dim MStaffNo, mStudno, Remarks As String
Dim mCRCode, mCRDesc, mCRGrp As String
Dim strVal, tempBuff As String

Dim Emark As Double, StudCtr As Integer
Dim mvMarks As Integer, MarkPtr As Integer
Dim strpos As Integer, Emask As Integer, i As Integer



Private Sub cmbRefL_Click()
   grdGeneral.Enabled = False
    On Error GoTo PError
   cmbrefS.Enabled = True
   cmdVoid.Enabled = False
   strpos = InStr(1, cmbRefL, ",", 1)
   GTBuff = Trim(Mid(cmbRefL, strpos + 2))
   tmpStr1 = Trim(Left(cmbRefL, strpos - 1))
   If mvAcad = True Then
      mvSql = "Select * from qryEvalstest Where evalref = '" & tmpStr1 & "' and staffno = '" & mvUserid & "'"
   Else
      mvSql = "Select * from qryEvalstest Where evalref = '" & tmpStr1 & "'"
   End If
   'Set rstSelClass = db1.OpenRecordset(mvSql, dbOpenDynaset)
   Set rstEvalstest = db1.OpenRecordset(mvSql, dbOpenDynaset)
   mvSql = "evalref = '" & Trim(Left(cmbRefL, strpos - 1)) & "'"
   'find subject reference end
   rstEvalstest.FindFirst mvSql
   If rstEvalstest.NOMATCH Then
      DoEvents
   Else
      evalSubjL = rstEvalstest!subject & ", " & GTBuff
   End If
   'find subject reference end
   cmbStud.Enabled = False
   StudScore.Enabled = False
   sComment.Enabled = False
   cmdUpd.Enabled = False
   cmdView.Enabled = True
   VReason.Enabled = True
   strpos = InStr(1, cmbRefL, ",", 1)
   GTBuff = Trim(Left(cmbRefL, strpos - 1))
   Viewref
PError:
End Sub

Private Sub cmbrefS_Click()
    On Error GoTo PError
   grdGeneral.Enabled = False
   strpos = InStr(1, cmbrefS, ",", 1)
   GTBuff = Trim(Left(cmbrefS, strpos - 1))
   evalSubjS = Trim(Mid(cmbrefS, strpos + 2))
   If mvAcad = True Then
      mvSql = "Select * from qryEvalstest Where evalref = '" & GTBuff & "' and staffno = '" & mvUserid & "'"
   Else
      mvSql = "Select * from qryEvalstest Where evalref = '" & GTBuff & "'"
   End If
   Set rstSelClass = db1.OpenRecordset(mvSql, dbOpenDynaset)
   rstSelClass.Requery
   cmbRefL.Enabled = True
   cmdView.Enabled = False
   VReason.Enabled = False
   cmdVoid.Enabled = False
   cmbStud.Enabled = True
   cmbStud.Clear
   rstSelClass.Requery
   Do While Not rstSelClass.EOF
      mvSql = "studno = '" & rstSelClass("studno") & "'"
      rstStudmast.Requery
      rstStudmast.FindFirst mvSql
      If rstStudmast.NOMATCH Then
         cmbStud.AddItem rstSelClass("studno") + ", "
      Else
         cmbStud.AddItem rstSelClass("studno") + ", " + rstStudmast("studnames")
         ' find class description
         mvSql = "code = '" & rstSelClass("class") & "'"
         rstCNames.Requery
         rstCNames.FindFirst mvSql
         If rstCNames.NOMATCH Then
             LClass = ""
         Else
            LClass = rstCNames!code & " , " & rstCNames!Desc & " " & rstCNames!NickName & " - " & rstCNames!CGrp
         End If
         ' end of class desc find
      End If
      rstSelClass.MoveNext
   Loop
   If rstSelClass.RecordCount <> 0 Then
     cmbSClass.ListIndex = 0
   End If
   strpos = InStr(1, cmbrefS, ",", 1)
   GTBuff = Trim(Left(cmbrefS, strpos - 1))
   Viewref
PError0:
    Exit Sub
PError:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0
End Sub

Private Sub cmbsClass_Click()
On Error GoTo PError
   grdGeneral.Enabled = True
   cmdSave.Enabled = True
   Brefresh
PError0:
    Exit Sub
PError:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0
End Sub
Private Sub cmbStud_Click()
On Error GoTo PError
   strpos = InStr(1, cmbStud, ",", 1)
   GTBuff = Trim(Left(cmbStud, strpos - 1))
   tmpStr1 = InStr(1, cmbrefS, ",", 1)
   tempBuff = Trim(Left(cmbrefS, tmpStr1 - 1))
   mvSql = "Select * from qryEvalstest where evalref = '" & tempBuff & "'"
   Set rstSelClass = db1.OpenRecordset(mvSql, dbOpenDynaset)
   mvSql = "studno = '" & GTBuff & "' and evalref = '" & tempBuff & "'"
   rstSelClass.Requery
   rstSelClass.FindFirst mvSql
   If rstSelClass.NOMATCH Then
      StudScore = 0
      sComment = ""
   Else
      StudScore.Enabled = True
      sComment.Enabled = True
      cmdUpd.Enabled = True
      StudScore = Str(rstSelClass!score)
      If IsNull(rstSelClass!Remarks) Then
         sComment = ""
      Else
         sComment = rstSelClass!Remarks
      End If
      cmbRefL.Enabled = True
   End If
   ' find the students name
    mvSql = "StudNo = '" & GTBuff & "'"
    rstStudmast.FindFirst mvSql
    If rstStudmast.NOMATCH Then
        StudNames.Caption = ""
        Exit Sub
    Else
        StudNames.Caption = rstStudmast!StudNames
    End If
PError:
End Sub

Private Sub cmbStud_LostFocus()
 On Error GoTo PError
    strpos = InStr(1, cmbStud, ",", 1)
    If strpos = 0 Then
        GTBuff = cmbStud
    Else
        GTBuff = Trim(Left(cmbStud, strpos - 1))
    End If
   tmpStr1 = InStr(1, cmbrefS, ",", 1)
   tempBuff = Trim(Left(cmbrefS, tmpStr1 - 1))
   mvSql = "Select * from qryEvalstest where evalref = '" & tempBuff & "'"
   Set rstSelClass = db1.OpenRecordset(mvSql, dbOpenDynaset)
   mvSql = "studno = '" & GTBuff & "' and evalref = '" & tempBuff & "'"
   rstSelClass.Requery
   rstSelClass.FindFirst mvSql
   If rstSelClass.NOMATCH Then
      StudNames.Caption = ""
      MsgBox "Student number not found", vbCritical, "Evaluation Update"
   Else
      StudScore.Enabled = True
      sComment.Enabled = True
      cmdUpd.Enabled = True
      StudScore = rstSelClass!score
      If IsNull(rstSelClass!Remarks) Then
         sComment = ""
      Else
         sComment = rstSelClass!Remarks
      End If
      cmbRefL.Enabled = True
   End If
   ' find the students name
    mvSql = "StudNo = '" & GTBuff & "'"
    rstStudmast.FindFirst mvSql
    If rstStudmast.NOMATCH Then
        StudNames.Caption = ""
        Exit Sub
    Else
        StudNames.Caption = rstStudmast!StudNames
    End If
PError:
End Sub

Private Sub cmbsubJ_click()
    On Error GoTo PError
   firstTime = False
   strpos = InStr(1, cmbsubJ, ",", 1)
   MsubJCode = Left(cmbsubJ, strpos - 1)
   mSubjDesc = Trim(Mid(cmbsubJ, strpos + 1))
   cmbSClass.Enabled = True
   If mvAcad = True Then
      mvSql = "Select * from mastclassroom Where Staffno = '" & mvUserid & "' and subjcode = '"
      mvSql = mvSql & MsubJCode & "' and extsubj = false"
   Else
      mvSql = "Select * from mastclassroom Where subjcode = '" & MsubJCode & "' and extsubj = false"
   End If
   Set rstSelClass = db1.OpenRecordset(mvSql, dbOpenDynaset)
   rstSelClass.Requery
   If rstSelClass.RecordCount <> 0 Then
      rstSelClass.MoveFirst
      SubjnamE = rstSelClass!subjdesc
      If mvAcad = True Then
         subjmasT = rstSelClass!StaffNames
      Else
         'subjmasT = mvUserid
         subjmasT = mvUserName
      End If
      CName.Caption = rstSelClass!roomDesc
      CGrp.Caption = rstSelClass!classid
      mCRCode = rstSelClass("roomid")
      mCRDesc = rstSelClass("roomdesc")
   Else
      SubjnamE.Caption = ""
      subjmasT.Caption = ""
      CName.Caption = ""
      CGrp.Caption = ""
      mCRCode = ""
      mCRDesc = ""
   End If
   cmbSClass.Clear
   Do While Not rstSelClass.EOF
      cmbSClass.AddItem rstSelClass("roomid") + ", " + rstSelClass("roomdesc") + " - " + rstSelClass("classid")
      rstSelClass.MoveNext
   Loop
   If rstSelClass.RecordCount <> 0 Then
     cmbSClass.ListIndex = 0
   End If
PError0:
    firstTime = True
    Exit Sub
PError:
    MsgBox "The error: " + Error$ + " has occurred", vbCritical, "Error Condition Detected"
    On Error GoTo 0
    Resume PError0
End Sub


Private Sub cmbSubj_LostFocus()
On Error GoTo PError
    firstTime = False
    strpos = InStr(1, cmbsubJ, ",", 1)
    If strpos = 0 Then
        tempBuff = cmbsubJ
    Else
        tempBuff = Left(cmbsubJ, strpos - 1)
    End If
    
   If mvAcad = True Then
      tempBuff = "subjcode = '" & tempBuff & "' and staffno = '" & mvUserid & "'"
   Else
      tempBuff = "subjcode = '" & tempBuff & "'"
   End If
    
    rstMCRoom.FindFirst tempBuff
    If rstMCRoom.NOMATCH Then
       MsgBox "Staff does not lecture entered subject", vbExclamation, "Evaluation"
       Exit Sub
    Else
        cmbsubJ = rstMCRoom!subjcode & ", " & rstMCRoom!subjdesc
    End If
   strpos = InStr(1, cmbsubJ, ",", 1)
   MsubJCode = Left(cmbsubJ, strpos - 1)
   mSubjDesc = Trim(Mid(cmbsubJ, strpos + 1))
   cmbSClass.Enabled = True
   
   If mvAcad = True Then
      mvSql = "Select * from mastclassroom Where Staffno = '" & mvUserid & "' and subjcode = '"
      mvSql = mvSql & MsubJCode & "' and extsubj = false"
   Else
      mvSql = "Select * from mastclassroom Where subjcode = '" & MsubJCode & "' and extsubj = false"
   End If
   Set rstSelClass = db1.OpenRecordset(mvSql, dbOpenDynaset)
   rstSelClass.Requery
   If rstSelClass.RecordCount <> 0 Then
      rstSelClass.MoveFirst
      SubjnamE = rstSelClass!subjdesc
      If mvAcad = True Then
         subjmasT = rstSelClass!StaffNames
      Else
         'subjmasT = mvUserid
         subjmasT = mvUserName
      End If
      CName.Caption = rstSelClass!roomDesc
      CGrp.Caption = rstSelClass!classid
      mCRCode = rstSelClass("roomid")
      mCRDesc = rstSelClass("roomdesc")
   Else
      SubjnamE.Caption = ""
      subjmasT.Caption = ""
      CName.Caption = ""
      CGrp.Caption = ""
      mCRCode = ""
      mCRDesc = ""
   End If
   cmbSClass.Clear
   Do While Not rstSelClass.EOF
      cmbSClass.AddItem rstSelClass("roomid") + ", " + rstSelClass("roomdesc") + " - " + rstSelClass("classid")
      rstSelClass.MoveNext
   Loop
   firstTime = False
   If rstSelClass.RecordCount <> 0 Then
      cmbSClass.ListIndex = 0
   End If
   firstTime = True
PError:
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub


Private Sub cmdList_Click()
On Error GoTo PError
     mvSql = "Select staffno,rcode,rdesc,tdesc,rapply From TabStaffResrc"
     mvSql = mvSql + " where staffno = '" & mvUserid & "' order by rapply,rdesc;"
     Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
     If (rst.BOF And rst.EOF) Then
         MsgBox "Current Teacher has no Responsibilities assigned", vbExclamation, "Staff Responsibilities"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "FULL RESPONSIBILITIES FOR " & UCase(subjmasT)
         frmBrowse.Caption = Me.Caption & " "
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         frmBrowse.grdGeneral.Columns(0).Width = 1400
         frmBrowse.grdGeneral.Columns(1).Width = 1000
         frmBrowse.grdGeneral.Columns(2).Width = 3000
         frmBrowse.grdGeneral.Columns(3).Width = 2200
         frmBrowse.grdGeneral.Columns(4).Width = 2500
         Col0.Caption = "Staff No."
         Col1.Caption = "Code"
         Col2.Caption = "Resource Allocation"
         Col3.Caption = "Class"
         Col4.Caption = "Responsibility "
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub cmdRefresh_Click()
    Brefresh
End Sub
Private Sub cmdok_Click()
On Error GoTo PError
grdGeneral.Enabled = False
NExRef = UCase(NExRef)
cmbStud.Enabled = False
cmdUpd.Enabled = False
StudScore.Enabled = False
sComment.Enabled = False
cmdVoid.Enabled = False
cmdView.Enabled = False
VReason.Enabled = False
If cmdOk.Caption = "&Accept" Then
   If NExRef = "" Or IsNull(NExRef) Or Len(evDate) = 0 Or evDate > Date Then
      MsgBox "Invalid reference or date, please correct", vbExclamation, "Evaluation"
      cmbsubJ.Enabled = False
      cmbSClass.Enabled = False
   Else
      mvSql = "evalref = '" & NExRef & "'"
      rstEvals.FindFirst mvSql
      If rstEvals.NOMATCH Then
         cmbsubJ.Enabled = True
         cmbsubJ.SetFocus
         cmdOk.Caption = "&Cancel"
         PanLR.Enabled = False
         PanSSA.Enabled = False
      Else
         MsgBox "Evaluation reference is already used or has already been entered, select a different Reference", vbExclamation, "Evaluation"
      End If
  End If
Else
   cmdOk.Caption = "&Accept"
   NExRef.Enabled = True
   evDate.Enabled = True
   PanLR.Enabled = True
   PanSSA.Enabled = True
   cmbsubJ.Enabled = False
   cmbSClass.Enabled = False
   grdGeneral.Enabled = False
   cmdSave.Enabled = False
End If
PError:
End Sub


Private Sub cmdRoles_Click()
On Error GoTo PError
     mvSql = "Select roomid,roomdesc,classid,subjcode,subjdesc,staffno, staffnames From mastclassroom "
     mvSql = mvSql + " where staffno = '" & mvUserid & "' order by subjcode;"
     Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
     If (rst.BOF And rst.EOF) Then
         MsgBox "Current user has no Roles defined as Teacher", vbExclamation, "Teachers Roles "
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "ROLES FOR " & UCase(subjmasT) & ": - CLASSES AND SUBJECTS TO TEACH"
         frmBrowse.Caption = Me.Caption & " "
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         frmBrowse.grdGeneral.Columns(0).Width = 600
         frmBrowse.grdGeneral.Columns(1).Width = 1200
         frmBrowse.grdGeneral.Columns(2).Width = 1000
         frmBrowse.grdGeneral.Columns(3).Width = 1000
         frmBrowse.grdGeneral.Columns(4).Width = 2800
         frmBrowse.grdGeneral.Columns(5).Width = 1400
         frmBrowse.grdGeneral.Columns(6).Width = 2800
         Col0.Caption = "Class Code."
         Col1.Caption = "Class Name"
         Col2.Caption = "Class Group"
         Col3.Caption = "Subject Code"
         Col4.Caption = "Subject Description "
         Col5.Caption = "Staff No. "
         Col6.Caption = "Staff Names"
     
   End If
   frmBrowse.Show
PError:
End Sub

Private Sub CmdUpd_Click()
On Error GoTo PError
   strpos = InStr(1, cmbStud, ",", 1)
   StudNames = Trim(Mid(cmbStud, strpos + 2))
   mStudno = Trim(Left(cmbStud, strpos - 1))
   tmpStr1 = InStr(1, cmbrefS, ",", 1)
   tempBuff = Trim(Left(cmbrefS, tmpStr1 - 1))
   mvSql = "studno = '" & mStudno & "' and evalref = '" & tempBuff & "'"
   rstSelClass.Requery
   rstSelClass.FindFirst mvSql
   If rstSelClass.NOMATCH Then
      MsgBox "Score Entry not found for student " & mStudno, vbExclamation, "Evaluation Update"
   Else
      If StudScore < 0 Or StudScore > mvMarks Then
         MsgBox "Score for Student number " & mStudno & " is less than zero or more than the maximum mark that can be earned in this subject, Reenter", vbExclamation, "Evaluation Input"
         StudScore.SetFocus
         Exit Sub
      End If
      If Len(Trim((sComment))) = 0 Then
         MsgBox "You must state the reason for changing scores.", vbExclamation, "Evaluation Update"
         sComment.SetFocus
         Exit Sub
      End If
      rstSelClass.Edit
      rstSelClass("score") = StudScore
      rstSelClass("remarks") = sComment
      rstSelClass.Update
      GenClass.fleLogin mvUserid, "Changed Student Result", Date, Time
      '__________add  review to student notes
      If Len(Trim((sComment))) > 0 Then
         rstAutoB.AddNew
         rstAutoB("studno") = Trim(mStudno)
         rstAutoB("autodate") = Date
         rstAutoB("autodet") = sComment
         rstAutoB("userid") = mvUserid
         rstAutoB.Update
      End If
      MsgBox "Update has been successful", vbExclamation, "Evaluation Update"
      '__________View the new result
      strpos = InStr(1, cmbrefS, ",", 1)
      GTBuff = Trim(Left(cmbrefS, strpos - 1))
      mvSql = "Select evdate,evalref,Studno,class,subject,score,staffno,remarks from  qryEvalstest where evalref = '" & GTBuff & "'"
      Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
      rst.Requery
      If (rst.BOF And rst.EOF) Then
         MsgBox "No Students in selected class", vbExclamation, "Evaluation"
         Set datGeneral.Recordset = rst
         grdGeneral.Caption = Me.Caption
      Else
         Set datGeneral.Recordset = rst
         grdGeneral.Caption = Me.Caption & " - Class Test Input"
         Set Col0 = grdGeneral.Columns(0)
         Set Col1 = grdGeneral.Columns(1)
         Set Col2 = grdGeneral.Columns(2)
         Set Col3 = grdGeneral.Columns(3)
         Set Col4 = grdGeneral.Columns(4)
         Set Col5 = grdGeneral.Columns(5)
         Set Col6 = grdGeneral.Columns(6)
         Set Col7 = grdGeneral.Columns(7)
         grdGeneral.Columns(0).Width = 1000
         grdGeneral.Columns(1).Width = 1200
         grdGeneral.Columns(2).Width = 1200
         grdGeneral.Columns(3).Width = 600
         grdGeneral.Columns(4).Width = 800
         grdGeneral.Columns(5).Width = 800
         grdGeneral.Columns(6).Width = 1400
         grdGeneral.Columns(7).Width = 4000
         Col0.Caption = "Eval. Date"
         Col1.Caption = "Reference"
         Col2.Caption = "Student No."
         Col3.Caption = "Class"
         Col4.Caption = "Subject"
         Col5.Caption = "Score"
         Col6.Caption = "Teacher's No."
         Col7.Caption = "Remarks"
         cmdVoid.Enabled = True
      End If
   End If
PError:
End Sub

Private Sub cmdView_Click()
   If mnu15 = False Then
      MsgBox "You do not have enough priviledge to Void Results, Approval denied. Contact the Administrator", vbExclamation, "Evaluation"
      cmdVoid.Enabled = False
   Else
      MsgBox "You have enough priviledge to Void Results, Approval Granted", vbExclamation, "Evaluation"
      cmdVoid.Enabled = True
   End If
End Sub
Private Sub Viewref()
On Error GoTo PError
     grdGeneral.Enabled = True
     mvSql = "Select evdate,evalref,Studno,class,subject,score,staffno,remarks from  qryEvalstest where evalref = '" & GTBuff & "'"
     Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
     rst.Requery
     If (rst.BOF And rst.EOF) Then
         MsgBox "No Students in selected class", vbExclamation, "Evaluation"
         Set datGeneral.Recordset = rst
         grdGeneral.Caption = Me.Caption
     Else
         Set datGeneral.Recordset = rst
         grdGeneral.Caption = Me.Caption & " - Class Test Input"
         Set Col0 = grdGeneral.Columns(0)
         Set Col1 = grdGeneral.Columns(1)
         Set Col2 = grdGeneral.Columns(2)
         Set Col3 = grdGeneral.Columns(3)
         Set Col4 = grdGeneral.Columns(4)
         Set Col5 = grdGeneral.Columns(5)
         Set Col6 = grdGeneral.Columns(6)
         Set Col7 = grdGeneral.Columns(7)
         grdGeneral.Columns(0).Width = 1200
         grdGeneral.Columns(1).Width = 1400
         grdGeneral.Columns(2).Width = 1200
         grdGeneral.Columns(3).Width = 600
         grdGeneral.Columns(4).Width = 800
         grdGeneral.Columns(5).Width = 800
         grdGeneral.Columns(6).Width = 1400
         grdGeneral.Columns(7).Width = 4000
         Col0.Caption = "Eval. Date"
         Col1.Caption = "Reference"
         Col2.Caption = "Student No."
         Col3.Caption = "Class"
         Col4.Caption = "Subject"
         Col5.Caption = "Score"
         Col6.Caption = "Teacher's No."
         Col7.Caption = "Remarks"
   End If
PError:
End Sub
Private Sub cmdVoid_Click()
On Error GoTo PError
   strpos = InStr(cmbRefL, ",")
   GTBuff = Trim(Left(cmbRefL, strpos - 1))
   Dim i As Integer
   If rstEvals.RecordCount < 1 Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
   If Len(Trim((VReason))) = 0 Then
      MsgBox "You must state the reason for voiding the entire student scores for the assessment.", vbExclamation, "Evaluation Update"
      VReason.SetFocus
      Exit Sub
   End If
   i = VoidCheck()
   If i = vbYes Then
       db1.Execute "DELETE * FROM TabEvals where evalref = '" & GTBuff & "'"
       ''_________filter evaluation records for user who is logging on
       If mvAcad = True Then
          mvSql = "Select * from qryEvalstest Where Staffno = '"
          mvSql = mvSql & mvUserid & "'"
       Else
          mvSql = "Select * from qryEvalstest "
       End If
       Set rstEvalstest = db1.OpenRecordset(mvSql, dbOpenDynaset)
      'end ____________________
       'GenClass.fleLogin mvUserid, "Deleted entire class result", Date, Time
       VReason = "Deleted " & evalSubjL & ": " & VReason
       VReason = Left(VReason, 60)
       GenClass.fleLogin mvUserid, VReason, Date, Time
       rstEvals.Requery
       MsgBox "Evaluation with reference " & UCase(GTBuff) & " has been removed", vbExclamation, "Evaluation Removal"
       ' clear the grid
       If Not (rstEvals.BOF And rstEvals.EOF) Then
          mvSql = "Select Studno,subject,score from qryEvalstest where evalref = '" & GTBuff & "'"
          Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
          rst.Requery
          Set datGeneral.Recordset = rst
        Else
          Set datGeneral.Recordset = rstEvals
        End If
        grdGeneral.Caption = Me.Caption
        Set datGeneral.Recordset = rst
        grdGeneral.Caption = Me.Caption & " - Class Test Input"
        Set Col0 = grdGeneral.Columns(0)
        Set Col1 = grdGeneral.Columns(1)
        Set Col2 = grdGeneral.Columns(2)
        grdGeneral.Columns(0).Width = 1600
        grdGeneral.Columns(1).Width = 3000
        grdGeneral.Columns(2).Width = 1000
        grdGeneral.Columns(2).NumberFormat = "##,###,###,###.00"
        Col0.Caption = "Student Number."
        Col1.Caption = "Student Names"
        Col2.Caption = "Score"
        cmdVoid.Enabled = False
    End If
    grdGeneral.Enabled = False
    ListRef
    evalSubjL.Caption = ""
PError:
End Sub

Private Sub Form_Load()
On Error GoTo PError
     Dim flag As Integer, mtable As String
     GenClass.fleLogin mvUserid, "Accessed Test Result Administration", Date, Time
     Call FormCentreMDI(Me)
     Set wrktemp = DBEngine.Workspaces(0)
     Set db1 = wrktemp.OpenDatabase(DBpath, True)
     Set rstSubj = db1.OpenRecordset("defsubj", dbOpenDynaset)
     Set rstClass = db1.OpenRecordset("defclass", dbOpenDynaset)
     Set rstEvals = db1.OpenRecordset("TabEvals", dbOpenDynaset)
     Set rstCNames = db1.OpenRecordset("defclassnames", dbOpenDynaset)
     Set rstAutoB = db1.OpenRecordset("autob", dbOpenDynaset)
     Set rstLMast = db1.OpenRecordset("lectmast", dbOpenDynaset)
     Set rstStudmast = db1.OpenRecordset("studmast", dbOpenDynaset)
     Set rstMCRoom = db1.OpenRecordset("MastClassRoom", dbOpenDynaset)
     Set rstMarkS = db1.OpenRecordset("TAbMarkScale", dbOpenDynaset)
     cTerm.Caption = GDescTerm
     Me.Caption = MVCoyname
     firstTime = False
     'create input table for each teacher
      tempTaB = Trim(mvUserid) + "Temp "
      db1.Execute "Drop Table " + tempTaB
Loadcont:
      mvSql = "Create Table " + Trim(mvUserid) + "Temp " _
         & "(studno Text, studnames Text, " _
         & "score long , Remarks text);"
      db1.Execute mvSql
     Set rstInpTab = db1.OpenRecordset(tempTaB, dbOpenDynaset)
     cmbSClass.Enabled = False


    ''_________filter evaluation records for user who is logging on
    If mvAcad = True Then
       mvSql = "Select * from qryEvalstest Where Staffno = '"
       mvSql = mvSql & mvUserid & "'"
    Else
       mvSql = "Select * from qryEvalstest "
    End If
    Set rstEvalstest = db1.OpenRecordset(mvSql, dbOpenDynaset)
    'end ____________________

    ''_________filter records for user who is logging on
    If mvAcad = True Then
       mvSql = "Select * from mastclassroom Where Staffno = '"
       mvSql = mvSql & mvUserid & "' and extsubj = false" & " order by subjcode"
    Else
       mvSql = "Select * from mastclassroom where extsubj = false order by subjcode"
    End If
    '____________________
    
      'populate combo remove
    Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
    rst.Requery
    If rst.RecordCount <> 0 Then
      rst.MoveFirst
    End If
    cmbsubJ.Clear
    Do While Not rst.EOF
         tmpStr1 = rst("SubjCode")
         cmbsubJ.AddItem rst("subjcode") + ", " + rst("subjdesc")
         rst.MoveNext
         If rst.EOF Then Exit Do
         tmpStr2 = rst("SubjCode")
         Do While Not rst.EOF And tmpStr1 = tmpStr2
            rst.MoveNext
            If rst.EOF Then Exit Do
            tmpStr2 = rst("SubjCode")
         Loop
    Loop
    If rst.RecordCount <> 0 Then
       cmbsubJ.ListIndex = 0
    End If
    rstStudmast.MoveFirst
    Do While Not rstStudmast.EOF
       cmbStudEval.AddItem rstStudmast!StudNo + ", " + rstStudmast!StudNames
       rstStudmast.MoveNext
    Loop
    If rstStudmast.RecordCount <> 0 Then
       cmbStudEval.ListIndex = 0
    End If
    cmbSClass.Enabled = False
    cmbsubJ.Enabled = False
    grdGeneral.Enabled = False
    cmbStud.Enabled = False
    cmdUpd.Enabled = False
    StudScore.Enabled = False
    sComment.Enabled = False
    cmdVoid.Enabled = False
    cmdView.Enabled = False
    VReason.Enabled = False
    cmdSave.Enabled = False
    evDate = Date
    ListRef
    Exit Sub
PError:
    If Err.Number = 3376 Then
        Resume Loadcont
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Error: " + Error$ + " Has Occurred", vbExclamation, Me.Caption
    End If
    'On Error Resume Next
    'Resume LoadClr
End Sub
Private Sub Form_Unload(Cancel As Integer)
    db1.Close
End Sub
Public Sub Brefresh()
On Error GoTo PError
      strpos = InStr(cmbSClass, ",")
      tmpStr1 = InStr(cmbSClass, "-")
      i = tmpStr1 - strpos
      mCRCode = Trim(Left(cmbSClass, strpos - 1))
      CName.Caption = Trim(Mid(cmbSClass, strpos + 2, i - 2))
      CGrp.Caption = Trim(Mid(cmbSClass, tmpStr1 + 2))
   If mvAcad = True Then
      mvSql = "Select * from studmast Where cclass = '" & mCRCode & "'"
   Else
      ' just a provision for more control
      mvSql = "Select * from studmast Where cclass = '" & mCRCode & "'"
   End If
    Set rstSelClass = db1.OpenRecordset(mvSql, dbOpenDynaset)
    rstSelClass.Requery
    If rstSelClass.RecordCount <> 0 Then
       rstSelClass.MoveFirst
       rstSelClass.Requery
       rstSelClass.MoveFirst
       db1.Execute "DELETE * FROM " & tempTaB
       StudCtr = 0
       Do While Not rstSelClass.EOF
          rstInpTab.AddNew
          rstInpTab("studno") = rstSelClass!StudNo
          rstInpTab("studnames") = rstSelClass!StudNames
          rstInpTab("score") = 0
          rstInpTab("remarks") = ""
          rstInpTab.Update
          MAStudno(StudCtr) = rstSelClass!StudNo
          StudCtr = StudCtr + 1
          rstSelClass.MoveNext
       Loop
       LCtr.Caption = StudCtr
    Else
       LCtr.Caption = "0"
       db1.Execute "DELETE * FROM " & tempTaB
    End If
    '---Display the grids
     mvSql = "Select Studno,studnames,score,remarks from " & tempTaB
     Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
     rst.Requery
     If (rst.BOF And rst.EOF) Then
         If firstTime = True Then
            MsgBox "No students in the selected class where the teacher should teach. Teacher will be unable to evaluate", vbExclamation, "Evaluation"
            firstTime = False
         End If
         Set datGeneral.Recordset = rst
         grdGeneral.Caption = Me.Caption
     Else
         Set datGeneral.Recordset = rst
         grdGeneral.Caption = Me.Caption & " - Class Test Input"
         Set Col0 = grdGeneral.Columns(0)
         Set Col1 = grdGeneral.Columns(1)
         Set Col2 = grdGeneral.Columns(2)
         Set Col3 = grdGeneral.Columns(3)
         grdGeneral.Columns(0).Width = 1600
         grdGeneral.Columns(1).Width = 3000
         grdGeneral.Columns(2).Width = 1000
         grdGeneral.Columns(2).NumberFormat = "##,###,###,###.00"
         grdGeneral.Columns(3).Width = 4000
         Col0.Caption = "Student Number."
         Col1.Caption = "Student Names"
         Col2.Caption = "Score"
         Col3.Caption = "Remarks"
         'FGrid.Columns(4).NumberFormat = "##,###,###,###.00"
         'FGrid.Columns(5).NumberFormat = "##,###,###,###.00"
   End If
   mvSql = "Select * from tabmarkscale where subjcode = '" & MsubJCode & "' and classroom = '" & mCRCode & "'"
   Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
   If GautoMs = True Then
      mvMarks = 100
      ssPM.Caption = Trim(Str(mvMarks))
      lblAms.Caption = "Unrestricted"
   Else
      If (rst.BOF And rst.EOF) Then
         mvMarks = 100
      Else
         mvMarks = rst!tesT
      End If
      ssPM.Caption = Trim(Str(mvMarks))
      lblAms.Caption = "Restricted"
   End If
PError:
End Sub


Private Sub grdGeneral_GotFocus()
  NExRef.Enabled = False
  evDate.Enabled = False
  cmdOk.Caption = "&Cancel"
  PanSSA.DragMode = False
  PanLR.Enabled = False
End Sub


Private Sub CmdSave_Click()
'On Error GoTo PError
Dim MarkPtr As Integer
    EFlag = False
    Fldeval
    If EFlag = True Then Exit Sub
    rstInpTab.Requery
    If Not (rstInpTab.BOF And rstInpTab.EOF) Then
       rstInpTab.MoveFirst
       MarkPtr = 0
       Do While Not rstInpTab.EOF
          Emark = rstInpTab!score
          Remarks = rstInpTab!Remarks
          Remarks = Left(Remarks, 70)
          rstEvals.AddNew
          rstEvals("Evtype") = "TEST"
          rstEvals("Evalref") = UCase(NExRef)
          rstEvals("studno") = MAStudno(MarkPtr)
          rstEvals("class") = UCase(mCRCode)
          rstEvals("subject") = UCase(MsubJCode)
          rstEvals("score") = Emark
          rstEvals("remarks") = Remarks
          rstEvals("subperc") = 0
          rstEvals("scoreval") = 0
          mvSql = "code = '" & MsubJCode & "'"
          rstSubj.FindFirst mvSql
          If rstSubj.NOMATCH Then
             MsgBox "Subject description not found", vbExclamation, "Evaluation Input"
             rstEvals("extint") = "XXXX"
          Else
             rstEvals("extint") = rstSubj!extint
          End If
          If Len(Trim(teaCher)) = 0 Then
             rstEvals("staffno") = UCase(mvUserid)
          Else
             rstEvals("staffno") = teaCher
          End If
          rstEvals("evdate") = evDate
          rstEvals("entdate") = Date
          rstEvals("schterm") = UCase(GCurrTerm)
          rstEvals.Update
          '__________add  review to student notes
          If Len(Trim((Remarks))) > 0 Then
             rstAutoB.AddNew
             rstAutoB("studno") = MAStudno(MarkPtr)
             rstAutoB("autodate") = Date
             rstAutoB("autodet") = Remarks
             rstAutoB("userid") = mvUserid
             rstAutoB.Update
          End If
          MarkPtr = MarkPtr + 1
          rstInpTab.MoveNext
       Loop
       cmdOk.Caption = "&Accept"
       cmdOk.Caption = "&Accept"
       NExRef.Enabled = True
       evDate.Enabled = True
       PanLR.Enabled = True
       PanSSA.Enabled = True
       cmbsubJ.Enabled = False
       cmbSClass.Enabled = False
       grdGeneral.Enabled = True
       cmdSave.Enabled = False
       teaCher = ""
       ListRef
    Else
       MsgBox "Empty class cannot be updated", vbExclamation, "Evaluation Input"
    End If
PError:
End Sub
Private Sub nEXrEF_gotfocus()
    teaCher = ""
    NExRef.SelStart = 0
    NExRef.SelLength = Len(NExRef)
    cmbsubJ.Enabled = False
    cmbSClass.Enabled = False
End Sub
Private Sub evdate_GotFocus()
    evDate.SelStart = 0
    evDate.SelLength = Len(evDate)
End Sub

Private Sub ListRef()
On Error GoTo PError
    If mvAcad = True Then
       mvSql = "Select * from qryEvalstest Where staffno = '" & mvUserid & "'"
    Else
       mvSql = "Select * from qryEvalstest "
    End If
    Set rstSelRef = db1.OpenRecordset(mvSql, dbOpenDynaset)
    rstSelRef.Requery
    If rstSelRef.RecordCount <> 0 Then
      rstSelRef.MoveFirst
    End If
    cmbRefL.Clear
    cmbrefS.Clear
    Do While Not rstSelRef.EOF
         tmpStr1 = rstSelRef("evalref")
         mvSql = "code = '" & rstSelRef("subject") & "'"
         rstSubj.Requery
         rstSubj.FindFirst mvSql
         If rstSubj.NOMATCH Then
            cmbRefL.AddItem rstSelRef("evalref") + ", " + rstSelRef("subject")
            cmbrefS.AddItem rstSelRef("evalref") + ", " + rstSelRef("subject")
         Else
            cmbRefL.AddItem rstSelRef("evalref") + ", " + rstSubj("desc")
            cmbrefS.AddItem rstSelRef("evalref") + ", " + rstSubj("desc")
         End If
         rstSelRef.MoveNext
         If rstSelRef.EOF Then Exit Sub
         tmpStr2 = rstSelRef("evalref")
         Do While Not rstSelRef.EOF And tmpStr1 = tmpStr2
            rstSelRef.MoveNext
            If rstSelRef.EOF Then Exit Sub
            tmpStr2 = rstSelRef("evalref")
         Loop
    Loop
PError:
End Sub
Private Sub Fldeval()
On Error GoTo PError
     rstInpTab.Requery
     If Not (rstInpTab.BOF And rstInpTab.EOF) Then rstInpTab.MoveFirst
     MarkPtr = 0
     Do While Not rstInpTab.EOF
        mStudno = MAStudno(MarkPtr)
        Emark = rstInpTab!score
        If Emark < 0 Or Emark > mvMarks Then
           MsgBox "Score for Student number " & mStudno & " is less than zero or more than the maximum mark that can be earned in this subject, Reenter", vbExclamation, "Evaluation Input"
           EFlag = True
           Exit Sub
        End If
        MarkPtr = MarkPtr + 1
        rstInpTab.MoveNext
     Loop
PError:
End Sub
Private Sub cmbStudEval_click()
   strpos = InStr(1, cmbStudEval, ",", 1)
   mStudno = Trim(Left(cmbStudEval, strpos - 1))
   cmdvResult.DoClick
End Sub

Private Sub cmbStudEval_LostFocus()
 On Error GoTo PError
   strpos = InStr(1, cmbStudEval, ",", 1)
   If strpos = 0 Then
       mStudno = cmbStudEval
   Else
       mStudno = Trim(Left(cmbStudEval, strpos - 1))
   End If
   mvSql = "studno = '" & mStudno & "'"
   rstStudmast.MoveFirst
   rstStudmast.FindFirst mvSql
   If rstStudmast.NOMATCH Then
      MsgBox "Student not a member of any class. ", vbCritical, "Student Performance List"
   Else
      cmbStudEval = rstStudmast("studno") + ", " + rstStudmast("studnames")
   End If
PError:
End Sub

Private Sub cmdVResult_Click()
On Error GoTo PError
     strpos = InStr(1, cmbStudEval, ",", 1)
     mStudno = Trim(Left(cmbStudEval, strpos - 1))
     tmpStr1 = Trim(Mid(cmbStudEval, strpos + 1))
     mvSql = "Select evdate,evalref,Studno,class,subject,score,staffno,remarks from  tabevals where studno = '" & mStudno & "' and Evtype = 'TEST'" & " order by subject"
     Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
     rst.Requery
     If (rst.BOF And rst.EOF) Then
         'MsgBox "No Results for Selected Student", vbExclamation, "Evaluation"
         Set datGeneral.Recordset = rst
         grdGeneral.Caption = Me.Caption
     Else
         Set datGeneral.Recordset = rst
         grdGeneral.Caption = "TEST RESULTS ON ALL SUBJECTS FOR " & tmpStr1
         Set Col0 = grdGeneral.Columns(0)
         Set Col1 = grdGeneral.Columns(1)
         Set Col2 = grdGeneral.Columns(2)
         Set Col3 = grdGeneral.Columns(3)
         Set Col4 = grdGeneral.Columns(4)
         Set Col5 = grdGeneral.Columns(5)
         Set Col6 = grdGeneral.Columns(6)
         Set Col7 = grdGeneral.Columns(7)
         grdGeneral.Columns(0).Width = 1200
         grdGeneral.Columns(1).Width = 1400
         grdGeneral.Columns(2).Width = 1200
         grdGeneral.Columns(3).Width = 600
         grdGeneral.Columns(4).Width = 800
         grdGeneral.Columns(5).Width = 800
         grdGeneral.Columns(6).Width = 1400
         grdGeneral.Columns(7).Width = 4000
         Col0.Caption = "Eval. Date"
         Col1.Caption = "Reference"
         Col2.Caption = "Student No."
         Col3.Caption = "Class"
         Col4.Caption = "Subject"
         Col5.Caption = "Score"
         Col6.Caption = "Teacher's No."
         Col7.Caption = "Remarks"
   End If
PError:
End Sub

Private Sub cmdCRooms_Click()
On Error GoTo PError
      mvSql = "Select * from  defclassnames order by code"
      Set rst = db1.OpenRecordset(mvSql, dbOpenDynaset)
      rst.Requery
      frmBrowse.Caption = MVCoyname
     If (rst.BOF And rst.EOF) Then
         MsgBox "No Classes defined", vbExclamation, "Evaluation"
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = Me.Caption
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "LIST OF ALL CLASSES ORDERED BY CLASS"
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         Set Col6 = frmBrowse.grdGeneral.Columns(6)
         Set Col7 = frmBrowse.grdGeneral.Columns(7)
         frmBrowse.grdGeneral.Columns(0).Width = 800
         frmBrowse.grdGeneral.Columns(1).Width = 1400
         frmBrowse.grdGeneral.Columns(2).Width = 1400
         frmBrowse.grdGeneral.Columns(3).Width = 1000
         frmBrowse.grdGeneral.Columns(4).Width = 800
         frmBrowse.grdGeneral.Columns(5).Width = 1200
         frmBrowse.grdGeneral.Columns(6).Width = 1200
         frmBrowse.grdGeneral.Columns(7).Width = 2600
         Col0.Caption = "Class Code"
         Col1.Caption = "Description"
         Col2.Caption = "Cklass Nickname"
         Col3.Caption = "Group"
         Col4.Caption = "Size"
         Col5.Caption = "Subject Count"
         Col6.Caption = "Captain No."
         Col7.Caption = "Class captain Name"
      End If
      frmBrowse.Show
PError:
End Sub
Private Sub teaCher_lostfocus()
 On Error GoTo PError
    mvSql = "StaffNo = '" & teaCher & "'"
    rstLMast.FindFirst mvSql
    If rstLMast.NOMATCH Then
       i = MsgBox("Teachers ID number not found, Do you want to correct it?", vbYesNo, "Staff Search")
       If i = vbYes Then
          teaCher.SetFocus
       Else
          teaCher = ""
       End If
    Else
       teaCher = UCase(teaCher)
    End If
PError:
End Sub


