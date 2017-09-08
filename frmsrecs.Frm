VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmsrecs 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll - Direct Recruitment"
   ClientHeight    =   8895
   ClientLeft      =   1305
   ClientTop       =   1080
   ClientWidth     =   12675
   DrawWidth       =   3
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8895
   ScaleWidth      =   12675
   Begin Threed.SSPanel SSPanel3 
      Height          =   8895
      Left            =   60
      TabIndex        =   21
      Top             =   60
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   15690
      _StockProps     =   15
      BackColor       =   11639171
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cmbState 
         Height          =   315
         Left            =   1680
         TabIndex        =   84
         Top             =   4356
         Width           =   3495
      End
      Begin VB.ComboBox mstatus 
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00004000&
         Height          =   315
         ItemData        =   "frmsrecs.frx":0000
         Left            =   1560
         List            =   "frmsrecs.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   7875
         Width           =   1335
      End
      Begin VB.ComboBox AppAvl 
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00004000&
         Height          =   315
         ItemData        =   "frmsrecs.frx":002C
         Left            =   1560
         List            =   "frmsrecs.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   7485
         Width           =   1335
      End
      Begin VB.TextBox PosAppl 
         Height          =   315
         Left            =   1680
         TabIndex        =   81
         Top             =   5910
         Width           =   3540
      End
      Begin VB.TextBox PTelCon 
         Height          =   315
         Left            =   1680
         TabIndex        =   80
         Top             =   5130
         Width           =   3495
      End
      Begin VB.TextBox Nation 
         Height          =   315
         Left            =   1680
         TabIndex        =   79
         Top             =   4743
         Width           =   3495
      End
      Begin VB.TextBox PBirth 
         Height          =   315
         Left            =   1680
         TabIndex        =   78
         Top             =   3969
         Width           =   1935
      End
      Begin VB.TextBox OthNames 
         Height          =   315
         Left            =   1680
         TabIndex        =   77
         Top             =   2808
         Width           =   1935
      End
      Begin VB.TextBox FName 
         Height          =   315
         Left            =   1680
         TabIndex        =   76
         Top             =   2421
         Width           =   1935
      End
      Begin VB.TextBox Surname 
         Height          =   315
         Left            =   1680
         TabIndex        =   75
         Top             =   2034
         Width           =   1935
      End
      Begin VB.TextBox Appid 
         Height          =   315
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   74
         Top             =   1260
         Width           =   1935
      End
      Begin VB.TextBox StaffNo 
         Height          =   315
         Left            =   1680
         TabIndex        =   72
         Top             =   1647
         Width           =   1935
      End
      Begin VB.TextBox Email 
         Height          =   315
         Left            =   1680
         TabIndex        =   71
         Top             =   5517
         Width           =   3495
      End
      Begin VB.Data datgenTab 
         Caption         =   "Gentab"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox HmAddr1 
         Height          =   1005
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   6360
         Width           =   4215
      End
      Begin Threed.SSFrame FrameSex 
         Height          =   1215
         Left            =   3840
         TabIndex        =   22
         Top             =   2880
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   2143
         _StockProps     =   14
         Caption         =   "Sex"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Begin Threed.SSOption Sex 
            Height          =   135
            Index           =   1
            Left            =   240
            TabIndex        =   2
            Top             =   660
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   238
            _StockProps     =   78
            Caption         =   "Female"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption Sex 
            Height          =   135
            Index           =   0
            Left            =   240
            TabIndex        =   1
            Top             =   360
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   238
            _StockProps     =   78
            Caption         =   "Male"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
      End
      Begin MSMask.MaskEdBox ApplDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   73
         Top             =   3585
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Birth 
         Height          =   315
         Left            =   1680
         TabIndex        =   85
         Top             =   3195
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin VB.Label Label27 
         BackColor       =   &H00B19983&
         Caption         =   "Email"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   65
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label staffName 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   360
         TabIndex        =   53
         Top             =   120
         Width           =   5535
      End
      Begin VB.Label Label26 
         BackColor       =   &H00B19983&
         Caption         =   "Date of 1st Appt:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   51
         Top             =   3675
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         Caption         =   "Staff Number:"
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
         Left            =   360
         TabIndex        =   38
         Top             =   1665
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00B19983&
         Caption         =   "Surname:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   2265
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00B19983&
         Caption         =   "First Name:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   2610
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00B19983&
         Caption         =   "Other Names:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   2970
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00B19983&
         Caption         =   "Rank:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00B19983&
         Caption         =   "Marital Status:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   7935
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00B19983&
         Caption         =   "Date of Birth:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   3315
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H00B19983&
         Caption         =   "Town:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   4020
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00B19983&
         Caption         =   "State of Origin:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   4380
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackColor       =   &H00B19983&
         Caption         =   "Nationality:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   4725
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackColor       =   &H00B19983&
         Caption         =   "Age:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   8280
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackColor       =   &H00B19983&
         Caption         =   "Home Address:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   6390
         Width           =   1215
      End
      Begin VB.Label Label15 
         BackColor       =   &H00B19983&
         Caption         =   "Contact Phone:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   5085
         Width           =   1215
      End
      Begin VB.Label Label19 
         BackColor       =   &H00B19983&
         Caption         =   "Availability:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   7575
         Width           =   1215
      End
      Begin VB.Label ApplId 
         BackColor       =   &H00808000&
         Caption         =   "File Number:"
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
         Left            =   360
         TabIndex        =   24
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label ApplAge 
         BackColor       =   &H00FFC0C0&
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
         Left            =   1560
         TabIndex        =   23
         Top             =   8280
         Width           =   615
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   5235
      Left            =   11040
      TabIndex        =   20
      Top             =   3360
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   9234
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
      BevelInner      =   2
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   300
         TabIndex        =   14
         Top             =   3660
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H0000FF00&
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
         Height          =   615
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdbott 
         Caption         =   "&Bottom"
         Height          =   375
         Left            =   300
         TabIndex        =   18
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   300
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   300
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdtop 
         Caption         =   "&Top"
         Height          =   375
         Left            =   300
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdbrow 
         Caption         =   "Bro&Wse"
         Height          =   375
         Left            =   300
         TabIndex        =   13
         Top             =   3300
         Width           =   975
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   375
         Left            =   300
         TabIndex        =   12
         Top             =   2940
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   300
         TabIndex        =   11
         Top             =   2580
         Width           =   975
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   300
         TabIndex        =   10
         Top             =   2220
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   300
         TabIndex        =   9
         Top             =   1860
         Width           =   975
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   8895
      Left            =   6240
      TabIndex        =   39
      Top             =   0
      Width           =   6375
      _Version        =   65536
      _ExtentX        =   11245
      _ExtentY        =   15690
      _StockProps     =   15
      BackColor       =   11639171
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel1 
         Height          =   3135
         Left            =   120
         TabIndex        =   69
         Top             =   120
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   5530
         _StockProps     =   15
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         BevelInner      =   2
         Begin VB.CommandButton cmdphoto 
            BackColor       =   &H0080FF80&
            Caption         =   "P&hoto"
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
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Image staffpic 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "StaffPic"
            DataSource      =   "StIdTab"
            Height          =   2295
            Left            =   120
            Picture         =   "frmsrecs.frx":0086
            Stretch         =   -1  'True
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.ComboBox cmbJob 
         Height          =   315
         Left            =   1920
         TabIndex        =   61
         Top             =   6744
         Width           =   2655
      End
      Begin VB.ComboBox cmbDept 
         Height          =   315
         Left            =   1920
         TabIndex        =   60
         Top             =   6390
         Width           =   2655
      End
      Begin VB.ComboBox cmbGrade 
         Height          =   315
         Left            =   1920
         TabIndex        =   59
         Top             =   6036
         Width           =   2655
      End
      Begin VB.ComboBox cmbTribe 
         Height          =   315
         Left            =   1920
         TabIndex        =   58
         Top             =   5682
         Width           =   2655
      End
      Begin VB.ComboBox cmbLoc 
         Height          =   315
         Left            =   1920
         TabIndex        =   57
         Top             =   5328
         Width           =   2655
      End
      Begin VB.ComboBox cmbProf 
         Height          =   315
         Left            =   1920
         TabIndex        =   56
         Top             =   7098
         Width           =   2655
      End
      Begin VB.ComboBox cmbLGA 
         Height          =   315
         Left            =   1920
         TabIndex        =   55
         Top             =   4974
         Width           =   2655
      End
      Begin VB.TextBox PayAcc 
         Height          =   375
         Left            =   1920
         TabIndex        =   54
         Top             =   8280
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog GetFile 
         Left            =   4080
         Top             =   7560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "Bmp"
         FilterIndex     =   2
      End
      Begin Threed.SSFrame SSstaff 
         Height          =   2235
         Left            =   2880
         TabIndex        =   52
         Top             =   960
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   3942
         _StockProps     =   14
         Caption         =   "Staff Status"
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
         Alignment       =   2
         Font3D          =   2
         Begin Threed.SSOption staffval 
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   5
            Top             =   960
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Probation"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption staffval 
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   6
            Top             =   1200
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Confirmed"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption staffval 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   4
            Top             =   720
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Contract"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption staffval 
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   3
            Top             =   480
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Temp. Staff"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption staffval 
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   8
            Top             =   1680
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Off Service"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption staffval 
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   7
            Top             =   1440
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Pensioned"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSMask.MaskEdBox PayBasic 
         Height          =   375
         Left            =   1920
         TabIndex        =   62
         Top             =   7452
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox CmADate 
         Height          =   375
         Left            =   1920
         TabIndex        =   63
         Top             =   4560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Freepay 
         Height          =   375
         Left            =   1920
         TabIndex        =   64
         Top             =   7866
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   1035
         Left            =   120
         TabIndex        =   66
         Top             =   3360
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   1826
         _StockProps     =   14
         Caption         =   "Appointment Class"
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
         Alignment       =   2
         Font3D          =   2
         Begin Threed.SSOption staffCls 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   67
            Top             =   600
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Non Academic Staff"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption staffCls 
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   68
            Top             =   360
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Academic Staff"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
      End
      Begin VB.Image Image2 
         Height          =   1335
         Left            =   5040
         Picture         =   "frmsrecs.frx":05C8
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label25 
         BackColor       =   &H00B19983&
         Caption         =   "Account Number:"
         Height          =   255
         Left            =   300
         TabIndex        =   50
         Top             =   8280
         Width           =   1575
      End
      Begin VB.Label Label24 
         BackColor       =   &H00B19983&
         Caption         =   "Free pay:"
         Height          =   255
         Left            =   300
         TabIndex        =   49
         Top             =   7920
         Width           =   1575
      End
      Begin VB.Label Label23 
         BackColor       =   &H00B19983&
         Caption         =   "Discipline Studied:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   300
         TabIndex        =   48
         Top             =   7095
         Width           =   1575
      End
      Begin VB.Label Label22 
         BackColor       =   &H00B19983&
         Caption         =   "Tribe:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   300
         TabIndex        =   47
         Top             =   5685
         Width           =   1575
      End
      Begin VB.Label Label21 
         BackColor       =   &H00B19983&
         Caption         =   "Location/Ministry:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   300
         TabIndex        =   46
         Top             =   5325
         Width           =   1575
      End
      Begin VB.Label Label20 
         BackColor       =   &H00B19983&
         Caption         =   "Local Government:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   300
         TabIndex        =   45
         Top             =   4980
         Width           =   1575
      End
      Begin VB.Label Label18 
         BackColor       =   &H00B19983&
         Caption         =   "Job Title:"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   300
         TabIndex        =   44
         Top             =   6750
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackColor       =   &H00B19983&
         Caption         =   "Department:"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   300
         TabIndex        =   43
         Top             =   6390
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00B19983&
         Caption         =   "Current Grade:"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   300
         TabIndex        =   42
         Top             =   6030
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00B19983&
         Caption         =   "Date of Present Appt:"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   300
         TabIndex        =   41
         Top             =   4620
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00B19983&
         Caption         =   "Yearly Basic:"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   7560
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmsrecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim db1 As Database, wrktemp As Workspace
Dim GenData As Recordset, PicName As String
Dim StTab As Recordset, stateBuff As String
Dim AddFlag As Boolean, AppStatBuff As String
Dim DateBuff As Date, SAppStat As String, mStaffCtr As Long
Dim EditFlag As Boolean, MaxErrors As Integer
Dim SSex As String, SDisab As String, rstJobtab As Recordset
Dim rstCarmanp As Recordset, FldChk As Boolean
Dim AppBuff As String, strpos As Integer, RegTab As Recordset
Dim rstGrade As Recordset, rstDeptTab As Recordset
Dim FirstTime As Integer, FirstPass As Integer
Dim GenClass As New QLog, rstPayTab As Recordset
Dim rstLGA As Recordset, rstLoc As Recordset, BType As String
Dim rstTribe As Recordset, rstProf As Recordset
Dim acadCls As Boolean
Private Sub AppAvl_gotfocus()
    AppAvl = "Immediate"
End Sub

Private Sub appid_Click()
    Appid.SelStart = 0
    Appid.SelLength = Len(Appid)
End Sub

Private Sub appldate_gotfocus()
    ApplDate.SelStart = 0
    ApplDate.SelLength = Len(ApplDate)
End Sub

Private Sub Birth_GotFocus()
    Birth.SelStart = 0
    Birth.SelLength = Len(Birth)
    Appid = UCase(Appid)
End Sub





Private Sub CmdAdd_Click()
On Error GoTo PError
If cmdAdd.Caption = "&Add" Then
   FldEnab
   PicName = Picpath & "picinit.bmp"
   staffpic = LoadPicture(PicName)
   stateBuff = "StateCode = '" & GState & "'"
   StTab.FindFirst stateBuff
   cmbState = StTab!StateCode & "," & StTab!StateName
   GenData.AddNew
   rstCarmanp.AddNew
   AddFlag = True
   Appid.SetFocus
   CmdDisab
   ClearData
   RegTab.Requery
   stateBuff = Trim(Str(RegTab("staffctr") + 1))
   If Len(stateBuff) < 5 Then
        stateBuff = "0" & stateBuff
   End If
   StaffNo = RegTab("coyid") & stateBuff
   cmdEdit.Enabled = False
   cmdSave.Enabled = True
   cmdAdd.Enabled = True
   cmdAdd.Caption = "&Cancel"
Else
    FldDisab
   AddFlag = False
   cmdAdd.Caption = "&Add"
   CmdEnab
   cmdEdit.Enabled = True
   cmdSave.Enabled = False
   'rstCarManP.MoveFirst
End If
PError:
End Sub

Private Sub CmdBott_Click()
On Error GoTo PError
    Dim Count As Long
    If rstCarmanp.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstCarmanp.MoveLast
    Showdata1
PError:
End Sub

Private Sub CmdEdit_Click()
On Error GoTo PError
If cmdEdit.Caption = "&Edit" Then
  FldEnab
  StaffNo.Enabled = False
  If rstCarmanp.EOF And rstCarmanp.BOF Then
      MsgBox ("Empty Table")
     Exit Sub
   End If
   cmdSave.Enabled = True
   AppBuff = "Appid = '" & Appid & "'"
   rstCarmanp.FindFirst AppBuff
   If rstCarmanp.NOMATCH Then
       MsgBox "Staff record cannot be edited", vbInformation, "Resume"
   Else
       GenData.Edit
       rstCarmanp.Edit
       EditFlag = True
   End If
   Appid.SetFocus
   cmdEdit.Caption = "&Cancel"
   Appid.Enabled = True
   CmdDisab
Else
   FldDisab
   cmdEdit.Caption = "&Edit"
   cmdSave.Enabled = False
   CmdEnab
End If
PError:
End Sub

Private Sub CmdDel_Click()
   Dim i As Integer
 On Error GoTo PError
   If GenData.RecordCount < 1 Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
    i = DeleteCheck()
    If i = vbYes Then
        If GenData("rind") = "X" Then
            AppBuff = "Appid = '" & GenData("appid") & "'"
            rstCarmanp.FindFirst AppBuff
            If rstCarmanp.NOMATCH Then
                MsgBox "Staff record not deleted", vbInformation, "Resume"
            Else
                GenData.Delete
                rstCarmanp.Delete
            End If
        Else
            MsgBox "Staff record has entries and cannnot be deleted", vbInformation, "Resume"
            Exit Sub
        End If
        Showdata1
        CmdEnab
    End If
PError:
End Sub

Private Sub cmdFind_Click()
  Dim sql As String, digitflag As Integer, length As Integer
   Dim buff As String, AN As String

   On Error GoTo FindError1

   AN = InputBox$("Enter the File Number to find:")
   length = Len(AN)
   If length = 0 Then
      ErrorMessages (NOINPUT)
      Exit Sub
   End If
   If length = 1 Then
      buff = "AppId LIKE " & Chr$(34) & AN & "*" & Chr$(34)
   Else
      buff = "AppId LIKE " & Chr$(34) & AN & "*" & Chr$(34)
   End If
   rstCarmanp.FindFirst buff
   If rstCarmanp.NOMATCH Then
      ErrorMessages (NOMATCH)
      Exit Sub
   Else
      Showdata1
   End If

FindError0:
   Exit Sub

FindError1:
   MsgBox "ERROR:   " + Error$ + " has occurred", vbCritical, Me.Caption
   On Error GoTo 0
   Resume FindError0

End Sub

Private Sub CmdNext_Click()
    Dim flag As Integer
    On Error GoTo NextError
    If rstCarmanp.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstCarmanp.EOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEPAST)
        rstCarmanp.MoveLast
    Else
        rstCarmanp.MoveNext
        If rstCarmanp.EOF Then
            rstCarmanp.MoveLast
        End If
    End If
NextClear:
    Showdata1
    Exit Sub
NextError:
    ErrorMessages (NOMOVEPAST)
    rstCarmanp.MoveLast
    On Error GoTo 0
    Resume NextClear
End Sub

Private Sub CmdPhoto_Click()
 On Error GoTo PError
   GetFile.Filter = "Picture Files (*.JPG)|*.jpg| (*.BMP)|*.bmp| All Files|*.*"
   GetFile.FilterIndex = 1
   GetFile.DefaultExt = "bmp"
   GetFile.ShowOpen
   PicName = GetFile.FileName
   ' Use the LoadPicture function to add the picture to the control
   staffpic = LoadPicture(GetFile.FileName)
PError:
End Sub

Private Sub CmdPrev_Click()
    Dim flag As Integer
    On Error GoTo PrevError
    If rstCarmanp.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstCarmanp.BOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEFRONT)
        rstCarmanp.MoveFirst
    Else
        rstCarmanp.MovePrevious
        If rstCarmanp.BOF Then
            rstCarmanp.MoveFirst
        End If
    End If
PrevClear:
    Showdata1
    Exit Sub
PrevError:
    ErrorMessages (NOMOVEFRONT)
    rstCarmanp.MoveFirst
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
    If AddFlag = True Then
        GenData.AddNew
        rstCarmanp.AddNew
        RegTab.Requery
        stateBuff = Trim(Str(RegTab("staffctr") + 1))
        StaffnoVal
        If Len(StaffNo) > 12 Then
           Beep
           MsgBox "The staff number " & StaffNo & " generated automatically exceeds 12 digits. Reduce the length of the company ID on the registration table to maximum 4 characters.", vbExclamation, "Recruitment"
           Exit Sub
        End If
        RegTab.Requery
        rstCarmanp("staffno") = RegTab("staffctr") + 1
        RegTab.Requery
        RegTab.Edit
        mStaffCtr = RegTab("staffctr") + 1
        RegTab("staffctr") = RegTab("staffctr") + 1
        GenNil
        TabInit
        GenData("rind") = "X"
    Else
        RegTab.Edit
    End If
    RegTab.Update
    GenData("AppId") = UCase(Appid)
    rstCarmanp("appid") = UCase(Appid)
    GenData("Surname") = UCase(Surname)
    GenData("FName") = UCase(FName)
    GenData("OthNames") = UCase(OthNames)
    DateBuff = Date - (16 * NoDYear)
    If Len(Appid) > 12 Then
        MsgBox "File Number must not be more than 12 characters long.", vbInformation, "Resume"
        Appid.SetFocus
        Exit Sub
    End If
    If Birth > DateBuff Then
        MsgBox ("Person must be at least 16 years old, Correct the Birthdate")
        Birth.SetFocus
        Exit Sub
    End If
    GenData("Birth") = Birth
    GenData("PBirth") = PBirth
    strpos = InStr(1, cmbState, ",", 1)
    GenData("State") = Left(cmbState, strpos - 1)
    strpos = InStr(1, cmbLGA, ",", 1)
    GenData("lgacode") = Left(cmbLGA, strpos - 1)
    strpos = InStr(1, cmbLoc, ",", 1)
    GenData("loccode") = Left(cmbLoc, strpos - 1)
    strpos = InStr(1, cmbTribe, ",", 1)
    GenData("tribecode") = Left(cmbTribe, strpos - 1)
    strpos = InStr(1, cmbProf, ",", 1)
    GenData("profcode") = Left(cmbProf, strpos - 1)
    GenData("Nation") = Nation
    GenData("email") = Email
    GenData("Ptelcon") = PTelCon
    GenData("PosAppl") = PosAppl
    GenData("HmAddr1") = HmAddr1
    GenData("ApplDate") = ApplDate
    GenData("Sex") = SSex
    GenData("Disab") = SDisab
    GenData("AppStat") = SAppStat
    GenData("MStatus") = mstatus
    GenData("Age") = DateDiff("yyyy", Birth, Date)
    GenData("yrserv") = DateDiff("yyyy", ApplDate, Date)
    GenData("AppAvl") = AppAvl
    GenData("Appstat") = "Employed"
    GenData("rind") = "A"
    GenData.Update
    Fileupd
    cmdAdd.Caption = "&Add"
    cmdEdit.Caption = "&Edit"
    cmdEdit.Enabled = True
    CmdEnab
    FldDisab
    AddFlag = False
    EditFlag = False
    'rstCarManP.MoveLast
    Showdata1
    cmdSave.Enabled = False
    
PError0:
    Exit Sub
    
PError:
    If Err.Number = 3022 Then
        MsgBox "File Number is already in use. ", vbCritical, "Resume"
        Appid.SetFocus
    Else
        MsgBox "Check the fields for invalid data.", vbCritical, "Resume"
    End If
    On Error GoTo 0
    Resume PError0

End Sub

Private Sub CmdTop_Click()
On Error GoTo PError
    If rstCarmanp.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstCarmanp.MoveFirst
    Showdata1
PError:
End Sub

Private Sub Form_Load()
On Error GoTo PError
     GenClass.fleLogin mvUserid, "Accessed Personal Data", Date, Time
     Dim AppStatBuff As String
     Call FormCentreMDI(Me)
     Set wrktemp = DBEngine.Workspaces(0)
     Set db1 = wrktemp.OpenDatabase(Dbpath, True)
     Set GenData = db1.OpenRecordset("Genrecs", dbOpenDynaset)
     Set StTab = db1.OpenRecordset("StateTab", dbOpenDynaset)
     Set db1 = wrktemp.OpenDatabase(Dbpath, True)
     Set rstGrade = db1.OpenRecordset("OrgGdTab", dbOpenDynaset)
     Set rstDeptTab = db1.OpenRecordset("DeptDTab", dbOpenDynaset)
     Set rstJobtab = db1.OpenRecordset("JobTDTab", dbOpenDynaset)
     Set rstCarmanp = db1.OpenRecordset("CarManP", dbOpenDynaset)
     Set RegTab = db1.OpenRecordset("Pimsreg", dbOpenDynaset)
     Set rstPayTab = db1.OpenRecordset("PayRoll", dbOpenDynaset)
     Set rstLGA = db1.OpenRecordset("LGA", dbOpenDynaset)
     Set rstLoc = db1.OpenRecordset("Location", dbOpenDynaset)
     Set rstTribe = db1.OpenRecordset("Tribe", dbOpenDynaset)
     Set rstProf = db1.OpenRecordset("ProfTab", dbOpenDynaset)
     Dim flag As Integer, mtable As String
     Me.Caption = mVCoyName
         If Not GenData.EOF Then
            GenData.MoveFirst
         End If
         If Not rstGrade.EOF Then
            rstGrade.MoveFirst
         End If
         If Not rstDeptTab.EOF Then
            rstDeptTab.MoveFirst
         End If
         If Not rstJobtab.EOF Then
            rstJobtab.MoveFirst
         End If
          If Not rstPayTab.EOF Then
            rstPayTab.MoveFirst
         End If
    ''__________________________________
    datGentab.DatabaseName = Dbpath
    datGentab.RecordSource = "JobTDTab"
    datGentab.ReadOnly = False
    datGentab.Exclusive = False
    datGentab.Refresh
    Do While Not datGentab.Recordset.EOF
       cmbJob.AddItem datGentab.Recordset("JobCode") + ", " + datGentab.Recordset("Jobdesc")
       datGentab.Recordset.MoveNext
    Loop
    If Not (rstJobtab.EOF And rstJobtab.BOF) Then
        cmbJob.ListIndex = 0
    End If
    ''__________________________________
    datGentab.DatabaseName = Dbpath
    datGentab.RecordSource = "DeptDTab"
    datGentab.ReadOnly = False
    datGentab.Exclusive = False
    datGentab.Refresh
    Do While Not datGentab.Recordset.EOF
       cmbDept.AddItem datGentab.Recordset("DeptCode") + ", " + datGentab.Recordset("Deptdesc")
       datGentab.Recordset.MoveNext
    Loop
    cmbDept.ListIndex = 0
    ''__________________________________
    datGentab.DatabaseName = Dbpath
    datGentab.RecordSource = "OrgGdTab"
    datGentab.ReadOnly = False
    datGentab.Exclusive = False
    datGentab.Refresh
    Do While Not datGentab.Recordset.EOF
       cmbGrade.AddItem datGentab.Recordset("GrdCode") + ", " + datGentab.Recordset("Grddesc")
       datGentab.Recordset.MoveNext
    Loop
     cmbGrade.ListIndex = 0
    
    cmdSave.Enabled = False
    ''__________________________________
    datGentab.DatabaseName = Dbpath
    datGentab.RecordSource = "StateTab"
    datGentab.ReadOnly = False
    datGentab.Exclusive = False
    datGentab.Refresh
    Do While Not datGentab.Recordset.EOF
       cmbState.AddItem datGentab.Recordset("Statecode") + ", " + datGentab.Recordset("StateName")
       datGentab.Recordset.MoveNext
    Loop
    cmbState.ListIndex = 0
    ''__________________________________
    datGentab.DatabaseName = Dbpath
    datGentab.RecordSource = "LGA"
    datGentab.ReadOnly = False
    datGentab.Exclusive = False
    datGentab.Refresh
    Do While Not datGentab.Recordset.EOF
       cmbLGA.AddItem datGentab.Recordset("code") + ", " + datGentab.Recordset("Desc")
       datGentab.Recordset.MoveNext
    Loop
    ''__________________________________
    datGentab.DatabaseName = Dbpath
    datGentab.RecordSource = "Location"
    datGentab.ReadOnly = False
    datGentab.Exclusive = False
    datGentab.Refresh
    Do While Not datGentab.Recordset.EOF
       cmbLoc.AddItem datGentab.Recordset("code") + ", " + datGentab.Recordset("Desc")
       datGentab.Recordset.MoveNext
    Loop
    ''__________________________________
    datGentab.DatabaseName = Dbpath
    datGentab.RecordSource = "Tribe"
    datGentab.ReadOnly = False
    datGentab.Exclusive = False
    datGentab.Refresh
    Do While Not datGentab.Recordset.EOF
       cmbTribe.AddItem datGentab.Recordset("code") + ", " + datGentab.Recordset("Desc")
       datGentab.Recordset.MoveNext
    Loop
    ''__________________________________
    datGentab.DatabaseName = Dbpath
    datGentab.RecordSource = "ProfTab"
    datGentab.ReadOnly = False
    datGentab.Exclusive = False
    datGentab.Refresh
    Do While Not datGentab.Recordset.EOF
       cmbProf.AddItem datGentab.Recordset("profcode") + ", " + datGentab.Recordset("profDesc")
       datGentab.Recordset.MoveNext
    Loop
    ''CmdAdd.SetFocus
    If Not StTab.EOF Then
       StTab.MoveFirst
    End If
    AddFlag = False
    EditFlag = False
    cmdSave.Enabled = False
    staffName = ""
    BType = ""
    If Not rstCarmanp.EOF Then
        rstCarmanp.Requery
        rstCarmanp.MoveFirst
        StaffNo = rstCarmanp("StaffNo")
        BType = rstCarmanp("Staffstat")
        Showdata1
    End If
    FldDisab
PError:
End Sub

Private Sub CmdBrow_Click()
On Error GoTo PError
    mvSql = "Select staffno,AppId, staffnames, cmadate, staffpic From carmanp"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    Set frmBrowse.datGeneral.Recordset = rst
    frmBrowse.Caption = Me.Caption & " - Records View"
    frmBrowse.Show
PError:
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub
Public Sub Showdata1()
    On Error GoTo PError
     If rstCarmanp.BOF And rstCarmanp.EOF Then
        ClearData
        Exit Sub
     End If
     AppBuff = "Appid = '" & rstCarmanp("AppId") & "'"
     rstCarmanp.FindFirst AppBuff
     If rstCarmanp.NOMATCH Then
          MsgBox "Staff number not found", vbInformation, "Resume"
     Else
        StaffNo = rstCarmanp("StaffNo")
        ApplDate = rstCarmanp("CMaDate")
        CmADate = rstCarmanp("CMLDprom")
        
        acadCls = rstCarmanp("acadcls")
        Select Case acadCls
        Case True
           staffCls(0).Value = True
        Case False
           staffCls(1).Value = True
        End Select
        BType = rstCarmanp("Staffstat")
        Select Case BType
        Case "Temp. Staff"
            staffval(0).Value = True
        Case "Contract"
            staffval(1).Value = True
        Case "Probation"
            staffval(2).Value = True
        Case "Confirmed"
            staffval(3).Value = True
        Case "Pensionable"
            staffval(4).Value = True
        Case "Off Service"
            staffval(5).Value = True
        End Select
        '________________________________________
        AppBuff = "GrdCode = '" & rstCarmanp("StaffGrd") & "'"
        rstGrade.FindFirst AppBuff
        cmbGrade = rstGrade!GrdCode & "," & rstGrade!GrdDesc
        AppBuff = "DeptCode = '" & rstCarmanp("StaffDept") & "'"
        rstDeptTab.FindFirst AppBuff
        cmbDept = rstDeptTab!DeptCode & "," & rstDeptTab!DeptDesc
        AppBuff = "JobCode = '" & rstCarmanp("StaffJobt") & "'"
        rstJobtab.FindFirst AppBuff
        cmbJob = rstJobtab!JobCode & "," & rstJobtab!JobDesc
        ''______________________________________________
        AppBuff = "StaffNo = '" & StaffNo & "'"
        rstPayTab.FindFirst AppBuff
           If rstPayTab.NOMATCH Then
               ClearData
               Exit Sub
               ''ErrorMessages (NOMATCH)
           Else
               PayBasic = rstPayTab("PayBasic")
               PayAcc = rstPayTab("PayAcc")
               Freepay = rstPayTab("Freepay")
           End If
           If rstCarmanp("staffpic") <> "" Then
               PicName = rstCarmanp("staffpic")
               staffName = rstCarmanp("staffnames")
           Else
               PicName = Picpath & "picinit.bmp"
               staffName = ""
           End If
           staffpic = LoadPicture(PicName)
           showdata2
       End If
PError0:
     
     Exit Sub
PError:
    'MsgBox "Default Picture Loaded", vbCritical, "Resume"
    PicName = Picpath & "picinit.bmp"
    staffpic = LoadPicture(PicName)
    showdata2
    On Error GoTo 0
    Resume PError0
End Sub

Public Sub showdata2()
On Error GoTo PError
  Dim AppStatBuff As String
  If rstCarmanp.EOF Then
    rstCarmanp.MoveLast
  End If
  AppBuff = "Appid = '" & rstCarmanp("AppId") & "'"
  GenData.FindFirst AppBuff
  If Not (GenData.BOF And GenData.EOF) Then
    Appid = GenData("AppId")
    Surname = GenData("Surname")
    FName = GenData("FName")
    OthNames = GenData("OthNames")
    Birth = GenData("Birth")
    PBirth = GenData("PBirth")
    stateBuff = "StateCode = '" & GenData("State") & "'"
    StTab.FindFirst stateBuff
    cmbState = StTab!StateCode & "," & StTab!StateName
    AppBuff = "Code = '" & GenData("lgacode") & "'"
    rstLGA.FindFirst AppBuff
    cmbLGA = rstLGA!code & "," & rstLGA!Desc
    AppBuff = "Code = '" & GenData("locCode") & "'"
    rstLoc.FindFirst AppBuff
    cmbLoc = rstLoc!code & "," & rstLoc!Desc
    AppBuff = "Code = '" & GenData("tribecode") & "'"
    rstTribe.FindFirst AppBuff
    cmbTribe = rstTribe!code & "," & rstTribe!Desc
    AppBuff = "profCode = '" & GenData("profcode") & "'"
    rstProf.FindFirst AppBuff
    cmbProf = rstProf!ProfCode & "," & rstProf!ProfDesc
     ''______________________________________________
    Nation = GenData("Nation")
    If IsNull(GenData("email")) Then
        Email = "None"
    Else
       Email = GenData("email")
    End If
    PTelCon = GenData("PTelCon")
    PosAppl = GenData("PosAppl")
    HmAddr1 = GenData("HmAddr1")
    SSex = GenData("Sex")
    SDisab = GenData("Disab")
    If SDisab = "" Then
        SDisab = "No"
    End If
    SAppStat = GenData("AppStat")
    mstatus = GenData("MStatus")
    ApplAge = GenData("Age")
    AppAvl = GenData("AppAvl")
    Select Case SSex
        Case "Male"
            Sex(0).Value = True
        Case "Female"
            Sex(1).Value = True
    End Select
    staffval(FirstTime).Value = True
  End If
PError:
End Sub

Public Sub ClearData()
On Error GoTo PError
        Appid = ""
        Surname = ""
        FName = ""
        OthNames = ""
        Birth = Date - (16 * NoDYear)
        PBirth = ""
        Nation = "Nigerian"
        PTelCon = ""
        PosAppl = ""
        HmAddr1 = ""
        ApplDate = Date
        SSex = "Male"
        SDisab = "No"
        mstatus = "Single"
        ApplAge = 0
        AppAvl = "1 Month"
        StaffNo = ""
        CmADate = Date
        Email = " "
        PayAcc = "Nil"
        PayBasic = 0
        Freepay = 0
        staffName = ""
PError:
End Sub
Private Sub Form_Unload(Cancel As Integer)
    db1.Close
End Sub


Private Sub Sex_Click(Index As Integer, Value As Integer)
On Error GoTo PError
    Select Case Index
        Case 0
            SSex = "Male"
        Case 1
            SSex = "Female"
    End Select
PError:
End Sub


Public Sub CmdEnab()
    cmdAdd.Enabled = True
    cmdDel.Enabled = True
    cmdBrow.Enabled = True
    cmdtop.Enabled = True
    cmdbott.Enabled = True
    cmdnext.Enabled = True
    cmdprev.Enabled = True
    cmdFind.Enabled = True
End Sub

Public Sub CmdDisab()
    cmdAdd.Enabled = False
    cmdDel.Enabled = False
    cmdBrow.Enabled = False
    cmdtop.Enabled = False
    cmdbott.Enabled = False
    cmdnext.Enabled = False
    cmdprev.Enabled = False
    cmdFind.Enabled = False

End Sub

Public Sub FldEnab()
On Error GoTo PError
        Appid.Enabled = True
        Surname.Enabled = True
        FName.Enabled = True
        OthNames.Enabled = True
        Birth.Enabled = True
        PBirth.Enabled = True
        Nation.Enabled = True
        cmbState.Enabled = True
        PTelCon.Enabled = True
        PosAppl.Enabled = True
        HmAddr1.Enabled = True
        ApplDate.Enabled = True
        FrameSex.Enabled = True
        mstatus.Enabled = True
        ApplAge.Enabled = True
        AppAvl.Enabled = True
        cmbGrade.Enabled = True
        cmbDept.Enabled = True
        cmbJob.Enabled = True
        PayBasic.Enabled = True
        Freepay.Enabled = True
        PayAcc.Enabled = True
        cmdphoto.Enabled = True
        CmADate.Enabled = True
        StaffNo.Enabled = False
        cmbLoc.Enabled = True
        Email.Enabled = True
        cmbTribe.Enabled = True
        cmbProf.Enabled = True
        cmbLGA.Enabled = True
        SSstaff.Enabled = True
PError:
End Sub

Public Sub FldDisab()
On Error GoTo PError
        Appid.Enabled = False
        Surname.Enabled = False
        FName.Enabled = False
        OthNames.Enabled = False
        Birth.Enabled = False
        PBirth.Enabled = False
        Nation.Enabled = False
        PTelCon.Enabled = False
        cmbState.Enabled = False
        PosAppl.Enabled = False
        HmAddr1.Enabled = False
        ApplDate.Enabled = False
        FrameSex.Enabled = False
        mstatus.Enabled = False
        ApplAge.Enabled = False
        AppAvl.Enabled = False
        Email.Enabled = False
        cmbGrade.Enabled = False
        cmbDept.Enabled = False
        cmbJob.Enabled = False
        PayBasic.Enabled = False
        Freepay.Enabled = False
        PayAcc.Enabled = False
        cmdphoto.Enabled = False
        CmADate.Enabled = False
        StaffNo.Enabled = False
        cmbLoc.Enabled = False
        cmbTribe.Enabled = False
        cmbProf.Enabled = False
        cmbLGA.Enabled = False
        SSstaff.Enabled = False
PError:
End Sub
Public Sub GenNil()
On Error GoTo PError
    GenData("lcomp") = "Nil"
    GenData("lcaddr") = "Nil"
    GenData("lcphone") = "Nil"
    GenData("lcjobtit") = "Nil"
    GenData("lcdesg") = "Nil"
    GenData("lcstdate") = Date
    GenData("lcbasic") = 0
    GenData("lchouse") = 0
    GenData("lctrans") = 0
    GenData("lcothers") = 0
    GenData("lclnout") = 0
    GenData("lcyr") = 0
    GenData("lc1name") = "Nil"
    GenData("lc1addr") = "Nil"
    GenData("lc1phone") = "Nil"
    GenData("lc1title") = "Nil"
    GenData("lc1desg") = "Nil"
    GenData("lc1stdate") = Date
    GenData("lc1enddate") = Date
    GenData("lc2name") = "Nil"
    GenData("lc2addr") = "Nil"
    GenData("lc2phone") = "Nil"
    GenData("lc2title") = "Nil"
    GenData("lc2desg") = "Nil"
    GenData("lc2stdate") = Date
    GenData("lc2enddate") = Date
    GenData("schdate") = Date
    GenData("schname") = "Nil"
    GenData("schqual") = "Nil"
    GenData("coldate") = Date
    GenData("colname") = "Nil"
    GenData("colqual") = "Nil"
    GenData("unvdate1") = Date
    GenData("unvname1") = "Nil"
    GenData("unvqual1") = "Nil"
    GenData("unvdate2") = Date
    GenData("unvname2") = "Nil"
    GenData("unvqual2") = "Nil"
    GenData("prfdate1") = Date
    GenData("prfdate2") = Date
    GenData("prfdate3") = Date
    GenData("prfdate4") = Date
    GenData("prfqual1") = "Nil"
    GenData("prfqual2") = "Nil"
    GenData("prfqual3") = "Nil"
    GenData("prfqual4") = "Nil"
    GenData("lang1") = "Nil"
    GenData("lang2") = "Nil"
    GenData("lang3") = "Nil"
    GenData("lang4") = "Nil"
    GenData("scholshp") = "Nil"
    GenData("exhbtns") = "Nil"
    GenData("prizes") = "Nil"
    GenData("projects") = "Nil"
    GenData("persint") = "Nil"
    GenData("socrec") = "Nil"
    GenData("genrem") = "Nil"
    GenData("referee1") = "Nil"
    GenData("refaddr1") = "Nil"
    GenData("reftitle1") = "Nil"
    GenData("referee2") = "Nil"
    GenData("refaddr2") = "Nil"
    GenData("reftitle2") = "Nil"
    GenData("referee3") = "Nil"
    GenData("refaddr3") = "Nil"
    GenData("reftitle3") = "Nil"
    GenData("nxkin") = "Nil"
    GenData("nxaddr1") = "Nil"
    GenData("nxreln") = "Nil"
    GenData("spname") = "Nil"
    GenData("spaddr1") = "Nil"
    GenData("faname") = "Nil"
    GenData("fastat") = "Nil"
    GenData("faoccup") = "Nil"
    GenData("faaddr1") = "Nil"
    GenData("maname") = "Nil"
    GenData("mastat") = "Nil"
    GenData("maoccup") = "Nil"
    GenData("maaddr1") = "Nil"
    GenData("child1") = "Nil"
    GenData("ch1birth") = Date
    GenData("ch1sex") = "Nil"
    GenData("child2") = "Nil"
    GenData("ch2birth") = Date
    GenData("ch2sex") = "Nil"
    GenData("child3") = "Nil"
    GenData("ch3birth") = Date
    GenData("ch3sex") = "Nil"
    GenData("child4") = "Nil"
    GenData("ch4birth") = Date
    GenData("ch4sex") = "Nil"
    GenData("chtot") = 0
    GenData("wifeno") = 0
    GenData("KPhone") = "Nil"
    GenData("KEmail") = "Nil"
    GenData("SPhone") = "Nil"
    GenData("SEmail") = "Nil"
    GenData("R1Phone") = "Nil"
    GenData("R2Phone") = "Nil"
    GenData("R3Phone") = "Nil"
PError:
End Sub
Public Sub TabInit()
 On Error GoTo PError
        Dim strpos As Integer, GrdBuff As String
        AppBuff = "StaffNo = '" & StaffNo & "'"
        rstPayTab.FindFirst AppBuff
        If rstPayTab.NOMATCH Then
           rstPayTab.AddNew
           rstPayTab("StaffNo") = StaffNo
           strpos = InStr(1, cmbGrade, ",", 1)
           rstPayTab("StaffGrd") = Left(cmbGrade, strpos - 1)
           rstPayTab("StDate") = Date
           rstPayTab("PayBasic") = PayBasic
           rstPayTab("PayAcc") = PayAcc
           rstPayTab("PayDedTot") = 0
           rstPayTab("corpint") = 0
           rstPayTab("PayGross") = 0
           rstPayTab("medLim") = 0
           rstPayTab("LvDaysEnt") = 0
           rstPayTab("LvTaken") = 0
           rstPayTab("Medlim") = 0
           rstPayTab("MedExp") = 0
           rstPayTab("TermChrg") = 0
           rstPayTab("Termchrg") = 0
           rstPayTab("TermGrat") = 0
           rstPayTab("TermGrat") = 0
           rstPayTab("TermDate") = Date
           rstPayTab("UpfPaid") = 0
           rstPayTab("UpfEarnd") = 0
           rstPayTab("TSubScript") = 0
           rstPayTab("TClubfee") = 0
           rstPayTab("LngBal") = 0
           rstPayTab("LngDed") = 0
           rstPayTab("rind") = "X"
           rstPayTab.Update
        End If
PError:
End Sub

Public Sub Fileupd()
 On Error GoTo PError
        Dim strpos As String
        Dim tLVDaysEnt As Single, tMedLim As Single
        rstCarmanp("StaffStat") = BType
        rstCarmanp("Comments") = "X"
        rstCarmanp("supervisor") = "X"
        rstCarmanp("rind") = "X"
        rstCarmanp("StaffNames") = UCase(FName) & " " & UCase(Surname)
        rstCarmanp("CMADate") = ApplDate
        rstCarmanp("CMLdProm") = CmADate
        rstCarmanp("yrserv") = DateDiff("yyyy", ApplDate, Date)
        strpos = InStr(1, cmbGrade, ",", 1)
        rstCarmanp("StaffGrd") = Left(cmbGrade, strpos - 1)
        AppBuff = "GrdCode = '" & Left(cmbGrade, strpos - 1) & "'"
        rstGrade.FindFirst AppBuff
        tMedLim = rstGrade("Medlim")
        tLVDaysEnt = rstGrade("LvdaysEnt")
        strpos = InStr(1, cmbJob, ",", 1)
        rstCarmanp("StaffJobT") = Left(cmbJob, strpos - 1)
        strpos = InStr(1, cmbDept, ",", 1)
        rstCarmanp("StaffDept") = Left(cmbDept, strpos - 1)
        rstCarmanp("StaffPic") = PicName
        rstCarmanp("acadcls") = acadCls
        rstCarmanp("StaffENat") = "Full"
        rstCarmanp("RFTotal") = 0
        AppBuff = "StaffNo = '" & StaffNo & "'"
        rstPayTab.FindFirst AppBuff
        If rstPayTab.NOMATCH Then
            TabInit
        Else
            rstPayTab.Edit
            rstPayTab("PayBasic") = PayBasic
            rstPayTab("Freepay") = Freepay
            rstPayTab("PayAcc") = PayAcc
            rstPayTab("Medlim") = tMedLim
            rstPayTab("LvDaysEnt") = tLVDaysEnt
            rstPayTab.Update
        End If
        rstCarmanp("MdRecomd") = "NoSalary"
        rstCarmanp.Update
PError:
End Sub
Public Sub FldVal()
  On Error GoTo PError
   If PayBasic < 0 Then
       Beep
       MsgBox "Yearly Basic Amount cannot be less than zero", vbExclamation, "Recruitment"
       PayBasic.SetFocus
       FldChk = False
       Exit Sub
    End If
    If EditFlag = False Then
        AppBuff = "Staffno = '" & StaffNo & "'"
        rstCarmanp.FindFirst AppBuff
        If rstCarmanp.NOMATCH Then
          DoEvents
        Else
          Beep
          MsgBox "Staff number is already allocated", vbExclamation, "Recruitment"
          StaffNo.SetFocus
          FldChk = False
          Exit Sub
        End If
    End If
    If Len(StaffNo) > 12 Then
       Beep
       MsgBox "Staff number must be 12 characters. Enter a valid Staff Number", vbExclamation, "Recruitment"
       PayBasic.SetFocus
       FldChk = False
       Exit Sub
    End If
    If Len(Trim(Surname)) = 0 Then
       Beep
       MsgBox "Surname cannot be blank. Enter a valid Surname", vbExclamation, "Recruitment"
       PayBasic.SetFocus
       FldChk = False
       Exit Sub
    End If
    If Len(Appid) > 10 Then
       Beep
       MsgBox "Application ID must be less than or equal 10 characters. Enter a valid Staff Number", vbExclamation, "Recruitment"
       PayBasic.SetFocus
       FldChk = False
       Exit Sub
    End If
    If (CDate(CmADate) < RegTab("RegDInc")) Or (CDate(CmADate) > Date) Then
       Beep
       MsgBox "Invalid Appointment Date", vbExclamation, "Recruitment"
       CmADate.SetFocus
       FldChk = False
       Exit Sub
    End If
PError:
End Sub

Private Sub staffCls_Click(Index As Integer, Value As Integer)
On Error GoTo PError
    Select Case Index
        Case 0
            staffCls(0).Value = True
            staffCls(1).Value = False
            acadCls = True
        Case 1
            staffCls(1).Value = True
            staffCls(0).Value = False
            acadCls = False
    End Select
PError:
End Sub

Private Sub StaffNo_Click()
    StaffNo.SelStart = 0
    StaffNo.SelLength = Len(StaffNo)
End Sub

Private Sub staffval_Click(Index As Integer, Value As Integer)
  On Error GoTo PError
   FirstTime = 9
    Select Case Index
        Case 0
            BType = "Temp. Staff"
            FirstTime = 0
        Case 1
            BType = "Contract"
            FirstTime = 1
        Case 2
            BType = "Probation"
            FirstTime = 2
        Case 3
            BType = "Confirmed"
            FirstTime = 3
        Case 4
            BType = "Pensionable"
            FirstTime = 4
       Case 5
            BType = "Off Service"
            FirstTime = 5
   End Select
   staffval(FirstTime).Value = True
PError:
End Sub

Private Sub staffno_lostfocus()
    StaffnoVal
End Sub
Public Sub StaffnoVal()
 On Error GoTo PError
    StaffNo.MaxLength = 12
    stateBuff = Trim(Str(RegTab("staffctr") + 1))
    If Len(stateBuff) = 7 Then
         stateBuff = "0" & stateBuff
    End If
    If Len(stateBuff) = 6 Then
         stateBuff = "00" & stateBuff
    End If
    If Len(stateBuff) = 5 Then
         stateBuff = "000" & stateBuff
    End If
    If Len(stateBuff) = 4 Then
         stateBuff = "0000" & stateBuff
    End If
    If Len(stateBuff) = 3 Then
         stateBuff = "00000" & stateBuff
    End If
    If Len(stateBuff) = 2 Then
         stateBuff = "000000" & stateBuff
    End If
    If Len(stateBuff) = 1 Then
         stateBuff = "0000000" & stateBuff
    End If
    If Len(stateBuff) = 0 Then
        MsgBox "Staff number automaticly generation is in valid - RegTab!staffctr", vbExclamation, "Recruitment"
        Exit Sub
    End If
    StaffNo = RegTab("CoyId") & stateBuff
PError:
End Sub
