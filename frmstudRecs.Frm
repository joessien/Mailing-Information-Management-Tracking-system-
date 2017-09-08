VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmStudRecs 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Tracking System"
   ClientHeight    =   8925
   ClientLeft      =   1305
   ClientTop       =   1080
   ClientWidth     =   15660
   DrawWidth       =   3
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8925
   ScaleWidth      =   15660
   Begin VB.PictureBox SSPanel5 
      BackColor       =   &H00B19983&
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   15555
      TabIndex        =   52
      Top             =   8040
      Width           =   15615
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H008080FF&
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
         Left            =   13920
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   1800
         TabIndex        =   56
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5400
         TabIndex        =   55
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   4320
         TabIndex        =   54
         Top             =   360
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
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.PictureBox SSPanel2 
      BackColor       =   &H00C0C0C0&
      Height          =   1875
      Left            =   13800
      ScaleHeight     =   1815
      ScaleWidth      =   1395
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
      Begin VB.CommandButton cmdbott 
         Caption         =   "&Bottom"
         Height          =   375
         Left            =   300
         TabIndex        =   6
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   300
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   300
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdtop 
         Caption         =   "&Top"
         Height          =   375
         Left            =   300
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.PictureBox SSPanel4 
      BackColor       =   &H00B19983&
      Height          =   7935
      Left            =   7800
      ScaleHeight     =   7875
      ScaleWidth      =   7755
      TabIndex        =   34
      Top             =   0
      Width           =   7815
      Begin VB.ComboBox cmbHouse 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   74
         Top             =   5040
         Width           =   3975
      End
      Begin VB.PictureBox cLCount 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         ScaleHeight     =   435
         ScaleWidth      =   795
         TabIndex        =   63
         Top             =   3480
         Width           =   855
      End
      Begin VB.ComboBox cmbState 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   49
         Top             =   5760
         Width           =   3975
      End
      Begin VB.ComboBox cmbSTerm 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   48
         Top             =   5400
         Width           =   3975
      End
      Begin VB.ComboBox cmbLGA 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   41
         Top             =   6120
         Width           =   3975
      End
      Begin VB.ComboBox cmbSkills 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   40
         Top             =   6480
         Width           =   3975
      End
      Begin VB.ComboBox cmbSports 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   39
         Top             =   6840
         Width           =   3975
      End
      Begin VB.ComboBox cmbTalent 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   38
         Top             =   7200
         Width           =   3975
      End
      Begin VB.ComboBox cmbClass 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   37
         Top             =   3600
         Width           =   3975
      End
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   3000
         ScaleHeight     =   3075
         ScaleWidth      =   2595
         TabIndex        =   35
         Top             =   240
         Width           =   2655
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
            TabIndex        =   36
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Image studPic 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "StaffPic"
            DataSource      =   "StIdTab"
            Height          =   2295
            Left            =   120
            Picture         =   "frmstudRecs.frx":0000
            Stretch         =   -1  'True
            Top             =   120
            Width           =   2415
         End
      End
      Begin MSComDlg.CommonDialog GetFile 
         Left            =   480
         Top             =   4440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "Bmp"
         FilterIndex     =   2
      End
      Begin VB.PictureBox SSPanel6 
         BackColor       =   &H00C0C0C0&
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
         Height          =   2475
         Left            =   6000
         ScaleHeight     =   2415
         ScaleWidth      =   1395
         TabIndex        =   58
         Top             =   5040
         Width           =   1455
         Begin VB.CommandButton cmdSubj 
            Caption         =   "Class &Subjects"
            Height          =   375
            Left            =   120
            TabIndex        =   62
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdCList 
            Caption         =   "&Class List"
            Height          =   375
            Left            =   120
            TabIndex        =   61
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CommandButton cmdStud 
            Caption         =   "&All Students"
            Height          =   375
            Left            =   120
            TabIndex        =   59
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find a Student"
            Height          =   375
            Left            =   120
            TabIndex        =   60
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.PictureBox mSize 
         BackColor       =   &H00FF80FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         ScaleHeight     =   435
         ScaleWidth      =   795
         TabIndex        =   64
         Top             =   4080
         Width           =   855
      End
      Begin VB.PictureBox SSPanel7 
         BackColor       =   &H00D8E9EC&
         Height          =   3135
         Left            =   240
         ScaleHeight     =   3075
         ScaleWidth      =   2595
         TabIndex        =   72
         Top             =   240
         Width           =   2655
         Begin VB.PictureBox cLogo 
            BackColor       =   &H00FF8080&
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
            Height          =   495
            Left            =   120
            ScaleHeight     =   435
            ScaleWidth      =   2355
            TabIndex        =   73
            Top             =   2520
            Width           =   2415
         End
         Begin VB.Image coyLogo 
            Height          =   2295
            Left            =   120
            Stretch         =   -1  'True
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.Label House 
         BackColor       =   &H00B19983&
         Caption         =   "House:"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   600
         TabIndex        =   75
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label CLRest 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Resticted"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   6000
         TabIndex        =   70
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label StaffNames 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1665
         TabIndex        =   66
         Top             =   4080
         Width           =   3990
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         BackColor       =   &H00B19983&
         Caption         =   "Class Supervisor:"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00B19983&
         Caption         =   "State of Origin:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   50
         Top             =   5880
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00B19983&
         Caption         =   "Start Term"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   600
         TabIndex        =   47
         Top             =   5520
         Width           =   855
      End
      Begin VB.Label Label16 
         BackColor       =   &H00B19983&
         Caption         =   "Talents:"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   600
         TabIndex        =   46
         Top             =   7320
         Width           =   975
      End
      Begin VB.Label Label17 
         BackColor       =   &H00B19983&
         Caption         =   "Class Admited:"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label20 
         BackColor       =   &H00B19983&
         Caption         =   "Local Govt:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   44
         Top             =   6240
         Width           =   855
      End
      Begin VB.Label Label21 
         BackColor       =   &H00B19983&
         Caption         =   "Skills"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   43
         Top             =   6600
         Width           =   735
      End
      Begin VB.Label Label22 
         BackColor       =   &H00B19983&
         Caption         =   "Sports:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   42
         Top             =   6960
         Width           =   975
      End
   End
   Begin VB.PictureBox SSPanel3 
      BackColor       =   &H00B19983&
      Height          =   7935
      Left            =   -120
      ScaleHeight     =   7875
      ScaleWidth      =   7755
      TabIndex        =   8
      Top             =   0
      Width           =   7815
      Begin VB.TextBox StudNo 
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
         Left            =   2160
         TabIndex        =   69
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox SNote 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   1680
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   68
         Top             =   6840
         Width           =   4455
      End
      Begin VB.Data DatTMP 
         Caption         =   "Tmptab"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4920
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.PictureBox FrameSex 
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Left            =   5400
         ScaleHeight     =   1275
         ScaleWidth      =   1755
         TabIndex        =   9
         Top             =   2280
         Width           =   1815
         Begin VB.PictureBox sSex 
            Height          =   375
            Index           =   1
            Left            =   360
            ScaleHeight     =   315
            ScaleWidth      =   795
            TabIndex        =   2
            Top             =   720
            Width           =   855
         End
         Begin VB.PictureBox Sex 
            Height          =   135
            Index           =   0
            Left            =   1000
            ScaleHeight     =   75
            ScaleWidth      =   945
            TabIndex        =   1
            Top             =   -840
            Width           =   1000
         End
         Begin VB.PictureBox sSex 
            Height          =   375
            Index           =   0
            Left            =   360
            ScaleHeight     =   315
            ScaleWidth      =   795
            TabIndex        =   51
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.TextBox PTelCon 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   32
         Top             =   5160
         Width           =   4455
      End
      Begin VB.TextBox Nation 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   31
         Top             =   4800
         Width           =   3375
      End
      Begin VB.TextBox RTown 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   30
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox OthNames 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   29
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox FName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   28
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox Surname 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   27
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox Fileno 
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
         Height          =   420
         Left            =   2160
         MaxLength       =   40
         TabIndex        =   26
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox Email 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   24
         Top             =   5520
         Width           =   4455
      End
      Begin VB.Data datgenTab 
         Caption         =   "Gentab"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox HmAddr1 
         Appearance      =   0  'Flat
         Height          =   885
         Left            =   1680
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   5880
         Width           =   4455
      End
      Begin MSMask.MaskEdBox RegDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   25
         Top             =   4080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox DOB 
         Height          =   315
         Left            =   1680
         TabIndex        =   33
         Top             =   3720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin VB.PictureBox cTerm 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1560
         ScaleHeight     =   675
         ScaleWidth      =   5475
         TabIndex        =   71
         Top             =   120
         Width           =   5535
      End
      Begin VB.Image Image2 
         Height          =   735
         Left            =   240
         Picture         =   "frmstudRecs.frx":0542
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   465
         Left            =   1680
         Picture         =   "frmstudRecs.frx":28304
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H00B19983&
         Caption         =   "Student Note:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   67
         Top             =   6840
         Width           =   1215
      End
      Begin VB.Label Label27 
         BackColor       =   &H00B19983&
         Caption         =   "Email"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label studNames 
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
         Left            =   1560
         TabIndex        =   22
         Top             =   960
         Width           =   5535
      End
      Begin VB.Label Label26 
         BackColor       =   &H00B19983&
         Caption         =   "Reg. Date:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00B19983&
         Caption         =   "Student Number:"
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
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00B19983&
         Caption         =   "Surname:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00B19983&
         Caption         =   "First Name:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00B19983&
         Caption         =   "Other Names:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00B19983&
         Caption         =   "Date of Birth:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H00B19983&
         Caption         =   "ResidenceTown:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackColor       =   &H00B19983&
         Caption         =   "Nationality:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackColor       =   &H00B19983&
         Caption         =   "Home Address:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label Label15 
         BackColor       =   &H00B19983&
         Caption         =   "Contact Phone:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label ApplId 
         BackColor       =   &H00B19983&
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
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label StudAge 
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
         Left            =   3720
         TabIndex        =   10
         Top             =   3720
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmStudRecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim db1 As Database, wrktemp As Workspace
Dim PicName As String, DateBuff, AgeBuff As String

Dim rstState, RegTab As Recordset
Dim rstClass, rstLGA, rstHouse As Recordset
Dim rstSkills, rstSports As Recordset
Dim rstTalent, rstTerms As Recordset
Dim rstStudmast, rstCRoom As Recordset
Dim rstSelClass, rstAutoB As Recordset

Dim AddFlag As Boolean, PBuff As Variant
Dim EditFlag As Boolean, MaxErrors As Integer
Dim mvSex As String, FldChk As Boolean
Dim strpos As Integer, tempBuff As String
Dim mstaffctr, mStudno, ClassCount As String
Dim FirstTime As Integer, FirstPass As Integer
Dim GenClass As New DProLog, buff As String
Dim BType As String

Private Sub cmbClass_Click()
  ClassDet
End Sub

Private Sub ClassDet()
    
    strpos = InStr(1, cmbClass, ",", 1)
    GTBuff = Trim(Left(cmbClass, strpos - 1))
    '_____________________________find class master
    rstCRoom.Requery
    mvSql = "Select * from qryclasssup Where rcode = '" & GTBuff & "'"
    Set rstSelClass = db1.OpenRecordset(mvSql, dbOpenDynaset)
    rstSelClass.Requery
    If rstSelClass.BOF And rstSelClass.EOF Then
       StaffNames.Caption = "No Supervisor assigned yet"
    Else
       StaffNames.Caption = rstSelClass!StaffNames
    End If
    '_________________________________find class current size
    
    mvSql = "Select * From studmast where Cclass = '" & GTBuff & "'"
    Set rstSelClass = db1.OpenRecordset(mvSql, dbOpenDynaset)
    rstSelClass.Requery
    Set DatTMP.Recordset = rstSelClass
    DatTMP.Refresh
    ClassCount = 0
    Do While Not DatTMP.Recordset.EOF
       ClassCount = ClassCount + 1
       DatTMP.Recordset.MoveNext
    Loop
    cLCount.Caption = Str(ClassCount)
    '_______________________________find class max size
    GTBuff = "Code = '" & GTBuff & "'"
    rstClass.FindFirst GTBuff
    If rstClass.NOMATCH Then
       MsgBox "Class is not defined", vbInformation, "Student Records"
   Else
       mSize.Caption = rstClass!CSize
   End If
   If AddFlag = True Or EditFlag = True Then
      If (Crest = True And Val(ClassCount) >= rstClass!CSize) Then
          MsgBox "Class is restricted and full. You will be unable to save the student record except you change the class.", vbExclamation, "Student Records"
      Else
         DoEvents
      End If
   End If
End Sub

Private Sub cmdCList_Click()
On Error GoTo PError
    cmbClass.Enabled = True
    strpos = InStr(1, cmbClass, ",", 1)
    GTBuff = Trim(Left(cmbClass, strpos - 1))
    mvSql = "Select Cclass,StudNo,stFileno, studnames, regdate, dob, sex From studmast where Cclass = '"
    mvSql = mvSql + GTBuff & "' order by Cclass;"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No students have been registered in selected Cclass", vbExclamation, "Student Records"
         Exit Sub
     Else
         Set frmBrwse.datGeneral.Recordset = rst
         frmBrwse.grdGeneral.Caption = "LIST OF ALL STUDENTS IN SELECTED CLASS"
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

Private Sub cmdSubj_Click()
On Error GoTo PError
    cmbClass.Enabled = True
    strpos = InStr(1, cmbClass, ",", 1)
    GTBuff = Trim(Left(cmbClass, strpos - 1))
    mvSql = "Select roomdesc,classid, subjdesc, staffnames From mastclassroom where roomid = '"
    mvSql = mvSql + GTBuff & "' order by subjdesc;"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No subjects have been assigned to selected class", vbExclamation, "Student Records"
         Exit Sub
     Else
         Set frmBrwse.datGeneral.Recordset = rst
         frmBrwse.grdGeneral.Caption = "LIST OF SUBJECTS FOR LECTURES IN SELECTED CLASS"
         frmBrwse.Caption = Me.Caption
         Set Col0 = frmBrwse.grdGeneral.Columns(0)
         Set Col1 = frmBrwse.grdGeneral.Columns(1)
         Set Col2 = frmBrwse.grdGeneral.Columns(2)
         Set Col3 = frmBrwse.grdGeneral.Columns(3)
         frmBrwse.grdGeneral.Columns(0).Width = 1200
         frmBrwse.grdGeneral.Columns(1).Width = 1200
         frmBrwse.grdGeneral.Columns(2).Width = 2800
         frmBrwse.grdGeneral.Columns(3).Width = 2800
         Col0.Caption = "Class Name"
         Col1.Caption = "Class Group"
         Col2.Caption = "Subject Description"
         Col3.Caption = "Subject Teacher"
   End If
   frmBrwse.Show
PError:
End Sub

Private Sub DOB_LostFocus()
   StudAge.Caption = Int((Val(Date - CDate(DOB))) / 365)
End Sub

Private Sub Email_gotfocus()
    Email.SelStart = 0
    Email.SelLength = Len(Email)
End Sub

Private Sub Fileno_gotfocus()
    Fileno.SelStart = 0
    Fileno.SelLength = Len(Fileno)
End Sub
Private Sub FName_gotfocus()
    FName.SelStart = 0
    FName.SelLength = Len(FName)
End Sub

Private Sub Nation_gotfocus()
    Nation.SelStart = 0
    Nation.SelLength = Len(Nation)
End Sub

Private Sub OthNames_gotfocus()
    OthNames.SelStart = 0
    OthNames.SelLength = Len(OthNames)
End Sub

Private Sub PTelCon_gotfocus()
    PTelCon.SelStart = 0
    PTelCon.SelLength = Len(PTelCon)
End Sub

Private Sub Regdate_gotfocus()
    regDate.SelStart = 0
    regDate.SelLength = Len(regDate)
End Sub

Private Sub dob_GotFocus()
    DOB.SelStart = 0
    DOB.SelLength = Len(DOB)
End Sub
Private Sub CmdAdd_Click()
On Error GoTo PError
If cmdadd.Caption = "&Add" Then
   FldEnab
   PicName = PicsPath & "picinit.bmp"
   StudPic = LoadPicture(PicName)
   tempBuff = "StateCode = '" & GlbState & "'"
   rstState.FindFirst tempBuff
   cmbState = rstState!StateCode & "," & rstState!StateName
   rstStudmast.AddNew
   AddFlag = True
   CmdDisab
   ClearData
   If GAutoCtr = False Then
      StudNo = ""
      StudNo.SetFocus
   Else
      Fileno.SetFocus
      RegTab.Requery
      GTBuff = Trim(Str(RegTab("staffctr") + 1))
      StudnoVal
      StudNo = PBuff
   End If
   cmdedit.Enabled = False
   cmdsave.Enabled = True
   cmdadd.Enabled = True
   cmdadd.Caption = "&Cancel"
   ClassDet
Else
   ClearData
   FldDisab
   AddFlag = False
   cmdadd.Caption = "&Add"
   CmdEnab
   cmdedit.Enabled = True
   cmdsave.Enabled = False
   'rstStudMast.MoveFirst
End If
PError:
End Sub

Private Sub CmdBott_Click()
On Error GoTo PError
    Dim Count As Long
    If rstStudmast.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstStudmast.MoveLast
    showdata
    ClassDet
PError:
End Sub

Private Sub CmdEdit_Click()
'On Error GoTo PError
If cmdedit.Caption = "&Edit" Then
  If rstStudmast.EOF And rstStudmast.BOF Then
      MsgBox ("No Records available for editing")
      FldDisab
      Exit Sub
   End If
   FldEnab
   StudNo.Enabled = False
   GTBuff = "studno = '" & StudNo & "'"
   mStudno = StudNo
    rstStudmast.FindFirst GTBuff
   If rstStudmast.NOMATCH Then
       MsgBox "Student record is selected", vbInformation, "Student Records"
   Else
       rstStudmast.Edit
       EditFlag = True
   End If
   Fileno.SetFocus
   cmdedit.Caption = "&Cancel"
   Fileno.Enabled = True
   CmdDisab
   cmdedit.Caption = "&Cancel"
   cmdadd.Enabled = False
   cmddel.Enabled = False
   cmdsave.Enabled = True
Else
   ClearData
   FldDisab
   EditFlag = False
   cmdedit.Caption = "&Edit"
   CmdEnab
   cmdedit.Enabled = True
   cmdsave.Enabled = False
   cmdadd.Enabled = True
   cmddel.Enabled = True
End If
PError:
End Sub

Private Sub CmdDel_Click()
   Dim i As Integer
 On Error GoTo PError
   If rstStudmast.RecordCount < 1 Then
      ErrorMessages (EMPTYTABLE)
      Exit Sub
   End If
    i = DeleteCheck()
    If i = vbYes Then
        If rstStudmast("rind") = "X" Then
            GTBuff = "Fileno = '" & rstStudmast("Fileno") & "'"
            rstStudmast.FindFirst GTBuff
            If rstStudmast.NOMATCH Then
                MsgBox "Student record not found", vbInformation, "Student Records"
            Else
                rstStudmast.Delete
            End If
        Else
            MsgBox "Student record has entries and cannnot be deleted", vbInformation, "Student Records"
            Exit Sub
        End If
        showdata
        CmdEnab
    End If
PError:
End Sub

Private Sub cmdFind_Click()
  Dim sql As String, digitflag As Integer, length As Integer
   Dim AN As String

   On Error GoTo FindError1

   AN = InputBox$("Enter the Student Number to find:")
   length = Len(AN)
   If length = 0 Then
      ErrorMessages (NOINPUT)
      Exit Sub
   End If
   If length = 1 Then
      buff = "studno LIKE " & Chr$(34) & AN & "*" & Chr$(34)
   Else
      buff = "studno LIKE " & Chr$(34) & AN & Chr$(34)
   End If
   rstStudmast.FindFirst buff
   If rstStudmast.NOMATCH Then
      ErrorMessages (NOMATCH)
      Exit Sub
   Else
      showdata
      ClassDet
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
    If rstStudmast.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstStudmast.EOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEPAST)
        rstStudmast.MoveLast
    Else
        rstStudmast.MoveNext
        If rstStudmast.EOF Then
            rstStudmast.MoveLast
        End If
    End If
NextClear:
    showdata
    ClassDet
    Exit Sub
NextError:
    ErrorMessages (NOMOVEPAST)
    rstStudmast.MoveLast
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
   StudPic = LoadPicture(GetFile.FileName)
PError:
End Sub

Private Sub CmdPrev_Click()
    Dim flag As Integer
    On Error GoTo PrevError
    If rstStudmast.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstStudmast.BOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEFRONT)
        rstStudmast.MoveFirst
    Else
        rstStudmast.MovePrevious
        If rstStudmast.BOF Then
            rstStudmast.MoveFirst
        End If
    End If
PrevClear:
    showdata
    ClassDet
    Exit Sub
PrevError:
    ErrorMessages (NOMOVEFRONT)
    rstStudmast.MoveFirst
    On Error GoTo 0
    Resume PrevClear
End Sub

Private Sub CmdSave_Click()
    Dim i As Integer
    'On Error GoTo PError
    FldChk = True
    FldVal
    If FldChk = False Then
        Exit Sub
    End If
    If AddFlag = True Then
        rstStudmast.AddNew
        If GAutoCtr = True Then
           RegTab.Requery
           GTBuff = Trim(Str(RegTab("staffctr") + 1))
           StudnoVal
           If Len(StudNo) > 12 Then
              Beep
              MsgBox "The Student number " & StudNo & " generated automatically exceeds 12 digits. Reduce the length of the company ID on the registration table to maximum 4 characters.", vbExclamation, "Student Records"
              Exit Sub
           End If
           rstStudmast("StudNo") = PBuff
           RegTab.Requery
           RegTab.Edit
              mstaffctr = RegTab("staffctr") + 1
           RegTab("staffctr") = RegTab("staffctr") + 1
           RegTab.Update
        Else
           If Len(StudNo) < 12 Then
              Beep
              MsgBox "Student number must not be less than 12 characters. Enter a valid Student Number", vbExclamation, "Student Records"
              StudNo.SetFocus
              Exit Sub
           End If
           rstStudmast("StudNo") = StudNo
        End If
        GenNil
    Else
        RegTab.Edit
    End If
    rstStudmast("stFileno") = RegTab("CoyId") & "/" & Trim(UCase(Fileno))
    rstStudmast("Surname") = UCase(Surname)
    rstStudmast("FName") = UCase(FName)
    rstStudmast("OthNames") = UCase(OthNames)
    rstStudmast("studNames") = UCase(Surname) & ", " & UCase(FName) & " " & UCase(OthNames)
    DateBuff = Date - (10 * 365)
    AgeBuff = Date - (20 * 365)
    If Len(Fileno) > 20 Then
        MsgBox "File Number must not be more than 20 characters long.", vbInformation, "Student Records"
        Fileno.SetFocus
        Exit Sub
    End If
    If CDate(DOB) > CDate(DateBuff) Then
        MsgBox ("Person must be at least 10 years old, Correct the Birthdate")
        DOB.SetFocus
        Exit Sub
    End If
    If CDate(DOB) < CDate(AgeBuff) Then
       i = MsgBox("Person may be too old to be admitted, Do you want to correct the age?", vbYesNo, "Student Records")
       If i = vbYes Then
          DOB.SetFocus
          Exit Sub
       Else
          DoEvents
       End If
    End If
    rstStudmast("dob") = DOB
    rstStudmast("rtown") = RTown
    
    strpos = InStr(1, cmbClass, ",", 1)
    rstStudmast("Cclass") = Trim(Left(cmbClass, strpos - 1))
    rstStudmast("Sclass") = Trim(Left(cmbClass, strpos - 1))
    
    strpos = InStr(1, cmbState, ",", 1)
    rstStudmast("State") = Trim(Left(cmbState, strpos - 1))
    
    strpos = InStr(1, cmbLGA, ",", 1)
    rstStudmast("lga") = Trim(Left(cmbLGA, strpos - 1))
    
    strpos = InStr(1, cmbSkills, ",", 1)
    rstStudmast("Skills") = Trim(Left(cmbSkills, strpos - 1))
    
    strpos = InStr(1, cmbSports, ",", 1)
    rstStudmast("sports") = Trim(Left(cmbSports, strpos - 1))
    
    strpos = InStr(1, cmbTalent, ",", 1)
    rstStudmast("talent") = Trim(Left(cmbTalent, strpos - 1))
    
    strpos = InStr(1, cmbSTerm, ",", 1)
    rstStudmast("sterm") = Trim(Left(cmbSTerm, strpos - 1))
    
    strpos = InStr(1, cmbHouse, ",", 1)
    rstStudmast("house") = Trim(Left(cmbHouse, strpos - 1))
    
    rstStudmast("Nation") = Nation
    rstStudmast("stemail") = Email
    rstStudmast("Ptelcon") = PTelCon
    rstStudmast("HmAddr1") = HMAddr1
    rstStudmast("Regdate") = regDate
    rstStudmast("gradyear") = Year(regDate) + 6
    rstStudmast("studpic") = PicName
    rstStudmast("Sex") = mvSex
    rstStudmast("Age") = DateDiff("yyyy", DOB, Date)
    rstStudmast.Update
    '__________add  review to student notes
    If Len(Trim((SNote))) > 0 Then
       rstAutoB.AddNew
       If AddFlag = True Then
          rstAutoB("studno") = Trim(PBuff)
       Else
          rstAutoB("studno") = Trim(mStudno)
       End If
       rstAutoB("autodate") = Date
       rstAutoB("autodet") = SNote
       rstAutoB("userid") = mvUserid
       rstAutoB.Update
    End If
    cmdadd.Caption = "&Add"
    cmdedit.Caption = "&Edit"
    cmdedit.Enabled = True
    CmdEnab
    FldDisab
    If AddFlag = True Then
       AddFlag = False
       rstStudmast.MoveLast
    Else
       EditFlag = False
    End If
    showdata
    cmdsave.Enabled = False
    SNote = ""
PError0:
    Exit Sub
    
PError:
    If Err.Number = 3022 Then
        MsgBox "File Number is already in use. ", vbCritical, "Student Records"
        Fileno.SetFocus
    Else
        MsgBox "Check the fields for invalid data.", vbCritical, "Student Records"
    End If
    On Error GoTo 0
    Resume PError0

End Sub

Private Sub CmdTop_Click()
On Error GoTo PError
    If rstStudmast.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    rstStudmast.MoveFirst
    showdata
    ClassDet
PError:
End Sub

Private Sub Form_Load()
'On Error GoTo PError
     GenClass.fleLogin mvUserid, "Accessed Student Records", Date, Time
     Dim AppStatBuff As String
     Call FormCentreMDI(Me)
     Set wrktemp = DBEngine.Workspaces(0)
     Set db1 = wrktemp.OpenDatabase(DBpath, True)
     
     Set rstState = db1.OpenRecordset("defstate", dbOpenDynaset)
     Set RegTab = db1.OpenRecordset("dproreg", dbOpenDynaset)
     Set rstClass = db1.OpenRecordset("defclassnames", dbOpenDynaset)
     Set rstLGA = db1.OpenRecordset("Deflga", dbOpenDynaset)
     Set rstSkills = db1.OpenRecordset("defskill", dbOpenDynaset)
     Set rstStudmast = db1.OpenRecordset("studmast", dbOpenDynaset)
     Set rstSports = db1.OpenRecordset("defsports", dbOpenDynaset)
     Set rstHouse = db1.OpenRecordset("defHouse", dbOpenDynaset)
     Set rstTalent = db1.OpenRecordset("deftalents", dbOpenDynaset)
     Set rstTerms = db1.OpenRecordset("defTerms", dbOpenDynaset)
     Set rstCRoom = db1.OpenRecordset("mastclassroom", dbOpenDynaset)
     Set rstAutoB = db1.OpenRecordset("autob", dbOpenDynaset)
     
     Dim flag As Integer, mtable As String
     Me.Caption = MVCoyname
     cLogo = RegTab!coyid
     cTerm.Caption = GDescTerm
    ''__________________________________
    datGenTab.DatabaseName = DBpath
    datGenTab.RecordSource = "defclassnames"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       If IsNull(datGenTab.Recordset("nickname")) Then
          cmbClass.AddItem datGenTab.Recordset("Code") + ", " + datGenTab.Recordset("desc") + " - " + datGenTab.Recordset("cgrp")
       Else
          cmbClass.AddItem datGenTab.Recordset("Code") + ", " + datGenTab.Recordset("desc") + " " + datGenTab.Recordset("nickname") + " - " + datGenTab.Recordset("cgrp")
      End If
      datGenTab.Recordset.MoveNext
    Loop
    If Not (datGenTab.Recordset.EOF And datGenTab.Recordset.BOF) Then
        cmbClass.ListIndex = 0
    End If
    ''__________________________________
    datGenTab.DatabaseName = DBpath
    datGenTab.RecordSource = "defstate"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       cmbState.AddItem datGenTab.Recordset("stateCode") + ", " + datGenTab.Recordset("statename")
       datGenTab.Recordset.MoveNext
    Loop
    If Not (datGenTab.Recordset.EOF And datGenTab.Recordset.BOF) Then
        cmbState.ListIndex = 0
    End If
    ''__________________________________
    datGenTab.DatabaseName = DBpath
    datGenTab.RecordSource = "deflga"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       cmbLGA.AddItem datGenTab.Recordset("Code") + ", " + datGenTab.Recordset("desc")
       datGenTab.Recordset.MoveNext
    Loop
    If Not (datGenTab.Recordset.EOF And datGenTab.Recordset.BOF) Then
        cmbLGA.ListIndex = 0
    End If
    
    cmdsave.Enabled = False
    ''__________________________________
    datGenTab.DatabaseName = DBpath
    datGenTab.RecordSource = "defskill"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       cmbSkills.AddItem datGenTab.Recordset("code") + ", " + datGenTab.Recordset("desc")
       datGenTab.Recordset.MoveNext
    Loop
    If Not (datGenTab.Recordset.EOF And datGenTab.Recordset.BOF) Then
        cmbSkills.ListIndex = 0
    End If
    ''__________________________________
    datGenTab.DatabaseName = DBpath
    datGenTab.RecordSource = "defsports"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       cmbSports.AddItem datGenTab.Recordset("code") + ", " + datGenTab.Recordset("Desc")
       datGenTab.Recordset.MoveNext
    Loop
    If Not (datGenTab.Recordset.EOF And datGenTab.Recordset.BOF) Then
        cmbSports.ListIndex = 0
    End If
    ''__________________________________
    datGenTab.DatabaseName = DBpath
    datGenTab.RecordSource = "deftalents"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       cmbTalent.AddItem datGenTab.Recordset("code") + ", " + datGenTab.Recordset("Desc")
       datGenTab.Recordset.MoveNext
    Loop
    If Not (datGenTab.Recordset.EOF And datGenTab.Recordset.BOF) Then
        cmbTalent.ListIndex = 0
    End If
    ''__________________________________
    datGenTab.DatabaseName = DBpath
    datGenTab.RecordSource = "defterms"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       cmbSTerm.AddItem datGenTab.Recordset("code") + ", " + datGenTab.Recordset("Desc")
       datGenTab.Recordset.MoveNext
    Loop
    If Not (datGenTab.Recordset.EOF And datGenTab.Recordset.BOF) Then
        cmbSTerm.ListIndex = 0
    End If
    ''__________________________________
    datGenTab.DatabaseName = DBpath
    datGenTab.RecordSource = "defhouse"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       cmbHouse.AddItem datGenTab.Recordset("code") + ", " + datGenTab.Recordset("Desc")
       datGenTab.Recordset.MoveNext
    Loop
    If Not (datGenTab.Recordset.EOF And datGenTab.Recordset.BOF) Then
        cmbHouse.ListIndex = 0
    End If
    ''__________________________________
    ''CmdAdd.SetFocus
    AddFlag = False
    EditFlag = False
    cmdsave.Enabled = False
    ClearData
    If Not rstStudmast.EOF Then
        rstStudmast.Requery
        rstStudmast.MoveFirst
        StudNo = rstStudmast("StudNo")
        showdata
        ClassDet
    End If
    FldDisab
    If Crest = True Then
       CLRest.Caption = "Restricted"
    Else
       CLRest.Caption = "Unrestricted"
    End If
    If IsNull(GLogo) Or Len(Trim(GLogo)) = 0 Then
        tmpStr1 = PicsPath & "rocks.bmp"
        coyLogo = LoadPicture(tmpStr1)
    Else
        coyLogo = LoadPicture(GLogo)
    End If
PError:
End Sub

Private Sub cmdStud_Click()
'On Error GoTo PError
    mvSql = "Select Cclass, StudNo,stFileno, studnames, regdate, dob, sex From studmast order by studno;"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No students records have been created ", vbExclamation, "Student Records"
         Exit Sub
     Else
         Set frmBrwse.datGeneral.Recordset = rst
         frmBrwse.grdGeneral.Caption = "LIST OF ALL STUDENTS ORDERED BY STUDENT NUMBER"
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

Private Sub cmdexit_Click()
    Unload Me
End Sub
Public Sub showdata()
  'On Error GoTo PError
  If rstStudmast.RecordCount = 1 Then
     rstStudmast.MoveFirst
  End If
  If Not (rstStudmast.BOF And rstStudmast.EOF) Then
        StudNo = rstStudmast("StudNo")
        strpos = Len(RegTab("CoyId"))
        Fileno = Trim(Mid(rstStudmast("stFileno"), strpos + 2))
        regDate = rstStudmast("regdate")
        If IsNull(rstStudmast("stemail")) Then
           Email = "None"
        Else
           Email = rstStudmast("stemail")
        End If
        Select Case mvSex
           Case "Male"
               sSex(0).Value = True
           Case "Female"
               sSex(1).Value = True
        End Select
        Surname = rstStudmast("Surname")
        FName = rstStudmast("FName")
        OthNames = rstStudmast("OthNames")
        StudNames = rstStudmast("studNames")
        DOB = rstStudmast("dob")
        RTown = rstStudmast("rtown")
        Nation = rstStudmast("Nation")
        Email = rstStudmast("stemail")
        PTelCon = rstStudmast("Ptelcon")
        HMAddr1 = rstStudmast("HmAddr1")
        StudAge = DateDiff("yyyy", rstStudmast("dob"), Date)
        '________________________________________
        GTBuff = "Code = '" & rstStudmast("Cclass") & "'"
        rstClass.FindFirst GTBuff
        If IsNull(rstClass!NickName) Then
           cmbClass = rstClass!code & ", " & rstClass!Desc + " - " + rstClass!CGrp
        Else
           cmbClass = rstClass!code & ", " & rstClass!Desc + " " & rstClass!NickName + " - " + rstClass!CGrp
        End If
        GTBuff = "stateCode = '" & rstStudmast("state") & "'"
        rstState.FindFirst GTBuff
        cmbState = rstState!StateCode & "," & rstState!StateName
        
        GTBuff = "Code = '" & rstStudmast("lga") & "'"
        rstLGA.FindFirst GTBuff
        cmbLGA = rstLGA!code & "," & rstLGA!Desc
        
        GTBuff = "Code = '" & rstStudmast("skills") & "'"
        rstSkills.FindFirst GTBuff
        cmbSkills = rstSkills!code & "," & rstSkills!Desc
        
        GTBuff = "Code = '" & rstStudmast("sports") & "'"
        rstSports.FindFirst GTBuff
        cmbSports = rstSports!code & "," & rstSports!Desc
        
        GTBuff = "Code = '" & rstStudmast("talent") & "'"
        rstTalent.FindFirst GTBuff
        cmbTalent = rstTalent!code & "," & rstTalent!Desc
        
        GTBuff = "Code = '" & rstStudmast("sterm") & "'"
        rstTerms.FindFirst GTBuff
        cmbSTerm = rstTerms!code & "," & rstTerms!Desc
        
        If IsNull(rstStudmast("house")) Then
           If Not (rstHouse.EOF And rstHouse.BOF) Then
              cmbHouse.ListIndex = 0
           End If
        Else
           GTBuff = "Code = '" & rstStudmast("house") & "'"
           rstHouse.FindFirst GTBuff
           cmbHouse = rstHouse!code & "," & rstHouse!Desc
        End If
        
        If rstStudmast("studpic") <> "" Then
            PicName = rstStudmast("studpic")
            StudNames = rstStudmast("studnames")
        Else
            PicName = PicsPath & "picinit.bmp"
            StaffNames = ""
        End If
        StudPic = LoadPicture(PicName)
   Else
        MsgBox "Student Records not found", vbCritical, "Student Records"
   End If
PError0:
     Exit Sub
PError:
    'MsgBox "Default Picture Loaded", vbCritical, "Student Records"
    PicName = PicsPath & "picinit.bmp"
    StudPic = LoadPicture(PicName)
    On Error GoTo 0
    Resume PError0
End Sub

Public Sub ClearData()
On Error GoTo PError
        Fileno = ""
        Surname = ""
        FName = ""
        OthNames = ""
        DOB = Date - (10 * 365)
        'RTown = ""
        Nation = "Nigerian"
        PTelCon = "000 0000 0000"
        HMAddr1 = ""
        SNote = ""
        regDate = Date
        If mvSex = "Male" Then
           sSex(0).Value = True
        Else
           sSex(1).Value = True
        End If
        StudAge.Caption = Int((Val(Date - CDate(DOB))) / 365)
        StudNo = ""
        regDate = Date
        Email = "None "
        StaffNames = ""
        StudNames = ""
        cLCount.Caption = ""
        mSize.Caption = ""
        SNote = ""
PError:
End Sub
Private Sub Form_Unload(Cancel As Integer)
    db1.Close
End Sub

Private Sub RTown_gotfocus()
    RTown.SelStart = 0
    RTown.SelLength = Len(RTown)
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

Public Sub FldEnab()
'On Error GoTo PError
        If GAutoCtr = False Then
           StudNo.Enabled = True
        End If
        Fileno.Enabled = True
        StudNo.Enabled = True
        Surname.Enabled = True
        FName.Enabled = True
        OthNames.Enabled = True
        DOB.Enabled = True
        RTown.Enabled = True
        Nation.Enabled = True
        PTelCon.Enabled = True
        HMAddr1.Enabled = True
        SNote.Enabled = True
        regDate.Enabled = True
        FrameSex.Enabled = True
        StudAge.Enabled = True
        cmbClass.Enabled = True
        cmbState.Enabled = True
        cmbLGA.Enabled = True
        cmbSkills.Enabled = True
        cmbSports.Enabled = True
        cmbTalent.Enabled = True
        cmbSTerm.Enabled = True
        cmbHouse.Enabled = True
        cmdphoto.Enabled = True
        StudAge.Enabled = True
        Email.Enabled = True
PError:
End Sub

Public Sub FldDisab()
On Error GoTo PError
        Fileno.Enabled = False
        StudNo.Enabled = False
        Surname.Enabled = False
        FName.Enabled = False
        OthNames.Enabled = False
        DOB.Enabled = False
        RTown.Enabled = False
        Nation.Enabled = False
        PTelCon.Enabled = False
        HMAddr1.Enabled = False
        SNote.Enabled = False
        regDate.Enabled = False
        FrameSex.Enabled = False
        StudAge.Enabled = False
        cmbClass.Enabled = False
        cmbState.Enabled = False
        cmbLGA.Enabled = False
        cmbSkills.Enabled = False
        cmbSports.Enabled = False
        cmbTalent.Enabled = False
        cmbSTerm.Enabled = False
        cmbHouse.Enabled = False
        cmdphoto.Enabled = False
        StudAge.Enabled = False
        Email.Enabled = False
PError:
End Sub
Public Sub GenNil()
On Error GoTo PError
    rstStudmast("guardian") = "Nil"
    rstStudmast("gdaddr1") = "Nil"
    rstStudmast("gdreln") = "Nil"
    rstStudmast("gdphone") = "Nil"
    rstStudmast("gdemail") = "Nil"
    rstStudmast("faname") = "Nil"
    rstStudmast("fastat") = True
    rstStudmast("faoccup") = "Nil"
    rstStudmast("faaddr1") = "Nil"
    rstStudmast("faphone") = "Nil"
    rstStudmast("maname") = "Nil"
    rstStudmast("mastat") = True
    rstStudmast("maoccup") = "Nil"
    rstStudmast("maaddr1") = "Nil"
    rstStudmast("maphone") = "Nil"
PError:
End Sub
Public Sub FldVal()
  On Error GoTo PError
    If EditFlag = False Then
        GTBuff = "StudNo = '" & StudNo & "'"
        rstStudmast.FindFirst GTBuff
        If rstStudmast.NOMATCH Then
          DoEvents
        Else
          Beep
          MsgBox "Student number is already allocated", vbExclamation, "Student Records"
          StudNo.SetFocus
          FldChk = False
          Exit Sub
        End If
    End If
    If Len(StudNo) > 12 Then
       Beep
       MsgBox "Student number must not be more than 12 characters. Enter a valid Student Number", vbExclamation, "Student Records"
       StudNo.SetFocus
       FldChk = False
       Exit Sub
    End If
    If Len(Trim(Surname)) = 0 Then
       Beep
       MsgBox "Surname cannot be blank. Enter a valid Surname", vbExclamation, "Student Records"
       Fileno.SetFocus
       FldChk = False
       Exit Sub
    End If
    If Len(Fileno) > 20 Then
       Beep
       MsgBox "Student File Number must be less than or equal 20 characters. Enter a valid Student Number", vbExclamation, "Student Records"
       Fileno.SetFocus
       FldChk = False
       Exit Sub
    End If
    If (CDate(regDate) < CDate(RegTab("RegDInc"))) Or ((CDate(regDate) > Date)) Then
       Beep
       MsgBox "Invalid Registration Date; Either Student Registration date is before school was incorporated or is greater than todays date", vbExclamation, "Student Records"
       regDate.SetFocus
       FldChk = False
       Exit Sub
    End If
    If (Crest = True And Val(ClassCount) >= rstClass!CSize) Then
       MsgBox "Class size is restricted and is full. Change the Admission Class and Retry ", vbExclamation, "Student Records"
       cmbClass.SetFocus
       FldChk = False
       Exit Sub
    End If
PError:
End Sub

Private Sub StudNo_gotfocus()
    StudNo.SelStart = 0
    StudNo.SelLength = Len(StudNo)
End Sub

Private Sub StudNo_lostfocus()
    If GAutoCtr = True Then
       StudnoVal
    End If
End Sub
Public Sub StudnoVal()
 On Error GoTo PError
    StudNo.MaxLength = 12
    PBuff = Trim(Str(RegTab("staffctr") + 1))
    If Len(PBuff) = 8 Then
         PBuff = RegTab("CoyId") & PBuff
    End If
    If Len(PBuff) = 7 Then
         PBuff = RegTab("CoyId") & "0" & PBuff
    End If
    If Len(PBuff) = 6 Then
         PBuff = RegTab("CoyId") & "00" & PBuff
    End If
    If Len(PBuff) = 5 Then
         PBuff = RegTab("CoyId") & "000" & PBuff
    End If
    If Len(PBuff) = 4 Then
         PBuff = RegTab("CoyId") & "0000" & PBuff
    End If
    If Len(PBuff) = 3 Then
         PBuff = RegTab("CoyId") & "00000" & PBuff
    End If
    If Len(PBuff) = 2 Then
         PBuff = RegTab("CoyId") & "000000" & PBuff
    End If
    If Len(PBuff) = 1 Then
         PBuff = RegTab("CoyId") & "0000000" & PBuff
    End If
    If Len(PBuff) = 0 Then
        MsgBox "Student number automaticly generation is in valid - RegTab!staffctr", vbExclamation, "Student Records"
        Exit Sub
    End If
    'StudNo = RegTab("CoyId") & PBuff
    StudNo = PBuff
    StudNo.Enabled = False
PError:
End Sub

Private Sub Surname_gotfocus()
    Surname.SelStart = 0
    Surname.SelLength = Len(Surname)
End Sub

Private Sub snote_gotfocus()
    SNote.SelStart = 0
    SNote.SelLength = Len(SNote)
End Sub
