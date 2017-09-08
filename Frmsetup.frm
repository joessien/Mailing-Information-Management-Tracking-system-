VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Tracking System"
   ClientHeight    =   7440
   ClientLeft      =   1110
   ClientTop       =   900
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   10665
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H00C0C0C0&
      Height          =   7575
      Index           =   0
      Left            =   0
      ScaleHeight     =   7515
      ScaleWidth      =   10635
      TabIndex        =   4
      Top             =   -120
      Width           =   10695
      Begin MSMask.MaskEdBox NodYear 
         Height          =   375
         Left            =   2640
         TabIndex        =   57
         Top             =   5760
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdReg 
         Caption         =   "Register"
         Height          =   495
         Left            =   6720
         TabIndex        =   55
         Top             =   6720
         Width           =   1935
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   495
         Left            =   9120
         TabIndex        =   54
         Top             =   6720
         Width           =   735
      End
      Begin VB.CommandButton cmdLog 
         Caption         =   "&Add Logo"
         Height          =   375
         Left            =   9120
         TabIndex        =   52
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Okay"
         Height          =   375
         Left            =   6720
         TabIndex        =   51
         Top             =   1350
         Width           =   615
      End
      Begin MSMask.MaskEdBox cVal 
         Height          =   375
         Left            =   3000
         TabIndex        =   50
         Top             =   6240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   3
         Format          =   "0"
         PromptChar      =   "_"
      End
      Begin VB.TextBox sControl 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7440
         TabIndex        =   49
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CheckBox AutoCtr 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   3840
         TabIndex        =   48
         Top             =   3780
         Width           =   255
      End
      Begin VB.ComboBox cmbBUnitD 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Height          =   315
         Left            =   6720
         TabIndex        =   47
         Top             =   4800
         Width           =   3135
      End
      Begin VB.CommandButton cmdRpt 
         BackColor       =   &H0000FF00&
         Caption         =   "Select Reports Path"
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
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Select the Database of the Source Table"
         Top             =   5520
         Width           =   2295
      End
      Begin VB.CommandButton cmdPics 
         BackColor       =   &H0000FF00&
         Caption         =   "Select Pictures Path"
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
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Select the Database of the Source Table"
         Top             =   5880
         Width           =   2295
      End
      Begin VB.CommandButton CmdDocs 
         BackColor       =   &H0000FF00&
         Caption         =   "Select Documents Path"
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
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Select the Database of the Source Table"
         Top             =   6240
         Width           =   2295
      End
      Begin VB.CommandButton cmdDB 
         BackColor       =   &H0000FF00&
         Caption         =   "Select Database Path"
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
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Select the Database of the Source Table"
         Top             =   5160
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
         Left            =   960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1080
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSMask.MaskEdBox Bnum 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Left            =   8280
         TabIndex        =   33
         Top             =   2160
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mvCoyId 
         Height          =   255
         Left            =   8280
         TabIndex        =   27
         Top             =   3840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   4
         PromptChar      =   "_"
      End
      Begin VB.TextBox YrEnd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   26
         Top             =   3480
         Width           =   1095
      End
      Begin MSMask.MaskEdBox reG 
         Height          =   285
         Left            =   3360
         TabIndex        =   24
         Top             =   3120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   40
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox PassExp 
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   5400
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox PassGrs 
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   5040
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   2
         PromptChar      =   "_"
      End
      Begin VB.PictureBox SSRibbon1 
         BackColor       =   &H00C8D0D4&
         Height          =   30
         Left            =   240
         ScaleHeight     =   30
         ScaleWidth      =   11175
         TabIndex        =   20
         Top             =   4680
         Width           =   11175
      End
      Begin VB.TextBox eMail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         MaxLength       =   40
         TabIndex        =   19
         Top             =   4200
         Width           =   5655
      End
      Begin VB.TextBox pHone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         MaxLength       =   30
         TabIndex        =   18
         Top             =   3480
         Width           =   2295
      End
      Begin MSMask.MaskEdBox AutomailNo 
         Height          =   255
         Left            =   4920
         TabIndex        =   17
         Top             =   3840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   6
         PromptChar      =   "_"
      End
      Begin VB.TextBox SName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   16
         Top             =   2160
         Width           =   3255
      End
      Begin MSMask.MaskEdBox YrInc 
         Height          =   285
         Left            =   7920
         TabIndex        =   3
         Top             =   3120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin VB.TextBox regAddr 
         Appearance      =   0  'Flat
         Height          =   525
         Left            =   3360
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   2520
         Width           =   5655
      End
      Begin VB.TextBox RegName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         MaxLength       =   70
         TabIndex        =   1
         Top             =   1800
         Width           =   5655
      End
      Begin VB.TextBox ActKey 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3360
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   1440
         Width           =   3255
      End
      Begin VB.PictureBox SSPanel3 
         BackColor       =   &H00C0C0C0&
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
         Height          =   915
         Left            =   3240
         Picture         =   "Frmsetup.frx":0000
         ScaleHeight     =   855
         ScaleWidth      =   5715
         TabIndex        =   5
         Top             =   120
         Width           =   5775
      End
      Begin MSComDlg.CommonDialog GetFile 
         Left            =   360
         Top             =   4080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "Bmp"
         FilterIndex     =   2
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry Date:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1680
         TabIndex        =   58
         Top             =   5760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label LblcVal 
         BackStyle       =   0  'Transparent
         Caption         =   "License:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1920
         TabIndex        =   56
         Top             =   6240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0C0C0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3120
         TabIndex        =   53
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label mDBPath 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6720
         TabIndex        =   46
         Top             =   5160
         Width           =   3135
      End
      Begin VB.Label mRptPath 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6720
         TabIndex        =   42
         Top             =   5520
         Width           =   3135
      End
      Begin VB.Label mPicsPath 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6720
         TabIndex        =   41
         Top             =   5880
         Width           =   3135
      End
      Begin VB.Label mDocsPath 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6720
         TabIndex        =   40
         Top             =   6240
         Width           =   3135
      End
      Begin VB.Label Label32 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Default Business Unit:"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   38
         Top             =   4800
         Width           =   2295
      End
      Begin VB.Image Image2 
         Height          =   855
         Left            =   960
         Picture         =   "Frmsetup.frx":10C62
         Stretch         =   -1  'True
         Top             =   120
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   960
         Stretch         =   -1  'True
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0C0C0&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   37
         Top             =   3480
         Width           =   135
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   36
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0C0C0&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5880
         TabIndex        =   35
         Top             =   3120
         Width           =   135
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6720
         TabIndex        =   34
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Building No.:"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   32
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   31
         Top             =   3120
         Width           =   135
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   30
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   29
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Company Prefix:"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   28
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Image Imgpic 
         BorderStyle     =   1  'Fixed Single
         Height          =   1215
         Left            =   9120
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Archive Month:"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   25
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Registration Number:"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Password Expiry Period:"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   5400
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Password Grace :"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Short Name:"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Auto Document Start No:"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   3840
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Email Address:"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Phone:"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Start Month:"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   9
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Company Name:"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Street Name/ Town/State:"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         TabIndex        =   7
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label label0 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Activation Key:"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   6
         Top             =   1440
         Width           =   1935
      End
   End
End
Attribute VB_Name = "FrmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MaxErrors As Integer
Dim db1 As Database, wrktemp As Workspace
Dim RegTab, rstBUnitD As Recordset
Dim ctr As Integer
Dim OkBuff1 As Integer, OkBuff2 As Integer
Dim OkBuff3 As Integer, OkBuff4 As Integer
Dim PicName As String, secKey As Long
Dim mvActkey As Long, mvRPM As String, stateBuff As String
Dim GenClass As New mProLog, strpos As Integer
Dim StTab As Recordset, StrPtr As Integer
Dim strlen, OkCtr As Integer

Private Sub ActKey_LostFocus()
On Error GoTo PError
If ActKey <> "" Then
    OkBuff1 = 1
Else
    OkBuff1 = 0
End If
PError:
End Sub

Private Sub cmbBUnitD_lostfocus()
On Error GoTo PError:
    strpos = InStr(1, cmbBUnitD, ",", 1)
    If strpos = 0 Then
       mvBuff = cmbBUnitD
    Else
      mvBuff = Left(cmbBUnitD, strpos - 1)
    End If
    If Len(cmbBUnitD) = 0 Then
       eMail.SetFocus
       Exit Sub
    End If
    mvSql = "rcode = '" & Trim(mvBuff) & "'"
    rstBUnitD.FindFirst mvSql
    If rstBUnitD.NOMATCH Then
        MsgBox "Invalid dropdown reference entered.", vbInformation, Me.Caption
        cmbBUnitD.SetFocus
        Exit Sub
    Else
        cmbBUnitD = rstBUnitD!RCode + ", " + rstBUnitD!RDesc
    End If
PError:
End Sub

Private Sub cmdLog_Click()
On Error GoTo PError
   GetFile.Filter = "Bitmap Files (*.BMP)|*.bmp| All Files|*.*"
   GetFile.FilterIndex = 1
   GetFile.DefaultExt = "bmp"
   GetFile.ShowOpen
   PicName = GetFile.FileName
   Imgpic = LoadPicture(GetFile.FileName)
PError:
End Sub

Private Sub cmdOk_Click()
On Error GoTo PError
 OkCtr = OkCtr + 1
 If OkCtr = 11 Then
    'sControl.Visible = True
    cVal.Visible = True
    LblcVal.Visible = True
    NodYear.Visible = True
    Label8.Visible = True
 Else
    'sControl.Visible = False
    cVal.Visible = False
    LblcVal.Visible = False
    NodYear.Visible = False
    Label8.Visible = False
 End If
 If OkBuff1 = 0 Then
    RegTab.Edit
    RegTab("rind") = "X"
    RegTab.Update
    Unload Me
 Else
    AkeyChk
 End If
PError:
End Sub
Public Sub secHarsh()
On Error GoTo PError
    secKey = 0
    If Bnum = "" Or Bnum = "0" Then
       Bnum = 1
    End If
    secKey = Int(ActKey / (100664 * Sqr(ActKey) * Bnum) * reG)
PError:
End Sub

Private Sub CmdDocs_Click()
    On Error GoTo ErrHandler
    GetFile.FileName = ""
    GetFile.Filter = "All Files (*.*)|*.*|Scanned Documents (*.bmp)|*.bmp"
    GetFile.FilterIndex = 2
    GetFile.DefaultExt = "MDB"
    'GetFile.Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
    GetFile.ShowOpen
    'Getfile.Action = 1
    mDocsPath = GetFile.FileName
    If Len(mDocsPath) = 0 Then
        MsgBox "No file was selected", vbInformation, Me.Caption
        Exit Sub
    Else
        MsgBox "File Selection successful", vbInformation, Me.Caption
    End If
Start0:
    Exit Sub
ErrHandler:
    MsgBox "The error: " + Error$ + " has occurred", "Error Condition Detected"
    On Error GoTo 0
    Resume Start0

End Sub


Private Sub cmdPics_Click()
    On Error GoTo ErrHandler
    GetFile.FileName = ""
    GetFile.Filter = "All Files (*.*)|*.*|Picture Files (*.jpg)|*.jpg"
    GetFile.FilterIndex = 2
    GetFile.DefaultExt = "MDB"
    'GetFile.Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
    GetFile.ShowOpen
    'Getfile.Action = 1
    mPicsPath = GetFile.FileName
    If Len(mPicsPath) = 0 Then
        MsgBox "No file was selected", vbInformation, Me.Caption
        Exit Sub
    Else
        MsgBox "File Selection successful", vbInformation, Me.Caption
    End If
Start0:
    Exit Sub
ErrHandler:
    MsgBox "The error: " + Error$ + " has occurred", "Error Condition Detected"
    On Error GoTo 0
    Resume Start0
End Sub
Private Sub cmdDB_Click()
    On Error GoTo ErrHandler
    GetFile.FileName = ""
    GetFile.Filter = "All Files (*.*)|*.*|Database Files (*.mdb)|*.mdb"
    GetFile.FilterIndex = 2
    GetFile.DefaultExt = "MDB"
    'GetFile.Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
    GetFile.ShowOpen
    'Getfile.Action = 1
    mDBPath = GetFile.FileName
    If Len(mDBPath) = 0 Then
        MsgBox "No file was selected", vbInformation, Me.Caption
        Exit Sub
    Else
        MsgBox "File Selection successful", vbInformation, Me.Caption
    End If
Start0:
    Exit Sub
ErrHandler:
    MsgBox "The error: " + Error$ + " has occurred", "Error Condition Detected"
    On Error GoTo 0
    Resume Start0
End Sub

Private Sub cmdRpt_Click()
   ' On Error GoTo ErrHandler
    GetFile.FileName = ""
    GetFile.Filter = "All Files (*.*)|*.*|Report Files (*.rpt)|*.rpt"
    GetFile.FilterIndex = 2
    GetFile.DefaultExt = "MDB"
    'GetFile.Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
    GetFile.ShowOpen
    'Getfile.Action = 1
    mRptPath = GetFile.FileName
    If Len(mRptPath) = 0 Then
        MsgBox "No file was selected", vbInformation, Me.Caption
        Exit Sub
    Else
        MsgBox "File Selection successful", vbInformation, Me.Caption
    End If
Start0:
    Exit Sub
ErrHandler:
    MsgBox "The error: " + Error$ + " has occurred", "Error Condition Detected"
    On Error GoTo 0
    Resume Start0

End Sub


Private Sub CmdReg_Click()
On Error GoTo PError
    If Len(Trim(RegName)) = 0 Then
       MsgBox "Invalid company Name", vbInformation, "Setup"
       RegName.SetFocus
       contEnab
       Exit Sub
    End If
    If Len(Trim(regAddr)) = 0 Then
       MsgBox "Invalid company Address", vbInformation, "Setup"
       regAddr.SetFocus
       contEnab
       Exit Sub
    End If
    If Len(Trim(reG)) = 0 Or reG = "0" Then
       MsgBox "Invalid company registration Number", vbInformation, "Setup"
       reG.SetFocus
       contEnab
       Exit Sub
    End If
    'If Int(Bnum) = 0 Or Bnum = 0 Then
       'MsgBox "Company building Number cannot be zero", vbInformation, "Setup"
       'Bnum.SetFocus
       'contEnab
       'Exit Sub
    'End If
    If Len(Trim(pHone)) = 0 Or pHone = "0" Then
       MsgBox "Invalid phone number", vbInformation, "Setup"
       Bnum.SetFocus
       contEnab
       Exit Sub
    End If
    If PassGrs < 0 Or PassGrs > 6 Or Len(Trim(PassGrs)) = 0 Then
       MsgBox "Invalid Password Grace, cannot be less than zero or greater than 6 , please Re-enter", vbInformation, "Setup"
       PassGrs.SetFocus
       contEnab
       Exit Sub
    End If
    If PassExp < 0 Or PassExp > 365 Or Len(Trim(PassExp)) = 0 Then
       MsgBox "Invalid Password Expiry period, cannot be less than zero or greater than 365 , please Re-enter", vbInformation, "Setup"
       PassExp.SetFocus
       contEnab
       Exit Sub
    End If
    If AutomailNo < 0 Or Len(Trim(AutomailNo)) = 0 Then
       MsgBox "Invalid AutoStaff Number Count Start, cannot be less than zero or blank , please Re-enter", vbInformation, "Setup"
       AutomailNo.SetFocus
       contEnab
       Exit Sub
    End If
    secHarsh
    RegTab.Edit
    RegTab("regName") = RegName
    RegTab("regnum") = reG
    RegTab("bnum") = Bnum
    RegTab("regAddr") = regAddr
    RegTab("regphone") = pHone
    RegTab("regemail") = eMail
    RegTab("regdInc") = YrInc
    RegTab("Staffctr") = AutomailNo
    RegTab("logo") = PicName
    GLogo = PicName
    RegTab("sname") = SName
    RegTab("coyId") = mvCoyId
    RegTab("scontrol") = sControl
    RegTab("cval") = cVal
    RegTab("nodyear") = NodYear
    RegTab("seckey") = secKey
    RegTab("passexp") = PassExp
    RegTab("passgrs") = PassGrs
    RegTab("autoctr") = AutoCtr
    ' nice little routine that extracts a path from a filename directory
    strlen = Len(mDBPath)
    StrPtr = 1
    strpos = 0
    Do While strlen > StrPtr
      strpos = InStr(StrPtr, mDBPath, "\")
      If strpos = 0 Then
         Exit Do
      Else
         StrPtr = strpos + 1
      End If
    Loop
    StrPtr = StrPtr - 1
    RegTab("dbpath") = Left(mDBPath, StrPtr)
    ' nice little routine that extracts a path from a filename directory
    strlen = Len(mRptPath)
    StrPtr = 1
    strpos = 0
    Do While strlen > StrPtr
      strpos = InStr(StrPtr, mRptPath, "\")
      If strpos = 0 Then
         Exit Do
      Else
         StrPtr = strpos + 1
      End If
    Loop
    StrPtr = StrPtr - 1
    RegTab("rptpath") = Left(mRptPath, StrPtr)
    Rptpath = Left(mRptPath, StrPtr)
    ' nice little routine that extracts a path from a filename directory
    strlen = Len(mPicsPath)
    StrPtr = 1
    strpos = 0
    Do While strlen > StrPtr
      strpos = InStr(StrPtr, mPicsPath, "\")
      If strpos = 0 Then
         Exit Do
      Else
         StrPtr = strpos + 1
      End If
    Loop
    StrPtr = StrPtr - 1
    RegTab("picspath") = Left(mPicsPath, StrPtr)
    PicsPath = Left(mPicsPath, StrPtr)
    ' nice little routine that extracts a path from a filename directory
    strlen = Len(mDocsPath)
    StrPtr = 1
    strpos = 0
    Do While strlen > StrPtr
      strpos = InStr(StrPtr, mDocsPath, "\")
      If strpos = 0 Then
         Exit Do
      Else
         StrPtr = strpos + 1
      End If
    Loop
    StrPtr = StrPtr - 1
    RegTab("docspath") = Left(mDocsPath, StrPtr)
    DocsPath = Left(mDocsPath, StrPtr)
  ' file extraction completed
    'strpos = InStr(1, cmbState, ",", 1)
    'RegTab("GState") = Left(cmbState, strpos - 1)
    strpos = InStr(1, cmbBUnitD, ",", 1)
    RegTab("Gbud") = Left(cmbBUnitD, strpos - 1)
    RegTab.Update
    'gstate = Left(cmbState, strpos - 1)
    gBunitD = Left(cmbBUnitD, strpos - 1)
    MVCoyname = RegName
    'GlbState = Left(cmbState, strpos - 1)
    FrmMtrack.Caption = RegName
    Me.Caption = RegName
    cmdReg.Enabled = False
    sControl.Visible = False
    cVal.Visible = False
    LblcVal.Visible = False
    NodYear.Visible = False
    Label8.Visible = False
    OkCtr = 0
    ActKey = ""
    contDisab
    Exit Sub
PError:
    MsgBox "Invalid entries detected, Correct and retry ", vbInformation, "System Setup"
    Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)
    db1.Close
End Sub




Private Sub RegAddr_LostFocus()
On Error GoTo PError
If regAddr <> "" Then
    OkBuff3 = 3
Else
    OkBuff3 = 0
End If
If OkBuff1 * OkBuff2 * OkBuff3 * OkBuff4 = 0 Then
    If Len(Trim(regAddr)) = 0 Then
       MsgBox "Invalid company Address", vbInformation, "Setup"
       regAddr.SetFocus
       contEnab
       Exit Sub
    End If
Else
    cmdReg.Enabled = True
End If
PError:
End Sub

Private Sub RegName_GotFocus()
    RegName.SelStart = 0
    RegName.SelLength = Len(RegName)
End Sub

Private Sub RegName_LostFocus()
On Error GoTo PError
If RegName <> "" Then
    OkBuff2 = 2
Else
    OkBuff2 = 0
End If
If OkBuff1 * OkBuff2 * OkBuff3 * OkBuff4 = 0 Then
    If Len(RegName) = 0 Then
       MsgBox "Invalid company Name", vbInformation, "Setup"
       RegName.SetFocus
       contEnab
       Exit Sub
    End If
Else
    cmdReg.Enabled = True
End If
PError:
End Sub

Private Sub Form_Load()
On Error GoTo PError
     GenClass.fleLogin mvUserid, "Accessed System Setup", Date, Time
     Call FormCentreMDI(Me)
     Set wrktemp = DBEngine.Workspaces(0)
     Set db1 = wrktemp.OpenDatabase(DBpath, True)
     Set RegTab = db1.OpenRecordset("DproReg", dbOpenDynaset)
     'Set StTab = db1.OpenRecordset("defState", dbOpenDynaset)
     Set rstBUnitD = db1.OpenRecordset("defbUnitD", dbOpenDynaset)
     Dim flag As Integer, mtable As String
         ''__________________________________
    
    If Not (rstBUnitD.EOF And rstBUnitD.BOF) Then rstBUnitD.MoveFirst
    Do While Not rstBUnitD.EOF
       cmbBUnitD.AddItem rstBUnitD("rcode") + ", " + rstBUnitD("rdesc")
       rstBUnitD.MoveNext
    Loop
    cmbBUnitD.ListIndex = 0
     If (RegTab.EOF And RegTab.BOF) Then
        RegTab.AddNew
        RegTab("regName") = ""
        RegTab("regnum") = 0
        RegTab("regAddr") = ""
        RegTab("regphone") = 0
        RegTab("regemail") = ""
        RegTab("yrInc") = Date
        RegTab("yrend") = Date
        RegTab("Staffctr") = 0
        RegTab("regDInc") = Date
        RegTab("bnum") = 0
        RegTab("seckey") = 0
        RegTab("state") = ""
        RegTab("AutoMailNo") = 0
        RegTab("logo") = ""
        RegTab("sname") = ""
        RegTab("passgrs") = 0#
        RegTab("medlim") = 0#
        RegTab("passexp") = 90
        RegTab("dbpath") = ""
        RegTab("rptpath") = ""
        RegTab("picspath") = ""
        RegTab("docspath") = ""
        RegTab.Update
        RegName.Enabled = False
        YrInc.Enabled = False
        ActKey.Enabled = False
        cmdReg.Enabled = False
        RegTab("actkey") = 0
        RegTab("autoctr") = AutoCtr
        RegTab.Update
        GAutoCtr = AutoCtr
    Else
        mvActkey = RegTab("actkey")
        showdata
    End If
    cmdReg.Enabled = False
    OkBuff1 = 0
    OkBuff2 = 0
    OkBuff3 = 0
    OkBuff4 = 0
    OkCtr = 0
    ctr = 1
    YrInc = Date
    contDisab
    sControl.Visible = False
    cVal.Visible = False
    LblcVal.Visible = False
    NodYear.Visible = False
    Label8.Visible = False
PError:
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub sControl_lostfocus()
    sControl.Enabled = False
End Sub
Private Sub cval_lostfocus()
    cVal.Enabled = False
    LblcVal.Visible = False
    NodYear.Visible = False
    Label8.Visible = False
End Sub


Private Sub yrInc_LostFocus()
On Error GoTo PError
If YrInc <> "" Then
    OkBuff4 = 4
Else
    OkBuff4 = 0
End If
If OkBuff1 * OkBuff2 * OkBuff3 * OkBuff4 = 0 Then
    If YrInc > Date Or YrInc = "" Then
       MsgBox "Invalid year of incorporation", vbInformation, "Setup"
       Bnum.SetFocus
       contEnab
       Exit Sub
    End If
Else
    cmdReg.Enabled = True
End If
PError:
End Sub

Public Sub AkeyChk()
On Error GoTo PError
     While ctr < 3
        If ActKey = "" Then
           ActKey = 0
        End If
        If ActKey = mvActkey Then
           contEnab
           showdata
           ActKey.Enabled = False
       Else
            Beep
            MsgBox "Wrong Activation key. Retry", vbCritical
            ActKey.Enabled = True
            ActKey.SetFocus
            ctr = ctr + 1
            ActKey = ""
            contDisab
        End If
        Exit Sub
    Wend
    MsgBox "Contact your Systems Administrator", vbCritical
    If ctr = 4 Then
       RegTab.Edit
       RegTab("rind") = "X"
       RegTab.Update
       Unload Me
    End If
    Unload Me
PError:
    MsgBox "Invalid data", vbInformation, Me.Caption
End Sub
Public Sub contDisab()
On Error GoTo PError
        RegName.Visible = False
        reG.Visible = False
        regAddr.Visible = False
        pHone.Visible = False
        eMail.Visible = False
        YrInc.Visible = False
        YrEnd.Visible = False
        AutomailNo.Visible = False
        Imgpic.Visible = False
        SName.Visible = False
        PassGrs.Visible = False
        PassExp.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        Label9.Visible = False
        Label12.Visible = False
        Label13.Visible = False
        Label14.Visible = False
        Imgpic.Visible = False
        cmdReg.Visible = False
        cmdReg.Visible = False
        cmdLog.Visible = False
        mvCoyId.Visible = False
        Label15.Visible = False
        Label19.Visible = False
        Bnum.Visible = False
        cmdExit.Visible = True
        Label21.Visible = False
        Label7.Visible = False
        Label22.Visible = False
        Label18.Visible = False
        Label23.Visible = False
        Label20.Visible = False
        Label17.Visible = False
        Label24.Visible = False
        Label32.Visible = False
        cmbBUnitD.Visible = False
        cmdDB.Visible = False
        cmdRpt.Visible = False
        CmdDocs.Visible = False
        cmdPics.Visible = False
        mDBPath.Visible = False
        mRptPath.Visible = False
        mPicsPath.Visible = False
        mDocsPath.Visible = False
        AutoCtr.Visible = False
PError:

End Sub
Public Sub contEnab()
On Error GoTo PError
        RegName.Visible = True
        reG.Visible = True
        regAddr.Visible = True
        pHone.Visible = True
        eMail.Visible = True
        YrInc.Visible = True
        YrEnd.Visible = True
        ActKey.Visible = True
        AutomailNo.Visible = True
        Imgpic.Visible = True
        SName.Visible = True
        PassExp.Visible = True
        PassGrs.Visible = True
        label1.Visible = True
        Label2.Visible = True
        Label3.Visible = True
        Label4.Visible = True
        Label5.Visible = True
        Label6.Visible = True
        Label7.Visible = True
        Label9.Visible = True
        Label12.Visible = True
        Label13.Visible = True
        Label14.Visible = True
        Label24.Visible = True
        mvCoyId.Visible = True
        Bnum.Visible = True
        Label15.Visible = True
        Label19.Visible = True
        Imgpic.Visible = True
        cmdReg.Visible = True
        cmdReg.Enabled = True
        cmdLog.Visible = True
        cmdExit.Visible = True
        Label21.Visible = True
        Label7.Visible = True
        Label22.Visible = True
        Label18.Visible = True
        Label23.Visible = True
        Label20.Visible = True
        Label17.Visible = True
        Label32.Visible = True
        cmbBUnitD.Visible = True
        cmdDB.Visible = True
        cmdRpt.Visible = True
        CmdDocs.Visible = True
        cmdPics.Visible = True
        mDBPath.Visible = True
        mRptPath.Visible = True
        mPicsPath.Visible = True
        mDocsPath.Visible = True
        AutoCtr.Visible = True
PError:
End Sub

Public Sub showdata()
 On Error GoTo PError
    RegTab.MoveFirst
    RegName = RegTab("regName")
    reG = RegTab("regnum")
    regAddr = RegTab("regaddr")
    pHone = RegTab("Regphone")
    'If IsNull(RegTab!gstate) = True Then
    '   cmbState.ListIndex = 0
    'Else
    '   stateBuff = "StateCode = '" & RegTab("GState") & "'"
    '   StTab.FindFirst stateBuff
    '   cmbState = StTab!StateCode & "," & StTab!StateName
    'End If
    If IsNull(RegTab!gbud) = True Then
       cmbBUnitD.ListIndex = 0
    Else
       mvBuff = "rCode = '" & RegTab("Gbud") & "'"
       rstBUnitD.FindFirst mvBuff
       cmbBUnitD = rstBUnitD!RCode & "," & rstBUnitD!RDesc
    End If
    If IsNull(RegTab!RegeMail) = True Then
        eMail = ""
    Else
        eMail = RegTab("Regemail")
    End If
    
    If IsNull(RegTab!RegDinc) = True Then
        YrInc = Date
    Else
        YrInc = RegTab("RegDinc")
    End If
    If IsNull(RegTab!coyid) = True Then
        mvCoyId = ""
    Else
        mvCoyId = RegTab("coyid")
    End If
    If IsNull(RegTab!sControl) = True Then
        sControl = ""
    Else
        sControl = RegTab("scontrol")
    End If
    If IsNull(RegTab!cVal) = True Then
        cVal = 0
    Else
        cVal = RegTab("cVal")
    End If
    If IsNull(RegTab!NodYear) = True Then
        NodYear = ""
    Else
        NodYear = RegTab("nodyear")
    End If
    If IsNull(RegTab!logo) = True Then
        GTBuff = PicsPath & "rocks.bmp"
        Imgpic = LoadPicture(GTBuff)
    Else
        GTBuff = RegTab!logo
        Imgpic = LoadPicture(GTBuff)
    End If
    If IsNull(RegTab!Bnum) Then
        DoEvents
    Else
        Bnum = RegTab("bnum")
    End If
    SName = RegTab("sname")
    RegName = RegTab("RegName")
    PassExp = RegTab("PassExp")
    PassGrs = RegTab("PassGrs")
    AutomailNo = RegTab("staffctr")
    mDBPath = RegTab("dbpath")
    mRptPath = RegTab("rptpath")
    mPicsPath = RegTab("picspath")
    mDocsPath = RegTab("docspath")
    If RegTab!AutoCtr = True Then
       AutoCtr.Value = 1
    Else
       AutoCtr.Value = 0
    End If
PError:
End Sub


