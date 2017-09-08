VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmUserLect 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datapro"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   13830
   Begin Threed.SSPanel SSPanel3 
      Height          =   1575
      Left            =   11520
      TabIndex        =   16
      Top             =   600
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   2778
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
      BevelInner      =   2
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
         Left            =   1200
         TabIndex        =   9
         Top             =   960
         Width           =   870
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
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   870
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
         Left            =   1200
         TabIndex        =   7
         Top             =   600
         Width           =   870
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
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   870
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
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   870
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
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   870
      End
   End
   Begin Threed.SSPanel SSPanel5 
      Height          =   1575
      Left            =   0
      TabIndex        =   18
      Top             =   600
      Width           =   11535
      _Version        =   65536
      _ExtentX        =   20346
      _ExtentY        =   2778
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
      BevelInner      =   2
      Begin VB.CheckBox ChkLock 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Account Lock"
         Height          =   255
         Left            =   5880
         TabIndex        =   24
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Data datgenTab 
         Caption         =   "Gentab"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   9120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox cmbStaffNo 
         Height          =   315
         Left            =   1200
         TabIndex        =   23
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txtConfm 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4080
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtUserPswd 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   1
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Confirm:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   610
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "User ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   250
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   990
         Width           =   975
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   975
      Left            =   0
      TabIndex        =   17
      Top             =   5040
      Width           =   13815
      _Version        =   65536
      _ExtentX        =   24368
      _ExtentY        =   1720
      _StockProps     =   15
      Caption         =   "Operations"
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
      Alignment       =   6
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   11
         Top             =   360
         Width           =   1305
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   13
         Top             =   360
         Width           =   1305
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   12
         Top             =   360
         Width           =   1305
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
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
         Left            =   3000
         TabIndex        =   10
         Top             =   360
         Width           =   1305
      End
      Begin VB.CommandButton cmdExit 
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
         Left            =   9600
         TabIndex        =   0
         Top             =   360
         Width           =   1305
      End
      Begin VB.CommandButton cmdCanc 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   14
         Top             =   360
         Width           =   1305
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   13815
      _Version        =   65536
      _ExtentX        =   24368
      _ExtentY        =   1085
      _StockProps     =   15
      Caption         =   "Teachers Access  Administration"
      BackColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   4
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   2655
      Left            =   0
      TabIndex        =   25
      Top             =   2280
      Visible         =   0   'False
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   4683
      _StockProps     =   15
      Caption         =   "Configuration"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   2
      Font3D          =   2
      Alignment       =   6
      Begin Threed.SSCheck chk1 
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   480
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Company Setup"
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
      Begin Threed.SSCheck chk2 
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   840
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "User Setup"
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
      Begin Threed.SSCheck chk5 
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   1920
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Scoring Percentages"
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
      Begin Threed.SSCheck chk3 
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   1200
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Structural References"
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
      Begin Threed.SSCheck chk4 
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   1560
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Interface Management"
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
   Begin Threed.SSPanel SSPanel7 
      Height          =   2655
      Left            =   2760
      TabIndex        =   31
      Top             =   2280
      Visible         =   0   'False
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   4683
      _StockProps     =   15
      Caption         =   "Administration"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   2
      Font3D          =   2
      Alignment       =   6
      Begin Threed.SSCheck chk6 
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   480
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Classroom Administration"
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
      Begin Threed.SSCheck chk8 
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   1200
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Staff Responsibilities"
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
      Begin Threed.SSCheck chk9 
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   1560
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Student Transfers"
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
      Begin Threed.SSCheck chk7 
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   840
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Subject Allocations"
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
      Begin Threed.SSCheck chk10 
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   1920
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "General Reference"
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
   Begin Threed.SSPanel SSPanel8 
      Height          =   2655
      Left            =   5520
      TabIndex        =   37
      Top             =   2280
      Visible         =   0   'False
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   4683
      _StockProps     =   15
      Caption         =   "Persons"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   2
      Font3D          =   2
      Alignment       =   6
      Begin Threed.SSCheck chk15 
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   1920
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Super User"
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
      Begin Threed.SSCheck chk14 
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   1560
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Student Withdrawal"
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
      Begin Threed.SSCheck chk11 
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   480
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Staff Management"
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
      Begin Threed.SSCheck chk12 
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   840
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Student Admissions"
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
      Begin Threed.SSCheck chk13 
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   1200
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Student Records"
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
   Begin Threed.SSPanel SSPanel9 
      Height          =   2655
      Left            =   11040
      TabIndex        =   43
      Top             =   2280
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   4683
      _StockProps     =   15
      Caption         =   "Reviews"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   2
      Font3D          =   2
      Alignment       =   6
      Begin Threed.SSCheck chk21 
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   480
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Enquiries"
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
      Begin Threed.SSCheck chk23 
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   1200
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Teachers Reports"
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
      Begin Threed.SSCheck chk22 
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   840
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "General Reports"
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
      Begin Threed.SSCheck chk25 
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   1920
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Notes"
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
      Begin Threed.SSCheck chk24 
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   1560
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Activity log"
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
   Begin Threed.SSPanel SSPanel2 
      Height          =   2655
      Left            =   8280
      TabIndex        =   49
      Top             =   2280
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   4683
      _StockProps     =   15
      Caption         =   "Data"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   2
      Font3D          =   2
      Alignment       =   6
      Begin Threed.SSCheck chk16 
         Height          =   255
         Left            =   360
         TabIndex        =   50
         Top             =   480
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Evaluation Input"
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
      Begin Threed.SSCheck chk20 
         Height          =   255
         Left            =   360
         TabIndex        =   51
         Top             =   1920
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Result Advice"
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
      Begin Threed.SSCheck chk18 
         Height          =   255
         Left            =   360
         TabIndex        =   52
         Top             =   1200
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Performance Reports"
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
      Begin Threed.SSCheck chk19 
         Height          =   255
         Left            =   360
         TabIndex        =   53
         Top             =   1560
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Intelligence Reports"
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
      Begin Threed.SSCheck chk17 
         Height          =   255
         Left            =   360
         TabIndex        =   54
         Top             =   840
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Processing Routines"
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
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   0
      Picture         =   "frmLAccess.frx":0000
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   8280
   End
End
Attribute VB_Name = "frmUserLect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mvOpt As Integer, mvReply As Integer
Dim db6 As Database, mvErrFlg As Integer
Dim wrktemp As Workspace
Dim rstUserFle As Recordset, rst, rstReg As Recordset
Dim GenClass As New DProLog
Dim mvBuff As String, mvLRec As String
Dim rstLMast As Recordset
Dim strpos As Integer, tempBuff As String

Private Sub FindAcc()
    mvBuff = "Select * from UserFle Where UserID = '" _
        & tempBuff & "'"
    Set rstUserFle = db6.OpenRecordset(mvBuff, dbOpenDynaset)
     Select Case mvOpt
        Case 1
            If Not (rstUserFle.EOF And rstUserFle.BOF) Then
                MsgBox "User ID Already Defined", vbInformation, Me.Caption
                cmbStaffNo.SetFocus
            End If
        Case Is > 1
            If rstUserFle.EOF And rstUserFle.BOF Then
                MsgBox "User ID Does Not Exits", vbInformation, Me.Caption
                cmbStaffNo.SetFocus
            Else
                showdata
            End If
    End Select
End Sub

Private Sub cmbStaffNo_Click()
    strpos = InStr(1, cmbStaffNo, ",", 1)
    tempBuff = Left(cmbStaffNo, strpos - 1)
    txtUserName = Trim(Mid(cmbStaffNo, strpos + 1))
End Sub

Private Sub cmbStaffNo_LostFocus()
 'On Error GoTo perror
    strpos = InStr(1, cmbStaffNo, ",", 1)
    If strpos = 0 Then
        tempBuff = cmbStaffNo
    Else
        tempBuff = Left(cmbStaffNo, strpos - 1)
    End If
    mvSql = "StaffNo = '" & tempBuff & "'"
    rstLMast.FindFirst mvSql
    If rstLMast.NOMATCH Then
        MsgBox "Lecturer's number entered not found", vbInformation, Me.Caption
        Exit Sub
    Else
        txtUserName = rstLMast!StaffNames
    End If
PError:
End Sub

Private Sub CmdAdd_Click()
    FldEnab
    ClearFlds
    CmdDisab
    cmdedit.Enabled = False
    cmddel.Enabled = False
    EnabSavCan
    mvOpt = 1
    cmdadd.Enabled = False
    cmbStaffNo.SetFocus
End Sub

Private Sub CmdBrow_Click()
    Screen.MousePointer = vbHourglass
    mvSql = "Select UserID, UserName From UserFle"
    Set rst = db6.OpenRecordset(mvSql, dbOpenSnapshot)
     If (rst.BOF And rst.EOF) Then
         MsgBox "No users are defined", vbExclamation, "User Administration"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "List of System Users"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         frmBrowse.grdGeneral.Columns(0).Width = 1400
         frmBrowse.grdGeneral.Columns(1).Width = 2800
         Col0.Caption = "User Id"
         Col1.Caption = "User Names"
   End If
   frmBrowse.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCanc_Click()
    ClearFlds
    disSavCan
    EnabCmd
    CmdEnab
    cmdadd.Enabled = True
    If (rstUserFle.EOF And rstUserFle.BOF) Then
        cmdedit.Enabled = False
        cmddel.Enabled = False
    End If
    FldDisab
End Sub

Private Sub CmdDel_Click()
    mvOpt = 3
    If tempBuff = "" Then
       MsgBox "User ID Not Specified", vbInformation, Me.Caption
       Exit Sub
    End If
    FindAcc
    mvReply = DeleteCheck()
    If mvReply = vbYes Then
        rstUserFle.Delete
        MsgBox "User Details Deleted", vbInformation, Me.Caption
        FldEnab
        ClearFlds
        FldDisab
    Else
        MsgBox "User Details Not Deleted", vbExclamation, Me.Caption
    End If
End Sub

Private Sub CmdEdit_Click()
   mvOpt = 2
   If tempBuff = "" Then
       MsgBox "User ID Not Specified", vbInformation, Me.Caption
       Exit Sub
   End If
   FindAcc
   FldEnab
   cmbStaffNo.Enabled = False
   CmdDisab
   cmdadd.Enabled = False
   cmddel.Enabled = False
   EnabSavCan
   txtUserName.SetFocus
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdTop_Click()
    mvBuff = "Select * from UserFle Order By UserID;"
    Set rstUserFle = db6.OpenRecordset(mvBuff, dbOpenDynaset)
    If Not rstUserFle.BOF Then
       rstUserFle.MoveFirst
       showdata
       FldDisab
    End If
End Sub

Private Sub CmdNext_Click()
    Dim flag As Integer
    On Error GoTo NextError
    If rstUserFle.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstUserFle.EOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEPAST)
        rstUserFle.MoveLast
    Else
        rstUserFle.MoveNext
        If rstUserFle.EOF Then
            rstUserFle.MoveLast
        End If
    End If
NextClear:
    showdata
    Exit Sub
NextError:
    ErrorMessages (NOMOVEPAST)
    rstUserFle.Requery
    rstUserFle.MoveLast
    On Error GoTo 0
    Resume NextClear
End Sub

Private Sub CmdPrev_Click()
  Dim flag As Integer
    
    On Error GoTo PrevError
    If rstUserFle.RecordCount < 1 Then
        ErrorMessages (EMPTYTABLE)
        Exit Sub
    End If
    flag = rstUserFle.BOF
    If flag Then
        Beep
        ErrorMessages (NOMOVEFRONT)
        rstUserFle.MoveFirst
    Else
        rstUserFle.MovePrevious
        If rstUserFle.BOF Then
            rstUserFle.MoveFirst
        End If
    End If
PrevClear:
    showdata
    Exit Sub
PrevError:
    ErrorMessages (NOMOVEFRONT)
    On Error GoTo 0
    Resume PrevClear
End Sub

Private Sub CmdBott_Click()
    mvBuff = "Select * from UserFle Order By UserID;"
    Set rstUserFle = db6.OpenRecordset(mvBuff, dbOpenDynaset)
    If Not (rstUserFle.BOF And rstUserFle.EOF) Then
       rstUserFle.MoveLast
       showdata
       FldDisab
    End If
End Sub

Private Sub cmdFind_Click()
    Dim mvLength As Integer
    Dim mvBuff As String, mvAns As String
    mvLength = 0
    mvAns = InputBox$("Enter User ID:")
    mvLength = Len(mvAns)
    If mvLength = 0 Then
       MsgBox "User Selected Cancel or Nothing To Find", vbInformation, Me.Caption
       Exit Sub
    End If
    mvBuff = "Select * from UserFle Where UserID = '" _
        & mvAns & "'"
    Set rstUserFle = db6.OpenRecordset(mvBuff, dbOpenDynaset)
    If rstUserFle.EOF And rstUserFle.BOF Then
        MsgBox "User ID Not Found", vbInformation, Me.Caption
        Exit Sub
    Else
        showdata
    End If
End Sub

Private Sub CmdSave_Click()
    mvErrFlg = 0
    ValFlds
    If mvErrFlg = 0 Then
        Select Case mvOpt
            Case 1
                rstUserFle.AddNew
                rstUserFle("UserID") = UCase(tempBuff)
                UpdTab
            Case 2
                rstUserFle.Edit
                UpdTab
        End Select
        cmdadd.Enabled = True
        CmdEnab
        disSavCan
        EnabCmd
        FldDisab
        strpos = InStr(cmbStaffNo, ",")
        tmpStr1 = Trim(Left(cmbStaffNo, strpos - 1))
        mvBuff = "Select * from UserFle Order By UserID;"
        Set rstUserFle = db6.OpenRecordset(mvBuff, dbOpenDynaset)
        GTBuff = "userid = '" & UCase(tmpStr1) & "'"
        rstUserFle.FindFirst GTBuff
    End If
End Sub

Private Sub showdata()
    If Not (rstUserFle.BOF And rstUserFle.EOF) Then
        mvLRec = tempBuff
        tempBuff = rstUserFle!userid
        txtUserPswd = rstUserFle!UserPswd
        txtConfm = txtUserPswd
        txtUserName = rstUserFle!UserName
        chk1.Value = rstUserFle!mnu1
        chk2.Value = rstUserFle!mnu2
        chk3.Value = rstUserFle!mnu3
        chk4.Value = rstUserFle!mnu4
        chk5.Value = rstUserFle!mnu5
        chk6.Value = rstUserFle!mnu6
        chk7.Value = rstUserFle!mnu7
        chk8.Value = rstUserFle!mnu8
        chk9.Value = rstUserFle!mnu9
        chk10.Value = rstUserFle!mnu10
        chk11.Value = rstUserFle!mnu11
        chk12.Value = rstUserFle!mnu12
        chk13.Value = rstUserFle!mnu13
        chk14.Value = rstUserFle!mnu14
        chk15.Value = rstUserFle!mnu15
        chk16.Value = rstUserFle!mnu16
        chk17.Value = rstUserFle!mnu17
        chk18.Value = rstUserFle!mnu18
        chk19.Value = rstUserFle!mnu19
        chk20.Value = rstUserFle!mnu20
        chk21.Value = rstUserFle!mnu21
        chk22.Value = rstUserFle!mnu22
        chk23.Value = rstUserFle!mnu23
        chk24.Value = rstUserFle!mnu24
        chk25.Value = rstUserFle!mnu25
        If rstUserFle!alock = True Then
           ChkLock.Value = 1
        Else
           ChkLock.Value = 0
        End If
        cmbStaffNo = rstUserFle!userid & ", " & rstUserFle!UserName
        FldDisab
    End If
End Sub

Private Sub CmdEnab()
    cmdFind.Enabled = True
    cmdbrow.Enabled = True
    cmdtop.Enabled = True
    cmdbott.Enabled = True
    cmdnext.Enabled = True
    cmdprev.Enabled = True
End Sub

Private Sub EnabCmd()
    cmdadd.Enabled = True
    cmdedit.Enabled = True
    cmddel.Enabled = True
End Sub

Private Sub CmdDisab()
    cmdFind.Enabled = False
    cmdbrow.Enabled = False
    cmdtop.Enabled = False
    cmdbott.Enabled = False
    cmdnext.Enabled = False
    cmdprev.Enabled = False
    If rstUserFle.EOF And rstUserFle.BOF Then
        cmdedit.Enabled = False
        cmddel.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Call FormCentreMDI(Me)
    ClearFlds
    disSavCan
    Screen.MousePointer = vbHourglass
    Set wrktemp = DBEngine.Workspaces(0)
    Set db6 = wrktemp.OpenDatabase(DBpath, True)
    Set rstUserFle = db6.OpenRecordset("userfle", dbOpenDynaset)
    Set rstLMast = db6.OpenRecordset("lectmast", dbOpenDynaset)
    Set rstReg = db6.OpenRecordset("DProReg", dbOpenDynaset)
    If rstUserFle.EOF And rstUserFle.BOF Then
        CmdDisab
        cmdedit.Enabled = False
        cmddel.Enabled = False
    End If
    ''__________________________________
    If Not (rstLMast.BOF And rstLMast.EOF) Then rstLMast.MoveFirst
    Do While Not rstLMast.EOF
       cmbStaffNo.AddItem rstLMast("staffno") + ", " + rstLMast("staffnames")
       rstLMast.MoveNext
    Loop
    If rstLMast.RecordCount <> 0 Then
       cmbStaffNo.ListIndex = 0
    End If
    FldDisab
    GenClass.fleLogin mvUserid, "Accessed Lecturer Access Administration", Date, Time
    Screen.MousePointer = vbDefault
End Sub

Private Sub ClearFlds()
    tempBuff = ""
    txtUserName = ""
    txtUserPswd = ""
    txtConfm = ""
    chk1.Value = False
    chk2.Value = False
    chk3.Value = False
    chk4.Value = False
    chk5.Value = False
    chk6.Value = False
    chk7.Value = False
    chk8.Value = False
    chk9.Value = False
    chk10.Value = False
    chk11.Value = False
    chk12.Value = False
    chk13.Value = False
    chk14.Value = False
    chk15.Value = False
    chk16.Value = False
    chk17.Value = False
    chk18.Value = False
    chk19.Value = False
    chk20.Value = False
    chk21.Value = False
    chk22.Value = False
    chk23.Value = False
    chk24.Value = False
    chk25.Value = False
    ChkLock.Value = False
   mvOpt = 0
End Sub

Private Sub disSavCan()
    cmdSave.Enabled = False
    cmdCanc.Enabled = False
End Sub

Private Sub EnabSavCan()
    cmdSave.Enabled = True
    cmdCanc.Enabled = True
End Sub

Private Sub UpdTab()
'On Error GoTo PError
    rstUserFle("UserName") = txtUserName
    rstUserFle("UserPswd") = txtUserPswd
    rstUserFle("PassExp") = Date + rstReg!PassExp
    rstUserFle("mnu1") = chk1.Value
    rstUserFle("mnu2") = chk2.Value
    rstUserFle("mnu3") = chk3.Value
    rstUserFle("mnu4") = chk4.Value
    rstUserFle("mnu5") = chk5.Value
    rstUserFle("mnu6") = chk6.Value
    rstUserFle("mnu7") = chk7.Value
    rstUserFle("mnu8") = chk8.Value
    rstUserFle("mnu9") = chk9.Value
    rstUserFle("mnu10") = chk10.Value
    rstUserFle("mnu11") = chk11.Value
    rstUserFle("mnu12") = chk12.Value
    rstUserFle("mnu13") = chk13.Value
    rstUserFle("mnu14") = chk14.Value
    rstUserFle("mnu15") = chk15.Value
    rstUserFle("mnu16") = chk16.Value
    rstUserFle("mnu17") = chk17.Value
    rstUserFle("mnu18") = chk18.Value
    rstUserFle("mnu19") = chk19.Value
    rstUserFle("mnu20") = chk20.Value
    rstUserFle("mnu21") = chk21.Value
    rstUserFle("mnu22") = chk22.Value
    rstUserFle("mnu23") = chk23.Value
    rstUserFle("mnu24") = chk24.Value
    rstUserFle("mnu25") = chk25.Value
    rstUserFle("alock") = ChkLock.Value
    If mvOpt = 1 Then
       rstUserFle("acad") = True
    End If
    rstUserFle.Update
    Exit Sub
PError:
    MsgBox "Record already exist", vbExclamation
End Sub

Private Sub ValFlds()
    If tempBuff = "" Then
        MsgBox "Invalid User ID", vbExclamation
        mvErrFlg = 1
        cmbStaffNo.SetFocus
        Exit Sub
    End If
    If txtUserName = "" Then
        MsgBox "Invalid User Name", vbExclamation
        mvErrFlg = 1
        txtUserName.SetFocus
        Exit Sub
    End If
    If txtUserPswd = "" Or txtUserPswd <> txtConfm Then
        MsgBox "Invalid Password", vbExclamation
        mvErrFlg = 1
        txtUserPswd.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    db6.Close
End Sub

Private Sub txtUserName_GotFocus()
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName)
End Sub

Private Sub txtUserPswd_gotFocus()
    txtUserPswd.SelStart = 0
    txtUserPswd.SelLength = Len(txtUserPswd)
End Sub

Private Sub txtConfm_GotFocus()
    txtConfm.SelStart = 0
    txtConfm.SelLength = Len(txtConfm)
End Sub

Private Sub FldEnab()
    cmbStaffNo.Enabled = True
    txtUserName.Enabled = True
    txtUserPswd.Enabled = True
    txtConfm.Enabled = True
    chk1.Enabled = True
    chk2.Enabled = True
    chk3.Enabled = True
    chk4.Enabled = True
    chk5.Enabled = True
    chk6.Enabled = True
    chk7.Enabled = True
    chk8.Enabled = True
    chk9.Enabled = True
    chk10.Enabled = True
    chk11.Enabled = True
    chk12.Enabled = True
    chk13.Enabled = True
    chk14.Enabled = True
    chk15.Enabled = True
    chk16.Enabled = True
    chk17.Enabled = True
    chk18.Enabled = True
    chk19.Enabled = True
    chk20.Enabled = True
    chk21.Enabled = True
    chk22.Enabled = True
    chk23.Enabled = True
    chk24.Enabled = True
    chk25.Enabled = True
    ChkLock.Enabled = True
End Sub

Private Sub FldDisab()
    cmbStaffNo.Enabled = False
    txtUserName.Enabled = False
    txtUserPswd.Enabled = False
    txtConfm.Enabled = False
    chk1.Enabled = False
    chk2.Enabled = False
    chk3.Enabled = False
    chk4.Enabled = False
    chk5.Enabled = False
    chk6.Enabled = False
    chk7.Enabled = False
    chk8.Enabled = False
    chk9.Enabled = False
    chk10.Enabled = False
    chk11.Enabled = False
    chk12.Enabled = False
    chk13.Enabled = False
    chk14.Enabled = False
    chk15.Enabled = False
    chk16.Enabled = False
    chk17.Enabled = False
    chk18.Enabled = False
    chk19.Enabled = False
    chk20.Enabled = False
    chk21.Enabled = False
    chk22.Enabled = False
    chk23.Enabled = False
    chk24.Enabled = False
    chk25.Enabled = False
    ChkLock.Enabled = False
End Sub
