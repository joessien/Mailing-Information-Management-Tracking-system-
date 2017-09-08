VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmRsltEnq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Tracking System"
   ClientHeight    =   6285
   ClientLeft      =   1305
   ClientTop       =   1380
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   9645
   Begin VB.PictureBox SSPanel3 
      BackColor       =   &H00C0C0C0&
      Height          =   5535
      Left            =   3840
      ScaleHeight     =   5475
      ScaleWidth      =   5715
      TabIndex        =   10
      Top             =   720
      Width           =   5775
      Begin VB.PictureBox cmdAccept 
         Height          =   375
         Left            =   3600
         ScaleHeight     =   315
         ScaleWidth      =   915
         TabIndex        =   46
         Top             =   1080
         Width           =   975
      End
      Begin MSMask.MaskEdBox PYear 
         Height          =   255
         Left            =   1800
         TabIndex        =   45
         Top             =   720
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   4
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmbStudNo 
         Height          =   315
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   3855
      End
      Begin VB.PictureBox StudNames 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   360
         ScaleHeight     =   555
         ScaleWidth      =   5235
         TabIndex        =   28
         Top             =   1560
         Width           =   5295
      End
      Begin VB.PictureBox cmdExit 
         Height          =   375
         Left            =   4680
         ScaleHeight     =   315
         ScaleWidth      =   915
         TabIndex        =   48
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Lclass 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   2640
         TabIndex        =   47
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label29 
         BackColor       =   &H00400000&
         Caption         =   "Enter Year:"
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
         Height          =   255
         Left            =   480
         TabIndex        =   44
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4440
         TabIndex        =   43
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Term1Pos 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4440
         TabIndex        =   42
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Term2Pos 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4440
         TabIndex        =   41
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Term3Pos 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4440
         TabIndex        =   40
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Term6Pos 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4440
         TabIndex        =   39
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Term5Pos 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4440
         TabIndex        =   38
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Term4Pos 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4440
         TabIndex        =   37
         Top             =   3840
         Width           =   1095
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "Grade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3120
         TabIndex        =   36
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Term1Grade 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   35
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Term2grade 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   34
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Term3Grade 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   33
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Term6Grade 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   32
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Term5Grade 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   31
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Term4Grade 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   30
         Top             =   3840
         Width           =   1095
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Score"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1800
         TabIndex        =   29
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Term4Score 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   26
         Top             =   3840
         Width           =   1095
         WordWrap        =   -1  'True
      End
      Begin VB.Label Term5Score 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   25
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Term6Score 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   24
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label term3score 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   23
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label term2score 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   22
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label term1Score 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   21
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sixth Term"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fifth Term:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Second Term:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "First Term:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00400000&
         Caption         =   "Student No:"
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
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Third Term:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fourth Term:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   3840
         Width           =   1335
      End
   End
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H00C0C0C0&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5475
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   720
      Width           =   3735
      Begin VB.Data datGenTab 
         Caption         =   "General Data"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label stFileNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Date of Registration:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sex:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Age:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Date of Birth:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "File Number:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Sex 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Age 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label DOB 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label regDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label StudNo 
         Alignment       =   2  'Center
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
         Height          =   495
         Left            =   600
         TabIndex        =   1
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Image StudPic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2055
         Left            =   600
         Picture         =   "frmRstlEnq.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.PictureBox SSPanel5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   9555
      TabIndex        =   27
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "FrmRsltEnq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim db1 As Database, wrktemp As Workspace
Dim FVal As Boolean
Dim PicName As String, MAppid As String
Dim EditFlag As Boolean, MaxErrors As Integer
Dim PayTval As Integer, studBuff, mStudno As String
Dim strpos As Integer

Dim rstStudmast As Recordset
Dim rstCNames As Recordset
Dim rstRslt As Recordset

Dim GenClass As New mProLog
Private Sub cmbStudno_LostFocus()
    StudFind
End Sub

Private Sub cmdAccept_Click()
'On Error GoTo Perror
If Len(Trim(PYear)) = 0 Then
    MsgBox "Invalid Academic Year, Renter", vbCritical, "Enquiry"
    PYear.SetFocus
    Exit Sub
End If
mvSql = "Select * From histclassrslt where studno= '" & mStudno & "' and pyear = " & PYear & " order by schterm"
Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
If Not (rst.BOF And rst.EOF) Then
   mvSql = "code = '" & rst!currclass & "'"
   rstCNames.FindFirst mvSql
   If rstCNames.NOMATCH Then
      Lclass = ""
   Else
      If IsNull(rstCNames!NickName) Then
         Lclass.Caption = rstCNames!CGrp & " ," & rstCNames!Desc
      Else
         Lclass.Caption = rstCNames!CGrp & " ," & rstCNames!Desc & " - " & rstCNames!NickName
      End If
   End If
   rst.MoveFirst
   Do While Not rst.EOF
      Select Case rst!schterm
      Case "TERM1"
         term1Score = rst!avgscore
         Term1Grade = rst!avggrade
         Term1Pos = rst!classpos
      Case "TERM2"
         term2score = rst!avgscore
         Term2grade = rst!avggrade
         Term2Pos = rst!classpos
      Case "TERM3"
         term3score = rst!avgscore
         Term3Grade = rst!avggrade
         Term3Pos = rst!classpos
      Case "TERM4"
         Term4Score = rst!avgscore
         Term4Grade = rst!avggrade
         Term4Pos = rst!classpos
      Case "TERM5"
         Term5Score = rst!avgscore
         Term5Grade = rst!avggrade
         Term5Pos = rst!classpos
      Case "TERM6"
         Term6Score = rst!avgscore
         Term6Grade = rst!avggrade
         Term6Pos = rst!classpos
      End Select
     rst.MoveNext
   Loop
Else
       MsgBox "Mail with the Registration Number entered and year of study is not found.", vbCritical, "Enquiry"
       ClearData
End If
PError:
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'On Error GoTo PError
     GenClass.fleLogin mvUserid, "Accessed Mail Enquiry", Date, Time
     Call FormCentreMDI(Me)
     Set wrktemp = DBEngine.Workspaces(0)
     Set db1 = wrktemp.OpenDatabase(DBpath, True)
     Set rstStudmast = db1.OpenRecordset("Studmast", dbOpenDynaset)
     Set rstCNames = db1.OpenRecordset("DefClassNames", dbOpenDynaset)
     Set rstRslt = db1.OpenRecordset("histclassrslt", dbOpenDynaset)
     Me.Caption = MVCoyname
    ''__________________________________
    datgenTab.DatabaseName = DBpath
    datgenTab.RecordSource = "studmast"
    datgenTab.ReadOnly = False
    datgenTab.Exclusive = False
    datgenTab.Refresh
    Do While Not datgenTab.Recordset.EOF
       cmbStudNo.AddItem datgenTab.Recordset("studno") + ", " + datgenTab.Recordset("StudNames")
       datgenTab.Recordset.MoveNext
    Loop
    cmbStudNo.ListIndex = 0
    '______________________________________
    ClearData
    PYear = ""
PError:
End Sub

Public Sub showdata()
'On Error GoTo perror1
     Dim StBuff As String
     If Not (rstStudmast.BOF And rstStudmast.EOF) Then
        PicName = rstStudmast("studpic")
        StudPic = LoadPicture(PicName)
        mStudno = rstStudmast("studno")
        StudNames = rstStudmast("studnames")
        stFileNo = rstStudmast("stfileno")
        regDate = rstStudmast("regdate")
        Sex = rstStudmast("Sex")
        DOB = rstStudmast("dob")
        Age = rstStudmast("age")
    End If
perror1:
End Sub

Private Sub cmbStudno_Click()
    StudFind
End Sub

Private Sub Form_Unload(Cancel As Integer)
    db1.Close
End Sub
Public Sub StudFind()
On Error GoTo PError
ClearData
   strpos = InStr(1, cmbStudNo, ",", 1)
   If strpos = 0 Then
     studBuff = "studno = '" & UCase(cmbStudNo) & "'"
     mStudno = UCase(cmbStudNo)
   Else
     studBuff = "studno = '" & Left(cmbStudNo, strpos - 1) & "'"
     mStudno = Left(cmbStudNo, strpos - 1)
   End If
    rstStudmast.FindFirst studBuff
    If rstStudmast.NOMATCH Then
       MsgBox "Mail with the Registration Number entered is not found.", vbCritical, "Enquiry"
       ClearData
    Else
        cmbStudNo = rstStudmast!StudNo & ", " & rstStudmast!StudNames
        showdata
    End If
PError:
End Sub

Public Sub ClearData()
On Error GoTo PError
        PicName = PicsPath & "picinit.bmp"
        StudPic = LoadPicture(PicName)
        StudNo.Caption = ""
        StudNames = ""
        stFileNo = ""
        regDate = ""
        Lclass = ""
        Sex = ""
        DOB = ""
        Age = ""
        mStudno = ""
        term1Score = ""
        Term1Grade = ""
        Term1Pos = ""
        term2score = ""
        Term2grade = ""
        Term2Pos = ""
        term3score = ""
        Term3Grade = ""
        Term3Pos = ""
        Term4Score = ""
        Term4Grade = ""
        Term4Pos = ""
        Term5Score = ""
        Term5Grade = ""
        Term5Pos = ""
        Term6Score = ""
        Term6Grade = ""
        Term6Pos = ""
PError:
End Sub

Private Sub PYear_gotfocus()
    PYear.SelStart = 0
    PYear.SelLength = Len(PYear)
    If Len(Trim(cmbStudNo)) = 0 Then
       If Not (rstStudmast.BOF And rstStudmast.EOF) Then
          cmbStudNo.ListIndex = 0
          strpos = InStr(1, cmbStudNo, ",", 1)
          mStudno = Left(cmbStudNo, strpos - 1)
       End If
    Else
       mStudno = Left(cmbStudNo, strpos - 1)
    End If
End Sub

