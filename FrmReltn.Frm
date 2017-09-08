VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmRelTn 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datapro"
   ClientHeight    =   6495
   ClientLeft      =   750
   ClientTop       =   600
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6495
   ScaleWidth      =   8160
   Begin Threed.SSPanel SSPanel1 
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8175
      _Version        =   65536
      _ExtentX        =   14420
      _ExtentY        =   1085
      _StockProps     =   15
      Caption         =   "Student Relationships"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   4
   End
   Begin VB.ComboBox cmbStudNo 
      Height          =   315
      Left            =   2520
      TabIndex        =   5
      Top             =   720
      Width           =   4635
   End
   Begin VB.Data datGenTab 
      Caption         =   "General Data"
      Connect         =   "Access"
      DatabaseName    =   "C:\vb\Kdhpims\kdhpims.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "GenRecs"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1620
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   765
      Left            =   0
      TabIndex        =   1
      Top             =   5760
      Width           =   8205
      _Version        =   65536
      _ExtentX        =   14473
      _ExtentY        =   1349
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
      BevelWidth      =   2
      BevelInner      =   1
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4260
         TabIndex        =   8
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton cmdUpd 
         Caption         =   "&Update"
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdBrow 
         Caption         =   "&List"
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   5580
         TabIndex        =   3
         Top             =   180
         Width           =   1155
      End
      Begin Threed.SSCommand SSCommand11 
         Height          =   30
         Left            =   450
         TabIndex        =   2
         Top             =   4455
         Width           =   30
         _Version        =   65536
         _ExtentX        =   53
         _ExtentY        =   53
         _StockProps     =   78
         Caption         =   "SSCommand11"
      End
   End
   Begin TabDlg.SSTab STReltn 
      Height          =   3765
      Left            =   0
      TabIndex        =   11
      Top             =   1560
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   6641
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabsPerRow      =   8
      TabHeight       =   529
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
      TabCaption(0)   =   "Guardian"
      TabPicture(0)   =   "FrmReltn.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "GDPhone"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "GDEmail"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "guardiaN"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "GdAddr1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "GDReln"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Father"
      TabPicture(1)   =   "FrmReltn.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label10"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "faALive"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Data1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "FaOccup"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "FaAddr1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "FaName"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "FAPhone"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Mother"
      TabPicture(2)   =   "FrmReltn.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label12"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label13"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label14"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label15"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "maAlive"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "maoCCUP"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "maaDDR1"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "mAnAME"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "mapHONE"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin VB.TextBox mapHONE 
         Height          =   300
         Left            =   -73320
         TabIndex        =   42
         Top             =   2880
         Width           =   4125
      End
      Begin VB.TextBox mAnAME 
         Height          =   300
         Left            =   -73320
         TabIndex        =   38
         Top             =   1020
         Width           =   4155
      End
      Begin VB.TextBox maaDDR1 
         Height          =   780
         Left            =   -73320
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Top             =   1920
         Width           =   4155
      End
      Begin VB.TextBox maoCCUP 
         Height          =   300
         Left            =   -73320
         TabIndex        =   36
         Top             =   1440
         Width           =   4125
      End
      Begin VB.TextBox FAPhone 
         Height          =   300
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   34
         Top             =   2880
         Width           =   4125
      End
      Begin VB.TextBox GDPhone 
         Height          =   285
         Left            =   -72960
         MaxLength       =   40
         TabIndex        =   31
         Top             =   2760
         Width           =   4935
      End
      Begin VB.TextBox GDEmail 
         Height          =   285
         Left            =   -72960
         MaxLength       =   40
         TabIndex        =   30
         Top             =   3240
         Width           =   4935
      End
      Begin VB.TextBox FaName 
         Height          =   300
         Left            =   1680
         TabIndex        =   20
         Top             =   1020
         Width           =   4155
      End
      Begin VB.TextBox FaAddr1 
         Height          =   780
         Left            =   1680
         MaxLength       =   200
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   1920
         Width           =   4155
      End
      Begin VB.TextBox FaOccup 
         Height          =   300
         Left            =   1680
         TabIndex        =   18
         Top             =   1440
         Width           =   4125
      End
      Begin VB.TextBox guardiaN 
         Height          =   285
         Left            =   -72960
         TabIndex        =   14
         Top             =   960
         Width           =   4935
      End
      Begin VB.TextBox GdAddr1 
         Height          =   645
         Left            =   -72960
         MaxLength       =   200
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1920
         Width           =   4935
      End
      Begin VB.TextBox GDReln 
         Height          =   285
         Left            =   -72960
         TabIndex        =   12
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4800
         Width           =   1455
      End
      Begin Threed.SSFrame faALive 
         Height          =   1275
         Left            =   6240
         TabIndex        =   21
         Top             =   960
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   2249
         _StockProps     =   14
         Caption         =   "Living"
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
         Begin Threed.SSOption faStat 
            Height          =   240
            Index           =   1
            Left            =   210
            TabIndex        =   22
            Top             =   810
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "No"
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
         Begin Threed.SSOption faStat 
            Height          =   285
            Index           =   0
            Left            =   195
            TabIndex        =   23
            Top             =   360
            Width           =   660
            _Version        =   65536
            _ExtentX        =   1164
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Yes"
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
      Begin Threed.SSFrame maAlive 
         Height          =   1320
         Left            =   -68760
         TabIndex        =   27
         Top             =   960
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   2328
         _StockProps     =   14
         Caption         =   "Living"
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
         Begin Threed.SSOption maStat 
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "No"
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
         Begin Threed.SSOption maStat 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   780
            _Version        =   65536
            _ExtentX        =   1376
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Yes"
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
      Begin VB.Label Label15 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
         Height          =   270
         Left            =   -74640
         TabIndex        =   43
         Top             =   2880
         Width           =   1200
      End
      Begin VB.Label Label14 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   300
         Left            =   -74640
         TabIndex        =   41
         Top             =   1020
         Width           =   1185
      End
      Begin VB.Label Label13 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   285
         Left            =   -74640
         TabIndex        =   40
         Top             =   1920
         Width           =   1185
      End
      Begin VB.Label Label12 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation:"
         Height          =   270
         Left            =   -74640
         TabIndex        =   39
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label Label10 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
         Height          =   270
         Left            =   360
         TabIndex        =   35
         Top             =   2880
         Width           =   1200
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         Height          =   270
         Left            =   -74040
         TabIndex        =   33
         Top             =   3240
         Width           =   780
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
         Height          =   270
         Left            =   -74040
         TabIndex        =   32
         Top             =   2760
         Width           =   780
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   300
         Left            =   360
         TabIndex        =   26
         Top             =   1020
         Width           =   1185
      End
      Begin VB.Label Label9 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   285
         Left            =   360
         TabIndex        =   25
         Top             =   1920
         Width           =   1185
      End
      Begin VB.Label Label11 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation:"
         Height          =   270
         Left            =   360
         TabIndex        =   24
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   300
         Left            =   -74040
         TabIndex        =   17
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Relationship:"
         Height          =   285
         Left            =   -74040
         TabIndex        =   16
         Top             =   1440
         Width           =   1170
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   270
         Left            =   -74040
         TabIndex        =   15
         Top             =   1920
         Width           =   780
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FF80&
      Caption         =   "Leave fields as Nil if not applicable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1920
      TabIndex        =   9
      Top             =   5520
      Width           =   3195
   End
   Begin VB.Label tesxt1 
      BackStyle       =   0  'Transparent
      Caption         =   "Student No:"
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
      Height          =   300
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1740
   End
End
Attribute VB_Name = "FrmRelTn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db6 As Database, wrktemp As Workspace
Dim rstStudmast As Recordset, strbuff As String
Dim strpos As Integer, FldChk As Boolean
Dim mMaStat As String, mFaStat As String
Dim GenClass As New DProLog, StudBuff As String

Private Sub cmbStudNo_Click()
    strpos = InStr(1, cmbStudNo, ",", 1)
    If strpos = 0 Then
      Exit Sub
    End If
    strbuff = "StudNo = '" & Left(cmbStudNo, strpos - 1) & "'"
    rstStudmast.FindFirst strbuff
    If rstStudmast.NOMATCH Then
       MsgBox "Student Record not Found", vbCritical, "Invalid Data"
       cmbStudNo.ListIndex = 0
    Else
        showdata
    End If

End Sub

Private Sub cmbStudNo_LostFocus()
 'On Error GoTo PError
    strpos = InStr(1, cmbStudNo, ",", 1)
    If strpos = 0 Then
        StudBuff = cmbStudNo
    Else
        StudBuff = Left(cmbStudNo, strpos - 1)
    End If
    StudBuff = "StudNo = '" & StudBuff & "'"
    rstStudmast.FindFirst StudBuff
    If rstStudmast.NOMATCH Then
        MsgBox "Student Record not Found", vbCritical, "Invalid Data"
    Else
        cmbStudNo = rstStudmast!StudNo & ", " & rstStudmast!StudNames
        showdata
    End If
perror:
End Sub

Private Sub Cmdclear_Click()
    ClearData
End Sub

Private Sub CmdUpd_Click()
If cmdUpd.Caption = "&Update" Then
   rstStudmast.Edit
   FldEnab
   guardiaN.SetFocus
   CmdDisab
   cmbStudNo.Enabled = False
   cmdUpd.Caption = "&Cancel"
Else
   cmbStudNo.Enabled = True
   cmbStudNo.SetFocus
   cmdUpd.Caption = "&Update"
   ClearData
   CmdEnab
   FldDisab
End If
End Sub

Private Sub CmdSave_Click()
    Dim i As Integer
   ' On Error GoTo PError
    FldChk = True
    FldVal
    If FldChk = False Then
       Exit Sub
    End If
    rstStudmast("Guardian") = guardiaN
    rstStudmast("GdAddr1") = GdAddr1
    rstStudmast("GDReln") = GDReln
    rstStudmast("GDPhone") = GDPhone
    rstStudmast("GDEmail") = GDEmail
    If faStat(0).Value = True Then
      rstStudmast("faStat") = True
    Else
      rstStudmast("faStat") = False
    End If
    rstStudmast("FaName") = FaName
    rstStudmast("FaAddr1") = FaAddr1
    rstStudmast("FAPhone") = FAPhone
    rstStudmast("FaOccup") = FaOccup
    If maStat(0).Value = True Then
      rstStudmast("maStat") = True
    Else
      rstStudmast("maStat") = False
    End If
    rstStudmast("mAnAME") = mAnAME
    rstStudmast("maoCCUP") = maoCCUP
    rstStudmast("maaDDR1") = maaDDR1
    rstStudmast("mapHONE") = mapHONE
    rstStudmast.Update
    cmdUpd.Caption = "&Update"
    cmbStudNo.Enabled = True
    CmdEnab
    FldDisab
PError0:
    Exit Sub
perror:
    MsgBox "Check the fields for invalid data", vbCritical, "Resume"
    On Error GoTo 0
    Resume PError0

End Sub

Private Sub fastat_Click(Index As Integer, Value As Integer)
    Select Case Index
        Case 0
            mFaStat = "Yes"
        Case 1
            mFaStat = "No"
    End Select
End Sub

Private Sub Form_Load()
     GenClass.fleLogin mvUserid, "Accessed Student Relationships", Date, Time
     Dim AppStatBuff As String
     Call FormCentreMDI(Me)
     Set wrktemp = DBEngine.Workspaces(0)
     Set db6 = wrktemp.OpenDatabase(DBpath, True)
     Set rstStudmast = db6.OpenRecordset("studmast", dbOpenDynaset)
    ''__________________________________
    datGenTab.DatabaseName = DBpath
    datGenTab.RecordSource = "studmast"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       cmbStudNo.AddItem datGenTab.Recordset.StudNo + ", " + datGenTab.Recordset.StudNames
       datGenTab.Recordset.MoveNext
    Loop
    If Not (datGenTab.Recordset.BOF And datGenTab.Recordset.EOF) Then
        cmbStudNo.ListIndex = 0
    End If
    ''__________________________________
    ClearData
    FldDisab
End Sub

Private Sub CmdBrow_Click()
On Error GoTo perror
    mvSql = "Select guardian, gdphone, faname,faphone,maname,maphone From studmast "
    Set rst = db6.OpenRecordset(mvSql, dbOpenSnapshot)
    If (rst.BOF And rst.EOF) Then
         MsgBox "No students recod found", vbExclamation, "Relationships"
         Exit Sub
     Else
         Set frmBrowse.datGeneral.Recordset = rst
         frmBrowse.grdGeneral.Caption = "Student Relationships"
         frmBrowse.Caption = Me.Caption
         Set Col0 = frmBrowse.grdGeneral.Columns(0)
         Set Col1 = frmBrowse.grdGeneral.Columns(1)
         Set Col2 = frmBrowse.grdGeneral.Columns(2)
         Set Col3 = frmBrowse.grdGeneral.Columns(3)
         Set Col4 = frmBrowse.grdGeneral.Columns(4)
         Set Col5 = frmBrowse.grdGeneral.Columns(5)
         frmBrowse.grdGeneral.Columns(0).Width = 1800
         frmBrowse.grdGeneral.Columns(1).Width = 2000
         frmBrowse.grdGeneral.Columns(2).Width = 1800
         frmBrowse.grdGeneral.Columns(3).Width = 2400
         frmBrowse.grdGeneral.Columns(4).Width = 1800
         frmBrowse.grdGeneral.Columns(5).Width = 2400
         Col0.Caption = "Guardian Name"
         Col1.Caption = "Guardian's Phone"
         Col2.Caption = "Father's Name"
         Col3.Caption = "Father's Phone"
         Col4.Caption = "Mother's Name"
         Col5.Caption = "Mother's Phone"
         Col6.Caption = "Sex"
   End If
   frmBrowse.Show
perror: End Sub

Private Sub cmdExit_Click()
    
    Unload Me
End Sub

Public Sub showdata()
  'On Error GoTo SError
  If Not (rstStudmast.EOF And rstStudmast.BOF) Then
    guardiaN = rstStudmast("Guardian")
    GdAddr1 = rstStudmast("GdAddr1")
    GDReln = rstStudmast("GDReln")
    GDPhone = rstStudmast("GDPhone")
    GDEmail = rstStudmast("GDEmail")
    FaName = rstStudmast("FaName")
    FaAddr1 = rstStudmast("FaAddr1")
    FAPhone = rstStudmast("FAPhone")
    FaOccup = rstStudmast("FaOccup")
    mAnAME = rstStudmast("mAnAME")
    maoCCUP = rstStudmast("maoCCUP")
    maaDDR1 = rstStudmast("maaDDR1")
    mapHONE = rstStudmast("mapHONE")
    
    If rstStudmast("fastat") = True Then
       faStat(0).Value = True
    Else
       faStat(1).Value = False
    End If
    If rstStudmast("maStat") = True Then
       maStat(0).Value = True
    Else
       maStat(1).Value = False
    End If
  End If
Exit Sub
SError:
  ClearData
  
End Sub

Public Sub ClearData()
    guardiaN = "NIL"
    GdAddr1 = "NIL"
    GDReln = "NIL"
    GDPhone = "NIL"
    GDEmail = "NIL"
    FaName = "NIL"
    FaAddr1 = "NIL"
    FAPhone = "NIL"
    FaOccup = "NIL"
    mAnAME = "NIL"
    maoCCUP = "NIL"
    maaDDR1 = "NIL"
    mMaStat = "Yes"
    maStat(0).Value = True
    faStat(0).Value = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    db6.Close
End Sub

Public Sub CmdEnab()
    cmdbrow.Enabled = True
    cmdSave.Enabled = False
    cmdClear.Enabled = True
End Sub

Public Sub CmdDisab()
    cmdbrow.Enabled = False
    cmdSave.Enabled = True
    cmdClear.Enabled = False
End Sub

Public Sub FldEnab()
    faALive.Enabled = True
    maAlive.Enabled = True
    guardiaN.Enabled = True
    GdAddr1.Enabled = True
    GDReln.Enabled = True
    GDPhone.Enabled = True
    GDEmail.Enabled = True
    faALive.Enabled = True
    FaName.Enabled = True
    FaAddr1.Enabled = True
    FAPhone.Enabled = True
    FaOccup.Enabled = True
    maAlive.Enabled = True
    mAnAME.Enabled = True
    maoCCUP.Enabled = True
    maaDDR1.Enabled = True
    mapHONE.Enabled = True
    
End Sub

Public Sub FldDisab()
    faALive.Enabled = False
    maAlive.Enabled = False
    guardiaN.Enabled = False
    GdAddr1.Enabled = False
    GDReln.Enabled = False
    GDPhone.Enabled = False
    GDEmail.Enabled = False
    FaName.Enabled = False
    FaAddr1.Enabled = False
    FAPhone.Enabled = False
    FaOccup.Enabled = False
    mAnAME.Enabled = False
    maoCCUP.Enabled = False
    maaDDR1.Enabled = False
    mapHONE.Enabled = False
    faALive.Enabled = False
    maAlive.Enabled = False
    
End Sub

Public Sub FldVal()
    If Len(guardiaN) = 0 Or Len(GdAddr1) = 0 Or Len(GDReln) = 0 Or Len(GDPhone) = 0 Or Len(GDEmail) = 0 Then
        MsgBox "One or more Fields on Guardian are Invalid", vbCritical, Me.Caption
        STReltn.Tab = 0
        guardiaN.SetFocus
        FldChk = False
        Exit Sub
    End If
    If Len(FaName) = 0 Or Len(mFaStat) = 0 Or Len(FaOccup) = 0 Or Len(FaAddr1) = 0 Or Len(FAPhone) = 0 Then
        MsgBox "One or more Fields on Father are Invalid", vbCritical, Me.Caption
        STReltn.Tab = 1
        FaName.SetFocus
        FldChk = False
        Exit Sub
    End If
    If Len(mAnAME) = 0 Or Len(mMaStat) = 0 Or Len(maoCCUP) = 0 Or Len(maaDDR1) = 0 Or Len(mapHONE) = 0 Then
        MsgBox "One or more Fields on Mother are Invalid", vbCritical, Me.Caption
        STReltn.Tab = 2
        mAnAME.SetFocus
        FldChk = False
        Exit Sub
    End If
End Sub

Private Sub Label18_Click()
End Sub

Private Sub maStat_Click(Index As Integer, Value As Integer)
    Select Case Index
        Case 0
            mMaStat = "Yes"
        Case 1
            mMaStat = "No"
    End Select
End Sub
