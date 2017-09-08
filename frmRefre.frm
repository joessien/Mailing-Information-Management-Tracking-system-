VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmRefre 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Human Resource"
   ClientHeight    =   6045
   ClientLeft      =   1410
   ClientTop       =   1575
   ClientWidth     =   7110
   ClipControls    =   0   'False
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6045
   ScaleWidth      =   7110
   Begin VB.ComboBox cmbAppid 
      Height          =   315
      Left            =   1620
      TabIndex        =   0
      Top             =   840
      Width           =   5355
   End
   Begin TabDlg.SSTab STRefDet 
      Height          =   3255
      Left            =   0
      TabIndex        =   10
      Top             =   1440
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5741
      _Version        =   393216
      Tab             =   2
      TabsPerRow      =   10
      TabHeight       =   520
      TabCaption(0)   =   "1st"
      TabPicture(0)   =   "frmRefre.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(4)=   "Referee1"
      Tab(0).Control(5)=   "RefTitle1"
      Tab(0).Control(6)=   "RefAddr1"
      Tab(0).Control(7)=   "R1Phone"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "2nd"
      TabPicture(1)   =   "frmRefre.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label8"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label9"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Referee2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "RefTitle2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "RefAddr2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "R2Phone"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "3rd"
      TabPicture(2)   =   "frmRefre.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label10"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label11"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label12"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label13"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Referee3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "RefTitle3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "RefAddr3"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "datGentab"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "R3Phone"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin VB.TextBox R3Phone 
         Height          =   285
         Left            =   1200
         TabIndex        =   34
         Top             =   2280
         Width           =   3975
      End
      Begin VB.TextBox R2Phone 
         Height          =   285
         Left            =   -73800
         TabIndex        =   32
         Top             =   2280
         Width           =   3975
      End
      Begin VB.TextBox R1Phone 
         Height          =   285
         Left            =   -73800
         TabIndex        =   30
         Top             =   2280
         Width           =   3975
      End
      Begin VB.Data datGentab 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   780
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.TextBox RefAddr3 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   1800
         Width           =   4035
      End
      Begin VB.TextBox RefTitle3 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Referee3 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox RefAddr2 
         Height          =   285
         Left            =   -73800
         TabIndex        =   6
         Top             =   1800
         Width           =   4035
      End
      Begin VB.TextBox RefTitle2 
         Height          =   285
         Left            =   -73800
         TabIndex        =   5
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Referee2 
         Height          =   285
         Left            =   -73800
         TabIndex        =   4
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox RefAddr1 
         Height          =   285
         Left            =   -73800
         TabIndex        =   3
         Top             =   1800
         Width           =   4035
      End
      Begin VB.TextBox RefTitle1 
         Height          =   285
         Left            =   -73800
         TabIndex        =   2
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Referee1 
         Height          =   285
         Left            =   -73800
         TabIndex        =   1
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label13 
         Caption         =   "Phone:"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Phone:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   33
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Phone:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   31
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Address:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Title:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Address:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   17
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Title:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   16
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Name:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   15
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Address:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   14
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Title:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   13
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   12
         Top             =   840
         Width           =   855
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   765
      Left            =   -120
      TabIndex        =   21
      Top             =   5160
      Width           =   7170
      _Version        =   65536
      _ExtentX        =   12647
      _ExtentY        =   1349
      _StockProps     =   15
      Caption         =   " "
      ForeColor       =   -2147483630
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
         Left            =   3960
         TabIndex        =   26
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   2820
         TabIndex        =   25
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton cmdUpd 
         Caption         =   "&Update"
         Height          =   375
         Left            =   1740
         TabIndex        =   24
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdBrow 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   300
         TabIndex        =   23
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   5460
         TabIndex        =   22
         Top             =   180
         Width           =   1155
      End
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   615
      Left            =   600
      TabIndex        =   28
      Top             =   0
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   1085
      _StockProps     =   15
      Caption         =   "Referee Details"
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
   Begin Threed.SSPanel SSPanel7 
      Height          =   615
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   15
      ForeColor       =   16576
      BackColor       =   11639171
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Font3D          =   2
   End
   Begin VB.Label Label1 
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
      Left            =   1200
      TabIndex        =   27
      Top             =   4920
      Width           =   4035
   End
   Begin VB.Label AppLbl 
      BackColor       =   &H80000004&
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
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmRefre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db1 As Database, wrktemp As Workspace
Dim GenData As Recordset, strbuff As String
Dim strpos As Integer, FldChk As Boolean
Dim GenClass As New DProLog

Private Sub cmbAppid_Click()
On Error GoTo PError
    strpos = InStr(1, cmbAppid, ",", 1)
    If strpos = 0 Then
      Exit Sub
    End If
    strbuff = "AppId = '" & Left(cmbAppid, strpos - 1) & "'"
    GenData.FindFirst strbuff
    If GenData.NOMATCH Then
       cmbAppid.ListIndex = 0
    Else
        showdata
    End If
PError:
End Sub

Private Sub Cmdclear_Click()
    ClearData
End Sub

Private Sub CmdUpd_Click()
On Error GoTo PError
If cmdUpd.Caption = "&Update" Then
   GenData.Edit
   fldEnab
   Referee1.SetFocus
   CmdDisab
   cmbAppid.Enabled = False
   cmdUpd.Caption = "&Cancel"
Else
   cmbAppid.Enabled = True
   cmbAppid.SetFocus
   cmdUpd.Caption = "&Update"
   ClearData
   CmdEnab
   fldDisab
End If
PError:
End Sub

Private Sub CmdSave_Click()
    Dim i As Integer
    On Error GoTo PError
    FldChk = True
    FldVal
    If FldChk = False Then
       Exit Sub
    End If
    GenData("referee1") = Referee1
    GenData("refaddr1") = RefAddr1
    GenData("reftitle1") = RefTitle1
    GenData("referee2") = Referee2
    GenData("refaddr2") = RefAddr2
    GenData("reftitle2") = RefTitle2
    GenData("referee3") = Referee3
    GenData("refaddr3") = RefAddr3
    GenData("reftitle3") = RefTitle3
    GenData("R1Phone") = R1Phone
    GenData("R2Phone") = R2Phone
    GenData("R3Phone") = R3Phone
    GenData.Update
    cmdUpd.Caption = "&Update"
    cmbAppid.Enabled = True
    CmdEnab
    fldDisab
PError0:
    Exit Sub
    
PError:
    MsgBox "Check the fields for invalid data", vbCritical, "Resume"
    On Error GoTo 0
    Resume PError0

End Sub

Private Sub Form_Load()
On Error GoTo PError
     GenClass.fleLogin mvUserid, "Accessed Referee Details", Date, Time
     Dim AppStatBuff As String
     Call FormCentreMDI(Me)
     Set wrktemp = DBEngine.Workspaces(0)
     Set db1 = wrktemp.OpenDatabase(DBpath, True)
     Set GenData = db1.OpenRecordset("Genrecs", dbOpenDynaset)
    ''__________________________________
    datGenTab.DatabaseName = DBpath
    datGenTab.RecordSource = "GenRecs"
    datGenTab.ReadOnly = False
    datGenTab.Exclusive = False
    datGenTab.Refresh
    Do While Not datGenTab.Recordset.EOF
       cmbAppid.AddItem datGenTab.Recordset.Appid + ", " + datGenTab.Recordset.FName + " " + datGenTab.Recordset.Surname
       datGenTab.Recordset.MoveNext
    Loop
    cmbAppid.ListIndex = 0
    ''__________________________________
    ClearData
    fldDisab
    Me.Caption = MVCoyname
PError:
End Sub

Private Sub CmdBrow_Click()
On Error GoTo PError
    mvSql = "Select referee1,refaddr1,reftitle1,referee2,refaddr2,reftitle2,referee3,refaddr3,reftitle3 From Genrecs"
    Set rst = db1.OpenRecordset(mvSql, dbOpenSnapshot)
    Set frmBrowse.datGeneral.Recordset = rst
    frmBrowse.Caption = Me.Caption & " - Records View"
    frmBrowse.Show
PError:
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Public Sub showdata()
  Dim AppStatBuff As String
  On Error GoTo SError
  If Not (GenData.EOF And GenData.BOF) Then
    Referee1 = GenData("referee1")
    RefAddr1 = GenData("refaddr1")
    RefTitle1 = GenData("reftitle1")
    Referee2 = GenData("referee2")
    RefAddr2 = GenData("refaddr2")
    RefTitle2 = GenData("reftitle2")
    Referee3 = GenData("referee3")
    RefAddr3 = GenData("refaddr3")
    RefTitle3 = GenData("reftitle3")
    R1Phone = GenData("R1Phone")
    R2Phone = GenData("R2Phone")
    R3Phone = GenData("R3Phone")
  End If
  Exit Sub
SError:
  ClearData
  
End Sub

Public Sub ClearData()
On Error GoTo PError
    Referee1 = "Nil"
    RefAddr1 = "Nil"
    RefTitle1 = "Nil"
    Referee2 = "Nil"
    RefAddr2 = "Nil"
    RefTitle2 = "Nil"
    Referee3 = "Nil"
    RefAddr3 = "Nil"
    RefTitle3 = "Nil"
    R1Phone = "Nil"
    R2Phone = "Nil"
    R3Phone = "Nil"
PError:
End Sub
Private Sub Form_Unload(Cancel As Integer)
    db1.Close
End Sub

Public Sub CmdEnab()
    cmdbrow.Enabled = True
    cmdsave.Enabled = False
    cmdClear.Enabled = True
End Sub

Public Sub CmdDisab()
    cmdbrow.Enabled = False
    cmdsave.Enabled = True
    cmdClear.Enabled = False
End Sub

Public Sub fldEnab()
    Referee1.Enabled = True
    RefAddr1.Enabled = True
    RefTitle1.Enabled = True
    Referee2.Enabled = True
    RefAddr2.Enabled = True
    RefTitle2.Enabled = True
    Referee3.Enabled = True
    RefAddr3.Enabled = True
    RefTitle3.Enabled = True
    R1Phone.Enabled = True
    R2Phone.Enabled = True
    R3Phone.Enabled = True

    
End Sub

Public Sub fldDisab()
    Referee1.Enabled = False
    RefAddr1.Enabled = False
    RefTitle1.Enabled = False
    Referee2.Enabled = False
    RefAddr2.Enabled = False
    RefTitle2.Enabled = False
    Referee3.Enabled = False
    RefAddr3.Enabled = False
    RefTitle3.Enabled = False
    R1Phone.Enabled = False
    R2Phone.Enabled = False
    R3Phone.Enabled = False

End Sub

Public Sub FldVal()
On Error GoTo PError
    If Len(Referee1) = 0 Or Len(RefAddr1) = 0 Or Len(RefTitle1) = 0 Then
        MsgBox "One or more Fields on the 1st Referee is Invalid", vbCritical, Me.Caption
        STRefDet.Tab = 0
        Referee1.SetFocus
        FldChk = False
        Exit Sub
    End If
    If Len(Referee2) = 0 Or Len(RefAddr2) = 0 Or Len(RefTitle2) = 0 Then
        MsgBox "One or more Fields on the 2nd Referee is Invalid", vbCritical, Me.Caption
        STRefDet.Tab = 1
        Referee2.SetFocus
        FldChk = False
        Exit Sub
    End If
    If Len(Referee3) = 0 Or Len(RefAddr3) = 0 Or Len(RefTitle3) = 0 Then
        MsgBox "One or more Fields on the 3rd Referee is Invalid", vbCritical, Me.Caption
        STRefDet.Tab = 2
        Referee3.SetFocus
        FldChk = False
        Exit Sub
    End If
PError:
End Sub

