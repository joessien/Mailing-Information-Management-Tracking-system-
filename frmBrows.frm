VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmBrowse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Tracking System Records Browser"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10575
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   10575
   Begin VB.CommandButton cmdok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Data datGeneral 
      Caption         =   "Data Buffer"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid grdGeneral 
      Align           =   1  'Align Top
      Bindings        =   "frmBrows.frx":0000
      Height          =   5775
      Left            =   0
      OleObjectBlob   =   "frmBrows.frx":0019
      TabIndex        =   0
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Call FormCentreMDI(Me)
End Sub

