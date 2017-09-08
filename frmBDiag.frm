VERSION 5.00
Begin VB.Form frmBDiag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   1650
   ClientTop       =   1050
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmBDiag.frx":0000
   ScaleHeight     =   6735
   ScaleWidth      =   8925
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   6240
      Width           =   1095
   End
End
Attribute VB_Name = "frmBDiag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdClose_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Caption = MVCoyname
End Sub
