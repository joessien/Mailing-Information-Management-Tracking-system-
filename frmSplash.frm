VERSION 5.00
Begin VB.Form frmsplash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spectrum Software Systems"
   ClientHeight    =   4635
   ClientLeft      =   2730
   ClientTop       =   2235
   ClientWidth     =   6150
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmSplash.frx":030A
   ScaleHeight     =   4635
   ScaleWidth      =   6150
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdCont 
      BackColor       =   &H8000000A&
      Caption         =   "&Continue"
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   1095
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCont_Click()
    cmdCont.Visible = False
    frmLogin.Show
    'Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    cmdCont.Visible = False
    frmLogin.Show
    'Unload Me
End Sub

Private Sub Form_Load()
    frmsplash.Top = 1300
    frmsplash.Left = 3000
End Sub

Private Sub Timer1_Timer()
    cmdCont.Visible = False
    frmLogin.Show
    Unload Me
End Sub
