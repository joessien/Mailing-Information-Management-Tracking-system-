Attribute VB_Name = "Module2"
Option Explicit

Public Sub FormCentreMDI(frmany As Form)
    frmany.Top = (FrmMtrack.ScaleHeight / 2) - (frmany.Height / 2)
    frmany.Left = (FrmMtrack.ScaleWidth / 2) - (frmany.Width / 2)
End Sub

Public Sub Main()
    'frmLogin.Show
End Sub

