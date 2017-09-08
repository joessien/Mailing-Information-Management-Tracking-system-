VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.MDIForm FrmMtrack 
   BackColor       =   &H8000000C&
   Caption         =   "Mail Tracking System Version 1.1 "
   ClientHeight    =   7545
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrptCtrl 
      Left            =   840
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowLeft      =   20
      WindowTop       =   40
      WindowWidth     =   750
      WindowHeight    =   500
      WindowState     =   2
   End
   Begin VB.Menu mnu10 
      Caption         =   "&Options"
      Begin VB.Menu mnu11 
         Caption         =   "Change &Password"
      End
      Begin VB.Menu mnu12 
         Caption         =   "&Log Out"
      End
      Begin VB.Menu mnu13 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnu20 
      Caption         =   "&Administration"
      Begin VB.Menu mnu21 
         Caption         =   "&Company Setup"
      End
      Begin VB.Menu mnu22 
         Caption         =   "-"
      End
      Begin VB.Menu mnu23 
         Caption         =   "&Admin User Access Setup"
      End
      Begin VB.Menu mnu24 
         Caption         =   "-"
      End
      Begin VB.Menu mnu25 
         Caption         =   "Activity Log Administration"
      End
   End
   Begin VB.Menu mnu30 
      Caption         =   "Mails &In"
      Begin VB.Menu mnu31 
         Caption         =   "Register New Mails"
      End
      Begin VB.Menu mnu32 
         Caption         =   "Update Mail Records"
      End
      Begin VB.Menu mnu33 
         Caption         =   "-"
      End
      Begin VB.Menu mnu34 
         Caption         =   "Remove Mail Record"
      End
   End
   Begin VB.Menu mnu40 
      Caption         =   "Mails &Out"
      Begin VB.Menu mnu41 
         Caption         =   "Forward Mails"
      End
      Begin VB.Menu mnu42 
         Caption         =   "Acknowledge Delivery "
      End
      Begin VB.Menu mnu43 
         Caption         =   "-"
      End
      Begin VB.Menu mnu44 
         Caption         =   "Domicile Mail"
      End
   End
   Begin VB.Menu mnu50 
      Caption         =   "&Reports"
      Begin VB.Menu mnu51 
         Caption         =   "Outgoing Mail Collection Confirmation slip"
      End
      Begin VB.Menu mnu53 
         Caption         =   "-"
      End
      Begin VB.Menu mnu54 
         Caption         =   "Pending Incoming Mails"
      End
      Begin VB.Menu mnu58 
         Caption         =   "In Mails on Hold"
      End
      Begin VB.Menu mnu54a 
         Caption         =   "-"
      End
      Begin VB.Menu mnu55 
         Caption         =   "Dispatched Outgoing Mails"
      End
      Begin VB.Menu mnu55a 
         Caption         =   "Received Outgoing Mails"
      End
      Begin VB.Menu mnu55b 
         Caption         =   "-"
      End
      Begin VB.Menu mnu56 
         Caption         =   "Deleted Mails"
      End
      Begin VB.Menu mnu57 
         Caption         =   "Domiciled Mails"
      End
      Begin VB.Menu mnu59 
         Caption         =   "-"
      End
      Begin VB.Menu mnu5a 
         Caption         =   "Business Diagram"
      End
   End
   Begin VB.Menu mnu60 
      Caption         =   "&Search Mails"
      Begin VB.Menu mnu61 
         Caption         =   "Incoming and Outgoing Mails Tracking"
      End
      Begin VB.Menu mnu62 
         Caption         =   "Domiciled and Removed Mails Tracking"
      End
      Begin VB.Menu mnu64 
         Caption         =   "-"
      End
      Begin VB.Menu mnu65 
         Caption         =   "Sorted Mail Listings"
      End
   End
   Begin VB.Menu mnu70 
      Caption         =   "Re&ferences"
      Begin VB.Menu mnu71 
         Caption         =   "Faculty / Division Definition"
      End
      Begin VB.Menu mnu72 
         Caption         =   "Academic Department Definition"
      End
      Begin VB.Menu mnu73 
         Caption         =   "Business Unit Definition"
      End
      Begin VB.Menu mnu74 
         Caption         =   "Business Unit Receivers"
      End
      Begin VB.Menu mnu75 
         Caption         =   "Subject Group Definition"
      End
      Begin VB.Menu mnu76 
         Caption         =   "Subclassification Definition"
      End
      Begin VB.Menu mnu77 
         Caption         =   "Mail Priority Definition"
      End
      Begin VB.Menu mnu78 
         Caption         =   "Status Flags"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu79 
         Caption         =   "State Definition"
      End
      Begin VB.Menu mnu7a 
         Caption         =   "Designations"
      End
   End
   Begin VB.Menu mnu90 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "FrmMtrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As Database, wrktemp As Workspace
Dim rstBinMailQry, rstDomMailQry As Recordset
Dim rstFinMailQry, rstInMailQry As Recordset
Dim rstLMailqry, rstOutMailQry As Recordset

Private Sub mmnu92_Click()
   FrmLectRecs.Show
End Sub

Private Sub mnu01_Click()
   FrmSetup.Show
End Sub

Private Sub mnu02_Click()
    frmuser.Show
End Sub

Private Sub mnu03_Click()
    frmUserLect.Show
End Sub

Private Sub mnu04_Click()
   FrmMScale.Show
End Sub

Private Sub mnu0c_Click()
  frmAlog.Show
End Sub

Private Sub MDIForm_Load()
     Set wrktemp = DBEngine.Workspaces(0)
     Set db1 = wrktemp.OpenDatabase(DBpath, True)
     Set rstBinMailQry = db1.OpenRecordset("BinMailQuery", dbOpenDynaset)
     Set rstDomMailQry = db1.OpenRecordset("DomMailQuery", dbOpenDynaset)
     Set rstFinMailQry = db1.OpenRecordset("FinMailQuery", dbOpenDynaset)
     Set rstInMailQry = db1.OpenRecordset("LMailQuery", dbOpenDynaset)
     Set rstLMailqry = db1.OpenRecordset("LMailQuery", dbOpenDynaset)
     Set rstOutMailQry = db1.OpenRecordset("OutMailQuery", dbOpenDynaset)
End Sub

Private Sub mnu11_Click()
    frmCPasswd.Show
End Sub

Private Sub mnu12_Click()
    With FrmMtrack
        '________________________clear access rights
        ' ----------User setup
        .mnu10.Enabled = False
        .mnu11.Enabled = False
        .mnu12.Enabled = False
        .mnu13.Enabled = False
        ' ----------company setup
        .mnu20.Enabled = False
        .mnu21.Enabled = False
        '.mnu22.Enabled = False
        .mnu23.Enabled = False
        '.mnu24.Enabled = False
        .mnu25.Enabled = False
        ' ----------Incoming Mail
        .mnu30.Enabled = False
        .mnu31.Enabled = False
        .mnu32.Enabled = False
        '.mnu33.Enabled = False
        .mnu34.Enabled = False
        ' ----------Out going Mail
        .mnu40.Enabled = False
        .mnu41.Enabled = False
        .mnu42.Enabled = False
        '.mnu43.Enabled = False
        .mnu44.Enabled = False
        ' ----------Reports Mail
        .mnu50.Enabled = False
        ' ----------Searching criteria
        .mnu60.Enabled = False
        .mnu61.Enabled = False
        .mnu62.Enabled = False
        .mnu65.Enabled = False
        ' ----------Referencing
        .mnu70.Enabled = False
        .mnu71.Enabled = False
        .mnu72.Enabled = False
        .mnu73.Enabled = False
        .mnu74.Enabled = False
        .mnu75.Enabled = False
        .mnu76.Enabled = False
        .mnu77.Enabled = False
        .mnu78.Enabled = False
        .mnu79.Enabled = False

   End With
     mnu1 = False
     mnu2 = False
     mnu3 = False
     mnu4 = False
     mnu5 = False
     mnu6 = False
     mnu7 = False
     mnu8 = False
     mnu9 = False
     mnu10 = False
     mnu11 = False
     mnu12 = False
     Unload Me
    frmLogin.Show

End Sub

Private Sub mnu13_Click()
    End
End Sub

Private Sub mnu21_Click()
    FrmSetup.Show
End Sub

Private Sub mnu22_Click()
    FrmRelTn.Show
End Sub

Private Sub mnu23_Click()
    frmuser.Show
End Sub

Private Sub mnu25_Click()
    frmAlog.Show
End Sub

Private Sub mnu26_Click()
    frmWDRRecs.Show
End Sub

Private Sub mnu27_Click()
    FrmWDRNotes.Show
End Sub
Private Sub mnu2a_Click()
    FrmAutoB.Show
End Sub

Private Sub mnu31_Click()
   frmMailIn.Show
End Sub

Private Sub mnu32_Click()
   frmMailInUpd.Show
End Sub

Private Sub mnu34_Click()
   frmMailInRem.Show
End Sub

Private Sub mnu41_Click()
    frmMailOut.Show
End Sub


Private Sub mnu42_Click()
   frmMailOutUpd.Show
End Sub

Private Sub mnu44_Click()
   frmMailDom.Show
End Sub

Private Sub mnu51_Click()
   rstOutMailQry.Requery
   CrptCtrl.ReportFileName = Rptpath & "outmailcon.rpt"
   'CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
End Sub

Private Sub mnu54_Click()
   rstInMailQry.Requery
   CrptCtrl.ReportFileName = Rptpath & "allinmail.rpt"
   'CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
End Sub

Private Sub mnu55_Click()
   rstOutMailQry.Requery
   CrptCtrl.ReportFileName = Rptpath & "ndoutmail.rpt"
   'CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
End Sub

Private Sub mnu55a_Click()
   rstOutMailQry.Requery
   CrptCtrl.ReportFileName = Rptpath & "doutmail.rpt"
   'CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
End Sub

Private Sub mnu56_Click()
   rstBinMailQry.Requery
   CrptCtrl.ReportFileName = Rptpath & "alldelmail.rpt"
   'CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
End Sub

Private Sub mnu57_Click()
   rstDomMailQry.Requery
   CrptCtrl.ReportFileName = Rptpath & "alldommail.rpt"
   'CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
End Sub

Private Sub mnu58_Click()
   rstInMailQry.Requery
   CrptCtrl.ReportFileName = Rptpath & "allheldmail.rpt"
   'CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
End Sub

Private Sub mnu5a_Click()
   frmBDiag.Show
End Sub

Private Sub mnu61_Click()
  FrmLquery.Show
End Sub

Private Sub mnu63_Click()
   FrmSubclass.Show
End Sub

Private Sub mnu62_Click()
   FrmDquery.Show
End Sub

Private Sub mnu64_Click()
   FrmSubj.Show
End Sub

Private Sub mnu64a_Click()
   FrmTerms.Show
End Sub

Private Sub mnu67_Click()
   FrmDept.Show
End Sub

Private Sub mnu68_Click()
   FrmQual.Show
End Sub

Private Sub mnu69_Click()
   FrmcmbFind.Show
End Sub

Private Sub mnu65_Click()
   FrmListings.Show
End Sub

Private Sub mnu71_Click()
   FrmFaculty.Show
End Sub

Private Sub mnu72_Click()
   FrmDept.Show
End Sub

Private Sub mnu72a_Click()
  FrmTalents.Show
End Sub

Private Sub mnu73_Click()
   FrmBUnitD.Show
End Sub

Private Sub mnu74_Click()
   FrmBUnitR.Show
End Sub

Private Sub mnu75_Click()
   FrmSubjG.Show
End Sub

Private Sub mnu75a_Click()
   FrmElect.Show
End Sub

Private Sub mnu75b_Click()
   FrmHouse.Show
End Sub

Private Sub mnu81_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "rsltslip.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu82_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "overperf.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu83_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "SubjPerf.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu84_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "Examhist.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu85_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "Cumrpt.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu86_Click()
On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "SubList.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu86a_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "studslist.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu86b_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "medhist.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu87_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "sysNotes.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu88_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "TSubjClass.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu89_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "Exexams.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu8ab_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "gradslip.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub


Private Sub mnu8c_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "PASSLIST.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu8d_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "FAILLIST.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu8da_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "evaldetails.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu8f_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "studrecsn.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu8g_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "gradlist.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:

End Sub

Private Sub mnu8h_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "wdrlist.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:

End Sub

Private Sub mnu8j_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "studnotes.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu8k_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "GRADNOTES.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:

End Sub

Private Sub mnu8l_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "WDRNOTES.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu8n_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "stafflist.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:

End Sub

Private Sub mnu8o_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "tsubjclass.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:

End Sub

Private Sub mnu8p_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "staffrespo.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:

End Sub

Private Sub mnu8q_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "teachperf.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu8r_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "subjteacheryes.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu8s_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "subjteacherno.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu8t_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "subjteachclass.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:

End Sub

Private Sub mnu91_Click()
   frmLRecsN.Show
End Sub

Private Sub mnu92_Click()
   frmLRecsU.Show
End Sub

Private Sub mnu93_Click()
    FrmSResrc.Show
End Sub

Private Sub mnu93a_Click()
   FrmClass.SSPanel5.Caption = "Class Room Review"
   FrmClass.SSAction.Enabled = False
   FrmClass.SSPelect.Enabled = False
   FrmClass.Show
End Sub

Private Sub mnuA1_Click()
   frmAppDef.Show
End Sub

Private Sub mnuA2_Click()
  frmImport.Show
End Sub

Private Sub mnuAI_Click()
   FrmAIPrt.Show
End Sub

Private Sub mnuBI1_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "EntStat.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnuEXRPrt_Click()
   FrmRsltEXPrt.Show
End Sub

Private Sub mnuGoP_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "gradgrpoperf.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnugoS_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "gradgrpsperf.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnuOP_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "cgrpoperf.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnuOS_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "cgrpsperf.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnuREnq_Click()
    FrmRsltEnq.Show
End Sub

Private Sub mnuRPrt_Click()
   FrmRsltPrt.Show
End Sub

Private Sub mnust1_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "stotclass.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnust2_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "stotglobal.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnust5_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "femalestats.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnust6_Click()
   On Error GoTo PError
   CrptCtrl.ReportFileName = Rptpath & "malestats.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnuStEnq_Click()
   FrmStudEnq.Show
End Sub

Private Sub mnu76_Click()
    FrmSubclass.Show
End Sub

Private Sub mnu77_Click()
    FrmPriority.Show
End Sub

Private Sub mnu79_Click()
    FrmState.Show
End Sub

Private Sub mnu7a_Click()
   FrmTitle.Show
End Sub
