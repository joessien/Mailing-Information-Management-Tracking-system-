VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.MDIForm FrmDataPro 
   BackColor       =   &H8000000C&
   Caption         =   "Students Information Management Expert - Standard Edition Version 2.1 "
   ClientHeight    =   7545
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   14235
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrptCtrl 
      Left            =   6240
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowLeft      =   0
      WindowTop       =   40
      WindowWidth     =   720
      WindowHeight    =   410
      WindowTitle     =   "PeopleWare Information Management System - Report"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu mnu10 
      Caption         =   "&Options"
      Begin VB.Menu mnu13 
         Caption         =   "Change &Password"
      End
      Begin VB.Menu mnu14 
         Caption         =   "&Log Out"
      End
      Begin VB.Menu mnu15 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnu00 
      Caption         =   "&Application Administration"
      Begin VB.Menu mnu01 
         Caption         =   "&Company Setup"
      End
      Begin VB.Menu mnu0a 
         Caption         =   "-"
      End
      Begin VB.Menu mnu02 
         Caption         =   "&Admin User Access Setup"
      End
      Begin VB.Menu mnu03 
         Caption         =   "&Teachers Access Setup"
      End
      Begin VB.Menu mnu0b 
         Caption         =   "-"
      End
      Begin VB.Menu mnu0c 
         Caption         =   "Activity Log Administration"
      End
   End
   Begin VB.Menu mnu40 
      Caption         =   "&Class Management"
      Begin VB.Menu mnu62a 
         Caption         =   "Class Room Administration"
      End
      Begin VB.Menu mnu04a 
         Caption         =   "-"
      End
      Begin VB.Menu mnu41 
         Caption         =   "Subjects Allocation by School Level"
      End
      Begin VB.Menu mnu41a 
         Caption         =   "Subjects Allocation by Class Group"
      End
      Begin VB.Menu mnu41b 
         Caption         =   "Subjects Allocation by Class Unit"
      End
      Begin VB.Menu mnu04b 
         Caption         =   "-"
      End
      Begin VB.Menu mnu04 
         Caption         =   "Scoring Percentages "
      End
      Begin VB.Menu mnu4c 
         Caption         =   "-"
      End
      Begin VB.Menu mnu42 
         Caption         =   "Students Class Transfer"
      End
      Begin VB.Menu mnu48 
         Caption         =   "Student Step Transfer"
      End
      Begin VB.Menu mnu49 
         Caption         =   "Departmental Transfer"
      End
   End
   Begin VB.Menu mnu90 
      Caption         =   "&Teachers Management"
      Begin VB.Menu mnu91 
         Caption         =   "New Employments"
      End
      Begin VB.Menu mnu92 
         Caption         =   "Update Employment Records"
      End
      Begin VB.Menu mnu92a 
         Caption         =   "-"
      End
      Begin VB.Menu mnu93 
         Caption         =   "Staff Responsibility Allocations"
      End
      Begin VB.Menu mnu93a 
         Caption         =   "Class Room Review"
      End
      Begin VB.Menu mnu94 
         Caption         =   "Prep Supervision"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu20 
      Caption         =   "&Student Administration"
      Begin VB.Menu mnu21 
         Caption         =   "Student Admissions"
      End
      Begin VB.Menu mnu22 
         Caption         =   "&Relationship Details"
      End
      Begin VB.Menu mnu22a 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu23a 
         Caption         =   "-"
      End
      Begin VB.Menu mnu24 
         Caption         =   "&Recreational  Activities"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu25 
         Caption         =   "&Student Notes"
      End
      Begin VB.Menu mnu23 
         Caption         =   "&Medical Information"
      End
      Begin VB.Menu mnu25a 
         Caption         =   "-"
      End
      Begin VB.Menu mnu26 
         Caption         =   "Withdraw Student"
      End
      Begin VB.Menu mnu27 
         Caption         =   "WithDrawal Notes"
      End
   End
   Begin VB.Menu mnuA0 
      Caption         =   "&Interfaces"
      Begin VB.Menu mnuA1 
         Caption         =   "Define Interfaces"
      End
      Begin VB.Menu mnuA2 
         Caption         =   "Import Staff Records"
      End
      Begin VB.Menu mnuA3 
         Caption         =   "Export Final Results"
      End
   End
   Begin VB.Menu mnu50 
      Caption         =   "&Evaluations"
      Begin VB.Menu mnu52 
         Caption         =   "Test Result Administration"
      End
      Begin VB.Menu mnu53 
         Caption         =   "Praticals Result Administration"
      End
      Begin VB.Menu mnu55 
         Caption         =   "Assignment Result Administration"
      End
      Begin VB.Menu mnu43 
         Caption         =   "Class Effort Administration"
      End
      Begin VB.Menu mnu5a 
         Caption         =   "-"
      End
      Begin VB.Menu mnu51 
         Caption         =   "&Examination Result Administration"
      End
      Begin VB.Menu mnu43a 
         Caption         =   "External Exams Result Administration"
      End
      Begin VB.Menu mnu58a 
         Caption         =   "-"
      End
      Begin VB.Menu mnu54d 
         Caption         =   "Cumulative Result Adjustment"
      End
      Begin VB.Menu mnu54a 
         Caption         =   "Student Result Adjustment"
      End
      Begin VB.Menu mnu54b 
         Caption         =   "-"
      End
      Begin VB.Menu mnu54c 
         Caption         =   "Term Processsing"
      End
      Begin VB.Menu mnu58 
         Caption         =   "Pre-Term Close Result Review"
      End
      Begin VB.Menu mnu59 
         Caption         =   "Students Promotion Administration"
      End
   End
   Begin VB.Menu mnu80 
      Caption         =   "&Reports"
      Begin VB.Menu mnu82 
         Caption         =   "Class Performance"
      End
      Begin VB.Menu mnu83 
         Caption         =   "Subject performance"
      End
      Begin VB.Menu mnu84 
         Caption         =   "Examination Performance"
      End
      Begin VB.Menu mnu85 
         Caption         =   "Student Cummulative Report"
      End
      Begin VB.Menu mnu89 
         Caption         =   "External Results Performance"
      End
      Begin VB.Menu mnu8q 
         Caption         =   "Teachers Class Performance Index"
      End
      Begin VB.Menu mnu89a 
         Caption         =   "-"
      End
      Begin VB.Menu mnu8c 
         Caption         =   "Passed Students List"
      End
      Begin VB.Menu mnu8d 
         Caption         =   "Failed Students List"
      End
      Begin VB.Menu mnu8da 
         Caption         =   "Students Evaluation Detail Report"
      End
      Begin VB.Menu mnu8db 
         Caption         =   "-"
      End
      Begin VB.Menu mnu86a 
         Caption         =   "Student Sequential List"
      End
      Begin VB.Menu mnu86 
         Caption         =   "Class Subject List"
      End
      Begin VB.Menu mnu86b 
         Caption         =   "Students Medical Report"
      End
      Begin VB.Menu mnu8a 
         Caption         =   "-"
      End
      Begin VB.Menu mnu81 
         Caption         =   "Current Students Result Advice"
      End
      Begin VB.Menu mnu8ab 
         Caption         =   "Graduated Students Result Advice"
      End
      Begin VB.Menu mnuOP 
         Caption         =   "Class Group Overall Performance"
      End
      Begin VB.Menu mnuOS 
         Caption         =   "Class Group Subject Performance"
      End
      Begin VB.Menu mnuGoP 
         Caption         =   "Graduants Class Group Overall Performance"
      End
      Begin VB.Menu mnugoS 
         Caption         =   "Graduants Class Group Subject Performance"
      End
      Begin VB.Menu mnu8e 
         Caption         =   "-"
      End
      Begin VB.Menu mnu8f 
         Caption         =   "&Current Students Records"
      End
      Begin VB.Menu mnu8g 
         Caption         =   "&Graduated Students Records"
      End
      Begin VB.Menu mnu8h 
         Caption         =   "Withdrawn Students Records"
      End
      Begin VB.Menu mnu8i 
         Caption         =   "-"
      End
      Begin VB.Menu mnu8j 
         Caption         =   "Current Student Notes"
      End
      Begin VB.Menu mnu8k 
         Caption         =   "Graduated  Student Notes"
      End
      Begin VB.Menu mnu8l 
         Caption         =   "Withdrawn Student Notes"
      End
      Begin VB.Menu mnu87 
         Caption         =   "System Processing Notes"
      End
      Begin VB.Menu mnu8m 
         Caption         =   "-"
      End
      Begin VB.Menu mnu8n 
         Caption         =   "Teachers Records"
      End
      Begin VB.Menu mnu8o 
         Caption         =   "Teachers Subjects Class"
      End
      Begin VB.Menu mnu8t 
         Caption         =   "Subjects Teachers Class"
      End
      Begin VB.Menu mnu8p 
         Caption         =   "Teachers Responsibilities"
      End
      Begin VB.Menu mnu8r 
         Caption         =   "Classroom Subjects with Teachers "
      End
      Begin VB.Menu mnu8s 
         Caption         =   "Classroom Subjects without Teachers "
      End
   End
   Begin VB.Menu mnuEnq 
      Caption         =   "En&quiries"
      Begin VB.Menu mnuStEnq 
         Caption         =   "Student Information Enquiry"
      End
      Begin VB.Menu mnuREnq 
         Caption         =   "Current Students Result Enquirry"
      End
      Begin VB.Menu mnuRPrt 
         Caption         =   "Current Students Result Reprint"
      End
      Begin VB.Menu mnuEXRPrt 
         Caption         =   "Ex-Students Result Reprint"
      End
      Begin VB.Menu mnuaa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAI 
         Caption         =   "Students Intelligence Report"
      End
      Begin VB.Menu mnuBI1 
         Caption         =   "Students Population Statistics"
      End
      Begin VB.Menu mnuSTa 
         Caption         =   "-"
      End
      Begin VB.Menu mnust1 
         Caption         =   "Student Total - Class Statistics"
      End
      Begin VB.Menu mnust2 
         Caption         =   "Student Total - Enterprise Statistics"
      End
      Begin VB.Menu mnuSTb 
         Caption         =   "-"
      End
      Begin VB.Menu mnust5 
         Caption         =   "Female Gender Statistics"
      End
      Begin VB.Menu mnust6 
         Caption         =   "Male Gender Statistics"
      End
   End
   Begin VB.Menu mnu60 
      Caption         =   "Re&ferences"
      Begin VB.Menu mnu61 
         Caption         =   "School Levels Definition"
      End
      Begin VB.Menu mnu62 
         Caption         =   "&Class Group Definition"
      End
      Begin VB.Menu mnu63 
         Caption         =   "Sub&jects Classification"
      End
      Begin VB.Menu mnu64 
         Caption         =   "Subject Definition"
      End
      Begin VB.Menu mnu64a 
         Caption         =   "School Terms"
      End
      Begin VB.Menu mnu64b 
         Caption         =   "-"
      End
      Begin VB.Menu mnu65 
         Caption         =   "&External Exam Subjects Definition"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu66 
         Caption         =   "&Scores and Grade Definition"
      End
      Begin VB.Menu mnu66a 
         Caption         =   "-"
      End
      Begin VB.Menu mnu67 
         Caption         =   "&Department Definition"
      End
      Begin VB.Menu mnu68 
         Caption         =   "Qualification Definition"
      End
      Begin VB.Menu mnu69 
         Caption         =   "Professional Disclipine Definition"
      End
      Begin VB.Menu mnu70 
         Caption         =   "Job Title Definition"
      End
      Begin VB.Menu mnu71 
         Caption         =   "Staff Category Definition"
      End
      Begin VB.Menu mnu71a 
         Caption         =   "-"
      End
      Begin VB.Menu mnu75a 
         Caption         =   "Elective Subject Group Definition"
      End
      Begin VB.Menu mnu72 
         Caption         =   "Sports Activites Definition"
      End
      Begin VB.Menu mnu72a 
         Caption         =   "Pupils Talent Definition"
      End
      Begin VB.Menu mnu74 
         Caption         =   "Special Skills Definition"
      End
      Begin VB.Menu mnu73 
         Caption         =   "&State Definition"
      End
      Begin VB.Menu mnu75 
         Caption         =   "Local Government Area"
      End
      Begin VB.Menu mnu75b 
         Caption         =   "School House Definition"
      End
   End
   Begin VB.Menu mnu99 
      Caption         =   "Window"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "FrmDataPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

Private Sub mnu13_Click()
   frmCPasswd.Show
End Sub

Private Sub mnu14_Click()
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
           mnu13 = False
           mnu14 = False
           mnu15 = False
           mnu16 = False
           mnu17 = False
           mnu18 = False
           mnu19 = False
           mnu20 = False
           mnu21 = False
           mnu22 = False
           mnu23 = False
           mnu24 = False
           mnu25 = False
    Unload Me
    frmLogin.Show
End Sub

Private Sub mnu15_Click()
    End
End Sub

Private Sub mnu21_Click()
    frmStudRecs.Show
End Sub

Private Sub mnu22_Click()
    FrmRelTn.Show
End Sub

Private Sub mnu23_Click()
    frmMeds.Show
End Sub

Private Sub mnu25_Click()
    FrmAutoB.Show
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

Private Sub mnu41_Click()
    FrmCSubAllL.Show
End Sub

Private Sub mnu41a_Click()
   FrmCSubAllG.Show
End Sub

Private Sub mnu41b_Click()
   FrmCSubAllC.Show
End Sub

Private Sub mnu42_Click()
    frmtransClass.Show
End Sub

Private Sub mnu43_Click()
    FrmEffort.Show
End Sub

Private Sub mnu45_Click()
    FrmJob.Show
End Sub

Private Sub mnu45b_Click()
    FrmSubclass.Show
End Sub

Private Sub mnu46_Click()
    Frmgrade.Show
End Sub

Private Sub mnu43a_Click()
 FrmExtExam.Show
End Sub

Private Sub mnu48_Click()
    frmtransStep.Show
End Sub

Private Sub mnu49_Click()
    frmtransClass.sstrans.Caption = "Student Departmental Transfer"
    frmtransClass.Show
End Sub

Private Sub mnu4b_Click()
    FrmDisc.Show
End Sub

Private Sub mnu4c_Click()
    FrmState.Show
End Sub

Private Sub mnu51_Click()
   FrmExam.Show
End Sub

Private Sub mnu52_Click()
   FrmTest.Show 0
End Sub

Private Sub mnu53_Click()
   FrmPract.Show
End Sub

Private Sub mnu54a_Click()
   FrmRsltAdj.Show
End Sub

Private Sub mnu54c_Click()
   FrmProc.Show
End Sub

Private Sub mnu54d_Click()
   FrmCumulAdj.Show
End Sub

Private Sub mnu55_Click()
   FrmAssign.Show
End Sub

Private Sub mnu58_Click()
   FrmPreCloseView.Show
End Sub

Private Sub mnu59_Click()
   FrmManprom.Show
End Sub

Private Sub mnu61_Click()
    FrmSlev.Show
End Sub
Private Sub mnu62_Click()
  FrmClasGrp.Show
End Sub

Private Sub mnu62a_Click()
   FrmClass.Show
End Sub

Private Sub mnu63_Click()
   FrmSubclass.Show
End Sub

Private Sub mnu64_Click()
   FrmSubj.Show
End Sub

Private Sub mnu64a_Click()
   FrmTerms.Show
End Sub

Private Sub mnu66_Click()
   FrmSGrading.Show
End Sub

Private Sub mnu67_Click()
   FrmDept.Show
End Sub

Private Sub mnu68_Click()
   FrmQual.Show
End Sub

Private Sub mnu69_Click()
   FrmDisc.Show
End Sub

Private Sub mnu70_Click()
   FrmJob.Show
End Sub

Private Sub mnu71_Click()
   FrmStaffcat.Show
End Sub

Private Sub mnu72_Click()
   FrmSports.Show
End Sub

Private Sub mnu72a_Click()
  FrmTalents.Show
End Sub

Private Sub mnu73_Click()
   FrmState.Show
End Sub

Private Sub mnu74_Click()
   FrmSkills.Show
End Sub

Private Sub mnu75_Click()
   FrmLGA.Show
End Sub

Private Sub mnu75a_Click()
   FrmElect.Show
End Sub

Private Sub mnu75b_Click()
   FrmHouse.Show
End Sub

Private Sub mnu81_Click()
   'On Error GoTo perror
   CrptCtrl.ReportFileName = Rptpath & "rsltslip.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu82_Click()
   'On Error GoTo perror
   CrptCtrl.ReportFileName = Rptpath & "overperf.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu83_Click()
   'On Error GoTo perror
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
   'On Error GoTo perror
   CrptCtrl.ReportFileName = Rptpath & "teachperf.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu8r_Click()
   'On Error GoTo perror
   CrptCtrl.ReportFileName = Rptpath & "subjteacheryes.rpt"
   CrptCtrl.DataFiles(0) = DBpath
   CrptCtrl.Action = 1
PError:
End Sub

Private Sub mnu8s_Click()
   'On Error GoTo perror
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
