Attribute VB_Name = "Module1"
' Attribute VB_Name = "globals"
Option Explicit

Global CurrentPeriod As Integer
Global tmpStr1, tmpStr2, tmpStr3, tmpStr4, tmpStr5, tmpStr6 As String
Global DproDBpath As String, GlString As String
Global DproPicpath As String, NodYear As Single
Global DproRPTpath As String
Global mvUserid As String, menu1 As Boolean
Global mvUserName As String
Global mvSql As String, rst As Recordset
Global Const HOURGLASSCURSOR = 11   ' Cursor stuff
Global Const NODELETE = 5
Global Const NOMOVEFRONT = 19
Global Const NOMOVEPAST = 20
Global Const NOMATCH = 21
Global Const EMPTYTABLE = 39
Global Const INPUTREQUIRED = 22
Global Const EMPTYTRANSACTION = 23
Global Const NOINPUT = 44
Global Const BADPASSWORD = 50
Global Const vbGrayText = 7
Global PMsg As String
Global GCodeTerm As String
Global GDescTerm As String
Global GCurrTerm As String
Global GPyear As Integer
Global GSControl As String
Global GmedLim As Currency
Global GAutoCtr As Boolean
Global DBpath As String
Global Rptpath As String
Global DocsPath As String
Global PicsPath As String
Global mvDbFile As String
Global MVCoyname As String
Global gstate As String
Global gBunitD As String
Global coyname As String
Global Crest As Boolean
Global GAutoProm As Boolean
Global GautoMs As Boolean
Global mvAcad As Boolean
Global GTBuff As String
Global GlbState As String
Global CInit As Boolean
Global mvBuff As String
Global GLogo As String
Global GTest As Boolean
'terms global names
Global GPromTerm, GPromdesc, GCumTerm As String
Global mnu1 As Boolean, mnu2 As Boolean, mnu3 As Boolean, mnu4 As Boolean, mnu5 As Boolean
Global mnu6 As Boolean, mnu7 As Boolean, mnu8 As Boolean, mnu9 As Boolean, mnu10 As Boolean
Global mnu11 As Boolean, mnu12 As Boolean, mnu13 As Boolean, mnu14 As Boolean, mnu15 As Boolean
Global mnu16 As Boolean, mnu17 As Boolean, mnu18 As Boolean, mnu19 As Boolean, mnu20 As Boolean
Global mnu21 As Boolean, mnu22 As Boolean, mnu23 As Boolean, mnu24 As Boolean, mnu25 As Boolean
Global Col0, Col1, Col2, Col3, Col4, Col5, Col6, Col7, Col8, Col9, col10 As Column

Sub AfterLastError()
   Beep
   MsgBox "Cannot move past the last record", vbCritical, "Access Error"
End Sub

Sub BeforeFirstError()
   Beep
   MsgBox "Cannot move in front of first record", vbCritical, "Access Error"
End Sub

Function DecodePassword(p As String, check As Integer) As String
   Dim i As Integer, temp As Integer
   Dim length As Integer
   Dim s As String * 30, w As String

   For i = 1 To 4
      Mid$(s, i, 1) = Chr$(Asc(Mid$(p, i, 1)) - 26)
   Next i
   length = Val(Mid$(s, 1, 2))
   check = Val(Mid$(s, 3, 2))

   For i = 5 To length + 4
      w = w + Chr$(Asc(Mid$(p, i, 1)) - check)
   Next i
   DecodePassword = w

End Function
Function CSCheck()
   PMsg = "You have added or removed a learning subject to a classroom. "
   PMsg = PMsg + "You need to update classroom for the subject to become available. "
   PMsg = PMsg + "Do you want to do that now?"
   CSCheck = MsgBox(PMsg, vbYesNo, "Verify Class Room Update")
End Function

Function CRCheck()
   PMsg = "You have added or removed a learning resource to a classroom. "
   PMsg = PMsg + "You need to update classroom for the resource to become available. "
   PMsg = PMsg + "Do you want to do that now?"
   CRCheck = MsgBox(PMsg, vbYesNo, "Verify Class Room Update")
End Function

Function VoidCheck()
   VoidCheck = MsgBox("Evaluations with the selected reference will be removed from the system. Are you sure you wish to continue?", vbYesNo, "Verify Void")
End Function
Function DomCheck()
   DomCheck = MsgBox("Are you sure you wish to domicile this mail?", vbYesNo, "Verify Delete")
End Function

Function DeleteCheck()
   DeleteCheck = MsgBox("Are you sure you wish to delete this record?", vbYesNo, "Verify Delete")
End Function
Function TermCheck()
   TermCheck = MsgBox("Are you sure you wish to Continue Termination?", vbYesNo, "Verify Termination")
End Function
Function ConfirmS()
   ConfirmS = MsgBox("Are you sure you wish to assign this Subject to Staff?", vbYesNo, "Verify Update")
End Function
Function ConfirmP()
   ConfirmP = MsgBox("Are you sure you wish to assign sports role to staff?", vbYesNo, "Verify Update")
End Function
Function ConfirmM()
   ConfirmM = MsgBox("Are you sure you wish to assign Class Master Role to staff?", vbYesNo, "Verify Update")
End Function
Function cont()
   cont = MsgBox("Are you sure you wish to continue?", vbYesNo, "Verify Update")
End Function
Function Confirm()
   Confirm = MsgBox("Are you sure you wish to assign this Role to staff?", vbYesNo, "Verify Update")
End Function
Function cInitChk()
   PMsg = "Each Class Room Subjects have been updated already. If you choose to update again, "
   PMsg = PMsg + "the subjects previously defined for classes will be removed and fresh update done using current settings. "
   PMsg = PMsg + " Do you want to continue? "
   cInitChk = MsgBox(PMsg, vbYesNo, "Verify Update")
End Function
Function mScaleChk()
   PMsg = "Some Subject markscales have been created already. If you choose to continue, "
   PMsg = PMsg + "all previous subjects markscales for all subjects will be removed and fresh update done using "
   PMsg = PMsg + "a uniform scale as specified on current input. Do you want to continue? "
   mScaleChk = MsgBox(PMsg, vbYesNo, "Verify Update")
End Function

Function classupD()
   classupD = MsgBox("Are you sure you wish to update class records and assign subjects to teachers?", vbYesNo, "Verify Update")
End Function
Function ConfirmPref()
   ConfirmPref = MsgBox("Are you sure you wish to update Classroom with changes?", vbYesNo, "Verify Update")
End Function
Function ConfirmLBin()
   ConfirmLBin = MsgBox("Are you sure you wish to empty activity log records?", vbYesNo, "Verify Update")
End Function

Sub EncodePassword(Password As String, D As String, check As Integer)
   Dim i As Integer, temp As Integer
   Dim s As String * 30, cs As String

   For i = 5 To Len(Password) + 4
      temp = Asc(Mid$(Password, i - 4, 1)) + check
      D = D + Chr$(temp)
   Next i
   s = D
   Randomize                     ' Make it random offsets
   For i = Len(D) + 1 To 30
      temp = 0
      Do While temp < 32 Or temp > 127
         temp = (Rnd * 127)         ' Force a small number
      Loop
      Mid$(s, i, 1) = Chr$(temp)
   Next i

   Password = s

End Sub

Sub ErrorMessages(Index As Integer)
   Dim i As Integer
   Beep
   Select Case Index
      Case NODELETE
         i = MsgBox("Cannot Delete Record.", vbCritical, "Delete")
      Case NOMOVEFRONT
         i = MsgBox("Cannot move before first record.", vbCritical, "Illegal Database Directive")
      Case NOMOVEPAST
         i = MsgBox("Cannot move after last record.", vbCritical, "Illegal Database Directive")
      Case NOMATCH
         i = MsgBox("Cannot find this record.", vbCritical, "Search Failure")
      Case EMPTYTABLE
         i = MsgBox("Empty Table, No records available.", vbCritical, "Empty Table(s)")
      Case INPUTREQUIRED
         i = MsgBox("You must supply data for this field.", vbCritical, "Required Data")
      Case NOINPUT
         i = MsgBox("No input given. Option cancelled", vbCritical, "No Input")
      Case EMPTYTRANSACTION
         i = MsgBox("No debit or credit balance given.", vbCritical, "No Balances")
   End Select
End Sub

Function LeapYear(y As Integer) As Integer

   If y Mod 4 = 0 And y Mod 100 <> 0 Or y Mod 400 = 0 Then
      LeapYear = 1
   Else
      LeapYear = 0
   End If

End Function


Sub SayDollar(Amount As Currency, result As String)
   Dim buff As String, done As String
   Static units(10) As String, teens(10) As String
   Static tens(10) As String, denoms(4) As String
   Dim Index As Integer, length As Integer
   Dim i As Integer, Passes As Integer
   Dim temp As Double, remains As Double

   units(0) = "0": units(1) = "one": units(2) = "two"
   units(3) = "three": units(4) = "four": units(5) = "five"
   units(6) = "six": units(7) = "seven": units(8) = "eight"
   units(9) = "nine": teens(0) = "ten"

   teens(1) = "eleven": teens(2) = "twelve": teens(3) = "thirteen"
   teens(4) = "fourteen": teens(5) = "fifteen": teens(6) = "sixteen"
   teens(7) = "seventeen": teens(8) = "eighteen": teens(9) = "nineteen"

   tens(1) = "ten": tens(2) = "twenty": tens(3) = "thirty"
   tens(4) = "forty": tens(5) = "fifty": tens(6) = "sixty"
   tens(7) = "seventy": tens(8) = "eighty": tens(9) = "ninety"
   
   denoms(1) = " hundred ": denoms(2) = " thousand ": denoms(3) = " million "

   buff = Format$(Amount, "#########.00")
   length = Len(buff) - 3                 ' Get past decimal and cents
   Passes = length
   Index = 0
   temp = 0
   remains = 0

   Do While (length > -1)
      Select Case (length Mod 3)
         Case 0
            If Mid$(buff, 1, 1) = "." Then
               done = done & " dollars and "
               length = 3
            Else
               If length > 0 Then
                  If Len(done) > 0 Then
                     done = done & denoms(length / 3 + 1)
                  End If
                  If Mid$(buff, 1, 1) <> "0" Then
                     i = Val(Mid$(buff, 1, 1))
                     done = done & units(i)
                     done = done & denoms(1)
                  End If
               Else
                  i = Len(done) - 2
                  If Mid$(done, i, 1) = "d" Then
                     done = done & "no"
                  End If
                  done = done & " cents"
               End If
            End If
            length = length - 1
         Case 1
            If Mid$(buff, 1, 1) <> "0" Then
               i = Val(Mid$(buff, 1, 1))
               done = done & units(i)
            End If
            If Mid$(buff, 1, 1) = "0" And Passes = 1 Then
               done = done & "no dollars"
            Else
               If length = 1 And Len(done) = 1 Then
                  done = done & "dollar"
               End If
            End If
            length = length - 1
         Case 2
            If Mid$(buff, 1, 1) = "1" Then
               i = Val(Mid$(buff, 2, 1))
               done = done & teens(i)
               buff = Mid$(buff, 2)
               length = length - 2
            Else
               i = Val(Mid$(buff, 1, 1))
               done = done & tens(i)
               If Mid$(buff, 1, 1) <> "0" And Mid$(buff, 2, 1) <> "0" Then
                  done = done & "-"
               End If
               length = length - 1
            End If
      End Select
      buff = Mid$(buff, 2)
   Loop
   Mid$(done, 1, 1) = UCase$(Mid$(done, 1, 1))
   result = done
End Sub


Public Sub CentreForm(X As Form)

   X.Move (Screen.Width - X.Width) / 2, _
   (Screen.Height - X.Height) / 2
   
End Sub


