<%
'          **************************************************
'          *  Copyright Fulcrum Computing, Inc.  2000       *
'          **************************************************
'          * This program is copyrighted under the laws of  *
'          * the United States. Any unautorized duplication *
'          * or use of this progrm is prohibited.           *
'          **************************************************
'**Start Encode**
%>

<%

   DIM strDueDay, intDueDay, intFirstDay, strDefDueDay

FUNCTION CheckDate(ReqDte)

'  Verify Date Format - Returning Date in the format "YYYY/MM/DD"

   DIM strY, strD, strM

   If isDate(ReqDte) = false then
      EXIT FUNCTION
   End If

   strY = YEAR(ReqDte)
   strM = MONTH(ReqDte)
   strD = DAY(ReqDte)

   If len(strM) = 1 Then
      strM = "0" & strM
   End If
   If len(strD) = 1 Then
      strD = "0" & strD
   End If

   CheckDate = strY & "/" & strM & "/" & strD

END FUNCTION

FUNCTION CutOffDays(OPT,ADJ)

   DIM strReqDay, strCurDay

   Select Case ucase(OPT)
      Case "SUN"
         intFirstDay = 1
      Case "MON"
         intFirstDay = 2
      Case "TUE"
         intFirstDay = 3
      Case "WED"
         intFirstDay = 4
      Case "THU"
         intFirstDay = 5
      Case "FRI"
         intFirstDay = 6
      Case "SAT"
         intFirstDay = 7
      Case Else
         intFirstDay = 0
   End Select

   CutOffDays = 0

   If intFirstDay > 0 then
      CutOffDays = WEEKDAY(DATE,intFirstDay) + ADJ
   ElseIf OPT <> "" and isnumeric(OPT) = true Then
      CutOffDays = OPT
   End If

END FUNCTION

%>
