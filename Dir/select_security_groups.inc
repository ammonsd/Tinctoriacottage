
<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/array_sort.inc" -->

<%

   DIM bolListOnly, bolOthGroups, arySecRec(5000), strPrevGroups, aryGrps(100), intGrpSub

SUB Process_Security_Groups

   DIM intSub, intSub2, intMaxCols, intCol

   strSpecSecGrp = "SecMaint"

   bolPerformHC = true

   strPrevGroups = "|"

   intGrpSub = 0
   For intSub = 1 to UBound(arySecRec)
      If trim(arySecRec(intSub)) = "" Then
         EXIT FOR
      End If
      Call Process_Groups(arySecRec(intSub))
   Next

   If not bolListOnly Then
      Response.Write "<form action='build_email_list.asp' method='post'>" & vbCrLf
      Response.Write "<input type=hidden name='SO' value='Y'>" & vbCrLf
      Response.Write "<input type=hidden name='SECUR' value='Y'>" & vbCrLf
   End If
   Response.Write "<table align='center' border='0' cellpadding='0' cellspacing='0'><tr>" & vbCrLf
   If bolListOnly Then
      call Setup_Web_Page("Currently Assigned Security Groups",1)
   Else
      call Setup_Web_Page("Select Security Groups to Email",1)
   End If
   If intGrpSub = 0 Then
      If bolListOnly Then
         Response.Write "<td align='center'><b>No Groups Have Been Assigned</b></td>" & vbCrLf
      Else
         Response.Write "<td align='center'><b>No Groups Have Been Assigned - All Members Will Be Selected</b></td>" & vbCrLf
      End If
   End If
   arySorted = SortArray(aryGrps,0)
   Erase aryGrps
   intMaxCols = ROUND((intNbrEntries / 16) + .5)
   If intMaxCols < 1 Then
      intMaxCols = 1
   End If
   intCol = 0
   If bolListOnly and not bolOthGroups Then
      CALL Get_Function_Security_Groups
      For intSub = 0 to intNbrEntries - 1
         ' Remove Assigned Special Groups
         For intSub2 = 1 to aryFunctionGrps(0)
            If ucase(aryFunctionGrps(intSub2)) = ucase(arySorted(intSub)) Then
               arySorted(intSub) = ""
               EXIT FOR
            End If
         Next
      Next
      For intSub = 1 to aryFunctionGrps(0)
         intNbrEntries = intNbrEntries + 1
         arySorted(intNbrEntries-1) = "<font color='red'>" & aryFunctionGrps(intSub) & "</font>"
      Next
   End If
   For intSub = 0 to intNbrEntries - 1
      If ucase(arySorted(intSub)) <> "ALL" and arySorted(intSub) <> "" Then
         intCol = intCol + 1
         If bolListOnly Then
            Response.Write "<td><b>" & arySorted(intSub) & "</b></td>" & vbCrLf
         Else
            Response.Write "<td><input type=checkbox name='KW' value='" & arySorted(intSub) & "'>&nbsp;&nbsp;<b>" & arySorted(intSub) & "</b></td>" & vbCrLf
         End If
         If intCol = intMaxCols and intSub < intNbrEntries - 1 Then
            intCol = 0
            Response.Write "</tr><tr>" & vbCrLf
         End If
      End If
   Next

   If bolListOnly Then
      Response.Write "</tr><tr><td align='center'><br><input type='button' value='Close Window' onClick='window.close()'>" & vbCrLf
   Else
      Response.Write "</tr><tr><td align='center'><br><input type='submit' value='Process'>" & vbCrLf
   End If
   Response.Write "</table>" & vbCrLf

   If bolListOnly and aryFunctionGrps(0) > 0 and not bolOthGroups Then
      Response.Write "<div align='center'><font size='-1' color='red'>" & vbCrLf
      Response.Write "<br><br>Groups in RED Control Function Access Only" & vbCrLf
      Response.Write "</font></div>" & vbCrLf
   End If

END SUB

SUB Process_Groups(ChkGroup)

   DIM intSub, strKW, strGroups

   bolDoPhrases = true
   strGroups = RemoveLineBreaks(ChkGroup)

   If trim(strGroups) <> "" and not isnull(strGroups) Then
      intSub = 1
      strKW = GetWords(strGroups,intSub,1)
      do while strKW <> ""
         If instr(strPrevGroups,"|"&lcase(strKW)&"|" ) = 0 Then
            strPrevGroups = strPrevGroups & lcase(strKW) & "|"
            aryGrps(intGrpSub) = strKW
            intGrpSub = intGrpSub + 1
         End If
         intSub = intSub + intRetWords
         strKW = GetWords(strGroups,intSub,1)
      Loop
   End If

END SUB
%>
