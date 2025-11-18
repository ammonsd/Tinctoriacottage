
<%
   DIM bolIncLinks, strTypeLink, intNbrPgs, strBackName, strEndingComments, bolNoEnd

SUB Display_Ending_Msg(Msg,NextUrl)

   DIM bolPopUpMsg, strHld

   If SESSION("POPUPMSG") = "Y" Then
      bolPopUpMsg = true
      SESSION("POPUPMSG") = ""
   End If

   If not bolPopUpMsg Then
      bolNoTable = true
      Call Setup_Web_Page("","")
   End If

   If not bolPopUpMsg and NextUrl <> "" Then
      strHld = strHld & "<meta http-equiv='Refresh' content='2; url=" & NextUrl & "'>"
   End If
   strHld = strHld & "<div align='center'>"
   If Not bolPopUpMsg Then
      strHld = strHld & "<br><br><br><font size=+2>"
   End If
   strHld = strHld & "<b>" & Msg & "</b>"
   If Not bolPopUpMsg Then
      strHld = strHld & "</font>"
   End If

   If strEndingComments <> "" Then
      strHld = strHld & "<br><br>" & strEndingComments
   End If
   strHld = strHld & "</div>"

   If bolPopUpMsg Then
      SESSION("POPUPMSG") = REPLACE(strHld,"'","")
      SESSION("POPUPMSG") = REPLACE(SESSION("POPUPMSG"),CHR(34),"")
      EXIT SUB
   End If

   Response.Write strHld

   IF isnumeric(intNbrPgs) and intNbrPgs <> "" then
      If strBackName = "" Then
         strBackName = "Back"
      End If
      Response.Write "<div align='center'><br>"
      Response.Write "<input type='button' value='" & strBackName & "' onClick='history.go(-" & intNbrPgs & ")'>"
      Response.Write "<br><br></div>"
   End If
   IF bolIncLinks = true then
      Response.Write "<table>"
      If strTypeLink = "S"  Then
         Call Security_Maintenance_Links(1)
      Else
         Call Maintenance_Links(1)
      End If
      Response.Write "</table>"
   End If
   Response.Write "</div></html></body>"
   If not bolNoEnd Then
   Response.End
   End If
End Sub
%>
