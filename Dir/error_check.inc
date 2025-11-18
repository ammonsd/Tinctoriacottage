
<!--#INCLUDE FILE="../dir/display_popup_msg.inc" -->

<%
    DIM errMsgs, aryErrorMsgs, aryErrorText(3), bolEditError

SUB Process_Errors

   DIM intSub

   SESSION("POPUPMSG") = ""

   For intSub = 0 To errMsgs.Count -1
      SESSION("POPUPMSG") = SESSION("POPUPMSG") & "<b>&nbsp;&nbsp;&nbsp;&#183;&nbsp;&nbsp;" & aryErrorMsgs(intSub) & "</b><br>"
   Next
   If SESSION("POPUPMSG") <> "" Then
      bolEditError = true
   Else
      bolEditError = true
   End If

END SUB

%>
