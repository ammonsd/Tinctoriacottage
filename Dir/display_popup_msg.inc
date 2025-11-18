
<!--#INCLUDE FILE="../dir/js_cs_scripts.inc" -->

<%

   DIM strPopUpTitle

SUB Check_For_PopUp_Msg

   If SESSION("POPUPMSG") <> "" Then
      Call SetUp_PopUp_Msg
      aryErrorText(1) = "<form name='msg'>"
      aryErrorText(2) = " onLoad=" & CHR(34) & "PopUpMsg('" & strPopUpTitle & "','" & SESSION("POPUPMSG") & "',1)" & CHR(34)
      aryErrorText(3) = "</form>"
      SESSION("POPUPMSG") = ""
   End If

End SUB

%>
