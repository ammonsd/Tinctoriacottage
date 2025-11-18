<%

   DIM bolErrFnd, strEntity

SUB Write_Errors

   bolErrFnd = true

   Response.Write "<b>Entity: " & strEntity & "</b><br>"
   Response.Write "VBScript Errors Occured:<br>"
   Response.Write "Error Number=" & err.number & "<br>"
   Response.Write "Error Descr.=" & err.description & "<br>"
   Response.Write "Help Context=" & err.helpcontext & "<br>"
   Response.Write "Source=" & err.source & "<br>"
   Err.Clear

   Response.End

End Sub
%>

