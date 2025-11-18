<%@ LANGUAGE="VBSCRIPT"%>
<%
OPTION EXPLICIT
Response.Buffer = true

%>

<%

DIM X

For Each x In SESSION.CONTENTS
   response.write "<b>"&x&"</b>" & ": " & SESSION(x) & "<br>"
Next

For Each x In Request.ServerVariables
   If left(x,4) <> "ALL_" Then
      response.write "<b>"&x&"</b>" & ": " & Request.ServerVariables(x) & "<br>"
   End If
Next

%>
