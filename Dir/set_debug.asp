<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>
<%
   SESSION("DEBUG") = Request.QueryString("DEBUG")
   Response.Write "<br><br><br><br><br><br><div align='center'><input type='button' value='Back' onClick='history.go(-1)'></div>"
%>
