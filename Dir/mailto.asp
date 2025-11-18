<%@ LANGUAGE="VBSCRIPT" %>


<%
OPTION EXPLICIT
Response.Buffer = True

   If Request.QueryString("MT") = "DA-R" Then
      Response.Redirect("mailto:consulting@deanammons.com?Subject=Resume")
   ElseIf Request.QueryString("MT") = "DA" Then
      Response.Redirect("mailto:dean@deanammons.com")
   ElseIf Request.QueryString("MT") = "TA-R" Then
      Response.Redirect("mailto:consulting@fulcrumcomputing.com?Subject=Resume")
   ElseIf Request.QueryString("MT") = "TA" Then
      Response.Redirect("mailto:ammonst@tinctoriacottage.com")
   ElseIf Request.QueryString("MT") = "TC" Then
      Response.Redirect("mailto:webmaster@tinctoriacottage.com")
   End If
   Response.END
%>
