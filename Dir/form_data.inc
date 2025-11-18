
<%
   DIM bolFormDataOnly, bolNoSessionVariables, bolNoFormData

FUNCTION FormData(FldName)

   If Request.QueryString(FldName) <> "" and not bolFormDataOnly Then
      FormData = Request.QueryString(FldName)
   ElseIf Request.Form(FldName) <> "" and not bolNoFormData Then
      FormData = Request.Form(FldName)
   ElseIf not bolFormDataOnly and not bolNoSessionVariables Then
      FormData = SESSION(FldName)
   End If

END FUNCTION

FUNCTION GetParmInfo(FldName)

   If SESSION(FldName) <> "" Then
      GetParmInfo = SESSION(FldName)
   ElseIf Request.QueryString(FldName) <> "" Then
      GetParmInfo = Request.QueryString(FldName)
   End If

END FUNCTION
%>
