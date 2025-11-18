
<%
FUNCTION AddCookie(Cname, CKey, CValue, CExpDays)

   If Cname = "" Then
      Exit FUNCTION
   End If
   If IsNumeric(CExpDays) = False Then
      CExpDays = 0
   End If
   If CExpDays < 1 Then
      CExpDays = 90
   End If

   If CValue = "" Then
      CValue = " "
   End If

   If CKey <> "" Then
      Response.Cookies(Cname)(CKey) = CValue
   Else
      Response.Cookies(Cname) = CValue
   End If

   Response.Cookies(CName).Expires = Date + CExpDays
   Response.Cookies(CName).Domain = GetCookieDomain()
   Response.Cookies(CName).Path = GetCookiePath()

End FUNCTION

FUNCTION KillCookie(Cname,CKey)

   If CKey <> "" Then
      Call AddCookie(Cname, Ckey, "", 0)
   Else
      Response.Cookies(Cname).Expires = Date - 365
      Response.Cookies(CName).Domain = GetCookieDomain()
      Response.Cookies(CName).Path = GetCookiePath()
   End If

End FUNCTION

FUNCTION GetCookie(Cname, Ckey)
   If CKey <> "" Then
      GetCookie = Request.Cookies(Cname)(Ckey)
   Else
      GetCookie = Request.Cookies(Cname)
   End If
End FUNCTION

FUNCTION GetCookieDomain()
   If Application("CookieDomain") <> "" Then
      GetCookieDomain = Application("CookieDomain")
   Else
      GetCookieDomain =  Request.ServerVariables("SERVER_NAME")
   End If
End FUNCTION

FUNCTION GetCookiePath()

   DIM strTmp, intCntr

   If Application("CookiePath") <> "" Then
      GetCookiePath = Application("CookiePath")
   Else
      strTmp = Request.ServerVariables("SCRIPT_NAME")
      strTmp = Split(strTmp, "/")
      For intCntr = 0 to Ubound(strTmp) - 1
         GetCookiePath = GetCookiePath & strTmp(intCntr) & "/"
      Next
      If GetCookiePath = "/" Then
         GetCookiePath = ""
      End If
   End If
End FUNCTION
%>
