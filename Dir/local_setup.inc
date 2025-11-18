<%

   DIM bolRexx, bolSecurityOnly, bolWebMail, strWebCompany, strDefBookLang

SUB Local_Setup(PGM)

   ' Variable values specific to the web site

   bolRexx = true
   bolSecurityOnly  = true
   bolWebMail = false
   strWebCompany = "Tinctoria Cottage"
   strDefBookLang = "English"

END SUB

FUNCTION GetDbPswd
   GetDbPswd = "Password=fci1948mbrs"
END FUNCTION
%>
