<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/directory_select.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/select_security_groups.inc" -->
<%

   DIM strPreSecGrps, intSub, intSub2

   Call System_Setup("NONE")
   Call Logon_Check(GetCurPath("")&"directory_security_groups.asp",5,strLogonGrp)

   bolListOnly = true
   bolOthGroups = true
   bolBuildAryOnly = true


   aryDbDtls(2) = "6"
   call Database_Setup
   Call Get_DB_Entries

   bolUpdFlds = true
   intSub2 = 0

   For intSub = 1 to aryRecData(0,0)
      Call Get_DB_Record(intSub)
      If trim(strSecGrps) <> "" and strPreSecGrps <> strSecGrps Then
         intSub2 = intSub2 + 1
         arySecRec(intSub2) = strSecGrps
         strPreSecGrps = strSecGrps
      End If
   Next

   Call Process_Security_Groups
%>
