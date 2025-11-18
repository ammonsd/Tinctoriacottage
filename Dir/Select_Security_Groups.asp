<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/security_select.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/select_security_groups.inc" -->
<%

   DIM strPreSecGrps, intSub

   If Request.QueryString("L") = "Y"  Then
      Response.Redirect("Directory_Security_Groups.asp")
   End If

   strSpecSecGrp = "SecMaint"
   Call System_Setup("NONE")
   Call Logon_Check(GetCurPath("")&"select_security_groups.asp",5,strLogonGrp)

   Call Database_Setup

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.ActiveConnection = strDbConn

   objRecordSet.Source = "SELECT [SECURITY GROUPS] FROM " & APPLICATION("SECURTBL") & " ORDER BY [SECURITY GROUPS]"

   objRecordSet.LockType = adLockReadOnly
   objRecordSet.Open

   intSub = 0

   If Not objRecordSet.EOF Then
      objRecordSet.MoveFirst
      Do While Not objRecordSet.EOF
         If trim(objRecordSet.Fields("SECURITY GROUPS")) <> "" and strPreSecGrps <> objRecordSet.Fields("SECURITY GROUPS") Then
            intSub = intSub + 1
            arySecRec(intSub) = UnformatSecurityGroups(objRecordSet.Fields("SECURITY GROUPS"))
            strPreSecGrps = objRecordSet.Fields("SECURITY GROUPS")
         End If
         objRecordSet.MoveNext
      Loop
   End If

   objRecordset.Close
   Set objRecordSet = Nothing

   Call Process_Security_Groups
%>
