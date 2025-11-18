<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
Server.ScriptTimeout = 99999
%>

<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/directory_select.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/db_gen_maint.inc" -->

<%

   DIM intSub, RS

   Call System_Setup("NONE")
   Call Logon_Check(GetCurPath("")&"convert_members.asp",5,strLogonGrp)

   Call Database_Setup

   Set RS = Server.CreateObject("ADODB.Recordset")
   RS.ActiveConnection = strDbConn

   RS.Source = "SELECT * FROM MEMBERS ORDER BY [Last Name]"

   RS.LockType = adLockReadOnly
   RS.Open

   intSub = 0

   If Not RS.EOF Then
      Do While Not RS.EOF
         intSub = intSub + 1
         aryRecData(intSub,1) = RS.Fields("ID")
         aryRecData(intSub,2) = RS.Fields("EMAIL")
         aryRecData(intSub,3) = RS.Fields("Company")
         aryRecData(intSub,9) = RS.Fields("First Name")
         aryRecData(intSub,10) = RS.Fields("Last Name")
         aryRecData(intSub,11) = RS.Fields("Security Groups")
         aryRecData(intSub,12) = RS.Fields("Keywords")
         aryRecData(intSub,13) = RS.Fields("Contact Info")
         RS.MoveNext
      Loop
   End If

   RS.Close
   Set RS = Nothing

   aryRecData(0,0) = intSub

   bolUpdFlds = true

   For intSub = 1 to aryRecData(0,0)
      Call Get_DB_Record(intSub)
      If trim(strCompany) = "" or isnull(strCompany)Then
         strCompany = "Not Assigned"
      End If
      call Add_Record
   Next

   Response.Write "Number of Members Processed: " & aryRecData(0,0) & "<br>"

SUB Add_Record

   DIM intSub, bolDbInactive

   bolDbInactive = false

   intLastSeq = 5
   aryGenRec(0,1) = strEmail
   aryGenRec(1,1) = strFName
   aryGenRec(2,1) = strLName
   aryGenRec(3,1) = strSecGrps
   aryGenRec(4,1) = strKeywords
   aryGenRec(5,1) = strContactInfo

   For intSub = 0 to intLastSeq
      aryGenRec(intSub,0) = intSub
      aryGenRec(intSub,2) = "T"
   Next

   call AddGenTextRecords("DIRECTORY",strCompany,bolDbInactive)

   Response.Write trim(strFName & " " & strLName) & " Added to Database<br>"

END SUB
%>
