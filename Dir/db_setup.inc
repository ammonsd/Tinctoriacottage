

<!--#INCLUDE FILE="../dir/adovbs.inc" -->

<%

   DIM strDbConn, objRecordSet

SUB Database_Setup

   DIM strDbUser, strDbPswd, strDbName

   If strCurDir = "" Then
      Call System_Setup("NONE")
   End If

   strDbName = SESSION("HOMEDIR")
   If strDbName = "" Then
      strDbName = mid(Request.ServerVariables("SCRIPT_NAME"),2)
      strDbName = left(strDbName,instr(strDbName,"/")-1)
   End If
   If lcase(strDbName) = lcase(APPLICATION("PGMDIR")) Then
      strDbName = ""
   Else
      strDbName = "_" & strDbName
   End If
   strDbName = "Directory" & strDbName & ".mdb"
   If strDbName <> SESSION("DBNAME") Then
      SESSION("DBNAME") = strDbName
      If CheckFileExist(Server.MapPath("/cgi-bin/" & SESSION("DBNAME"))) <> 0 then
         SESSION("DBNAME") = APPLICATION("DBNAME")
      End If
   End If

   strDbUser = ""
   strDbPswd = GetDbPswd
   strDbConn = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.Mappath("\" & strFileLoc & "\") & "\"  & SESSION("DBNAME") & ";" & strDbUser & ";" & strDbPswd

END SUB

%>
