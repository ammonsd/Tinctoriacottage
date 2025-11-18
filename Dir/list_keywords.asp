<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/build_where_sql.inc" -->
<!--#INCLUDE FILE="../dir/array_sort.inc" -->
<%

   DIM strPrevKeys, aryKws(100), intKwSub, intSub, intMaxCols, intCol, strIE, bolNoDisplay, bolNoKeysFnd

   If SESSION("BPG") = "" Then
      SESSION("BPG") = 1
   End If

   If SESSION("RQL") = "" Then
      SESSION("RQL") = 3
   End If

   Call System_Setup("NONE")
   Call Logon_Check(GetCurPath("")&"list_keywords.asp",SESSION("RQL"),strLogonGrp)

   Call Database_Setup

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.ActiveConnection = strDbConn

   bolSesKWOnly = true

   Call Build_Where_SQL

   objRecordSet.Source = "SELECT * FROM " & Request.QueryString("TN") & strWhrSQL & " ORDER BY [KEYWORDS]"

   objRecordSet.LockType = adLockReadOnly
   objRecordSet.Open

   strPrevKeys = "|"

   bolDoPhrases = true
   bolNoKeysFnd = true

   If SESSION("KW") <> "" Then
      bolNoDisplay = true
      Call Process_Keys(ReplChars(SESSION("KW")," ",","))
   End If

   bolNoDisplay = false

   If Not objRecordSet.EOF Then
      objRecordSet.MoveFirst
      intKwSub = 0
      Do While Not objRecordSet.EOF
         Call Process_Keys(objRecordSet.Fields("Keywords"))
         Call Process_Keys(objRecordSet.Fields("Company"))
         objRecordSet.MoveNext
      Loop
      If bolNoKeysFnd Then
         Response.Write "<div align='center'><b>No Keywords Assigned</b><br><br></div>" & vbCrLf
      Else
         Response.Write "<div align='center'><font color='red'>" & vbCrLf
         Response.Write "Besides names and email domains, any combination of the following keywords can be used for <b>Directory Searches</b>.<br>" & vbCrLf
         Response.Write "Search using phrases by enclosing the phrase with quotation marks" & strIE & ".<br><br>" & vbCrLf
         Response.Write "</font></div>" & vbCrLf
         Response.Write "<table align='center' border='0' cellpadding='0' cellspacing='0'><tr>" & vbCrLf
         arySorted = SortArray(aryKws,0)
         Erase aryKws
         intMaxCols = ROUND((intNbrEntries / 16) + .5)
         If intMaxCols < 1 Then
            intMaxCols = 1
         End If
         intCol = 0
         For intSub = 0 to intNbrEntries - 1
            intCol = intCol + 1
            Response.Write "<td>" & arySorted(intSub) & "&nbsp;&nbsp;&nbsp;&nbsp;</td>" & vbCrLf
            If intCol = intMaxCols and intSub < intNbrEntries - 1 Then
               intCol = 0
               Response.Write "</tr><tr>" & vbCrLf
            End If
         Next
         Response.Write "</tr></table>" & vbCrLf
      End If
   End If
   objRecordset.Close
   Set objRecordSet = Nothing

   If not bolNoKeysFnd Then
      Response.Write "<br><br><br><div align='center'>" & vbCrLf
      If Request.QueryString("NC") = "Y"  Then
         Response.Write "<input type='button' value='Back' onClick='history.go(-" & SESSION("BPG") & ")'>" & vbCrLf
      Else
         Response.Write "<input type='button' value='Close Window' onClick='window.close()'>" & vbCrLf
      End If
      Response.Write "</div>" & vbCrLf
   End If
   SESSION("BPG") = ""

SUB Process_Keys(ChkKey)

   DIM intSub, strKW, strKeys

   strKeys = RemoveLineBreaks(ChkKey)

   If trim(strKeys) <> "" and not isnull(strKeys) Then
      intSub = 1
      strKW = GetWords(strKeys,intSub,1)
      do while strKW <> ""
         If left(strKW,1) = "-" Then
            strKW = mid(strKW,2)
         End If
         If instr(strPrevKeys,"|"&lcase(strKW)&"|" ) = 0 Then
            strPrevKeys = strPrevKeys & lcase(strKW) & "|"
            If not bolNoDisplay Then
               aryKws(intKwSub) = strKW
               bolNoKeysFnd = false
               intKwSub = intKwSub + 1
               If strIE = "" Then
                  If instr(strKW," ") > 0 Then
                     strIE = " (i.e. " & CHR(34) & strKW & CHR(34) & ")"
                  End If
               End If
            End If
         End If
         intSub = intSub + intRetWords
         strKW = GetWords(strKeys,intSub,1)
      Loop
   End If

END SUB
%>
