
<%

FUNCTION EmailDuplicate(EMAIL,ID)

    DIM strSQL, intLoc

    EmailDuplicate = false

    If EMAIL = "" or APPLICATION("DIRTBL") = "" Then
       EXIT FUNCTION
    End If

    intLoc = instr(EMAIL,"<")
    If intLoc > 0 Then
       EMAIL = trim(REPLACE(mid(EMAIL,intLoc+1),">"," "))
    End If

    Call Database_Setup

    strSQL =  " WHERE [eMail] = '" & lcase(EMAIL) & "'"
    If ID <> "" and isNumeric(ID) Then
       strSQL =  strSQL & " AND [ID] <> " & int(ID)
    End If

    Set objRecordSet = Server.CreateObject("ADODB.Recordset")
    objRecordSet.ActiveConnection = strDbConn

    strSQL = "SELECT Count(*) AS RecCnt FROM " & APPLICATION("DIRTBL") & strSQL
    objRecordSet.Open strSQL
    If objRecordSet("RecCnt") > 0 Then
       EmailDuplicate = true
    End If
    objRecordSet.Close
    Set objRecordSet = Nothing

END FUNCTION
%>

