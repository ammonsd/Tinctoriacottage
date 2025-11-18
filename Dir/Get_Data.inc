
<%

    DIM aryData(4,500), strFldAlign, strGetData

    '  Aray Values
    ' 1 - Prompt
    ' 2 - Default Data
    ' 3 - Field Type
    ' 4 - Request Form Name (same as 1 if blank)

SUB Get_Data(PGM,TITLE,BOXSIZE)

   DIM intSub, strTyp, strFormName, strColAlign, strColSize, strSep

   If ucase(strFldAlign) = "R" Then
      strColAlign = " align='right'"
   ElseIf ucase(strFldAlign) = "L" Then
      strColAlign = " align='left'"
   ElseIf ucase(strFldAlign) = "C" Then
      strColAlign = " align='center'"
   End If

   If not isnumeric(BOXSIZE) or int(BOXSIZE) = 0 Then
      BOXSIZE = 30
   End If
   SESSION("GETDATA") = "Y"

%>

<html>
<head>
<title><%=Title%></title>
<%Check_For_PopUp_Msg%>
</head>
<%Response.Write aryErrorText(1) & vbCrLf%>
<body bgcolor="#6D6D6D"<%=aryErrorText(2)%>>
<%Response.Write aryErrorText(3) & vbCrLf%>
<div align="center">
<div align="CENTER">
<br><br><br><br><br><br>
<table border="1" cellspacing="0" cellpadding="2" bordercolor="#000080">
<tr>
<td align='CENTER' valign="TOP" bgcolor="#000095">
<font color="white">
<b><%=Title%></b>
</font>
</td></tr>
<tr>
<td valign="TOP" bgcolor="#B4B4B4">
<font face="Arial">
<div align="CENTER">
<table border="0" cellspacing="0" cellpadding="2">
<form action="<%=PGM%>" method="post">

<%
For intSub = 1 to aryData(0,0)
   If aryData(3,intSub) = "H" Then
      If instr(aryData(2,intSub),"'") > 0 Then
         strSep = CHR(34)
      Else
         strSep = "'"
      End If
      strFormName = aryData(4,intSub)
      If strFormName = "" Then
         strFormName = aryData(1,intSub)
      End If
      Response.Write "<input type='hidden' name='" & strFormName & "' value=" & strSep & aryData(2,intSub) & strSep & "> " & vbCrLf
   End If
Next
For intSub = 1 to aryData(0,0)
   If aryData(3,intSub) <> "H" Then
      If instr(aryData(2,intSub),"'") > 0 Then
         strSep = CHR(34)
      Else
         strSep = "'"
      End If
      strColSize = " size='" & BOXSIZE & "' "
      If aryData(3,intSub) = "P" Then
         strTyp = "password"
      ElseIf aryData(3,intSub) = "C" Then
         strTyp = "checkbox"
      ElseIf aryData(3,intSub) = "R" Then
         strTyp = "radio"
         strColSize = ""
      ElseIf aryData(3,intSub) = "TA" Then
         strColSize = Round(BOXSIZE * .75)
      Else
         strTyp = "text"
      End If
      strFormName = aryData(4,intSub)
      If strFormName = "" Then
         strFormName = aryData(1,intSub)
      End If
      Response.Write "<tr><td no wrap" & strColAlign & ">" & vbCrLf
      Response.Write "<font size='-1'>" & vbCrLf
      Response.Write "<b>" & aryData(1,intSub) & ":</b>" & vbCrLf
      Response.Write "</font>" & vbCrLf
      Response.Write "</td><td nowrap>" & vbCrLf
      If aryData(3,intSub) = "TA" Then
         Response.Write "<textarea cols=" & strColSize & " rows=3 name='" & strFormName & "'>" & aryData(2,intSub) & "</textarea>" & vbCrLf
      Else
         Response.Write "<input type='" & strTyp & "' name='" & strFormName & "'" & strColSize & "value=" & strSep & aryData(2,intSub) & strSep & "> " & vbCrLf
      End If
      Response.Write "</td></tr>" & vbCrLf
   End If
Next%>
<tr><td colspan='2'><hr size='2'></td></tr>
<tr><td align='center' colspan='2'>
<input type='Submit' name='PROCESS' value='Process'>
</td></tr>
</form>
</td></tr>
</table>
</td></tr>
</table>
</td></table>
</td></table>
</div>
</body>
</html>

<%
End SUB
%>
