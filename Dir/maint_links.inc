<%
    DIM strMaint, strPreSR, strTypeMaint, intCutOff, aryMaintLinks(5,2), bolNewTable, bolNoSepLine

SUB Maintenance_Links(Cols)

    DIM strCols, strOpts, strMenu, intSub

    If SESSION("SECLEVL") < 3 Then
        EXIT SUB
    End If

    If Cols > 1 Then
       strCols = " colspan=" & Cols & "'"
    End If

    If bolNewTable Then
       Response.Write "<table align='center'>" & vbCrLf
    End If

    If not bolNoSepLine Then
       Response.Write "<tr><td" & strCols & "><hr width='100%' size='5' noshade></td></tr>" & vbCrLf
    End If
    Response.Write "<tr bgcolor='#990000'>" & vbCrLf
    Response.Write "<td" & strCols & " align='center' valign='bottom' nowrap>" & vbCrLf

    strOpts = "?MT=Y"

    If SESSION("MENUFN") = "" Then
       Call System_Setup("NONE")
    End If
    strMenu = SESSION("MENUFN")
    If strMenu = "" Then
       strMenu = GetCurPath("")
    End If

    If SESSION("SECLEVL") > 2 and SESSION("MENUFN") = "" Then
       Response.Write "&nbsp;" & vbCrLf
       Response.Write "<a href='directory_maint.asp?TM=A'><font size='-1' color='White'>Add Address</font></a>" & vbCrLf
       Response.Write "&nbsp;<font color='White'><b>|</b></font>&nbsp;" & vbCrLf
       Response.Write "<a href='directory_select.asp?MT=Y&TM=U'><font size='-1' color='White'>Update Address</font></a>" & vbCrLf
       Response.Write "&nbsp;<font color='White'><b>|</b></font>&nbsp;" & vbCrLf

       bolDisplayMsg = false
       strSpecSecGrp = "SecMaint"
       Call Verify_Group_Access("N/R","",5)
       If not bolSecError Then
          Response.Write "<a href='select_security.asp" & strOpts & "'><font size='-1' color='White'>Update Security</font></a>" & vbCrLf
          Response.Write "&nbsp;<font color='White'><b>|</b></font>&nbsp;" & vbCrLf
       End If
    End If

    Response.Write "<a href='" & strMenu & "'><font size='-1' color='White'>Menu</font></a>" & vbCrLf
    If aryMaintLinks(0,0) > 0 Then
       For intSub = 1 to aryMaintLinks(0,0)
          Response.Write "&nbsp;<font color='White'><b>|</b></font>&nbsp;" & vbCrLf
          Response.Write "<a href='" & aryMaintLinks(intSub,1) & "'><font size='-1' color='White'>" & aryMaintLinks(intSub,2) & "</font></a>" & vbCrLf
       Next
    End If
    Response.Write "&nbsp;<font color='White'><b>|</b></font>&nbsp;" & vbCrLf
    Response.Write "<a href='logoff.asp'><font size='-1' color='White'>Logoff</font></a>" & vbCrLf
    Response.Write "</td></tr>" & vbCrLf
    If bolNewTable Then
       Response.Write "</table'>" & vbCrLf
    End If

END SUB
%>

