
<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->

<%

SUB Security_Maintenance_Links(Cols)

    DIM strCols, strMenu

    If Cols > 1 Then
       strCols = " colspan=" & Cols & "'"
    End If

    Response.Write "<tr bgcolor='#990000'>" & vbCrLf
    Response.Write "<td" & strCols & " align='center' valign='bottom' nowrap>" & vbCrLf

    If SESSION("MENUFN") = "" Then
       Call System_Setup("NONE")
    End If
    strMenu = SESSION("MENUFN")
    If strMenu = "" Then
       strMenu = GetCurPath("")
    End If

    If SESSION("MENUFN") = "" Then
       bolDisplayMsg = false
       strSpecSecGrp = "SecMaint"
       Call Verify_Group_Access("N/R","",5)
       If not bolSecError Then
          Response.Write "&nbsp;" & vbCrLf
          Response.Write "<a href='security_maint.asp?TM=A'><font size='-1' color='White'>Add Profile</font></a>" & vbCrLf
          Response.Write "&nbsp;<font color='White'><b>|</b></font>&nbsp;" & vbCrLf
          Response.Write "<a href='select_security.asp?TM=U'><font size='-1' color='White'>Update Profile</font></a>" & vbCrLf
          Response.Write "&nbsp;<font color='White'><b>|</b></font>&nbsp;" & vbCrLf
          Response.Write "<a href='list_security.asp'><font size='-1' color='White'>List Profiles</font></a>" & vbCrLf
          Response.Write "&nbsp;<font color='White'><b>|</b></font>&nbsp;" & vbCrLf
       End If
    End If

    Response.Write "<a href='" & strMenu & "'><font size='-1' color='White'>Menu</font></a>" & vbCrLf
    If strSecurHelpFn <> "" Then
       Response.Write "&nbsp;<font color='White'><b>|</b></font>&nbsp;" & vbCrLf
       Response.Write "<a href='" & strSecurHelpFn & "' target='_blank'><font size='-1' color='White'>Help</font></a>" & vbCrLf
    End If
    Response.Write "&nbsp;<font color='White'><b>|</b></font>&nbsp;" & vbCrLf
    Response.Write "<a href='logoff.asp'><font size='-1' color='White'>Logoff</font></a>" & vbCrLf
    Response.Write "&nbsp;" & vbCrLf

    ' Bypass Maintenance Linkds
    bolNoLinks = true

END SUB
%>

