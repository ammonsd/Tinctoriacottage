
<!--#INCLUDE FILE="../dir/maint_links.inc" -->
<!--#INCLUDE FILE="../dir/error_check.inc" -->

<%
    DIM bolNoTable, intTblCols, intCP, intCS, bolGoToUrl, strTypeGoto, strAskYN, bolNoLinks, strTblClr, strFSize
    DIM bolNoBGImg, strBodyScript

SUB Setup_Web_Page(PgTitle,iCols)

    DIM strCols, strHld

    if isnumeric(intCP) = false then
       intCP = 0
    End If
    if isnumeric(intCS) = false then
       intCS = 0
    End If

    intTblCols = iCols
    If intTblCols > "1" Then
       strCols = " colspan='" & intTblCols & "'"
    End If

    Response.Write "<html>" & vbCrLf
    Response.Write "<head>" & vbCrLf
    Response.Write "<title></title>" & vbCrLf

    If bolEditError Then
       strPopUpTitle = "Data Input Errors"
    End If
    Call Check_For_PopUp_Msg

    If bolGoToUrl Then
       If strTypeGoTo = "NW" Then
          Call GoUrl_Code_NewWindow
       ElseIf strTypeGoto = "NF" Then
          Call GoUrl_Code
       Else
          Call GoUrl_Code_Frame
       End If
    End If
    If strAskYN <> "" Then
          Call AskYN_Form(strAskYN)
    End If
    Response.Write "</head>" & vbCrLf

    Response.Write aryErrorText(1) & vbCrLf
    If Application("BGCLR") <> "" and (bolNoBGImg or SESSION("NOBGIMG") = true) Then
       strHld = "<body bgcolor='" & Application("BGCLR") & "' link='Navy' vlink='Navy' alink='Navy' text='Black' " & aryErrorText(2)
    Else
       strHld = "<body bgcolor='" & Application("BGCLR") & "' link='Navy' vlink='Navy' alink='Navy' text='Black' background='" & Application("BGIMG") & "'" & aryErrorText(2)
    End If
    If strBodyScript <> "" and aryErrorText(1) = "" Then
        strHld = strHld & " onload = " & strBodyScript
    End IF

    Response.Write strHld & ">" & vbCrLf

    Response.Write aryErrorText(3) & vbCrLf

    If bolNoTable Then
        Exit SUB
    End IF
    If strTblClr <> "" Then
       strHld = " bgcolor='" & strTblClr & "'"
    Else
       strHld = ""
    End If
    If strFSize = "" Then
       strFSize = "+2"
    End If
    Response.Write "<table align='center' border='0' cellpadding='" & intCP & "' cellspacing='" & intCS & "'" & strHld & ">" & vbCrLf
    If PgTitle <> "" Then
       Response.Write "<tr><td align='center'" & strCols & ">" & vbCrLf
       Response.Write "<font size='" & strFSize & "' color='#990000'>" & vbCrLf
       Response.Write "<b>" & PgTitle & "</b>" & vbCrLf
       Response.Write "</font>" & vbCrLf
       Response.Write "</td></tr>" & vbCrLf
       Response.Write "<tr><td" & strCols & " align='center'>" & vbCrLf
       Response.Write "<img src='/" & Application("PGMDIR") & "/graphics/red_line.gif' width='400' height='3' border='0' alt=''>" & vbCrLf
       Response.Write "</td></tr>" & vbCrLf
       Response.Write "<tr><td>" & vbCrLf
       Response.Write "<p>&nbsp;</p>" & vbCrLf
       Response.Write "</td></tr>" & vbCrLf
    End If

End SUB

SUB Wrapup_Web_Page

    If not bolNoLinks Then
       If intTblCols = "" Then
          intTblCols = 1
       End If

       If strMaint = "Y" and SESSION("PSWD") <> "" then
          Call Maintenance_Links(intTblCols)
       End If
    End If

    If not bolNoTable Then
       Response.Write "</table>" & vbCrLf
    End IF

    Response.Write "</body>" & vbCrLf
    Response.Write "</html>" & vbCrLf

End SUB

%>
