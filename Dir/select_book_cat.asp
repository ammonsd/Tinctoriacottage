<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/book_select.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->

<%
    DIM intSub, strHld, X, intCntr, strGoTo, strChk

    Call System_Setup("NONE")

    If (Request.Form("UpdFile") = "Y" or Request.QueryString("UpdFile") = "Y") Then
       strGoTo = Request.Form("P")
       For intSub = 1 to Request.Form("Cnt")
         X = "Cat" & intSub
         If Request.Form(X) <> "" Then
            If strHld = "" Then
               strHld = "?RKW=" & Request.Form(X)
            Else
               strHld = strHld & "%20%OR%20%" & Request.Form(X)
            End If
         End If
       Next
       If strHld = "" Then
          strHld = "?RKW=*ALL*"
       End If
       Response.Redirect(strGoTo & strHld)
    End If

    strGoTo = Request.QueryString("P")
    If strGoTo = "" Then
       strGoTo = "select_book_cat.asp"
    End If

    bolBuildAryOnly = true
    call Build_Category_Selection(strCategory)

    strTblClr = "#b0c4de"
    strFSize = "+1"

    Response.Write "<img src='/" & Application("PGMDIR") & "/graphics/spacer.gIf' width='0' height='100' border='0' alt=''>"

    Response.Write "<form action='select_book_cat.asp' method='post'>" & vbCrLf
    Response.Write "<input type='hidden' name='UpdFile' value='Y'>" & vbCrLf
    Response.Write "<input type='hidden' name='Cnt' value='" & aryCategory(0) &  "'>" & vbCrLf
    Response.Write "<input type='hidden' name='P' value='" & strGoTo &  "'>" & vbCrLf

    If Request.QueryString("P") = "" Then
       strChk = " checked"
    End If

    call Setup_Web_Page("Available Categories",3)

    intCntr = 0
    For intSub = 1 to aryCategory(0)
      If intCntr = 0 Then
         Response.Write "<tr>" & vbCrLf
      End If
      intCntr = intCntr + 1
      Response.Write "<td>" & vbCrLf
      Response.write "<input type='checkbox' name='Cat" & intSub & "' value='" & aryCategory(intSub) & "'" & strChk & ">" & vbCrLf
      Response.Write trim(aryCategory(intSub))& vbCrLf
      Response.write "</td>" & vbCrLf
      If intCntr = 3 Then
         Response.Write "</tr>" & vbCrLf
         intCntr = 0
      End If
    Next
    If intCntr > 0 Then
       Response.Write "</tr>" & vbCrLf
    End If

    Response.Write "<tr><td width='100%' colspan='3'>" & vbCrLf
    Response.Write "<hr width='100%' size='5' noshade>" & vbCrLf
    Response.Write "</td></tr>" & vbCrLf
    Response.Write "<tr><td align='center' colspan=3>" & vbCrLf
    Response.Write "<input type='submit' value='Select'></td></tr>" & vbCrLf

    call Wrapup_Web_Page
    Response.Write "</form>" & vbCrLf
%>
