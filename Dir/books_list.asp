<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/book_select.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/book_navigation_details.inc" -->

<%
    DIM intSub, strHld, bolExpFnd, bolDisAll, intCols, strWidth

    Call System_Setup("NONE")

    bolBuildAryOnly = true

    strMaint = "N"

    If Request.QueryString("ALL") = "Y" Then
       bolDisAll = true
       intCols = 550
       intCS = 7
    Else
       intCS = 4
       intCols = 350
    End If

    If Request.QueryString("ND") = "Y" Then
       intCS = intCS - 1
       strWidth = ""
    Else
       strWidth = " width=300"
    End If

    If Application("EnableSrch") = "Y" Then
       If Request.Form("SEARCH") <> "Y" and Request.QueryString("SRCH") = "Y" Then
          SESSION("AS") = "Y" 'Turn on Advance Search Options
          If Request.QueryString("SRCH") = "Y" Then
             bolSrchAryOnly = true
             Call Check_Search_Criteria("Y")
          End If
          SESSION("CAT") = ""
          call Setup_Web_Page("Enter Search Criteria",1)
          Response.Write "<tr><td>" & vbCrLf
          intSrchRows = 5
          bolBottomButton = true
          Call Insert_Search_Prompt("books_list.asp?ND="&Request.QueryString("ND")&"&ALL="&Request.QueryString("ALL"))
          Response.Write "</td></tr>" & vbCrLf
          Call Insert_Navigation_Details("S")
          strMaint = "Y"
          call Wrapup_Web_Page
          Response.End
       End If
    End If

    Call Check_Search_Criteria("Y")

    call Setup_Web_Page(strWebCompany & " Book Listing",intCS)

    aryDbDtls(2) = Request.QueryString("S")

    Call Get_DB_Entries

    strHld = "<tr><td align='center'><b>Title</b></td></td><td width=10></td><td align='center'><b>Author</b></td>"
    If Request.QueryString("ND") <> "Y" Then
       strHld = strHld & "<td align='center'><b>Description</b></td>"
    End If
    If bolDisAll Then
       strHld = strHld & "<td align='center'><b>Picture File</b></td><td align='center'><b>Category</b></td><td align='center'><b>Status</b></td>"
    End If
    Response.Write strHld & "</tr>" & vbCrLf
    Response.Write "<tr><td colspan='" & intCS & "'><hr align='left' size='5' noshade></td></tr>" & vbCrLf

    bolUpdFlds = true
    bolNoFormat = true

    For intSub = 1 to aryRecData(0,0)
       call Get_DB_Record(intSub)
       call Build_Book_Details
       Response.Write "<tr>" & vbCrLf
       Response.Write "<td" & strWidth & " valign='top'>" & strTitle & "</td>" & vbCrLf
       Response.Write "<td></td>" & vbCrLf
       Response.Write "<td" & strWidth & " valign='top'>" & strAuthor & "</td>" & vbCrLf
       If Request.QueryString("ND") <> "Y" Then
          Response.Write "<td width=" & intCols & " >" & strDetails & "</td>" & vbCrLf
       End If
       If bolDisAll Then
          Response.Write "<td nowrap valign='top'>" & strPicFile & "</td>" & vbCrLf
          Response.Write "<td valign='top'>" & strCategory & "</td>" & vbCrLf
          If bolInactive Then
             strHld = "<font color='red'>Inactive</font>"
          Else
             strHld = "Active"
          End If
          Response.Write "<td valign='top'>" & strHld & "</td>" & vbCrLf
       End If
       Response.Write "</tr>" & vbCrLf
       If Request.QueryString("ND") <> "Y" Then
          Response.Write "<tr><td colspan=" & intCS & "><hr size='1'></td></tr>" & vbCrLf
       End If
    Next

    Response.Write "<tr><td><p>&nbsp;</p></td></tr>" & vbCrLf

    bolNoLinks = true
    call Wrapup_Web_Page

%>
