<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/member_select.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/member_navigation_details.inc" -->

<%
    DIM intSub, strHld, bolExpFnd, bolDisAll, strAddress

    Call System_Setup("NONE")

    bolBuildAryOnly = true

    strMaint = "N"

    intCS = 5

    If Application("EnableSrch") = "Y" Then
       If Request.Form("SEARCH") <> "Y" and Request.QueryString("SRCH") = "Y" Then
          SESSION("AS") = "Y" 'Turn on Advance Search Options
          If Request.QueryString("SRCH") = "Y" Then
             bolSrchAryOnly = true
             Call Check_Search_Criteria
          End If
          call Setup_Web_Page("Enter Search Criteria",1)
          Response.Write "<tr><td>" & vbCrLf
          intSrchRows = 5
          bolBottomButton = true
          Call Insert_Search_Prompt("members_list.asp")
          Response.Write "</td></tr>" & vbCrLf
          Call Insert_Navigation_Details("S")
          strMaint = "Y"
          call Wrapup_Web_Page
          Response.End
       End If
    End If

    Call Check_Search_Criteria

    call Setup_Web_Page(strCompany & " Member Listing",intCS)

    Call Get_DB_Entries

    strHld = "<tr><td align='center'><b>Name</b></td></td><td align='center'><b>Email</b></td>"
    strHld = strHld & "<td align='center'><b>Address</b></td><td align='center'><b>Member<br>Since</b></td>"
    Response.Write strHld & "</tr>" & vbCrLf
    Response.Write "<tr><td colspan='" & intCS & "'><hr align='left' size='5' noshade></td></tr>" & vbCrLf

    bolUpdFlds = true

    For intSub = 1 to aryRecData(0,0)
       call Get_DB_Record(intSub)
       Response.Write "<tr>" & vbCrLf
       Response.Write "<td valign='top'>" & trim(strFName & " " & strLName) & "</td>" & vbCrLf
       Response.Write "<td valign='top'>" & strEmail & "</td>" & vbCrLf
       strAddress = ""
       If strStreet <> "" Then
          strAddress = strAddress & strStreet & "<br>"
       End If
       strHld = " "
       If strCity <> "" Then
          strAddress = strAddress & strCity
          strHld = ", "
       End If
       If strState <> "" Then
          strAddress = strAddress & strHld & strState
       End If
       If strZip <> "" Then
          strAddress = strAddress & " " & strZip
       End If
       If strCountry <> "United States" Then
          If strAddress <> "" and right(strAddress,1) <> ">" Then
             strAddress = strAddress & "<br>"
          End If
          strAddress = strAddress & strCountry
       End If
       Response.Write "<td>" & strAddress & "</td>" & vbCrLf
       Response.Write "<td valign='top'>" & GetWords(strDateAdded,1,1) & "</td>" & vbCrLf
       Response.Write "</tr>" & vbCrLf
    Next

    Response.Write "<tr><td><p>&nbsp;</p></td></tr>" & vbCrLf

    bolNoLinks = true
    call Wrapup_Web_Page

%>
