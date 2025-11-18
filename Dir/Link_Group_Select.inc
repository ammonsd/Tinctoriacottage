

<!--#INCLUDE FILE="../includes/setup_link_groups.inc" -->

<%

    DIM aryLinkGroup(50,2)

SUB Build_Link_Group_Select(ID,PGM)

    DIM strCurGrp, intSub, strHld, strParms, X

    If PGM <> "" Then
       For Each X in Request.QueryString
          strParms = strParms & "&" & X & "=" & Request.QueryString(X)
       Next
    End If

    call Build_Link_Group_Array
    call GoUrl_Code

    For intSub = 1 to aryLinkGroup(0,0)
        If aryLinkGroup(intSub,2) = ID Then
           strCurGrp = aryLinkGroup(intSub,1)
           aryLinkGroup(intSub,1) = ""
           EXIT FOR
        End If
    Next

    If strCurGrp = "" Then
       ID = ""
    End IF
    If PGM <> "" Then
       Response.Write "<select name='Url' onChange='goUrl(this)'>" & vbCrLf
    Else
       Response.Write "<select name='GID'>" & vbCrLf
   End If

   Response.Write "<option selected value='" & ID & "'>" & strCurGrp & "</option>"  & vbCrLf

    For intSub = 1 to aryLinkGroup(0,0)
        If aryLinkGroup(intSub,1) <> "" Then
           If PGM <> "" Then
              strHld = ReplParmChars(aryLinkGroup(intSub,1))
              Response.Write "<option value='" & PGM & aryLinkGroup(intSub,2) & "&GRP=" & strHld & strParms & "'>" & aryLinkGroup(intSub,1) & "</option>" & vbCrLf
           Else
              Response.Write "<option value='" & aryLinkGroup(intSub,2) & "'>" & aryLinkGroup(intSub,1) & "</option>" & vbCrLf
           End If
        End If
    Next
    Response.Write "</select>" & vbCrLf
END SUB
%>
