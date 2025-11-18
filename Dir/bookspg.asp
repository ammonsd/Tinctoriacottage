

<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/book_select.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->

<%

    DIM intSub, intSub2, strHld, bolPrev, strPopUpDetails, intCnt, intFactor, aryFactors(30,2), bolDisplayTnPics
    DIM strAuthor2, strTitle2, strDetails2

    Call System_Setup("NONE")

    If APPLICATION("BOOKLOGON") = "Y" Then
       strSpecSecGrp = "BOOKSPG"
       Call Logon_Check("bookspg.asp",3,"BOOKSPG")
    End If

    If Request.QueryString("P") = "" and Request.Form("SEARCH") <> "Y" and Request.QueryString("RKW") = "" Then
       Session.Contents.Remove("SRCHAREA")
       Session.Contents.Remove("SRCHTYPE")
       Session.Contents.Remove("SRCHKW")
       Session.Contents.Remove("SRCHSQL")
       Session.Contents.Remove("PG")
       Session.Contents.Remove("LID")
       Session.Contents.Remove("PGS")
    End If

    If Application("TnPics") = "N" and SESSION("DEMO") <> "Y" Then
       bolDisplayTnPics = false
    Else
       bolDisplayTnPics = true
    End If

    intMaxRecs = Application("BOOKPG")

    aryFactors(0,0) = 30
    intCnt = 1
    For intSub = 1 to aryFactors(0,0)
       If bolDisplayTnPics Then
          intCnt = intCnt + 1.25
       Else
          For intSub2 = 1 to intMaxRecs step 7
             If intSub2 > intSub Then
                EXIT FOR
             End If
             intCnt = intSub + 5 + ((intSub2 - 1) * 7)
          Next
       End If
       aryFactors(intSub,0) = intCnt
       aryFactors(intSub,1) = intSub
    Next

    Call Check_Search_Criteria("N")

    bolBuildAryOnly = true
    strPageDir = Request.QueryString("P")
    intPgNbr = SESSION("PG")
    intLID = SESSION("LID")
    intPgs = SESSION("PGS")
    aryDbDtls(2) = "3,4" 'Sort by Tile, Author

    do while strPageDir = strPageDir 'Simulate Forever
       call Get_DB_Entries
       If strPageDir <> "L" or (strPageDir = "L" and not bolMore) Then
          EXIT DO
       End If
    loop

   bolPrev = false
   IF intPgNbr <> "" Then
      IF int(intPgNbr) > 1 Then
         bolPrev = true
      End if
   End if

   SESSION("PG") = intPgNbr
   SESSION("LID") = intLID
   SESSION("PGS") = intPgs
%>

<html><!-- #BeginTemplate "/Templates/template.dwt" -->
<head>
<!-- #BeginEditable "doctitle" -->
<title>Books - <%=strWebCompany%></title>
<!-- #EndEditable -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-------------------- BEGIN COPYING THE JAVASCRIPT SECTION HERE ------------------->

<%
   strTitleClr = "#a5a673"
   strBgClr = "#c6c79c" 'Description Popup
   intBoxLeft = 200
   intBoxWidth = 750
   Call SetUp_PopUp_Msg
%>

<script language="JavaScript">
<!-- hide this script from non-javascript-enabled browsers

/* Function that swaps images. */

function di20(id, newSrc) {
    var theImage = FWFindImage(document, id, 0);
    if (theImage) {
        theImage.src = newSrc;
    }
}

/* Functions that track and set toggle group button states. */

function FWFindImage(doc, name, j) {
    var theImage = false;
    if (doc.images) {
        theImage = doc.images[name];
    }
    if (theImage) {
        return theImage;
    }
    if (doc.layers) {
        for (j = 0; j < doc.layers.length; j++) {
            theImage = FWFindImage(doc.layers[j].document, name, 0);
            if (theImage) {
                return (theImage);
            }
        }
    }
    return (false);
}


function setCookie(name, value) {
    document.cookie = name + "=" + escape(value);
}

function getCookie(Name) {
    var search = Name + "=";
    var retVal = "";
    if (document.cookie.length > 0) {
        offset = document.cookie.indexOf(search);
        if (offset != -1) {
            end = document.cookie.indexOf(";", offset);
            offset += search.length;
            if (end == -1) {
                end = document.cookie.length;
            }
            retVal = unescape(document.cookie.substring(offset, end));
        }
    }
    return (retVal);
}

function InitGrp(grp) {
    var cmd = false;
    if (getCookie) {
        cmd = getCookie(grp);
    }
    if (cmd) {
        eval("GrpDown(" + cmd + ")");
        eval("GrpRestore(" + cmd + ")");
    }
}

function FindGroup(grp, imageName) {
    var img = FWFindImage(document, imageName, 0);
    if (!img) {
        return (false);
    }
    var docGroup = eval("document.FWG_" + grp);
    if (!docGroup) {
        docGroup = new Object();
        eval("document.FWG_" + grp + " = docGroup");
        docGroup.theImages = new Array();
    }
    if (img) {
        var i;
        for (i = 0; i < docGroup.theImages.length; i++) {
            if (docGroup.theImages[i] == imageName) {
                break;
            }
        }
        docGroup.theImages[i] = imageName;
        if (!img.atRestSrc) {
            img.atRestSrc = img.src;
            img.initialSrc = img.src;
        }
    }
    return (docGroup);
}

function GrpDown(grp, imageName, downSrc, downOver) {
    if (!downOver) {
        downOver = downSrc;
    }
    var cmd = "'" + grp + "','" + imageName + "','" + downSrc + "','" + downOver + "'";
    setCookie(grp, cmd);
    var docGroup = FindGroup(grp, imageName, false);
    if (!docGroup || !downSrc) {
        return;
    }
    obj = FWFindImage(document, imageName, 0);
    var theImages = docGroup.theImages;
    if (theImages) {
        for (i = 0; i < theImages.length; i++) {
            var curImg = FWFindImage(document, theImages[i], 0);
            if (curImg && curImg != obj) {
                curImg.atRestSrc = curImg.initialSrc;
                curImg.isDown = false;
                obj.downOver = false;
                curImg.src = curImg.initialSrc;
            }
        }
    }
    obj.atRestSrc = downSrc;
    obj.downOver = downOver;
    obj.src = downOver;
    obj.isDown = true;
}

function GrpSwap(grp) {
    var i, j = 0, newSrc, objName;
    var docGroup = false;
    for (i = 1; i < (GrpSwap.arguments.length - 1); i += 2) {
        objName = GrpSwap.arguments[i];
        newSrc = GrpSwap.arguments[i + 1];
        docGroup = FindGroup(grp, objName);
        if (!docGroup) {
            continue;
        }
        obj = FWFindImage(document, objName, 0);
        if (!obj) {
            continue;
        }
        if (obj.isDown) {
            if (obj.downOver) {
                obj.src = obj.downOver;
            }
        } else {
            obj.src = newSrc;
            obj.atRestSrc = obj.initialSrc;
        }
        obj.skipMe = true;
        j++;
    }
    if (!docGroup) {
        return;
    }
    theImages = docGroup.theImages;
    if (theImages) {
        for (i = 0; i < theImages.length; i++) {
            var curImg = FWFindImage(document, theImages[i], 0);
            if (curImg && curImg.atRestSrc && !curImg.skipMe) {
                curImg.src = curImg.atRestSrc;
            }
            curImg.skipMe = false;
        }
    }
}

function GrpRestore(grp) {
    var docGroup = eval("document.FWG_" + grp);
    if (!docGroup) {
        return;
    }
    theImages = docGroup.theImages;
    if (theImages) {
        for (i = 0; i < theImages.length; i++) {
            var curImg = FWFindImage(document, theImages[i], 0);
            if (curImg && curImg.atRestSrc) {
                curImg.src = curImg.atRestSrc;
            }
        }
    }
}

// stop hiding -->
</script>

<!-------------------------- STOP COPYING THE JAVASCRIPT HERE -------------------------->

<body bgcolor="#FFFFCC" text="#666633" link="#666633" vlink="#666633" background="/graphics/tincbg.jpg">
<table width="100%" border="0">
  <tr>
    <td width="110"><img src="/graphics/shim.gif" width="110" height="1"></td>
    <td width="850" align="center"><!-- #BeginEditable "header" --><img src="/graphics/bookban.gif" width="300" height="90"><!-- #EndEditable --></td>
  </tr>
  <tr>
    <td width="110"></td>
    <td width='850'>
    <table>
    <tr><td nowrap width='450'>
<%IF SESSION("DEMO") <> "Y" Then%>
        <a href="/new/home.html" onMouseOut="GrpRestore('FwSimpleGroup');"  onMouseOver="GrpSwap('FwSimpleGroup','nav5_r1_c1','/Graphics/nav5_r1_c1_F2.gif')"  onClick="GrpDown('FwSimpleGroup','nav5_r1_c1','/Graphics/nav5_r1_c1_F3.gif','/Graphics/nav5_r1_c1_F3.gif')" ><img name="nav5_r1_c1" src="/Graphics/nav5_r1_c1.gif" width="47" height="27" border="0"></a>
        <a href="/new/linenpg.shtml" onMouseOut="GrpRestore('FwSimpleGroup');"  onMouseOver="GrpSwap('FwSimpleGroup','nav5_r1_c3','/Graphics/nav5_r1_c3_F2.gif')"  onClick="GrpDown('FwSimpleGroup','nav5_r1_c3','/Graphics/nav5_r1_c3_F3.gif','/Graphics/nav5_r1_c3_F3.gif')" ><img name="nav5_r1_c3" src="/Graphics/nav5_r1_c3.gif" width="95" height="27" border="0"></a>
        <a href="/new/wovenpg.shtml" onMouseOut="GrpRestore('FwSimpleGroup');"  onMouseOver="GrpSwap('FwSimpleGroup','nav5_r1_c5','/Graphics/nav5_r1_c5_F2.gif')"  onClick="GrpDown('FwSimpleGroup','nav5_r1_c5','/Graphics/nav5_r1_c5_F3.gif','/Graphics/nav5_r1_c5_F3.gif')" ><img name="nav5_r1_c5" src="/Graphics/nav5_r1_c5.gif" width="86" height="27" border="0"></a>
        <a href="/new/creationpg.shtml" onMouseOut="GrpRestore('FwSimpleGroup');"  onMouseOver="GrpSwap('FwSimpleGroup','nav5_r1_c4','/Graphics/nav5_r1_c4_F2.gif')"  onClick="GrpDown('FwSimpleGroup','nav5_r1_c4','/Graphics/nav5_r1_c4_F3.gif','/Graphics/nav5_r1_c4_F3.gif')" ><img name="nav5_r1_c4" src="/Graphics/nav5_r1_c4.gif" width="109" height="27" border="0"></a>
        <a href="/new/linkspg.shtml" onMouseOut="GrpRestore('FwSimpleGroup');"  onMouseOver="GrpSwap('FwSimpleGroup','nav5_r1_c7','/Graphics/nav5_r1_c7_F2.gif')"  onClick="GrpDown('FwSimpleGroup','nav5_r1_c7','/Graphics/nav5_r1_c7_F3.gif','/Graphics/nav5_r1_c7_F3.gif')" ><img name="nav5_r1_c7" src="/Graphics/nav5_r1_c7.gif" width="51" height="27" border="0"></a>
<%End If%>
    </td></tr>
    <%If strTotLine <> "" Then %>
         <tr><td align='left' nowrap>
         <font face="Times New Roman, Times, serif" color='#666633' size=2>
         <%If intPgNbr > 1 and intLID > 0 Then%>
         <a href='/<%=APPLICATION("PGMDIR")%>/bookspg.asp?P=T' onmouseover="window.status='First Page';return true">First</a>
         <%Else%>
         First
         <%End If%>
         &nbsp;|&nbsp;
         <%IF bolPrev Then%>
         <a href='/<%=APPLICATION("PGMDIR")%>/bookspg.asp?P=B' onmouseover="window.status='Previous Page';return true">Previous</a>
         <%Else%>
         Previous
         <%End If%>
         &nbsp;|&nbsp;
         <%IF bolMore Then%>
         <a href='/<%=APPLICATION("PGMDIR")%>/bookspg.asp?P=F' onmouseover="window.status='Next Page';return true">Next</a>
         <%Else%>
         Next
         <%End If%>
         &nbsp;|&nbsp;
         <%IF bolMore Then%>
         <a href='/<%=APPLICATION("PGMDIR")%>/bookspg.asp?P=L' onmouseover="window.status='Last Page';return true">Last</a>
         <%Else%>
         Last
         <%End If%>
         &nbsp;|&nbsp;
         <%=strTotLine%>
         </font>
        </td></tr>
       <tr><td nowrap align='left'>
       <font face="Times New Roman, Times, serif" color='#666633' size=2>
       Printable Formats:
       &nbsp;
       <a href='/<%=APPLICATION("PGMDIR")%>/books_list.asp?ND=Y&S=<%=aryDbDtls(2)%>' onmouseover="window.status='Display Author and Title for all <%=SESSION("RECCNT")%> books in a printable format.';return true" oncontextmenu="PopUpMsg('Printer-Friendly Format','Display Author and Title for all <%=SESSION("RECCNT")%> books in a format suitable for printing. ',1);return false" >Summary</a>
       &nbsp;&nbsp;
       <a href='/<%=APPLICATION("PGMDIR")%>/books_list.asp?S=<%=aryDbDtls(2)%>' onmouseover="window.status='Display cataloged details for all <%=SESSION("RECCNT")%> books in a printable format.';return true" oncontextmenu="PopUpMsg('Printer-Friendly Format','Display cataloged details for all <%=SESSION("RECCNT")%> books in a format suitable for printing. ',1);return false" >Detail</a>
       </font>
       </td></tr>
    <%End If%>
    </table>
  <tr>
    <td width="110">&nbsp;</td>
    <td width="850"><img src="/graphics/rule.gif" width="850" height="21"><!-- #BeginEditable "body%20text" -->
<table>
<%

   bolUpdFlds = true

   For intSub = 1 to aryRecData(0,0)
      call Get_DB_Record(intSub)
      call Build_Book_Details
      For intCnt = 1 to aryFactors(0,0)
         If intSub < aryFactors(intCnt,0) Then
            intFactor = aryFactors(intCnt,1)
            EXIT FOR
         End If
      Next
      Response.Write "<tr>" & vbCrLf
      If not bolDisplayTnPics Then
         strPicFileSm = ""
      End If
      strTitle2 = trim(Replace(strTitle,"'","`"))
      strAuthor2 = trim(Replace(strAuthor,"'","`"))
      strDetails2 = trim(Replace(strDetails,"'","`"))
      strPopUpDetails = " onclick=" & CHR(34) & "PopUpMsg('" & strTitle2 & "','<table><tr><td valign=top>"
      If strPicFile = "" Then
         strPopUpDetails = strPopUpDetails & "&nbsp;"
      Else
         If left(strPicFile,5) <> "http:" and left(strPicFile,1) <> "/" Then
            strHld = "/" & APPLICATION("PICDIR") & "/books/"
         Else
            strHld = ""
         End If
         strPopUpDetails = strPopUpDetails & "<img src=" & strHld & strPicFile & " border=0 alt=Book_Picture>"
      End If
      strPopUpDetails = strPopUpDetails & "</td><td>&nbsp;</td><td><b>" & strTitle2 & "</b><br>" & strAuthor2 & "<br><br>"
      strPopUpDetails = strPopUpDetails & strDetails2 & "</td></tr></table>'," & intFactor-1 & ");return false" & CHR(34) & " "
      If strPicFileSm = "" Then
         Response.Write "<td width='40'>&nbsp;</td>" & vbCrLf
      Else
         If left(strPicFileSm,5) <> "http:" and left(strPicFileSm,1) <> "/" Then
            strHld = "/" & APPLICATION("PICDIR") & "/books/"
         Else
            strHld = ""
         End If
         Response.Write "<td><img src ='" & strHld & strPicFileSm & "' border='0' alt='Book Picture' onmouseover=" & CHR(34) & "window.status='Display Book Details';return true" & CHR(34) & strPopUpDetails & "></td>" & vbCrLf
      End If
      Response.Write "<td><a href="" onmouseover=" & CHR(34) & "window.status='Display Book Details';return true" & CHR(34) & strPopUpDetails & "><b>" & strTitle & "</b></a><br>" & vbCrLf
      Response.Write "&nbsp;&nbsp;&nbsp;by " & strAuthor & "<br>" & vbCrLf
      strHld = ""
      If strHld <> "" Then
         Response.Write "&nbsp;&nbsp;&nbsp;" & strHld & "<br>" & vbCrLf
      End If
      Response.Write "</td></tr>" & vbCrLf
      If bolDisplayTnPics Then
         Response.Write "<tr><td>&nbsp;</td></tr>" & vbCrLf
      End If
   Next
%>

</table>
<table>
<%
IF Application("EnableSrch") = "Y" Then
   If SESSION("CAT") <> "" and SESSION("RECCNT") = 0 and SESSION("SRCHKW") = "" Then
      'No Records in Database for "forced" Book Category ("CAT")
      Response.Write "<tr><td width=850 colspan=2 align='center'>" & vbCrLf
      strHld = "<br><br>No Book Entries Found for Requested "
      If instr(SESSION("CAT"),",") = 0 Then
         strHld = strHld & "Category"
      Else
         strHld = strHld & "Categories"
      End If
      Response.Write strHld & vbCrLf
      Response.Write "</tr></td>" & vbCrLf
   Else
      Response.Write "<tr><td width='850' colspan=2><img src='/graphics/rule.gif' width='850' height='10'></tr>" & vbCrLf
      If strTotLine <> "" Then %>
         <tr><td width=825 align='center' nowrap>
         <div align='middle'>
         <font face="Times New Roman, Times, serif" color='#666633' size=2>
         <%If intPgNbr > 1 and intLID > 0 Then%>
         <a href='/<%=APPLICATION("PGMDIR")%>/bookspg.asp?P=T' onmouseover="window.status='First Page';return true">First</a>
         <%Else%>
         First
         <%End If%>
         &nbsp;|&nbsp;
         <%IF bolPrev Then%>
         <a href='/<%=APPLICATION("PGMDIR")%>/bookspg.asp?P=B' onmouseover="window.status='Previous Page';return true">Previous</a>
         <%Else%>
         Previous
         <%End If%>
         &nbsp;|&nbsp;
         <%IF bolMore Then%>
         <a href='/<%=APPLICATION("PGMDIR")%>/bookspg.asp?P=F' onmouseover="window.status='Next Page';return true">Next</a>
         <%Else%>
         Next
         <%End If%>
         &nbsp;|&nbsp;
         <%IF bolMore Then%>
         <a href='/<%=APPLICATION("PGMDIR")%>/bookspg.asp?P=L' onmouseover="window.status='Last Page';return true">Last</a>
         <%Else%>
         Last
         <%End If%>
         &nbsp;|&nbsp;
         <%=strTotLine%>
         </font>
         </div>
        </td></tr>
      <%End If
      Response.Write "<tr><td width=850 colspan=2 align='center'>" & vbCrLf
      Call Insert_Search_Prompt("bookspg.asp")
      Response.Write "</tr></td>" & vbCrLf
   End If
End If
%>
</table>

      <!-- #EndEditable --></td>
</tr>
</table>
</body>
<!-- #EndTemplate --></html>
