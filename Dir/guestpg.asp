<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/system_setup.inc" -->
<!--#INCLUDE FILE="../dir/process_setup.inc" -->
<!--#INCLUDE FILE="../dir/state_country_select.inc" -->
<!--#INCLUDE FILE="../dir/privacy_notice.inc" -->
<%

    DIM intCnt

    Call System_Setup("NONE")

    strTitleClr = "#a5a673"
    strBgClr = "#c6c79c" 'Description Popup
    intBoxLeft = 200
    intBoxWidth = 500
    Call SetUp_PopUp_Msg
%>

<html><!-- #BeginTemplate "/Templates/template.dwt" -->
<head>
<!-- #BeginEditable "doctitle" -->
<title>Guestbook -Tinctoria Cottage</title>
<!-- #EndEditable -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-------------------- BEGIN COPYING THE JAVASCRIPT SECTION HERE ------------------->

<script language="JavaScript">
<!-- hide this script from non-javascript-enabled browsers

if (document.images) {
nav5_r1_c1_F1 = new Image(47,27); nav5_r1_c1_F1.src = "/Graphics/nav5_r1_c1.gif";
nav5_r1_c1_F2 = new Image(47,27); nav5_r1_c1_F2.src = "/Graphics/nav5_r1_c1_F2.gif";
nav5_r1_c1_F3 = new Image(47,27); nav5_r1_c1_F3.src = "/Graphics/nav5_r1_c1_F3.gif";
nav5_r1_c3_F1 = new Image(95,27); nav5_r1_c3_F1.src = "/Graphics/nav5_r1_c3.gif";
nav5_r1_c3_F2 = new Image(95,27); nav5_r1_c3_F2.src = "/Graphics/nav5_r1_c3_F2.gif";
nav5_r1_c3_F3 = new Image(95,27); nav5_r1_c3_F3.src = "/Graphics/nav5_r1_c3_F3.gif";
nav5_r1_c4_F1 = new Image(109,27); nav5_r1_c4_F1.src = "/Graphics/nav5_r1_c4.gif";
nav5_r1_c4_F2 = new Image(109,27); nav5_r1_c4_F2.src = "/Graphics/nav5_r1_c4_F2.gif";
nav5_r1_c4_F3 = new Image(109,27); nav5_r1_c4_F3.src = "/Graphics/nav5_r1_c4_F3.gif";
nav5_r1_c5_F1 = new Image(86,27); nav5_r1_c5_F1.src = "/Graphics/nav5_r1_c5.gif";
nav5_r1_c5_F2 = new Image(86,27); nav5_r1_c5_F2.src = "/Graphics/nav5_r1_c5_F2.gif";
nav5_r1_c5_F3 = new Image(86,27); nav5_r1_c5_F3.src = "/Graphics/nav5_r1_c5_F3.gif";
nav5_r1_c6_F1 = new Image(49,27); nav5_r1_c6_F1.src = "/Graphics/nav5_r1_c6.gif";
nav5_r1_c6_F2 = new Image(49,27); nav5_r1_c6_F2.src = "/Graphics/nav5_r1_c6_F2.gif";
nav5_r1_c6_F3 = new Image(49,27); nav5_r1_c6_F3.src = "/Graphics/nav5_r1_c6_F3.gif";
nav5_r1_c7_F1 = new Image(51,27); nav5_r1_c7_F1.src = "/Graphics/nav5_r1_c7.gif";
nav5_r1_c7_F2 = new Image(51,27); nav5_r1_c7_F2.src = "/Graphics/nav5_r1_c7_F2.gif";
nav5_r1_c7_F3 = new Image(51,27); nav5_r1_c7_F3.src = "/Graphics/nav5_r1_c7_F3.gif";
}

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

<%

    strTitleClr = "#a5a673"
    strBgClr = "#c6c79c"
    Call Check_For_PopUp_Msg
    Response.Write aryErrorText(1) & vbCrLf
%>

<body bgcolor="#FFFFCC" text="#666633" link="#CC9900" vlink="#666666" background="/Graphics/tincbg.jpg"<%=aryErrorText(2)%>>
<table width="100%" border="0">
  <tr>
    <td width="110"><img src="/Graphics/shim.gif" width="110" height="1"></td>
    <td width="500"><!-- #BeginEditable "header" --><img src="/Graphics/guestban.gif" width="300" height="90"><!-- #EndEditable --></td>
  </tr>
  <tr>
    <td width="110"><br>
      <br>
      <br>
      <br>
    </td>
    <td width="500">
      <p align="left">
       <a href="/new/home.html" onMouseOut="GrpRestore('FwSimpleGroup');"  onMouseOver="GrpSwap('FwSimpleGroup','nav5_r1_c1','/Graphics/nav5_r1_c1_F2.gif')"  onClick="GrpDown('FwSimpleGroup','nav5_r1_c1','/Graphics/nav5_r1_c1_F3.gif','/Graphics/nav5_r1_c1_F3.gif')" ><img name="nav5_r1_c1" src="/Graphics/nav5_r1_c1.gif" width="47" height="27" border="0"></a>
       <a href="/new/linenpg.shtml" onMouseOut="GrpRestore('FwSimpleGroup');"  onMouseOver="GrpSwap('FwSimpleGroup','nav5_r1_c3','/Graphics/nav5_r1_c3_F2.gif')"  onClick="GrpDown('FwSimpleGroup','nav5_r1_c3','/Graphics/nav5_r1_c3_F3.gif','/Graphics/nav5_r1_c3_F3.gif')" ><img name="nav5_r1_c3" src="/Graphics/nav5_r1_c3.gif" width="95" height="27" border="0"></a>
       <a href="/new/creationpg.shtml" onMouseOut="GrpRestore('FwSimpleGroup');"  onMouseOver="GrpSwap('FwSimpleGroup','nav5_r1_c4','/Graphics/nav5_r1_c4_F2.gif')"  onClick="GrpDown('FwSimpleGroup','nav5_r1_c4','/Graphics/nav5_r1_c4_F3.gif','/Graphics/nav5_r1_c4_F3.gif')" ><img name="nav5_r1_c4" src="/Graphics/nav5_r1_c4.gif" width="109" height="27" border="0"></a>
       <a href="/new/wovenpg.shtml" onMouseOut="GrpRestore('FwSimpleGroup');"  onMouseOver="GrpSwap('FwSimpleGroup','nav5_r1_c5','/Graphics/nav5_r1_c5_F2.gif')"  onClick="GrpDown('FwSimpleGroup','nav5_r1_c5','/Graphics/nav5_r1_c5_F3.gif','/Graphics/nav5_r1_c5_F3.gif')" ><img name="nav5_r1_c5" src="/Graphics/nav5_r1_c5.gif" width="86" height="27" border="0"></a>
       <a href="/new/linkspg.shtml" onMouseOut="GrpRestore('FwSimpleGroup');"  onMouseOver="GrpSwap('FwSimpleGroup','nav5_r1_c7','/Graphics/nav5_r1_c7_F2.gif')"  onClick="GrpDown('FwSimpleGroup','nav5_r1_c7','/Graphics/nav5_r1_c7_F3.gif','/Graphics/nav5_r1_c7_F3.gif')" ><img name="nav5_r1_c7" src="/Graphics/nav5_r1_c7.gif" width="51" height="27" border="0"></a>
       <img src="/Graphics/shim.gif" width="1" height="27" border="0">
      </p>
    </td>
  </tr>
  <tr>
    <td width="110">&nbsp;</td>
    <td width="550"><img src="/Graphics/rule.gif" width="550" height="21"><!-- #BeginEditable "body%20text" -->
      <div align="left">
        <p><font face="Times New Roman, Times, serif" size="5"><img src="/Graphics/buttons/shim.gif" width="50" height="1">Please
          sign our Guest Book </font></p>
      </div>
      <form method="post" action="member_maint.asp">
      <input type="hidden" name="GB" value="http://<%=Application("HOMEDIR")%>">
      <%If APPLICATION("BOOKLOGON") = "Y" Then%>
      <input type="hidden" name="LogAccess" value="Y">
      <input type="hidden" name="Allow PSWD Update" value="Y">
      <%End If%>
        <table border="1" bgcolor="#CCCC66" cellpadding="4">
          <tr>
            <td height="33">
              <p><font face="Times New Roman, Times, serif"><b>First Name</b></font></p>
            </td>
            <td height="33"> <font face="Times New Roman, Times, serif">
              <input type="text" name="FirstName" size="60" maxlength="50">
              </font></td>
          </tr>
          <tr>
            <td height="33"><b>Last Name</b></td>
            <td height="33"><font face="Times New Roman, Times, serif">
              <input type="text" name="LastName" size="60" maxlength="50">
              </font></td>
          </tr>
          <tr>
            <td><font face="Times New Roman, Times, serif"><b>Email address</b></font></td>
            <td> <font face="Times New Roman, Times, serif">
              <input type="text" name="e-Mail" size="60" maxlength="50">
              </font></td>
          </tr>
          <tr>
            <td>Street Address</td>
            <td>
              <input type="text" name="Street" size="60" maxlength="50">
            </td>
          </tr>
          <tr>
            <td height="33"><font face="Times New Roman, Times, serif">City </font></td>
            <td height="33"> <font face="Times New Roman, Times, serif">
              <input type="text" name="City" size="30" maxlength="50">
              </font>
            <font face="Times New Roman, Times, serif">State</font>
            <font face="Times New Roman, Times, serif">
            <%call Build_States_Select_List("")%>
              </font>
            </td></tr>
          <tr>
            <td height="33"><font face="Times New Roman, Times, serif">Zip</font></td>
            <td height="33"> <font face="Times New Roman, Times, serif">
              <input type="text" name="Zip" size="5" maxlength="10">
              </font>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <font face="Times New Roman, Times, serif">Country</font>
            <font face="Times New Roman, Times, serif">
            <%call Build_Country_Select_List("")%>
            </font></td>
          </tr>
          <%' If APPLICATION("BOOKLOGON") = "Y" Then%>
             <tr><td>
             <font face="Times New Roman, Times, serif">User Name:</font>
             </td><td nowrap>
             <input type="text" name="UserName" size="15" maxlength="30" value="">
             <font face="Times New Roman, Times, serif" size=-1>
             (Required to access member-only areas)
             </font>
             </td></tr>
             <tr><td>
             <font face="Times New Roman, Times, serif">Password:</font>
             </td><td nowrap>
             <input type="password" name="Password" size="15" maxlength="10" value="">
             &nbsp;&nbsp;
             <font face="Times New Roman, Times, serif">Re-Enter Password:</font>
             <input type="password" name="Password2" size="15" maxlength="10" value="">
             </td></tr>
          <%' End If%>
          <tr>
            <td><font face="Times New Roman, Times, serif"><img src="/Graphics/buttons/shim.gif" width="1" height="1">
              Comments /<br>Areas of Interest
              </font></td>
              <td> <font face="Times New Roman, Times, serif">
              <textarea name="Comments" cols="45" rows="3" wrap="YES")></textarea>
              </font></td>
          </tr>
          <tr>
             <td colspan="2" nowrap face="Times New Roman, Times, serif">
             <input type="Checkbox" checked name="e-Mail List">
             <font size="-1">
             Please check here if you would like to be on our e-Mail List.
             <br>
             &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
             <%=PrivacyNoticeLink(3)%>
             </font>
             </td>
          </tr>
          <tr>
            <td colspan="2" height="29">
              <div align="center"><font size="6" face="Times New Roman, Times, serif">
                <input type="submit" name="ADD" value="Submit">
                T</font><font size="4" face="Times New Roman, Times, serif">hank
                you for visiting Tinctoria Cottage! </font></div>
            </td>
          </tr>
        </table>
        <p> <br>
          <br>
          <br>
          <br>
          <br>
          <br>
        </p>
        <p>&nbsp;</p>
        </form>
      <!-- #EndEditable --></td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
</body>
<!-- #EndTemplate --></html>
