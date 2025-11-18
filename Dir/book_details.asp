
<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/system_setup.inc" -->
<!--#INCLUDE FILE="../dir/book_select.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->

<%

    DIM intSub

    Call System_Setup("NONE")
    Call Get_DB_Record(Request.QueryString("ID"))

%>

<html><!-- #BeginTemplate "/Templates/template.dwt" -->
<head>
<!-- #BeginEditable "doctitle" -->
<title>Books - <%=strWebCompany%></title>
<!-- #EndEditable -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-------------------- BEGIN COPYING THE JAVASCRIPT SECTION HERE ------------------->

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

<body bgcolor="#FFFFCC" text="#666633" link="#CC9900" vlink="#666666" background="/Graphics/tincbg.jpg">
<table width="100%" border="0">
  <tr>
    <td width="110"><img src="/Graphics/shim.gif" width="110" height="1"></td>
    <td width="500" align="center"><!-- #BeginEditable "header" --><img src="/Graphics/bookban.gif" width="300" height="90"><!-- #EndEditable --></td>
  </tr>
  <tr>
    <td width="110">&nbsp;</td>
    <td width="500"><img src="/Graphics/rule.gif" width="500" height="21">
<table>
<%


    If strPicFile <> "" and left(strPicFile,5) <> "http:" and left(strPicFile,1) <> "/" then
       strPicFile = "/" & APPLICATION("PICDIR") & "/books/" & strPicFile
    End If

    call Build_Book_Details

    Response.Write "<tr>" & vbCrLf
    If trim(strPicFile) = "" Then
       Response.Write "<td>&nbsp;</td>" & vbCrLf
    Else
       Response.Write "<td><img src ='" & strPicFile & "' border='0' alt='Book Picture'></td>" & vbCrLf
    End If
    Response.Write "<td>&nbsp;</td>" & vbCrLf
    Response.Write "<td><b>" & strTitle & "</b><br>" & vbCrLf
    Response.Write "by " & strAuthor & "<br><br>" & vbCrLf
    Response.Write strDetails & "</td></tr>" & vbCrLf
%>

</table>
<div align = 'center'><input type=image src='/graphics/buttons/tincbookbut.jpg' onClick='history.go(-1)'</div>
      <!-- #EndEditable --></td>
  </tr>
</table>
</body>
<!-- #EndTemplate --></html>
