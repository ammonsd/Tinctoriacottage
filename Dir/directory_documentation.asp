<%@ LANGUAGE="VBSCRIPT"%>
<%
OPTION EXPLICIT
Response.Buffer = true
%>

<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/process_setup.inc" -->

<%
    DIM strCurLoc, bolSecMaintAccess, bolVCFAccess, bolWebMailAccess, strTitle, strSrchFlds, strCatComment, strBgClr
    DIM strNameTxt

    Call System_Setup("NONE")

    strCurLoc = GetPgmPath("")

    strBgClr = Application("BgCLR")
    If strBgClr = "" Then
       strBgClr = "#e6e6e6"
    End If

    bolDisplayMsg = false
    strSpecSecGrp = "SecMaint"
    Call Verify_Group_Access("N/R","",5)
    If not bolSecError Then
       bolSecMaintAccess =  true
    End If
    strSpecSecGrp = "VCF"
    Call Verify_Group_Access("N/R","",9)
    If not bolSecError Then
       bolVCFAccess =  true
    End If
    strSpecSecGrp = "WebMail"
    Call Verify_Group_Access("N/R","",9)
    If not bolSecError Then
       bolWebMailAccess =  true
    End If
    If SESSION("GROUP") <> "" Then
       strTitle = " for entries assigned to the " & CHR(34) & SESSION("GROUP") & CHR(34) & " membership group"
    End If
    strCatComment = "."
    strSrchFlds = "every key field in the database"
    strNameTxt = "Name"
    If SESSION("SRCHTYP") = "B" Then
       strSrchFlds = "the <b>Title</b>, <b>Author</b>"
       If SESSION("BC") = "" Then
          strSrchFlds = strSrchFlds & ", <b>Description</b> and <b>Category</b> book areas.  "
       Else
          strSrchFlds = strSrchFlds & " and <b>Description</b> book areas.  "
       End If
       strSrchFlds = strSrchFlds & "The search criteria only needs to be satisfied within one of these areas to "
       strSrchFlds = strSrchFlds & "qualify as a match.  "
       strSrchFlds = strSrchFlds & "The <b>Advance Search</b> option allows applying the search criteria against a "
       strSrchFlds = strSrchFlds & "a single book area only (Title, Author"
       IF SESSION("BC") = "" Then
          strSrchFlds = strSrchFlds & ", Description or Category)"
       Else
          strSrchFlds = strSrchFlds & " or Description)"
       End If
       If SESSION("BC") = "" Then
          strCatComment = ", except <b>Category</b> searches, which require an exact match (i.e. " & CHR(34) & "Miscellaneous" & CHR(34)
          strCatComment = strCatComment & " matches but " & CHR(34) & "Misc" & CHR(34) & " would not)."
       End If
       strNameTxt = "Author"
    End If

%>

<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 10">
<meta name=Originator content="Microsoft Word 10">
<link rel=File-List href="Address_Directory_Doc_files/filelist.xml">
<title><%=SESSION("GROUP")%> Address Directory Documentation</title>
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
  <o:Author>Dean Ammons</o:Author>
  <o:Template>Normal</o:Template>
  <o:LastAuthor>Dean Ammons</o:LastAuthor>
  <o:Revision>135</o:Revision>
  <o:TotalTime>923</o:TotalTime>
  <o:LastPrinted>2003-09-18T11:50:00Z</o:LastPrinted>
  <o:Created>2003-09-21T00:39:00Z</o:Created>
  <o:LastSaved>2004-02-03T19:45:00Z</o:LastSaved>
  <o:Pages>3</o:Pages>
  <o:Words>1253</o:Words>
  <o:Characters>7147</o:Characters>
  <o:Company>.</o:Company>
  <o:Lines>59</o:Lines>
  <o:Paragraphs>16</o:Paragraphs>
  <o:CharactersWithSpaces>8384</o:CharactersWithSpaces>
  <o:Version>10.4219</o:Version>
 </o:DocumentProperties>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <w:WordDocument>
  <w:SpellingState>Clean</w:SpellingState>
  <w:GrammarState>Clean</w:GrammarState>
  <w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel>
 </w:WordDocument>
</xml><![endif]-->
<style>
<!--
 /* Font Definitions */
 @font-face
	{font-family:Wingdings;
	panose-1:5 0 0 0 0 0 0 0 0 0;
	mso-font-charset:2;
	mso-generic-font-family:auto;
	mso-font-pitch:variable;
	mso-font-signature:0 268435456 0 0 -2147483648 0;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
h1
	{mso-style-next:Normal;
	margin:0in;
	margin-bottom:.0001pt;
	text-align:center;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:1;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-font-kerning:0pt;
	mso-bidi-font-weight:normal;}
h2
	{mso-style-next:Normal;
	margin:0in;
	margin-bottom:.0001pt;
	text-align:center;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:2;
	font-size:10.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Arial;}
h3
	{mso-style-next:Normal;
	margin-top:12.0pt;
	margin-right:0in;
	margin-bottom:3.0pt;
	margin-left:0in;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:3;
	font-size:13.0pt;
	font-family:Arial;}
h4
	{mso-style-next:Normal;
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:4;
	font-size:10.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Arial;}
p.MsoHeader, li.MsoHeader, div.MsoHeader
	{margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	tab-stops:center 3.0in right 6.0in;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
p.MsoFooter, li.MsoFooter, div.MsoFooter
	{margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	tab-stops:center 3.0in right 6.0in;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
p.MsoBodyText, li.MsoBodyText, div.MsoBodyText
	{margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:10.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Arial;
	mso-fareast-font-family:"Times New Roman";}
a:link, span.MsoHyperlink
	{color:blue;
	text-decoration:underline;
	text-underline:single;}
a:visited, span.MsoHyperlinkFollowed
	{color:purple;
	text-decoration:underline;
	text-underline:single;}
@page Section1
	{size:8.5in 11.0in;
	margin:1.0in 1.25in 1.0in 1.25in;
	mso-header-margin:.5in;
	mso-footer-margin:.5in;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}
 /* List Definitions */
 @list l0
	{mso-list-id:466357293;
	mso-list-type:hybrid;
	mso-list-template-ids:-1382382290 808073880 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}
@list l0:level1
	{mso-level-number-format:bullet;
	mso-level-text:\F0B7;
	mso-level-tab-stop:.6in;
	mso-level-number-position:left;
	margin-left:.6in;
	text-indent:-.15in;
	font-family:Symbol;
	mso-ansi-font-weight:normal;
	mso-ansi-font-style:normal;}
@list l0:level2
	{mso-level-tab-stop:1.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l0:level3
	{mso-level-tab-stop:1.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l0:level4
	{mso-level-tab-stop:2.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l0:level5
	{mso-level-tab-stop:2.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l0:level6
	{mso-level-tab-stop:3.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l0:level7
	{mso-level-tab-stop:3.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l0:level8
	{mso-level-tab-stop:4.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l0:level9
	{mso-level-tab-stop:4.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l1
	{mso-list-id:604272427;
	mso-list-type:hybrid;
	mso-list-template-ids:264524744 67698689 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}
@list l1:level1
	{mso-level-number-format:bullet;
	mso-level-text:\F0B7;
	mso-level-tab-stop:.5in;
	mso-level-number-position:left;
	text-indent:-.25in;
	font-family:Symbol;}
@list l1:level2
	{mso-level-tab-stop:1.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l1:level3
	{mso-level-tab-stop:1.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l1:level4
	{mso-level-tab-stop:2.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l1:level5
	{mso-level-tab-stop:2.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l1:level6
	{mso-level-tab-stop:3.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l1:level7
	{mso-level-tab-stop:3.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l1:level8
	{mso-level-tab-stop:4.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l1:level9
	{mso-level-tab-stop:4.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l2
	{mso-list-id:700933070;
	mso-list-type:hybrid;
	mso-list-template-ids:1610250440 808073880 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}
@list l2:level1
	{mso-level-number-format:bullet;
	mso-level-text:\F0B7;
	mso-level-tab-stop:.6in;
	mso-level-number-position:left;
	margin-left:.6in;
	text-indent:-.15in;
	font-family:Symbol;
	mso-ansi-font-weight:normal;
	mso-ansi-font-style:normal;}
@list l2:level2
	{mso-level-tab-stop:1.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l2:level3
	{mso-level-tab-stop:1.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l2:level4
	{mso-level-tab-stop:2.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l2:level5
	{mso-level-tab-stop:2.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l2:level6
	{mso-level-tab-stop:3.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l2:level7
	{mso-level-tab-stop:3.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l2:level8
	{mso-level-tab-stop:4.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l2:level9
	{mso-level-tab-stop:4.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l3
	{mso-list-id:1172796221;
	mso-list-type:hybrid;
	mso-list-template-ids:1729279996 808073880 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}
@list l3:level1
	{mso-level-number-format:bullet;
	mso-level-text:\F0B7;
	mso-level-tab-stop:.6in;
	mso-level-number-position:left;
	margin-left:.6in;
	text-indent:-.15in;
	font-family:Symbol;
	mso-ansi-font-weight:normal;
	mso-ansi-font-style:normal;}
@list l3:level2
	{mso-level-tab-stop:1.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l3:level3
	{mso-level-tab-stop:1.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l3:level4
	{mso-level-tab-stop:2.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l3:level5
	{mso-level-tab-stop:2.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l3:level6
	{mso-level-tab-stop:3.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l3:level7
	{mso-level-tab-stop:3.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l3:level8
	{mso-level-tab-stop:4.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l3:level9
	{mso-level-tab-stop:4.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l4
	{mso-list-id:1606309778;
	mso-list-type:hybrid;
	mso-list-template-ids:-1486831890 808073880 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}
@list l4:level1
	{mso-level-number-format:bullet;
	mso-level-text:\F0B7;
	mso-level-tab-stop:.6in;
	mso-level-number-position:left;
	margin-left:.6in;
	text-indent:-.15in;
	font-family:Symbol;
	mso-ansi-font-weight:normal;
	mso-ansi-font-style:normal;}
@list l4:level2
	{mso-level-number-format:bullet;
	mso-level-text:o;
	mso-level-tab-stop:1.0in;
	mso-level-number-position:left;
	text-indent:-.25in;
	font-family:"Courier New";}
@list l4:level3
	{mso-level-tab-stop:1.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l4:level4
	{mso-level-tab-stop:2.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l4:level5
	{mso-level-tab-stop:2.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l4:level6
	{mso-level-tab-stop:3.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l4:level7
	{mso-level-tab-stop:3.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l4:level8
	{mso-level-tab-stop:4.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l4:level9
	{mso-level-tab-stop:4.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l5
	{mso-list-id:1853640534;
	mso-list-type:hybrid;
	mso-list-template-ids:-854175194 -1299910572 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}
@list l5:level1
	{mso-level-number-format:bullet;
	mso-level-text:\F0B7;
	mso-level-tab-stop:.5in;
	mso-level-number-position:left;
	text-indent:-.25in;
	mso-ansi-font-size:10.0pt;
	font-family:Symbol;
	color:windowtext;}
@list l5:level2
	{mso-level-number-format:bullet;
	mso-level-text:o;
	mso-level-tab-stop:0in;
	mso-level-number-position:left;
	margin-left:0in;
	text-indent:-.25in;
	font-family:"Courier New";
	mso-bidi-font-family:"Times New Roman";}
@list l5:level3
	{mso-level-number-format:bullet;
	mso-level-text:\F0A7;
	mso-level-tab-stop:.5in;
	mso-level-number-position:left;
	margin-left:.5in;
	text-indent:-.25in;
	font-family:Wingdings;}
@list l5:level4
	{mso-level-number-format:bullet;
	mso-level-text:\F0B7;
	mso-level-tab-stop:1.0in;
	mso-level-number-position:left;
	margin-left:1.0in;
	text-indent:-.25in;
	font-family:Symbol;}
@list l5:level5
	{mso-level-tab-stop:2.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l5:level6
	{mso-level-tab-stop:3.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l5:level7
	{mso-level-tab-stop:3.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l5:level8
	{mso-level-tab-stop:4.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l5:level9
	{mso-level-tab-stop:4.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l6
	{mso-list-id:1901862037;
	mso-list-type:hybrid;
	mso-list-template-ids:-854175194 -1299910572 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}
@list l6:level1
	{mso-level-number-format:bullet;
	mso-level-text:\F0B7;
	mso-level-tab-stop:.5in;
	mso-level-number-position:left;
	text-indent:-.25in;
	mso-ansi-font-size:10.0pt;
	font-family:Symbol;
	color:windowtext;}
@list l6:level2
	{mso-level-number-format:bullet;
	mso-level-text:o;
	mso-level-tab-stop:1.0in;
	mso-level-number-position:left;
	text-indent:-.25in;
	font-family:"Courier New";
	mso-bidi-font-family:"Times New Roman";}
@list l6:level3
	{mso-level-number-format:bullet;
	mso-level-text:\F0A7;
	mso-level-tab-stop:.5in;
	mso-level-number-position:left;
	margin-left:.5in;
	text-indent:-.25in;
	font-family:Wingdings;}
@list l6:level4
	{mso-level-number-format:bullet;
	mso-level-text:\F0B7;
	mso-level-tab-stop:1.0in;
	mso-level-number-position:left;
	margin-left:1.0in;
	text-indent:-.25in;
	font-family:Symbol;}
@list l6:level5
	{mso-level-tab-stop:2.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l6:level6
	{mso-level-tab-stop:3.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l6:level7
	{mso-level-tab-stop:3.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l6:level8
	{mso-level-tab-stop:4.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l6:level9
	{mso-level-tab-stop:4.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l7
	{mso-list-id:2081979465;
	mso-list-type:hybrid;
	mso-list-template-ids:-2103253462 808073880 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}
@list l7:level1
	{mso-level-number-format:bullet;
	mso-level-text:\F0B7;
	mso-level-tab-stop:.6in;
	mso-level-number-position:left;
	margin-left:.6in;
	text-indent:-.15in;
	font-family:Symbol;
	mso-ansi-font-weight:normal;
	mso-ansi-font-style:normal;}
@list l7:level2
	{mso-level-tab-stop:1.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l7:level3
	{mso-level-tab-stop:1.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l7:level4
	{mso-level-tab-stop:2.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l7:level5
	{mso-level-tab-stop:2.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l7:level6
	{mso-level-tab-stop:3.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l7:level7
	{mso-level-tab-stop:3.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l7:level8
	{mso-level-tab-stop:4.0in;
	mso-level-number-position:left;
	text-indent:-.25in;}
@list l7:level9
	{mso-level-tab-stop:4.5in;
	mso-level-number-position:left;
	text-indent:-.25in;}
ol
	{margin-bottom:0in;}
ul
	{margin-bottom:0in;}
-->
</style>
<!--[if gte mso 10]>
<style>
 /* Style Definitions */
 table.MsoNormalTable
	{mso-style-name:"Table Normal";
	mso-tstyle-rowband-size:0;
	mso-tstyle-colband-size:0;
	mso-style-noshow:yes;
	mso-style-parent:"";
	mso-padding-alt:0in 5.4pt 0in 5.4pt;
	mso-para-margin:0in;
	mso-para-margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:10.0pt;
	font-family:"Times New Roman";}
</style>
<![endif]--><!--[if gte mso 9]><xml>
 <o:shapedefaults v:ext="edit" spidmax="1026"/>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout></xml><![endif]-->
</head>

<body bgcolor='<%=strBgClr%>' lang=EN-US link=blue vlink=purple style='tab-interval:.5in'>

<div class=Section1>

<p class=MsoNormal align=center style='text-align:center'>&nbsp;</p>

<div align=center>

<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0
 style='border-collapse:collapse;mso-padding-alt:0in 5.4pt 0in 5.4pt'>
 <tr style='mso-yfti-irow:0'>
  <td width=590 valign=top style='width:6.15in;padding:0in 5.4pt 0in 5.4pt'>
  <h3><u>Display Assigned Keywords</u></h3>
  <p class=MsoNormal>&nbsp;</p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:1'>
  <td width=590 valign=top style='width:6.15in;padding:0in 5.4pt 0in 5.4pt'>
  <p class=MsoBodyText>Opens a separate window with a list of the available keywords
  and phrases that can be copied and pasted into the search criteria prompt.
  Along with member names and email address domains, these keywords can be used
  to limit the number of members extracted.<span style='mso-spacerun:yes'> 
  </span>Copied phrases must be enclosed in quotes or the search will be done
  on each individual word.</p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:2'>
  <td width=590 valign=top style='width:6.15in;padding:0in 5.4pt 0in 5.4pt'>
  <h3><a name=EmailList><u>Email Mailing List</u></a></h3>
  <p class=MsoNormal>&nbsp;</p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:3'>
  <td width=590 valign=top style='width:6.15in;padding:0in 5.4pt 0in 5.4pt'>
  <p class=MsoBodyText>Instead of displaying a list of matching entries, this
  option will open a selection window containing each name that matches the
  search criteria.<span style='mso-spacerun:yes'>  </span>From this window,
  entries can be selected and designated as either a “CC”, “BCC” or “TO”
  address.<span style='mso-spacerun:yes'>  </span>Clicking on the <b
  style='mso-bidi-font-weight:normal'>Assign</b> button will concatenate the
  email address for the selected names in a format suitable for inputting into
  the address area of a new email message.<span style='mso-spacerun:yes'> 
  </span></p>
  <p class=MsoBodyText>&nbsp;</p>
  <table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0
   style='border-collapse:collapse;mso-padding-alt:0in 5.4pt 0in 5.4pt'>
   <tr style='mso-yfti-irow:0'>
    <td width=131 style='width:98.6pt;border:solid windowtext 1.0pt;mso-border-alt:
    solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
    <h2>Assign CC</h2>
    </td>
    <td width=459 valign=top style='width:344.2pt;border:solid windowtext 1.0pt;
    border-left:none;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
    solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
    <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
    font-family:Arial'>Available if the search returns at least two matching
    names.<span style='mso-spacerun:yes'>  </span>When selected, all of the
    checked names will be formatted for the “CC” address area.<span
    style='mso-spacerun:yes'>   </span><o:p></o:p></span></p>
    </td>
   </tr>
   <tr style='mso-yfti-irow:1'>
    <td width=131 style='width:98.6pt;border:solid windowtext 1.0pt;border-top:
    none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
    padding:0in 5.4pt 0in 5.4pt'>
    <h1><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:
    Arial'>Assign BCC<o:p></o:p></span></h1>
    </td>
    <td width=459 valign=top style='width:344.2pt;border-top:none;border-left:
    none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
    mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
    mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
    <p class=MsoBodyText>Works exactly the same as the “CC” option, except the results
    are formatted for the “BCC” address area.</p>
    </td>
   </tr>
   <tr style='mso-yfti-irow:2;mso-yfti-lastrow:yes'>
    <td width=131 style='width:98.6pt;border:solid windowtext 1.0pt;border-top:
    none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
    padding:0in 5.4pt 0in 5.4pt'>
    <h1><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:
    Arial'>Assign TO &amp; Build<o:p></o:p></span></h1>
    </td>
    <td width=459 valign=top style='width:344.2pt;border-top:none;border-left:
    none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
    mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
    mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
    <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
    font-family:Arial'>Formats the remaining selected names for the “To”
    address area, builds the address strings and displays the results
    window.<span style='mso-spacerun:yes'>  </span>At least one name must be
    checked for the TO address.<o:p></o:p></span></p>
    </td>
   </tr>
  </table>
  <p class=MsoBodyText><b><o:p></o:p></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:4'>
  <td width=590 valign=top style='width:6.15in;padding:0in 5.4pt 0in 5.4pt'>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText>The results window will display the concatenated “To”,
  “CC” and “BCC” email address strings, which can be copied and pasted into an
  email address area.<span style='mso-spacerun:yes'>   </span>If the <b
  style='mso-bidi-font-weight:normal'>Open Mail</b> option is selected, a new
  message window for the default email program will be opened with the “To”,
  “CC” and “BCC” areas pre-filled with the appropriate address string.<span
  style='mso-spacerun:yes'>   </span>An address area will <b style='mso-bidi-font-weight:
  normal'>not</b> be pre-filled if the corresponding address string exceeds the
  maximum size allowed (approximately 2000 characters).<span
  style='mso-spacerun:yes'>   </span>A new message will still be opened, but
  these addresses will have to be manually copied and pasted into the
  appropriate address areas.</p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:5'>
  <td width=590 valign=top style='width:6.15in;padding:0in 5.4pt 0in 5.4pt'>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText>To maintain multiple mailing lists, store the search
  criteria for each list in a text file with a notation explaining the mailing
  list purpose.<span style='mso-spacerun:yes'>  </span>For example,</p>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText style='margin-left:.5in'><b style='mso-bidi-font-weight:
  normal'>Project Team Members<o:p></o:p></b></p>
  <p class=MsoBodyText style='margin-left:.5in'>“Maggie May” OR “Eleanor Rigby”
  OR Infrastructure OR “Tech Support”</p>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText>To broadcast an email to members of a saved list:</p>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l3 level1 lfo4;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Select the <b style='mso-bidi-font-weight:
  normal'>Email Mailing</b> <b style='mso-bidi-font-weight:normal'>List</b>
  option,</p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l3 level1 lfo4;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Copy and paste the search criteria from the text
  file into the search input prompt,</p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l3 level1 lfo4;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Select and assign the “CC”, “BCC” and “To”
  recipients from the list of matching names, </p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l3 level1 lfo4;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Click <b style='mso-bidi-font-weight:normal'>Open
  Mail</b> to create a new email message pre-filled with the selected
  addresses.<span style='mso-spacerun:yes'>  </span></p>
  <p class=MsoBodyText>&nbsp;</p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:6'>
  <td width=590 valign=top style='width:6.15in;padding:0in 5.4pt 0in 5.4pt'>
  <p class=MsoBodyText>&nbsp;</p>
  <h3><u>Build Web Page Display</u></h3>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText>Instead of displaying a list of matching entries, this
  option will open a selection window containing each name that matches the
  search criteria.<span style='mso-spacerun:yes'>  </span>Clicking on the <b
  style='mso-bidi-font-weight:normal'>Create Web Page</b> button will build a
  web page of the selected entries that can be saved as an HTML file for
  offline access.</p>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText>The web page provides two functions:</p>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l0 level1 lfo6;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>Left</b>
  click on any displayed name to open a new email message for the selected
  entry, or</p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l0 level1 lfo6;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>Right</b>
  click on a name to view the person’s directory details.</p>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText>For easy access to the web page information:</p>
  <p class=MsoNormal>&nbsp;</p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l4 level1 lfo8;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Create a desktop shortcut:</p>
  <p class=MsoBodyText style='margin-left:1.0in;text-indent:-.25in;mso-list:
  l4 level2 lfo8;tab-stops:list 1.0in'><![if !supportLists]><span
  style='font-family:"Courier New";mso-fareast-font-family:"Courier New"'><span
  style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Open the web page,</p>
  <p class=MsoBodyText style='margin-left:1.0in;text-indent:-.25in;mso-list:
  l4 level2 lfo8;tab-stops:list 1.0in'><![if !supportLists]><span
  style='font-family:"Courier New";mso-fareast-font-family:"Courier New"'><span
  style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Right click anywhere on the page, and </p>
  <p class=MsoBodyText style='margin-left:1.0in;text-indent:-.25in;mso-list:
  l4 level2 lfo8;tab-stops:list 1.0in'><![if !supportLists]><span
  style='font-family:"Courier New";mso-fareast-font-family:"Courier New"'><span
  style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Select the “create shortcut” option.</p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l4 level1 lfo8;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Add quick key access to the shortcut:</p>
  <p class=MsoBodyText style='margin-left:1.0in;text-indent:-.25in;mso-list:
  l4 level2 lfo8;tab-stops:list 1.0in'><![if !supportLists]><span
  style='font-family:"Courier New";mso-fareast-font-family:"Courier New"'><span
  style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Right click on the new shortcut,</p>
  <p class=MsoBodyText style='margin-left:1.0in;text-indent:-.25in;mso-list:
  l4 level2 lfo8;tab-stops:list 1.0in'><![if !supportLists]><span
  style='font-family:"Courier New";mso-fareast-font-family:"Courier New"'><span
  style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Select properties, </p>
  <p class=MsoBodyText style='margin-left:1.0in;text-indent:-.25in;mso-list:
  l4 level2 lfo8;tab-stops:list 1.0in'><![if !supportLists]><span
  style='font-family:"Courier New";mso-fareast-font-family:"Courier New"'><span
  style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Click on the web document tab,</p>
  <p class=MsoBodyText style='margin-left:1.0in;text-indent:-.25in;mso-list:
  l4 level2 lfo8;tab-stops:list 1.0in'><![if !supportLists]><span
  style='font-family:"Courier New";mso-fareast-font-family:"Courier New"'><span
  style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Click anywhere in the shortcut key area,</p>
  <p class=MsoBodyText style='margin-left:1.0in;text-indent:-.25in;mso-list:
  l4 level2 lfo8;tab-stops:list 1.0in'><![if !supportLists]><span
  style='font-family:"Courier New";mso-fareast-font-family:"Courier New"'><span
  style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Press and hold the &lt;Ctrl&gt; key,</p>
  <p class=MsoBodyText style='margin-left:1.0in;text-indent:-.25in;mso-list:
  l4 level2 lfo8;tab-stops:list 1.0in'><![if !supportLists]><span
  style='font-family:"Courier New";mso-fareast-font-family:"Courier New"'><span
  style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Press any key that makes the
  &lt;Ctrl&gt;&lt;Alt&gt; combination unique, such as “=”, and</p>
  <p class=MsoBodyText style='margin-left:1.0in;text-indent:-.25in;mso-list:
  l4 level2 lfo8;tab-stops:list 1.0in'><![if !supportLists]><span
  style='font-family:"Courier New";mso-fareast-font-family:"Courier New"'><span
  style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Click “OK” to apply the changes and close the
  properties window.</p>
  <p class=MsoNormal>&nbsp;</p>
  <p class=MsoBodyText>Pressing the “&lt;Ctrl&gt;&lt;Alt&gt;= “ key combination
  at the same time will open the directory web page, assuming the current
  application does not use the same key combination.</p>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText>Since this function requires opening a new window for
  the generated web page, the process will fail if the browser is set to block
  popup windows.<span style='mso-spacerun:yes'>  </span>The steps required for
  allowing the new web page window to open depends on the browser being used:</p>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l2 level1 lfo10;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>Internet
  Explorer</b> – Press the &lt;Ctrl&gt; key <b style='mso-bidi-font-weight:
  normal'>while clicking</b> on the “Create Web Page” button to temporarily
  override the blocking feature.</p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l2 level1 lfo10;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>Netscape</b>
  – Open the Popup Manager and select “Allow Popups from This Site” to
  permanently override the blocking feature for the directory site only.</p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l2 level1 lfo10;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>Others</b>
  – Review the browser’s documentation for overriding the popup blocking
  feature.</p>
  <p class=MsoBodyText>&nbsp;</p>
<%IF SESSION("SECLEVL") < 3 Then%>
  <h3><u>Requesting New Directory Entries</u></h3>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText>Filling out the <b style='mso-bidi-font-weight:normal'><a
  href="http://<%=strCurLoc%>new_dir_entry.asp">New Entry Request</a></b> form will send an email
  request to add the entry to the directory database.<span
  style='mso-spacerun:yes'>  </span>The only required fields are the first and last
  names and either an email address or telephone number.<span
  style='mso-spacerun:yes'>  </span>The remaining fields are optional.<span
  style='mso-spacerun:yes'>  </span>A notification will be sent once the entry
  as been added.</p>
  <p class=MsoBodyText>&nbsp;</p>
<%End If%>
  <p class=MsoBodyText>&nbsp;</p>
<%IF SESSION("SECLEVL") > 2 Then%>
  <h3><u>Directory Maintenance</u></h3>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText>All of the following fields are included in a search <u>except</u>
  “Contact Info”.</p>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText style='margin-left:.5in;text-indent:-.25in;mso-list:
  l1 level1 lfo12;tab-stops:list .5in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b>First &amp; Last Name</b>:<span
  style='mso-spacerun:yes'>  </span>At least one of the fields must be
  entered.<span style='mso-spacerun:yes'>  </span></p>
  <p class=MsoBodyText style='margin-left:.5in;text-indent:-.25in;mso-list:
  l1 level1 lfo12;tab-stops:list .5in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b>Email Address</b>:<span
  style='mso-spacerun:yes'>  </span>Must be unique.<span
  style='mso-spacerun:yes'>  </span></p>
<%IF bolSecmaintAccess Then%>
  <p class=MsoBodyText style='margin-left:.5in;text-indent:-.25in;mso-list:
  l1 level1 lfo12;tab-stops:list .5in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b>Security Groups</b>: Used to identify who
  has access to the member’s record.<span style='mso-spacerun:yes'>  </span>At
  lease one group entry is required with multiple groups separated by a
  comma.<span style='mso-spacerun:yes'>  </span>Only with users who have access
  to at least one of the assigned groups will have access to the member’s
  details.</p>
<%End If%>
  <p class=MsoBodyText style='margin-left:.5in;text-indent:-.25in;mso-list:
  l1 level1 lfo12;tab-stops:list .5in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b>Company</b>:<span
  style='mso-spacerun:yes'>  </span>Do not include the quote character.</p>
  <p class=MsoBodyText style='margin-left:.5in;text-indent:-.25in;mso-list:
  l1 level1 lfo12;tab-stops:list .5in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b>Contact Info</b>: This is a free form field
  that can include any contact information.<span style='mso-spacerun:yes'> 
  </span>For readability, multiple phone numbers should be entered on separated
  lines followed by a type indicator, such as (W) for work, (H) for home or (M)
  for mobile.<span style='mso-spacerun:yes'>   </span>The contact information
  will be included on the Web Page directory details pop-up window.</p>
  <p class=MsoBodyText style='margin-left:.5in;text-indent:-.25in;mso-list:
  l1 level1 lfo12;tab-stops:list .5in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b>Keywords</b>: Enter the keywords that will
  be used to group this member in a search.<span style='mso-spacerun:yes'> 
  </span>A comma must separate multiple keywords and phrases.<span
  style='mso-spacerun:yes'>  </span>For example, entering the keywords <u>Team
  Leader, Support Team, DBA</u> would place the member in three separate search
  groups.<span style='mso-spacerun:yes'>  </span>Whenever the directory search
  criteria included “Team Leader”, “Support Team” or DBA, this member would be
  selected.</p>
<%End If%>
  <p class=MsoBodyText>&nbsp;</p>
<%IF bolSecmaintAccess Then%>
  <h3><u>Security Maintenance</u></h3>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText>A user must have a security account to access any of the
  email address directory functions.<span style='mso-spacerun:yes'> 
  </span>Besides the user’s name, the following input is required:</p>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText style='margin-left:.5in;text-indent:-.25in;mso-list:
  l5 level1 lfo14;tab-stops:list .5in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>User
  Name</b>: Used to log into the system.<span style='mso-spacerun:yes'> 
  </span>Must be unique.<span style='mso-spacerun:yes'>  </span>User’s email
  name (the portion left of the “@” symbol) is a good choice.</p>
  <p class=MsoBodyText style='margin-left:.5in;text-indent:-.25in;mso-list:
  l5 level1 lfo14;tab-stops:list .5in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b>Email Address</b>:<span
  style='mso-spacerun:yes'>  </span>Must be unique</p>
  <p class=MsoBodyText style='margin-left:.5in;text-indent:-.25in;mso-list:
  l5 level1 lfo14;tab-stops:list .5in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>Password:</b><span
  style='mso-spacerun:yes'>  </span>New users will be assigned the temporary
  password “Pass + word”, which will have to be changed at the first logon.</p>
  <p class=MsoBodyText style='margin-left:.5in;text-indent:-.25in;mso-list:
  l5 level1 lfo14;tab-stops:list .5in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>Security
  Level</b>: Determines what areas the user can access that are not controlled
  by Groups:</p>
  <p class=MsoBodyText style='margin-left:1.0in;text-indent:-.25in;mso-list:
  l6 level2 lfo16;tab-stops:list 1.0in'><![if !supportLists]><span
  style='font-family:"Courier New";mso-fareast-font-family:"Courier New"'><span
  style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>None</b>
  - Access denied to all areas.<span style='mso-spacerun:yes'>  </span>Normally
  used to deactivate an account without deleting it.</p>
  <p class=MsoBodyText style='margin-left:1.0in;text-indent:-.25in;mso-list:
  l6 level2 lfo16;tab-stops:list 1.0in'><![if !supportLists]><span
  style='font-family:"Courier New";mso-fareast-font-family:"Courier New"'><span
  style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>Access</b>
  – Access to all non-maintenance directory functions for assigned security
  groups</p>
  <p class=MsoBodyText style='margin-left:1.0in;text-indent:-.25in;mso-list:
  l6 level2 lfo16;tab-stops:list 1.0in'><![if !supportLists]><span
  style='font-family:"Courier New";mso-fareast-font-family:"Courier New"'><span
  style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>User</b>
  – Same as Access with ability to perform directory maintenance for assigned
  security groups</p>
  <p class=MsoBodyText style='margin-left:1.0in;text-indent:-.25in;mso-list:
  l6 level2 lfo16;tab-stops:list 1.0in'><![if !supportLists]><span
  style='font-family:"Courier New";mso-fareast-font-family:"Courier New"'><span
  style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>Administrator</b>
  – Same as User with ability to perform security maintenance for assigned
  security groups</p>
  <p class=MsoBodyText style='margin-left:1.0in;text-indent:-.25in;mso-list:
  l6 level2 lfo16;tab-stops:list 1.0in'><![if !supportLists]><span
  style='font-family:"Courier New";mso-fareast-font-family:"Courier New"'><span
  style='mso-list:Ignore'>o<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>Owner</b>
  - Access to all areas</p>
  <p class=MsoBodyText style='margin-left:.5in;text-indent:-.25in;mso-list:
  l6 level1 lfo16;tab-stops:list .5in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>Password
  Cycle:</b><span style='mso-spacerun:yes'>  </span>The number of days before
  the user will have to enter a new password</p>
  <p class=MsoBodyText style='margin-left:.5in;text-indent:-.25in;mso-list:
  l6 level1 lfo16;tab-stops:list .5in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>Expire
  Password:</b> Force the user to change their password during the next logon.</p>
  <p class=MsoBodyText style='margin-left:.5in;text-indent:-.25in;mso-list:
  l6 level1 lfo16;tab-stops:list .5in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>User
  Expiration</b>: The date when the user’s account will expire, changing their
  access level to none.</p>
  <p class=MsoBodyText style='margin-left:.5in;text-indent:-.25in;mso-list:
  l6 level1 lfo16;tab-stops:list .5in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>Security
  Groups:</b> Defines which directory security groups the user has access
  too.<span style='mso-spacerun:yes'>  </span>If assigning multiple groups, the
  first group entered <b style='mso-bidi-font-weight:normal'>MUST </b>be the
  user’s primary security group and each secondary group separated with a
  comma.<span style='mso-spacerun:yes'>  </span>Leave blank for access to all
  groups.<span style='mso-spacerun:yes'>  </span></p>
  <p class=MsoBodyText style='margin-left:.5in;text-indent:-.25in;mso-list:
  l6 level1 lfo16;tab-stops:list .5in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>Log
  Access</b>: If checked, an entry will be made into the directory access log
  file each time the user logs into the system.</p>
  <p class=MsoBodyText style='margin-left:.5in;text-indent:-.25in;mso-list:
  l6 level1 lfo16;tab-stops:list .5in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b style='mso-bidi-font-weight:normal'>Allow
  Password Change</b>: Check this box to allow the user to change their
  password when logging into the system.<span style='mso-spacerun:yes'> 
  </span>Even if this option is disabled, the user will be able to change their
  temporary password the first time they log into the system.</p>
<%End If%>
  </td>
 </tr>
 <tr style='mso-yfti-irow:7'>
  <td width=590 valign=top style='width:6.15in;padding:0in 5.4pt 0in 5.4pt'>
  <p class=MsoBodyText>&nbsp;</p>
<%IF bolVCFAccess Then%>
  <h3><u>Build VCF List</u></h3>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText>This option takes the data from the selected members and
  displays a list of data VCF formatted for importing into an address
  directory.<span style='mso-spacerun:yes'>  </span>The top of the display will
  contain the steps required for importing the data.</p>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText>&nbsp;</p>
<%End If%>
  </td>
 </tr>
 <tr style='mso-yfti-irow:8;mso-yfti-lastrow:yes'>
  <td width=590 valign=top style='width:6.15in;padding:0in 5.4pt 0in 5.4pt'>
<%IF bolWebMailAccess Then%>
  <h3><u>Web Mail</u></h3>
  <p class=MsoBodyText>&nbsp;</p>
  <p class=MsoBodyText>This option allows sending a web email message to the
  selected addresses using any web browser instead of the workstation default
  email program.<span style='mso-spacerun:yes'>  </span>For additional details,
  review the <a href="http://<%=strCurLoc%>Send%20Web%20Mail%20Documentation.htm">Web Mail
  documentation.</a></p>
  <p class=MsoBodyText>&nbsp;</p>
<%End If%>
  </td>
 </tr>
</table>

</div>

<p class=MsoNormal>&nbsp;</p>

</div>

</body>

</html>

