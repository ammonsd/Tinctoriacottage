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
<link rel=File-List href="Search_Help_files/filelist.xml">
<title>Database Searches</title>
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
  <o:Author>Dean Ammons</o:Author>
  <o:Template>Normal</o:Template>
  <o:LastAuthor>Dean Ammons</o:LastAuthor>
  <o:Revision>149</o:Revision>
  <o:TotalTime>1019</o:TotalTime>
  <o:LastPrinted>2004-01-19T19:03:00Z</o:LastPrinted>
  <o:Created>2003-09-21T00:39:00Z</o:Created>
  <o:LastSaved>2004-01-20T13:37:00Z</o:LastSaved>
  <o:Pages>1</o:Pages>
  <o:Words>427</o:Words>
  <o:Characters>2439</o:Characters>
  <o:Company>.</o:Company>
  <o:Lines>20</o:Lines>
  <o:Paragraphs>5</o:Paragraphs>
  <o:CharactersWithSpaces>2861</o:CharactersWithSpaces>
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
	{mso-list-id:1177038518;
	mso-list-type:hybrid;
	mso-list-template-ids:1442594416 808073880 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}
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
	{mso-list-id:1971746534;
	mso-list-type:hybrid;
	mso-list-template-ids:1101461418 808073880 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}
@list l1:level1
	{mso-level-number-format:bullet;
	mso-level-text:\F0B7;
	mso-level-tab-stop:.6in;
	mso-level-number-position:left;
	margin-left:.6in;
	text-indent:-.15in;
	font-family:Symbol;
	mso-ansi-font-weight:normal;
	mso-ansi-font-style:normal;}
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
  <p class=MsoBodyText><o:p>&nbsp;</o:p></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:1'>
  <td width=590 valign=top style='width:6.15in;padding:0in 5.4pt 0in 5.4pt'>
  <p class=MsoBodyText>The <b style='mso-bidi-font-weight:normal'>Search</b>
  function allows creating specific queries to limit the number of entries
  displayed.<span style='mso-spacerun:yes'>  </span></p>
  <p class=MsoBodyText><o:p>&nbsp;</o:p></p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l0 level1 lfo2;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Any number of terms can be entered, each
  separated by a space.<span style='mso-spacerun:yes'>  </span></p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l0 level1 lfo2;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Search parameters are not case sensitive and
  are applied against <%=strSrchFlds%>.</p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l0 level1 lfo2;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>Only a partial match is required to satisfy
  the criteria (i.e. SAN would match Sandy, Sanderson and Susan)<%=strCatComment%> </p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l0 level1 lfo2;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]>The search criteria are applied against each
  area individually, not as combined areas.<span style='mso-spacerun:yes'> 
  </span>For example, the request Rock AND Roll would not produce a match if
  “rock” was found in <%=strNameTxt%> and “roll” was found in Description, but would
  produce a match if both words were found in Description.</p>
  <p class=MsoBodyText><o:p>&nbsp;</o:p></p>
  <p class=MsoBodyText>Search results can be controlled using any combination
  of the following options:</p>
  <p class=MsoBodyText>&nbsp;</p>
  <table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0
   style='border-collapse:collapse;mso-padding-alt:0in 5.4pt 0in 5.4pt'>
   <tr style='mso-yfti-irow:0;height:22.0pt'>
    <td width=71 style='width:53.6pt;border:solid windowtext 1.0pt;mso-border-alt:
    solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;height:22.0pt'>
    <h4 align=center style='text-align:center'>Phrases</h4>
    </td>
    <td width=519 valign=top style='width:389.2pt;border:solid windowtext 1.0pt;
    border-left:none;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:
    solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;height:22.0pt'>
    <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
    font-family:Arial'>Search for complete phrases by enclosing multiple words
    within quotation marks.<span style='mso-spacerun:yes'>  </span><o:p></o:p></span></p>
    </td>
   </tr>
   <tr style='mso-yfti-irow:1'>
    <td width=71 style='width:53.6pt;border:solid windowtext 1.0pt;border-top:
    none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
    padding:0in 5.4pt 0in 5.4pt'>
    <h2>AND</h2>
    </td>
    <td width=519 valign=top style='width:389.2pt;border-top:none;border-left:
    none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
    mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
    mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
    <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
    font-family:Arial'>To retrieve entries that include both word A and word B,
    insert AND between words or phrases.<span style='mso-spacerun:yes'> 
    </span>If a logical operator is not specified, the operator AND will be
    assumed.<o:p></o:p></span></p>
    </td>
   </tr>
   <tr style='mso-yfti-irow:2'>
    <td width=71 style='width:53.6pt;border:solid windowtext 1.0pt;border-top:
    none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
    padding:0in 5.4pt 0in 5.4pt'>
    <h2>OR</h2>
    </td>
    <td width=519 valign=top style='width:389.2pt;border-top:none;border-left:
    none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
    mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
    mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
    <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
    font-family:Arial'>To retrieve entries that include either word A or word
    B, insert OR between words or phrases.<o:p></o:p></span></p>
    </td>
   </tr>
   <tr style='mso-yfti-irow:3'>
    <td width=71 style='width:53.6pt;border:solid windowtext 1.0pt;border-top:
    none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
    padding:0in 5.4pt 0in 5.4pt'>
    <p class=MsoNormal align=center style='text-align:center'><b
    style='mso-bidi-font-weight:normal'><span style='font-size:10.0pt;
    mso-bidi-font-size:12.0pt;font-family:Arial'>dash</span></b><span
    style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
    </td>
    <td width=519 valign=top style='width:389.2pt;border-top:none;border-left:
    none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
    mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
    mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
    <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
    font-family:Arial'>Prefix a word or phrase with a dash to select entries
    that <b style='mso-bidi-font-weight:normal'>do not</b> contain the
    requested text.<o:p></o:p></span></p>
    </td>
   </tr>
   <tr style='mso-yfti-irow:4'>
    <td width=71 style='width:53.6pt;border:solid windowtext 1.0pt;border-top:
    none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
    padding:0in 5.4pt 0in 5.4pt'>
    <p class=MsoNormal align=center style='text-align:center'><b
    style='mso-bidi-font-weight:normal'><span style='font-size:10.0pt;
    mso-bidi-font-size:12.0pt;font-family:Arial'>( )</span></b><span
    style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
    </td>
    <td width=519 valign=top style='width:389.2pt;border-top:none;border-left:
    none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
    mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
    mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
    <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
    font-family:Arial'>Use parentheses to group complex Boolean phrases, with
    no spaces between the parenthesis and the enclosed terms.<o:p></o:p></span></p>
    </td>
   </tr>
   <tr style='mso-yfti-irow:5'>
    <td width=71 style='width:53.6pt;border:solid windowtext 1.0pt;border-top:
    none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
    padding:0in 5.4pt 0in 5.4pt'>
    <p class=MsoNormal align=center style='text-align:center'><b
    style='mso-bidi-font-weight:normal'><span style='font-size:10.0pt;
    mso-bidi-font-size:12.0pt;font-family:Arial'>asterisk<o:p></o:p></span></b></p>
    </td>
    <td width=519 valign=top style='width:389.2pt;border-top:none;border-left:
    none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
    mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
    mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
    <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
    font-family:Arial'>Attach an asterisk to the end of a word or phrase to
    select entries that <b style='mso-bidi-font-weight:normal'>begin</b> with
    the requested text or prefix the request with an asterisk to select entries
    that <b style='mso-bidi-font-weight:normal'>end</b> with the text.<o:p></o:p></span></p>
    </td>
   </tr>
   <tr style='mso-yfti-irow:6'>
    <td width=71 style='width:53.6pt;border:solid windowtext 1.0pt;border-top:
    none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
    padding:0in 5.4pt 0in 5.4pt'>
    <p class=MsoNormal align=center style='text-align:center'><b
    style='mso-bidi-font-weight:normal'><span style='font-size:10.0pt;
    mso-bidi-font-size:12.0pt;font-family:Arial'>underscore<o:p></o:p></span></b></p>
    </td>
    <td width=519 valign=top style='width:389.2pt;border-top:none;border-left:
    none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
    mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
    mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
    <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
    font-family:Arial'>The underscore can be used to match on single
    characters.<o:p></o:p></span></p>
    </td>
   </tr>
   <tr style='mso-yfti-irow:7;mso-yfti-lastrow:yes'>
    <td width=71 style='width:53.6pt;border:solid windowtext 1.0pt;border-top:
    none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
    padding:0in 5.4pt 0in 5.4pt'>
    <p class=MsoNormal align=center style='text-align:center'><b
    style='mso-bidi-font-weight:normal'><span style='font-size:10.0pt;
    mso-bidi-font-size:12.0pt;font-family:Arial'>brackets<o:p></o:p></span></b></p>
    </td>
    <td width=519 valign=top style='width:389.2pt;border-top:none;border-left:
    none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
    mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
    mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
    <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
    font-family:Arial'>Used with asterisks, brackets specify a set of
    characters where at least one must be matched for an entry to be selected.<o:p></o:p></span></p>
    </td>
   </tr>
  </table>
  <p class=MsoBodyText><o:p></o:p></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:2;mso-yfti-lastrow:yes'>
  <td width=590 valign=top style='width:6.15in;padding:0in 5.4pt 0in 5.4pt'>
  <p class=MsoBodyText><o:p>&nbsp;</o:p></p>
  <p class=MsoBodyText><o:p>&nbsp;</o:p></p>
  <p class=MsoBodyText><b style='mso-bidi-font-weight:normal'><span
  style='mso-bidi-font-size:10.0pt'>Search Examples<o:p></o:p></span></b></p>
  <p class=MsoBodyText><o:p>&nbsp;</o:p></p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l1 level1 lfo4;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b>“Spinning Wheel”</b> would retrieve every
  entry that contains the word “Spinning” followed by the word “Wheel”,
  separated by a space.<span style='mso-spacerun:yes'>   </span></p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l1 level1 lfo4;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b>“Spinning Wheel” or Spinning-Wheel</b><span
  style='mso-bidi-font-weight:bold'> would retrieve all entries that contain
  either the phrase “Spinning Wheel” or the word “Spinning-Wheel”.</span></p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l1 level1 lfo4;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b>Rock and Roll</b> would retrieve all
  entries that contain both the words “Rock” and “Roll” anywhere in the
  searched field.<span style='mso-spacerun:yes'>  </span>Entering <b>Rock Roll</b>
  would produce the same results since “and” is assumed.</p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l1 level1 lfo4;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b>Rock and -Roll</b> would retrieve all
  entries that contain the word “Rock” but do not contain the word “Roll”
  anywhere in the searched field.<span style='mso-spacerun:yes'>  </span></p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l1 level1 lfo4;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b>(Peanut AND Butter) AND (Jelly OR Jam)</b>
  retrieves all entries that contain the words “peanut” and “butter” and
  include either the word “jelly” or the word “jam”.</p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l1 level1 lfo4;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b>*Spinning </b><span style='mso-bidi-font-weight:
  bold'>would retrieve every entry that ends with the word “spinning”.<span
  style='mso-spacerun:yes'>  </span></span></p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l1 level1 lfo4;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b>Spinning*</b><span style='mso-bidi-font-weight:
  bold'> would retrieve every entry that begins with the word “spinning”.</span></p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l1 level1 lfo4;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b>damas_</b><span style='mso-bidi-font-weight:
  bold'> would retrieve every entry that contains the word “damas” followed by
  any single character.</span></p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l1 level1 lfo4;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b>[0123456789]*</b><span style='mso-bidi-font-weight:
  bold'> would retrieve every entry that begins with a number</span></p>
  <p class=MsoBodyText style='margin-left:.6in;text-indent:-.15in;mso-list:
  l1 level1 lfo4;tab-stops:list .6in'><![if !supportLists]><span
  style='font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:
  Symbol'><span style='mso-list:Ignore'>·<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;
  </span></span></span><![endif]><b>*[?!]</b><span style='mso-bidi-font-weight:
  bold'> would retrieve every entry that ends with a “!” or a “?”</span></p>
  <p class=MsoBodyText><o:p>&nbsp;</o:p></p>
  </td>
 </tr>
</table>

</div>

<p class=MsoNormal><o:p>&nbsp;</o:p></p>

</div>

</body>

</html>

