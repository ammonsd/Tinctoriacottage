
<%

SUB Setup_ShowDetails_Script

   Response.Write "<script type='text/javascript'>" & vbCrLf
   Response.Write "function ShowDetails(EM) " & vbCrLf
   Response.Write "{" & vbCrLf
   Response.Write "alert(EM)" & vbCrLf
   Response.Write "return false;" & vbCrLf
   Response.Write "}" & vbCrLf
   Response.Write "</script>" & vbCrLf

END SUB

SUB GoUrl_Code%>

<SCRIPT LANGUAGE="javascript">
<!-- hide from old browsers
function goUrl(Selected)
   {
   /* change page location to option selected from navigation drop-down */
   var newURL = Selected.options[Selected.selectedIndex].value ;
   if (newURL!=" ") {
      parent.location.href = newURL;
   }
}
// End of javascript -->
</SCRIPT>

<%END SUB%>

<%SUB GoBack_Pages

   DIM intNbrPages

   intNbrPages = SESSION("BPG")+1
   SESSION("BPG") = ""

%>
<html>
<SCRIPT LANGUAGE="javascript">
<!-- hide from old browsers
function goBack(Pages)
   {
    history.go(-Pages)
   }
// End of javascript -->
</SCRIPT>
<body onload="goBack(<%=intNbrPages%>)">
</body></html>
<%END SUB%>

<%SUB GoUrl_Code_Frame%>

<SCRIPT LANGUAGE="javascript">
<!-- hide from old browsers
function goUrl(Selected)
   {
   /* change page location to option selected from navigation drop-down */
   var newURL = Selected.options[Selected.selectedIndex].value ;
   if (newURL!=" ") {
      parent.main.location.href = newURL;
   }
}
// End of javascript -->
</SCRIPT>

<%END SUB%>


<%SUB GoUrl_Code_NewWindow%>

<SCRIPT LANGUAGE="javascript">
<!-- hide from old browsers
function goUrl(Selected)
   {
   /* change page location to option selected from navigation drop-down */
   var newURL = Selected.options[Selected.selectedIndex].value ;
   if (newURL!=" ") {
      window.open(newURL);
   }
}
// End of javascript -->
</SCRIPT>

<%END SUB%>


<%

   DIM bolCSSDone, strTitleClr, strTitleClr2, strBgClr, strBgClr2

SUB SetUp_Window_CSS

   If bolCSSDone Then
      EXIT SUB
   End If

   If strTitleClr = "" Then
      strTitleClr = "#000095"
   End If
   If strTitleClr2 = "" Then
      strTitleClr2 = strTitleClr
   End If
   If strBgClr = "" Then
      strBgClr = "#B4B4B4"
   End If
   If strBgClr2 = "" Then
      strBgClr2 = strBgClr
   End If

   bolCSSDone = true
%>

<style type="text/css">
   input.prompt {border:1 solid transparent; background-color:<%=strBgClr%>;width:70;font-family:arial;font-size:12; color:black;}
   input.prompt2 {border:1 solid transparent; background-color:<%=strBgClr2%>;width:70;font-family:arial;font-size:12; color:black;}
   td.titlebar { background-color:<%=strTitleClr%>; color:white; font-weight:bold;font-family:arial; font-size:12;}
   td.titlebar2 { background-color:<%=strTitleClr2%>; color:white; font-weight:bold;font-family:arial; font-size:12;}
   table.promptbox {border:1 solid #ccccff; background-color:<%=strBgClr%>; color:black;padding-left:2;padding-right:2;padding-bottom:2;font-family:arial; font-size:12;}
   table.promptbox2 {border:1 solid #ccccff; background-color:<%=strBgClr2%>; color:black;padding-left:2;padding-right:2;padding-bottom:2;font-family:arial; font-size:12;}
   input.promptbox {border:1 solid #0000FF; background-color:white;width:100%;font-family:arial;font-size:12; color:black; }
</style>

<%END SUB%>

<%
SUB SetUp_Data_Prompt

   Call SetUp_Window_CSS
%>

<SCRIPT language='javascript'>

function DataPrompt(prompttitle, message, prompttext, sendto) {
   promptbox = document.createElement('div');
   promptbox.setAttribute ('id' , 'prompt')

   document.getElementsByTagName('body')[0].appendChild(promptbox)
   promptbox = eval("document.getElementById('prompt').style")
   promptbox.position = 'absolute'
   promptbox.top = 350
   promptbox.left = 225
   promptbox.width = 300
   promptbox.border = 'outset 1 #bbbbbb'
   document.getElementById('prompt').innerHTML = "<table cellspacing='0' cellpadding='0' border='0' width='100%'><tr valign='center'><td width='22' height='22' style='text-indent:2;' class='titlebar'></td><td class='titlebar'>" + prompttitle + "</td></tr></table>"
   document.getElementById('prompt').innerHTML = document.getElementById('prompt').innerHTML + "<table cellspacing='0' cellpadding='0' border='0' width='100%' class='promptbox'><tr><td>" + message + "</td></tr><tr><td><input type='text' id='promptbox' onblur='this.focus()' class='promptbox' value='" + prompttext + "'></td></tr><tr><td align='center'><br><input type='button' class='prompt' value='Process' onMouseOver='this.style.border=\"1 outset #dddddd\"' onMouseOut='this.style.border=\"1 solid transparent\"' onClick='" + sendto + "(document.getElementById(\"promptbox\").value); document.getElementsByTagName(\"body\")[0].removeChild(document.getElementById(\"prompt\"))'> <input type='button' class='prompt' value='Cancel' onMouseOver='this.style.border=\"1 outset transparent\"' onMouseOut='this.style.border=\"1 solid transparent\"' onClick='" + sendto + "(\"\"); document.getElementsByTagName(\"body\")[0].removeChild(document.getElementById(\"prompt\"))'></td></tr></table>"
   document.getElementById("promptbox").focus()
}
</SCRIPT>

<SCRIPT LANGUAGE="javascript">
<!--
function CallDataPrompt(Title,Msg,DefPrompt){
   DataPrompt('<div align="center">' + Title + '</div>', Msg, DefPrompt, 'GetAttachFileSpec');
}
//-->
</SCRIPT>

<%END SUB%>

<%


   DIM intBoxTop, intBoxLeft, intBoxWidth

SUB SetUp_PopUp_Msg

   Call SetUp_Window_CSS

   If intBoxTop = "" or not isNumeric(intBoxTop) Then
      intBoxTop = 125
   End If
   If intBoxLeft = "" or not isNumeric(intBoxLeft) Then
      intBoxLeft = 325
   End If
   If intBoxWidth = "" or not isNumeric(intBoxWidth) Then
      intBoxWidth = 500
   End If
%>

<SCRIPT language='javascript'>

function PopUpMsg(Title, Msg,LocFactor) {

   if (LocFactor < 1) {
      LocFactor = 1;
   }

   promptbox = document.createElement('div');
   promptbox.setAttribute ('id' , 'prompt')

   sendto = 'OK'
   document.getElementsByTagName('body')[0].appendChild(promptbox)
   promptbox = eval("document.getElementById('prompt').style")
   promptbox.position = 'absolute'
   promptbox.top = <%=intBoxTop%>*LocFactor
   promptbox.left = <%=intBoxLeft%>
   promptbox.width = <%=intBoxWidth%>
   promptbox.border = 'outset 1 #bbbbbb'
   document.getElementById('prompt').innerHTML = "<table cellspacing='0' cellpadding='0' border='0' width='100%'><tr valign='center'><td width='22' height='22' style='text-indent:2;' class='titlebar'></td><td class='titlebar' align='center'><font size=2>" + Title + "</font></td></tr></table>"
   document.getElementById('prompt').innerHTML = document.getElementById('prompt').innerHTML + "<table cellspacing='10' cellpadding='0' border='0' width='100%' class='promptbox'><tr><td><br>" + Msg + "</td></tr><tr><td align='center'><br><input type='button' class='prompt' value='OK' onMouseOver='this.style.border=\"1 outset transparent\"' onMouseOut='this.style.border=\"1 solid transparent\"' onClick='" + sendto + "(\"\"); document.getElementsByTagName(\"body\")[0].removeChild(document.getElementById(\"prompt\"))'></td></tr></table>"
}
function OK() {
}
</script>

<%END SUB %>

<%

SUB SetUp_PopUp_Pic

   Call SetUp_Window_CSS
%>

<SCRIPT language='javascript'>

function PopUpPic(Pic,LocFactor) {

   if (LocFactor < 1) {
      LocFactor = 1;
   }

   promptbox = document.createElement('div');
   promptbox.setAttribute ('id' , 'prompt')

   sendto = 'OK'
   document.getElementsByTagName('body')[0].appendChild(promptbox)
   promptbox = eval("document.getElementById('prompt').style")
   promptbox.position = 'absolute'
   promptbox.top = 125*LocFactor
   promptbox.left = 325
   promptbox.width = 500
   promptbox.border = 'outset 1 #bbbbbb'
   document.getElementById('prompt').innerHTML = document.getElementById('prompt').innerHTML + "<table cellspacing='10' cellpadding='0' border='0' width='100%' class='promptbox2'><tr><td align='center'><img src='" + Pic + "' alt='Loading Picture' border=0></td></tr><tr><td align='center'><br><input type='button' class='prompt2' value='OK' onMouseOver='this.style.border=\"1 outset transparent\"' onMouseOut='this.style.border=\"1 solid transparent\"' onClick='" + sendto + "(\"\"); document.getElementsByTagName(\"body\")[0].removeChild(document.getElementById(\"prompt\"))'></td></tr></table>"
}

function OK() {
}
</script>

<%END SUB %>
