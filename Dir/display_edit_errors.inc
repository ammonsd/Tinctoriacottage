
<!--#INCLUDE FILE="../dir/box_it.inc" -->

<%
   DIM strBackMsg, bolNonEditErr, strDisplayMsg

SUB Display_Errors

   DIM intSub, intBackPage

   intBackPage=1
   If SESSION("BPG") <> "" Then
      intBackPage = intBackPage + SESSION("BPG")
   End If

   If strBackMsg = "" Then
      strBackMsg = "<b>Click the ""Back"" button to make corrections.</b> "
      strBackMsg = ""
   ElseIf strBackMsg = "." Then
      strBackMsg = ""
   End If

   Response.Clear
   bolNoTable = true
   bolNoTable = true
   Call Setup_Web_Page("Data Input Errors","")
   Call Box_Top_Section
   If strDisplayMsg <> "" Then
      Response.Write strDisplayMsg & "<br><br>"
   End If%>

   <font color="#FF0000">
   <%
   For intSub = 0 To errMsgs.Count -1
      response.Write ( "<b>&nbsp;&nbsp;&nbsp;&#183;&nbsp;&nbsp;" & aryErrorMsgs(intSub) & "</b><br>" )
   Next
   %>
    </font>
    <br>
    <div align="center">
    <%=strBackMsg%>
    <br>
    <INPUT TYPE="button" VALUE="BACK" onClick="history.go(-<%=intBackPage%>)">
    </div>
</td>

<%
Call Box_Bottom_Section
Response.End
END SUB
%>
