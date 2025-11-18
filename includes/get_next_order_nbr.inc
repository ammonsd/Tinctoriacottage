
<%FUNCTION GetNextOrdNbr

   Dim objFn, objTextFn, strData, strFileName

   strFileName = Server.Mappath("\CGI-BIN\") & "\last_order_nbr.TXT"

   Set objFn = CreateObject("Scripting.FileSystemObject")
   on error resume next
   Set objTextFn = objFn.OpenTextFile(strFileName)
   GetNextOrdNbr = 1000

   If err.Number = 0 then
      strData = objTextFn.ReadLine
      if isnumeric(strData) = true then
         GetNextOrdNbr = strData + 1
      end if
   End If

   objTextFn.Close
   Set objTextFn = Nothing
   Set objFn = Nothing

   Set objFn = CreateObject("Scripting.FileSystemObject")
   Set ObjTextFn = ObjFn.CreateTextFile(strFileName, true)

   ObjTextFn.WriteLine GetNextOrdNbr

   ObjTextFn.close
   set ObjTextFn=nothing
   set ObjFn=nothing

END FUNCTION%>
