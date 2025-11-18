
<%
FUNCTION ReplChars(DATA, FINDCHAR, REPCHAR)

   DIM intLoc, intFactor

   ReplChars = DATA

   IF DATA="" or isnull(DATA) or FINDCHAR="" or isnull(FINDCHAR) Then
      EXIT FUNCTION
   End If

   If FINDCHAR = REPCHAR Then
      FINDCHAR = ""
   End If

   ReplChars=REPLACE(DATA,FINDCHAR,REPCHAR)

END FUNCTION

FUNCTION ReplParmChars(DATA)

'Replace " & "
   ReplParmChars = ReplChars(DATA," & ","%20%26%20")

'Replace Spaces
   ReplParmChars = ReplChars(ReplParmChars," ","%20")

'Replace Quote
   ReplParmChars = ReplChars(ReplParmChars,CHR(34),"%22")

'Replace Line Breaks
   ReplParmChars = ReplChars(ReplParmChars,CHR(13),"%0d")

END FUNCTION

FUNCTION RemoveParmChars(DATA)

   RemoveParmChars = ReplChars(DATA,"%20"," ")
   RemoveParmChars = ReplChars(RemoveParmChars,"%22",CHR(34))
   RemoveParmChars = ReplChars(RemoveParmChars,"&lt;","<")
   RemoveParmChars = ReplChars(RemoveParmChars,"&gt;",">")

END FUNCTION

FUNCTION ReplBrackets(DATA)

   ReplBrackets = ReplChars(DATA,"<","&lt;")
   ReplBrackets = ReplChars(ReplBrackets,">","&gt;")

END FUNCTION

FUNCTION ReplLineBreaks(DATA)

   ReplLineBreaks = ReplChars(DATA,CHR(13),"<br>")
   ReplLineBreaks = RemoveLineBreaks(ReplLineBreaks)

END FUNCTION

FUNCTION ReplLineBreaksJS(DATA)

   ReplLineBreaksJS = ReplChars(DATA,CHR(13),"\n")
   ReplLineBreaksJS = RemoveLineBreaks(ReplLineBreaksJS)

END FUNCTION

FUNCTION RemoveLineBreaks(DATA)
    RemoveLineBreaks = ReplChars(DATA,CHR(13),"")
    RemoveLineBreaks = ReplChars(RemoveLineBreaks,CHR(10),"")
END FUNCTION

FUNCTION ReplMultiBlanks(DATA)

   ReplMultiBlanks = ReplChars(DATA,"  ","&nbsp;&nbsp;")

END FUNCTION

FUNCTION InsertDoubleBS(DATA)

   DIM intLoc

   InsertDoubleBS = DATA
   If left(InsertDoubleBS,1) = "\" Then
      InsertDoubleBS = "\" & InsertDoubleBS
      intLoc = 3
   Else
      intLoc = 1
   End If

   Do while intLoc > 0
      intLoc = instr(intLoc,InsertDoubleBS,"\")
      If intLoc = 0 Then
        EXIT DO
      End If
      InsertDoubleBS = left(InsertDoubleBS,intLoc-1) & "\\" & mid(InsertDoubleBS,intLoc+1)
      intLoc = intLoc + 2
   Loop

END FUNCTION

FUNCTION RemoveSpaces(DATA)

   RemoveSpaces = TRIM(ReplChars(DATA," ",""))

END FUNCTION

FUNCTION TrimSpaces(DATA,NBR)

   DIM intWord, strWord

   If isnumeric(NBR) Then
      NBR = int(NBR)
   Else
      NBR = 1
   End If

   If instr(DATA," ") = 0 Then
      TrimSpaces = DATA
      EXIT FUNCTION
   End If

   bolDoPhrases = false
   TrimSpaces = GetWords(DATA,1,1)
   intWord = 2
   strWord = GetWords(DATA,intWord,1)

   Do While strWord <> ""
      TrimSpaces = TrimSpaces & SPACE(NBR) & strWord
      intWord = intWord + 1
      strWord = GetWords(DATA,intWord,1)
   Loop

END FUNCTION

FUNCTION CountChars(DATA,CHAR)

   DIM strData, intLoc, strChar

   CountChars = 0
   strData = ucase(DATA)
   strChar = ucase(CHAR)

   do while strData <> ""
      intLoc = instr(strData,strChar)
      If intLoc = 0 Then
         EXIT DO
      Else
         CountChars = CountChars + 1
         strData=mid(strData,intLoc+len(strChar))
      End if
   Loop

END FUNCTION
 %>
