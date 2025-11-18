<%
FUNCTION GetSelection(SELECTIONS,TYP,PRTITLE,PGM)

    DIM intSub, strWord

    GetSelection = SELECTIONS

    If instr(SELECTIONS,";") > 0 Then
       GetSelection = REPLACE(GetSelection,";"," ")
       bolDoPhrases = false
       intSub = 1
       strWord = GetWords(GetSelection,intSub,1)
       do while strWord <> ""
          aryData(1,intSub) = strWord
          aryData(2,intSub) = strWord
          aryData(3,intSub) = "R"
          aryData(4,intSub) = TYP
          intSub = intSub + 1
          strWord = GetWords(GetSelection,intSub,1)
       loop
       aryData(0,0) = intSub - 1
       strFldAlign = "R"
       call Get_Data(PGM,PRTITLE,60)
       Response.End
    End If

END FUNCTION
%>
