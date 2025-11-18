
<%

SUB Insert_Navigation_Details(TYP)

    DIM intSub

    intSub = aryMaintLinks(0,0)

    If strLstKwFn <> "" and TYP = "S" Then
      intSub = intSub + 1
      aryMaintLinks(intSub,1) = strLstKwFn & Application("DIRTBL") & "&NC=Y'"
      aryMaintLinks(intSub,2) = "Assigned Keywords"
    End If

    If strDirHelpFn <> "" Then
      intSub = intSub + 1
      aryMaintLinks(intSub,1) = strDirHelpFn
      aryMaintLinks(intSub,2) = "Documentation"
    End If
    aryMaintLinks(0,0) = intSub

End SUB
%>
