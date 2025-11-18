
<%

SUB Insert_Navigation_Details(TYP)

    DIM intSub
    intSub = aryMaintLinks(0,0)

    If strMbrHelpFn <> "" Then
      intSub = intSub + 1
      aryMaintLinks(intSub,1) = strMbrHelpFn
      aryMaintLinks(intSub,2) = "Documentation"
    End If
    aryMaintLinks(0,0) = intSub

End SUB
%>
