
<%
    DIM aryFunctionGrps(20)

SUB Get_Function_Security_Groups

    DIM intSub

    intSub = 1
    aryFunctionGrps(intSub) = "SecMaint"
    intSub = intSub + 1
    aryFunctionGrps(intSub) = "DirDelete"
    intSub = intSub + 1
    aryFunctionGrps(intSub) = "VCF"
    intSub = intSub + 1
    aryFunctionGrps(intSub) = "WebMail"
    aryFunctionGrps(0) = intSub

END SUB
%>

