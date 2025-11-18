<%

FUNCTION FormatSecurityGroups(KW)

    DIM intLoc, strKW

    If trim(KW) = "" Then
       EXIT FUNCTION
    End If


    strKW = KW & ","
    FormatSecurityGroups = ":"

    do while strKW <> ""
       intLoc = instr(strKW,",")
       FormatSecurityGroups = FormatSecurityGroups & trim(left(StrKW,intLoc-1)) & ":"
       strKW = trim(mid(strKW,intLoc+1))
    loop

END FUNCTION


FUNCTION UnFormatSecurityGroups(KW)

    DIM intLoc

    If trim(KW) = "" Then
       EXIT FUNCTION
    End If

    UnFormatSecurityGroups = ReplChars(mid(KW,2,len(KW)-2),":",",")

END FUNCTION


%>
