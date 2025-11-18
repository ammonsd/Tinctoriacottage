
<% FUNCTION Encrypt(DATA,KEY)

    DIM intLoc, intCntr, intFactor, strKey

    Do while len(strKey) < 60
       strKey = strKey & KEY
    Loop

    Encrypt = ""
    If DATA = "" or isnull(DATA) Then
       EXIT FUNCTION
    End If

    intFactor = Month(Date()) + Day(Date())
    intCNtr = 0

    For intLoc = 1 to len(DATA)
       if intCntr = 3 and intLoc < len(strkey) Then
          on error resume next
          Encrypt = Encrypt & CHR(ASC(mid(strKey,len(strKey)-intLoc,1)) + (len(strKey) + intFactor))
          intCntr = 0
       End If
       intCntr = intCntr + 1
       on error resume next
       Encrypt = Encrypt & CHR(ASC(mid(DATA,intLoc,1)) + (len(strKey) + intFactor))
    Next

    Encrypt = left(Encrypt,1) & CHR(intFactor) & mid(Encrypt,2)

End FUNCTION %>


<% FUNCTION Decrypt(DATA,KEY)

    DIM intLoc, intCntr, intFactor, strEnData, strKey

    Do while len(strKey) < 60
       strKey = strKey & KEY
    Loop

    Decrypt = ""
    If DATA = "" or isnull(DATA) Then
       EXIT FUNCTION
    End If

    on error resume next
    intFactor = ASC(mid(DATA,2,1))
    intCntr = 0
    strEnData = left(DATA,1) & mid(DATA,3)

    If strEnData <> "" Then
       For intLoc = 1 to len(strEnData)
          intCntr = intCntr + 1
          if intCntr = 4 and intLoc < len(strkey) Then
             intCntr = 0
          Else
             on error resume next
             Decrypt = Decrypt & CHR(ASC(mid(strEnData,intLoc,1)) - (len(strKey) + intFactor))
          End If
    Next
    End If

End FUNCTION %>
