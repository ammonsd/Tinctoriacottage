<%

     DIM aryFDtls(4,300), strFolderName, intFolderFileCount, intFolderSize, bolFullFileName, strFilePath, strSelDate

     '1 - Name
     '2 - Path
     '3 - Size
     '4 - Last Modified Date

SUB Get_Folder_File_Details(PATH,FSPEC)

     DIM objFS, objFolder, file, objItem, strUrl, intSub, strWord, bolOK, intWordNbr, strPath, strFSpec

     aryFDtls(0,0) = 0

     strFSpec = ucase(FSPEC)

     if left(strFSpec,1) = "*" Then
        strFSpec = mid(strFSpec,2)
     end if

     strFSpec = Trim(Replace(strFSpec,"*", " "))

     strPath = Replace(PATH, "\", "/")
     If left(strPath,1) <> "/" Then
        strPath = "/" & strPath
     End If

     on error resume next
     strPath = Server.MapPath(strPath)

     If err.Number <> 0 Then
        EXIT SUB
     End If

     set objFS = CreateObject("Scripting.FileSystemObject")
     set objFolder = objFS.GetFolder(strPath)

     strFolderName = objFolder.Name
     intFolderFileCount = objFolder.Files.Count
     intFolderSize = objFolder.Size

     intSub = 0
     bolDoPhrases = false

     For each objItem in objFolder.Files
       bolOK = false
       If strFSpec = "" Then
          bolOK = true
       Else
          bolOK = true
          intWordNbr = 1
          strWord = GetWords(strFSpec,intWordNbr,1)

          do while strWord <> ""
             If instr(ucase(objItem.Name),strWord) = 0 Then
                bolOK = false
                EXIT DO
             End If
             intWordNbr = intWordNbr + 1
             strWord = GetWords(strFSpec,intWordNbr,1)
          Loop
       End If
       If bolOk and strSelDate <> "" Then
          If strSelDate <> left(objItem.DateLastModified,instr(objItem.DateLastModified," ")-1) Then
             bolOk = false
          End If
       End If
       If bolOk Then
          intSub = intSub + 1
          aryFDtls(1,intSub) = objItem.Name
          aryFDtls(2,intSub) = objItem.Path
          aryFDtls(3,intSub) = objItem.Size
          aryFDtls(4,intSub) = objItem.DateLastModified
       End If
     Next
     aryFDtls(0,0) = intSub

     set objFS = nothing
     set objFolder = nothing
     set objItem = nothing

END SUB

FUNCTION MapURL(PATH)

     'Convert the file path to a URL address

     MapUrl = Right(PATH, Len(PATH) - Len(Server.MapPath("/")))
     If bolFullFileName Then
        MapUrl = Request.ServerVariables("SERVER_NAME") & MapUrl
     End If
     MapURL = Replace(MapUrl, "\", "/")

END FUNCTION

FUNCTION GetFileName(FN)

     DIM intLoc

     GetFileName = Replace(FN,"\","/")

     For intLoc = len(GetFileName) to 1 Step -1
        If mid(GetFileName,intLoc,1) = "/" Then
           strFilePath = left(GetFileName,intLoc)
           GetFileName = mid(GetFileName,intLoc+1)
           EXIT FOR
        End If
     Next

END FUNCTION

FUNCTION FileCount(PATH)

     DIM strFSpec, intLoc, strPath

     strPath = Replace(PATH, "\", "/")

     intLoc = instr(strPath,"*")
     If intLoc > 0 Then
        strFSpec = mid(strPath,intLoc)
        strPath = left(strPath,intLoc-1)
     End if

     Call Get_Folder_File_Details(strPath, strFSpec)

     FileCount = aryFDtls(0,0)

END FUNCTION

%>

