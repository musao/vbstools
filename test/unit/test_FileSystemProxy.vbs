' FileSystemProxy.vbs: test.
' @import ../../lib/com/FileSystemProxy.vbs
' @import ../../lib/com/ArrayList.vbs
' @import ../../lib/com/Broker.vbs
' @import ../../lib/com/BufferedReader.vbs
' @import ../../lib/com/BufferedWriter.vbs
' @import ../../lib/com/Calendar.vbs
' @import ../../lib/com/CharacterType.vbs
' @import ../../lib/com/CssGenerator.vbs
' @import ../../lib/com/HtmlGenerator.vbs
' @import ../../lib/com/ReadOnlyObject.vbs
' @import ../../lib/com/ReturnValue.vbs
' @import ../../lib/com/libCom.vbs

' @import ../../lib/libEnum.vbs

Option Explicit

Const MY_NAME = "test_FileProxy.vbs"
Dim PsPathTempFolder
Dim PeEntryType : Set PeEntryType = dicOf( _
    Array( _
        "ENTRY", 0 _
        , "FILE", 1 _
        , "FILE_EXCLUDING_ARCHIVE", 2 _
        , "FOLDER", 3 _
        , "CONTAINER", 4 _
    ) _
)
Dim PeDataType : Set PeDataType = dicOf( _
    Array( _
        "SHORTCUT_FILE", 0 _
        , "URL_SHORTCUT_FILE", 1 _
        , "TEXT_FILE", 2 _
        , "FOLDER", 3 _
    ) _
)
'###################################################################################################
'SetUp()/TearDown()
Sub SetUp()
    '実行スクリプト直下に当ファイル名で一時フォルダ作成
    PsPathTempFolder = fso().BuildPath(fso().GetParentFolderName(WScript.ScriptFullName), MY_NAME)
    If Not (fso().FolderExists(PsPathTempFolder)) Then fso().CreateFolder(PsPathTempFolder)
End Sub
Sub TearDown()
    '当テストで作成した一時フォルダを削除する
    fso().DeleteFolder PsPathTempFolder
End Sub

'###################################################################################################
'FileSystemProxy
Sub Test_FileProxy
    Dim a : Set a = new FileSystemProxy
    AssertEqualWithMessage 1, VarType(a), "VarType"
    AssertEqualWithMessage "FileSystemProxy", TypeName(a), "TypeName"
End Sub

'###################################################################################################
'FileSystemProxy.<properties>_initial
'   .allContainers
'   .allContainersIncludingSelf
'   .allEntries
'   .allEntriesIncludingSelf
'   .allFilesExcludingArchives
'   .allFilesExcludingArchivesIncludingSelf
'   .baseName
'   .containers
'   .dateLastModified
'   .entries
'   .extension
'   .filesExcludingArchives
'   .hasContainers
'   .hasEntries
'   .hasFilesExcludingArchives
'   .isBrowsable
'   .isFileSystem
'   .isFolder
'   .isLink
'   .name
'   .parentFolder
'   .path
'   .size
'   .toString
'   .type
Sub Test_FileProxy_proeerties_initial
    Dim ao : Set ao = (new FileSystemProxy)
    Dim ele, target, expected, actual
    For Each ele In Array( _
            dicOf( Array(  "target","allContainers"                         ,"expected",Null               , "actual", ao.allContainers) ) _
            , dicOf( Array("target","allContainersIncludingSelf"            ,"expected",Null               , "actual", ao.allContainersIncludingSelf) ) _
            , dicOf( Array("target","allEntries"                            ,"expected",Null               , "actual", ao.allEntries) ) _
            , dicOf( Array("target","allEntriesIncludingSelf"               ,"expected",Null               , "actual", ao.allEntriesIncludingSelf) ) _
            , dicOf( Array("target","allFilesExcludingArchives"             ,"expected",Null               , "actual", ao.allFilesExcludingArchives) ) _
            , dicOf( Array("target","allFilesExcludingArchivesIncludingSelf","expected",Null               , "actual", ao.allFilesExcludingArchivesIncludingSelf) ) _
            , dicOf( Array("target","baseName"                              ,"expected",Null               , "actual", ao.baseName) ) _
            , dicOf( Array("target","containers"                            ,"expected",Null               , "actual", ao.containers) ) _
            , dicOf( Array("target","dateLastModified"                      ,"expected",Null               , "actual", ao.dateLastModified) ) _
            , dicOf( Array("target","entries"                               ,"expected",Null               , "actual", ao.entries) ) _
            , dicOf( Array("target","extension"                             ,"expected",Null               , "actual", ao.extension) ) _
            , dicOf( Array("target","filesExcludingArchives"                ,"expected",Null               , "actual", ao.filesExcludingArchives) ) _
            , dicOf( Array("target","hasContainers"                         ,"expected",Null               , "actual", ao.hasContainers) ) _
            , dicOf( Array("target","hasEntries"                            ,"expected",Null               , "actual", ao.hasEntries) ) _
            , dicOf( Array("target","hasFilesExcludingArchives"             ,"expected",Null               , "actual", ao.hasFilesExcludingArchives) ) _
            , dicOf( Array("target","isBrowsable"                           ,"expected",Null               , "actual", ao.isBrowsable) ) _
            , dicOf( Array("target","isFileSystem"                          ,"expected",Null               , "actual", ao.isFileSystem) ) _
            , dicOf( Array("target","isFolder"                              ,"expected",Null               , "actual", ao.isFolder) ) _
            , dicOf( Array("target","isLink"                                ,"expected",Null               , "actual", ao.isLink) ) _
            , dicOf( Array("target","name"                                  ,"expected",Null               , "actual", ao.name) ) _
            , dicOf( Array("target","parentFolder"                          ,"expected",Null               , "actual", ao.parentFolder) ) _
            , dicOf( Array("target","path"                                  ,"expected",Null               , "actual", ao.path) ) _
            , dicOf( Array("target","size"                                  ,"expected",Null               , "actual", ao.size) ) _
            , dicOf( Array("target","toString"                              ,"expected","<FileSystemProxy>", "actual", ao.toString) ) _
            , dicOf( Array("target","type"                                  ,"expected",Null               , "actual", ao.type) ) _
    )
        target = ele("target")
        expected = ele("expected")
        actual = ele("actual")
        AssertEqualWithMessage expected, actual, target&"(initial)"
    Next
End Sub
Sub Test_FileProxy_properties
    Dim data : data = createData()
    Dim cs,path,ao
    For cs=0 To Ubound(data)
        path = data(cs)
        Set ao = new FileSystemProxy : ao.of(path)

'       .baseName
'       .dateLastModified
'       .extension
'       .isBrowsable
'       .isFileSystem
'       .isFolder
'       .isLink
'       .name
'       .parentFolder
'       .path
'       .size
'       .toString
'       .type
        assertFsProperties ao,path,cs

'       .allContainers
'       .allContainersIncludingSelf
'       .allEntries
'       .allEntriesIncludingSelf
'       .allFilesExcludingArchives
'       .allFilesExcludingArchivesIncludingSelf
'       .containers
'       .entries
'       .filesExcludingArchives
'       .hasContainers
'       .hasEntries
'       .hasFilesExcludingArchives
        assertFsEntries ao,path,cs
    Next
End Sub

'###################################################################################################
'FileSystemProxy.of()
Sub Test_FileProxy_of
    Dim p,a
    p = WScript.ScriptFullName
    Set a = (new FileSystemProxy).of(p)

    AssertEqualWithMessage 8, VarType(a), "VarType"
    AssertEqualWithMessage "FileSystemProxy", TypeName(a), "TypeName"
    AssertEqualWithMessage p, a.path, "path"
End Sub
Sub Test_FileProxy_of_Err
    Dim a,d : d = vbNullString
    On Error Resume Next
    Call (new FileSystemProxy).of(d)

    AssertEqualWithMessage "FileSystemProxy+of()", Err.Source, "Err.Source"
    AssertMatchWithMessage "^invalid argument.*", Err.Description, "Err.Description"
End Sub
Sub Test_FileProxy_of_ErrImmutable
    On Error Resume Next
    Dim ao
    Set ao = (new FileSystemProxy).of(createShortCutFile())
    Call ao.of(createUrlShortCutFile())

    AssertEqualWithMessage "FileSystemProxy+of()", Err.Source, "Err.Source"
    AssertEqualWithMessage "Because it is an immutable variable, its value cannot be changed.", Err.Description, "Err.Description"
End Sub

'###################################################################################################
'FileSystemProxy.setParent()
Sub Test_FileProxy_setParent
    Dim p,a,e
    p = WScript.ScriptFullName
    Set a = (new FileSystemProxy).of(p)
    Set e = (new FileSystemProxy).of(fso.GetParentFolderName(p))

    a.setParent(e)
    AssertEqualWithMessage e, a.parentFolder, "parentFolder"
End Sub
Sub Test_FileProxy_setParent_2times
    Dim p,a,e
    p = WScript.ScriptFullName
    Set a = (new FileSystemProxy).of(p)
    Dim tmp:Set tmp = a.parentFolder

    Set e = (new FileSystemProxy).of(fso.GetParentFolderName(p))
    a.setParent(e)
    AssertEqualWithMessage e, a.parentFolder, "parentFolder"
End Sub
Sub Test_FileProxy_setParent_Err_Initial
    On Error Resume Next
    Dim p,a
    p = WScript.ScriptFullName
    Call (new FileSystemProxy).setParent(fso.GetParentFolderName(p))

    AssertEqualWithMessage "FileSystemProxy+setParent()", Err.Source, "Err.Source"
    AssertEqualWithMessage "Please set the value before setting the parent folder.", Err.Description, "Err.Description"
End Sub
Sub Test_FileProxy_setParent_Err_NotFileSystemProxy
    On Error Resume Next
    Dim p,a
    p = WScript.ScriptFullName
    Set a = (new FileSystemProxy).of(p)
    Call a.setParent(dictionary)

    AssertEqualWithMessage "FileSystemProxy+setParent()", Err.Source, "Err.Source"
    AssertEqualWithMessage "This is not FileSystemProxy.", Err.Description, "Err.Description"
End Sub
Sub Test_FileProxy_setParent_Err_NotPrentFolder
    On Error Resume Next
    Dim p,a
    p = WScript.ScriptFullName
    Set a = (new FileSystemProxy).of(p)
    Call a.setParent(a)

    AssertEqualWithMessage "FileSystemProxy+setParent()", Err.Source, "Err.Source"
    AssertEqualWithMessage "This is not a parent folder.", Err.Description, "Err.Description"
End Sub


'###################################################################################################
'common
Function createData
    Dim defs
'    defs=Array( _
'    Array( defShortCutFile()                         , defFolder(Empty)                               , defFolder("defTextFile,defShortCutFile")        ) _
'    , Array( defShortCutFile()                         , defFolder("defFolder(defTextFile),defShortCutFile")                               , defFolder("defTextFile,defShortCutFile")        ) _
'    )

    defs=Array( _
    Array() _
    , Array(                                             defFolder(Empty)                                                                                 ) _
    , Array( defUrlShortCutFile()                      , defFolder("defTextFile")                                                                         ) _
    , Array( defShortCutFile()   , defUrlShortCutFile(), defFolder("defTextFile,defShortCutFile")                                                         ) _
    , Array(                                             defFolder(Empty)                               , defFolder(Empty)                                ) _
    , Array( defTextFile()                             , defFolder(Empty)                               , defFolder("defShortCutFile")                    ) _
    , Array( defShortCutFile()   , defUrlShortCutFile(), defFolder(Empty)                               , defFolder("defTextFile,defShortCutFile")        ) _
    , Array(                                             defFolder("defShortCutFile")                   , defFolder("defUrlShortCutFile")                 ) _
    , Array( defShortCutFile()                         , defFolder("defUrlShortCutFile")                , defFolder("defTextFile,defShortCutFile")        ) _
    , Array( defUrlShortCutFile(), defTextFile()       , defFolder("defShortCutFile,defUrlShortCutFile"), defFolder("defTextFile,defShortCutFile")        ) _
    )
'covers all patterns
'    defs=Array( _
'    Array() _
'    , Array(                                             defFolder(Empty)                                                                                 ) _
'    , Array(                                             defFolder("defShortCutFile")                                                                     ) _
'    , Array(                                             defFolder("defUrlShortCutFile,defTextFile")                                                      ) _
'    , Array(                                             defFolder(Empty)                               , defFolder(Empty)                                ) _
'    , Array(                                             defFolder(Empty)                               , defFolder("defTextFile")                        ) _
'    , Array(                                             defFolder(Empty)                               , defFolder("defShortCutFile,defUrlShortCutFile") ) _
'    , Array(                                             defFolder("defShortCutFile")                   , defFolder("defUrlShortCutFile")                 ) _
'    , Array(                                             defFolder("defTextFile")                       , defFolder("defShortCutFile,defUrlShortCutFile") ) _
'    , Array(                                             defFolder("defUrlShortCutFile,defTextFile")    , defFolder("defShortCutFile,defUrlShortCutFile") ) _
'    , Array( defTextFile()                                                                                                                                ) _
'    , Array( defShortCutFile()                         , defFolder(Empty)                                                                                 ) _
'    , Array( defUrlShortCutFile()                      , defFolder("defTextFile")                                                                         ) _
'    , Array( defShortCutFile()                         , defFolder("defUrlShortCutFile,defTextFile")                                                      ) _
'    , Array( defUrlShortCutFile()                      , defFolder(Empty)                               , defFolder(Empty)                                ) _
'    , Array( defTextFile()                             , defFolder(Empty)                               , defFolder("defShortCutFile")                    ) _
'    , Array( defShortCutFile()                         , defFolder(Empty)                               , defFolder("defTextFile,defShortCutFile")        ) _
'    , Array( defShortCutFile()                         , defFolder("defUrlShortCutFile")                , defFolder("defTextFile")                        ) _
'    , Array( defShortCutFile()                         , defFolder("defUrlShortCutFile")                , defFolder("defTextFile,defShortCutFile")        ) _
'    , Array( defShortCutFile()                         , defFolder("defUrlShortCutFile,defTextFile")    , defFolder("defShortCutFile,defUrlShortCutFile") ) _
'    , Array( defTextFile()       , defShortCutFile()                                                                                                      ) _
'    , Array( defUrlShortCutFile(), defTextFile()       , defFolder(Empty)                                                                                 ) _
'    , Array( defTextFile()       , defShortCutFile()   , defFolder("defUrlShortCutFile")                                                                  ) _
'    , Array( defShortCutFile()   , defUrlShortCutFile(), defFolder("defTextFile,defShortCutFile")                                                         ) _
'    , Array( defUrlShortCutFile(), defTextFile()       , defFolder(Empty)                               , defFolder(Empty)                                ) _
'    , Array( defShortCutFile()   , defUrlShortCutFile(), defFolder(Empty)                               , defFolder("defTextFile")                        ) _
'    , Array( defShortCutFile()   , defUrlShortCutFile(), defFolder(Empty)                               , defFolder("defTextFile,defShortCutFile")        ) _
'    , Array( defUrlShortCutFile(), defTextFile()       , defFolder("defShortCutFile")                   , defFolder("defUrlShortCutFile")                 ) _
'    , Array( defTextFile()       , defShortCutFile()   , defFolder("defUrlShortCutFile")                , defFolder("defTextFile,defShortCutFile")        ) _
'    , Array( defUrlShortCutFile(), defTextFile()       , defFolder("defShortCutFile,defUrlShortCutFile"), defFolder("defTextFile,defShortCutFile")        ) _
'    )
    Dim cs,cases
    For Each cs In defs
        pusha cases,caseNormal(cs)
    Next
    For Each cs In extractTargetForZip(defs)
        pusha cases,caseZip(cs)
    Next
    createData = cases
End Function
Function caseNormal(cs)
    caseNormal=createFolderAt(PsPathTempFolder,cs)
End Function
Function caseZip(cs)
    Dim flPath : flPath=createFolderAt(PsPathTempFolder,cs)
    Dim ele,paths
    For Each ele in getFolderItem2(flPath(0)).GetFolder.Items
        push paths,ele.path
    Next
    
    Dim zipPath : zipPath=getTempFilePath(PsPathTempFolder,"zip")
    zip paths,zipPath
    
    Dim ret
    For Each ele In flPath
        push ret,Replace(ele,flPath(0),zipPath)
    Next
    caseZip=ret
End Function

Function extractTargetForZip(defs)
    Dim cs,ret : ret=Array()
    For Each cs In defs
        If Not containEmpty(cs) Then push ret,cs
    Next
    extractTargetForZip=ret
End Function
Function containEmpty(cs)
    containEmpty=True
    If isEmptyArray(cs) Then Exit Function
    If IsArray(cs) Then
        Dim ele
        For Each ele In cs
            If containEmpty(ele) Then Exit Function
        Next
    End If
    containEmpty=False
End Function
Function isEmptyArray(ar)
    isEmptyArray=False
    If Not IsArray(ar) Then Exit Function
    If Ubound(ar)<0 Then isEmptyArray=True
End Function

Function createShortCutFile
    createShortCutFile = createShortCutFileAt(PsPathTempFolder)
End Function
Function createUrlShortCutFile
    createUrlShortCutFile = createUrlShortCutFileAt(PsPathTempFolder)
End Function
Function createTextFile
    createTextFile = createTextFileAt(PsPathTempFolder)
End Function

Function createShortCutFileAt(basePath)
    createShortCutFileAt = createShortCutCommon(basePath,"lnk",WScript.ScriptFullName)
End Function
Function createUrlShortCutFileAt(basePath)
    createUrlShortCutFileAt = createShortCutCommon(basePath,"url","https://www.google.com/")
End Function
Function createTextFileAt(basePath)
    Dim path : path = getTempFilePath(basePath,"txt")
    With fso.OpenTextFile(path, tsMode.FOR_WRITING, True, tsFormat.UNICODE)
        .Write "hoge"
        .Close
    End With
    createTextFileAt = path
End Function
Function createSomeFileAt(tp, basePath)
    Select Case tp
    Case PeDataType("SHORTCUT_FILE")
        createSomeFileAt = createShortCutFileAt(basePath)
    Case PeDataType("URL_SHORTCUT_FILE")
        createSomeFileAt = createUrlShortCutFileAt(basePath)
    Case PeDataType("TEXT_FILE")
        createSomeFileAt = createTextFileAt(basePath)
    End Select
End Function
Function createEmptyFolderAt(basePath)
    Dim path : path = getTempFolderPath(basePath)
    fso.CreateFolder path
    createEmptyFolderAt=path
End Function
Function createFolderAt(basePath,content)
    Dim path : path = createEmptyFolderAt(basePath)
    
    Dim ret : push ret, path
    Dim ele, tp, cnt
    For Each ele In content
        tp = ele(0)
        If tp=PeDataType("FOLDER") Then
            cnt = ele(1)
            pushA ret, createFolderAt(path, cnt)
        Else
            pushA ret, createSomeFileAt(tp, path)
        End If
    Next
    createFolderAt = ret
End Function
Function defShortCutFile
    defShortCutFile = Array(PeDataType("SHORTCUT_FILE"))
End Function
Function defUrlShortCutFile
    defUrlShortCutFile = Array(PeDataType("URL_SHORTCUT_FILE"))
End Function
Function defTextFile
    defTextFile = Array(PeDataType("TEXT_FILE"))
End Function
Function defFolder(d)
    Dim pram : pram=Array()
    If Not IsEmpty(d) Then
        Dim ele,re,func,arg
        Set re = reOf("([a-zA-Z0-9_]+)\(([^)]+)\)", "igm")
        For Each ele In Split(d,",")
            If re.Test(ele) Then
                func = re.Replace(ele, "$1")
                arg = re.Replace(ele, "$2")
                If StrComp(arg, "Empty", vbTextCompare)=0 Then arg = Empty
                push pram,GetRef(func)(arg)
            Else
                func = ele
                push pram,GetRef(func)()
            End If
        Next
    End If
    defFolder = Array(PeDataType("FOLDER"), pram)
End Function

'for verify the following properties
'   .hasContainers
'   .hasEntries
'   .hasFilesExcludingArchives
Function expectHasEntries(path,entryType)
    expectHasEntries = False
    
    Dim ret,obj,ele
    Set obj = getFolderItem2(path)
    ret = False
        Select Case entryType
        Case PeEntryType("CONTAINER")
            If fso.FolderExists(path) Then
                ret=(fso.GetFolder(path).SubFolders.Count>0)
            ElseIf obj.IsFolder Then
                For Each ele In obj.GetFolder.Items
                    ret = ele.IsFolder
                    If ret Then Exit For
                Next
            End If
            expectHasEntries=ret
        Case PeEntryType("FILE_EXCLUDING_ARCHIVE")
            If fso.FolderExists(path) Then
                ret=(fso.GetFolder(path).Files.Count>0)
            ElseIf obj.IsFolder Then
                For Each ele In obj.GetFolder.Items
                    ret = Not(ele.IsFolder)
                    If ret Then Exit For
                Next
            End If
            expectHasEntries=ret
        Case Else
        'PeEntryType("ENTRY") or others
            If obj.IsFolder Then expectHasEntries = obj.GetFolder.Items.Count>0
    End Select
End Function

'to verify the following properties
'   .baseName
'   .dateLastModified
'   .extension
'   .isBrowsable
'   .isFileSystem
'   .isFolder
'   .isLink
'   .name
'   .parentFolder
'   .path
'   .size
'   .toString
'   .type
Sub assertFsProperties(actualObj,path,caseNo)
    '準備
    Dim fi2, fo, flg
    Set fi2 = getFolderItem2(path)
    flg = True
    If fso.FolderExists(path) Then
        Set fo = fso.GetFolder(path)
    ElseIf fso.FileExists(path) Then
        Set fo = fso.GetFile(path)
    Else
        flg = False
    End IF

    With actualObj
        
        Dim ele, target, expected, actual
        For Each ele In Array( _
                dicOf(Array(  "target", "baseName"             , "expected", fso.GetBaseName(path)     , "actual", .baseName)) _
                , dicOf(Array("target", "extension"            , "expected", fso.GetExtensionName(path), "actual", .extension)) _
                , dicOf(Array("target", "isBrowsable"          , "expected", fi2.IsBrowsable           , "actual", .isBrowsable)) _
                , dicOf(Array("target", "isFileSystem"         , "expected", fi2.IsFileSystem          , "actual", .isFileSystem)) _
                , dicOf(Array("target", "isFolder"             , "expected", fso.FolderExists(path)    , "actual", .isFolder)) _
                , dicOf(Array("target", "isLink"               , "expected", fi2.IsLink                , "actual", .isLink)) _
                , dicOf(Array("target", "parentFolder TypeName", "expected", "FileSystemProxy"         , "actual", TypeName(.parentFolder))) _
                , dicOf(Array("target", "path"                 , "expected", path                      , "actual", .path)) _
                , dicOf(Array("target", "default"              , "expected", path                      , "actual", actualObj)) _
                , dicOf(Array("target", "toString"             , "expected", "<FileSystemProxy>"&path  , "actual", .toString)) _
                )
            target = ele("target")
            expected = ele("expected")
            If IsObject(ele("actual")) Then Set actual = ele("actual") Else actual = ele("actual")
            AssertEqualWithMessage expected, actual, "caseNo="&caseNo&"("&target&")" &", path="&path
        Next
        
        Dim data
        If flg Then
            data = Array( _
                dicOf(  Array("target","dateLastModified", "expected", fo.DateLastModified          , "actual", .dateLastModified)) _
                , dicOf(Array("target","name"            , "expected", fo.Name                      , "actual", .name)) _
                , dicOf(Array("target","parentFolder"    , "expected", fo.ParentFolder.path         , "actual", .parentFolder.path)) _
                , dicOf(Array("target","size"            , "expected", fo.Size                      , "actual", .size)) _
                , dicOf(Array("target","type"            , "expected", fo.Type                      , "actual", .type)) _
                )
        Else
            data = Array( _
                dicOf(  Array("target","dateLastModified", "expected", fi2.ModifyDate               , "actual", .dateLastModified)) _
                , dicOf(Array("target","name"            , "expected", fso.GetFileName(path)        , "actual", .name)) _
                , dicOf(Array("target","parentFolder"    , "expected", fso.GetParentFolderName(path), "actual", .parentFolder.path)) _
                , dicOf(Array("target","size"            , "expected", fi2.Size                     , "actual", .size)) _
                , dicOf(Array("target","type"            , "expected", fi2.Type                     , "actual", .type)) _
                )
        End If
        For Each ele In data
            target = ele("target")
            expected = ele("expected")
            actual = ele("actual")
            AssertEqualWithMessage expected, actual, "caseNo="&caseNo&"("&target&")" &", path="&path
        Next
    End With

    Set fo = Nothing
    Set fi2 = Nothing
End Sub
'to verify the following properties
'   .allContainers
'   .allContainersIncludingSelf
'   .allEntries
'   .allEntriesIncludingSelf
'   .allFilesExcludingArchives
'   .allFilesExcludingArchivesIncludingSelf
'   .containers
'   .entries
'   .filesExcludingArchives
'   .hasContainers
'   .hasEntries
'   .hasFilesExcludingArchives
Sub assertFsEntries(actualObj,path,cs)
    Dim ele, tp, et, has, items, allItems, allItemsIncludingSelf, text
    With actualObj
        For Each ele In Array( _
                dicOf(  Array("tp", "FilesExcludingArchives", "et", PeEntryType("FILE_EXCLUDING_ARCHIVE"), "has" ,.hasFilesExcludingArchives, "items", .filesExcludingArchives, "allItems", .allFilesExcludingArchives, "allItemsIncludingSelf", .allFilesExcludingArchivesIncludingSelf)) _
                , dicOf(Array("tp", "Containers"            , "et", PeEntryType("CONTAINER")             , "has" ,.hasContainers            , "items", .containers            , "allItems", .allContainers            , "allItemsIncludingSelf", .allContainersIncludingSelf)) _
                , dicOf(Array("tp", "Entries"               , "et", PeEntryType("ENTRY")                 , "has" ,.hasEntries               , "items", .entries               , "allItems", .allEntries               , "allItemsIncludingSelf", .allEntriesIncludingSelf)) _
                )
            tp = ele("tp")
            et = ele("et")
            has = ele("has")
            items = ele("items")
            allItems = ele("allItems")
            allItemsIncludingSelf = ele("allItemsIncludingSelf")
            text = "caseNo="&cs&"("&tp
    
            AssertEqualWithMessage expectHasEntries(path,et), has, text&",has)"
            assertFsEntriesProc items                , path, text&",items)"                , False, False, et
            assertFsEntriesProc allItems             , path, text&",allItems)"             , False, True , et
            assertFsEntriesProc allItemsIncludingSelf, path, text&",allItemsIncludingSelf)", True , True , et
    Next
    End With
End Sub
Sub assertFsEntriesProc(entries,path,caseNo,self,recursive,entryType)
    Dim ele, dic
    Set dic = dictionary
    For Each ele In entries
        If dic.Exists(ele) Then AssertFailWithMessage "caseNo="&caseNo&", '"&ele&"' Entries Duplication!"
        dic.Add ele.path,False
    Next

    assertFsEntriesProcEachEntry path,caseNo,dic,self,recursive,entryType
    
    For Each ele In dic.Keys
        If Not (dic(ele)=True) Then AssertFailWithMessage "caseNo="&caseNo&", '"&ele&"' Not Found !"
    Next

    AssertWithMessage True, "all ok"
    Set dic = Nothing
End Sub
Sub assertFsEntriesProcEachEntry(path,caseNo,dic,self,recursive,entryType)
    If self Then existsEntry path,caseNo,dic,entryType
    Dim ele
    If fso.FolderExists(path) Then
        For Each ele in fso.GetFolder(path).Files
            existsEntry ele.path,caseNo,dic,entryType
        Next
        For Each ele in fso.GetFolder(path).SubFolders
            existsEntry ele.path,caseNo,dic,entryType
            If recursive Then assertFsEntriesProcEachEntry ele.path,caseNo,dic,self,recursive,entryType
        Next
    ElseIf getFolderItem2(path).IsFolder Then
        For Each ele in getFolderItem2(path).GetFolder.Items
            existsEntry ele.path,caseNo,dic,entryType
            If recursive And ele.IsFolder Then
                If ele.GetFolder.Items.Count>0 Then assertFsEntriesProcEachEntry ele.path,caseNo,dic,self,recursive,entryType
            End If
        Next
    End If
End Sub
Sub existsEntry(path,caseNo,dic,entryType)
    Dim flg : flg = False
    Dim sEntryName : sEntryName = "entries"
    Select Case entryType
    Case PeEntryType("FILE_EXCLUDING_ARCHIVE")
        sEntryName = "filesExcludingArchives"
        If Not getFolderItem2(path).IsFolder Then flg = True
    Case PeEntryType("CONTAINER")
        sEntryName = "containers"
        If getFolderItem2(path).IsFolder Then flg = True
    Case Else
        flg = True
    End Select
    
    If flg Then
        If Not dic.Exists(path) Then AssertFailWithMessage "caseNo="&caseNo&","&sEntryName&" Not Exists " & path
        dic(path)=True
    End If
End Sub

Function fso
    Set fso = CreateObject("Scripting.FileSystemObject")
End Function
Function shellApp
    Set shellApp = CreateObject("Shell.Application")
End Function
Function shell
    Set shell = CreateObject("WScript.Shell")
End Function
Function dictionary
    Set dictionary = CreateObject("Scripting.Dictionary")
End Function
Function dicOf( _
    byVal avParams _
    )
    Dim oDict : Set oDict = dictionary()
    
    Dim vItem, vKey, boIsKey
    boIsKey = True
    For Each vItem In avParams
        If boIsKey Then
            vKey = vItem
            oDict(vKey)= Empty
        Else
            If IsObject(vItem) Then
                Set oDict(vKey) = vItem
            Else
                oDict(vKey) = vItem
            End If
        End If
        boIsKey = Not boIsKey
    Next
    
    Set dicOf = oDict
    Set oDict = Nothing
End Function
Function reOf(pattern, opt)
    Dim oRe, sOpts
    
    Set oRe = New RegExp
    oRe.Pattern = pattern
    
    sOpts = LCase(opt)
    If InStr(sOpts, "i") > 0 Then oRe.IgnoreCase = True
    If InStr(sOpts, "g") > 0 Then oRe.Global = True
    If InStr(sOpts, "m") > 0 Then oRe.Multiline = True
    
    Set reOf = oRe
    Set oRe = Nothing
End Function
Function getFolderItem2(path)
    Set getFolderItem2 = shellApp.Namespace(fso.GetParentFolderName(path)).Items().Item(fso.GetFileName(path))
End Function
Function getTempFilePath(parent,extention)
    getTempFilePath = fso.BuildPath(parent, getTempFile(extention))
End Function
Function getTempFolderPath(parent)
    getTempFolderPath = fso.BuildPath(parent, getTempName())
End Function
Function getTempFile(extention)
    getTempFile = getTempName()&"."&extention
End Function
Function getTempName()
    getTempName = fso.GetBaseName(fso.GetTempName())
End Function
Sub zip(paths,zpath)
    Dim i,path : path=""
    For Each i In paths
        If path="" Then
            path=wrapInQuotes(i)
        Else
            path=path&","& wrapInQuotes(i)
        End If
    Next

    Dim cmd : cmd = _
        "powershell -NoProfile -ExecutionPolicy Unrestricted Compress-Archive" _
        & " -Path " & path _
        & " -DestinationPath " & zpath
    Call shell.Run(cmd, 0, True)
End Sub
Function wrapInQuotes(target)
    wrapInQuotes = Chr(34) & Replace(target, Chr(34), Chr(34)&Chr(34)) & Chr(34)
End Function
Function createShortCutCommon(basePath,extension,target)
    Dim path : path = getTempFilePath(basePath,extension)
    With shell.CreateShortcut(path)
        .TargetPath = target
        .Save
    End With
    createShortCutCommon = path
End Function

Sub push(arr,ele)
    On Error Resume Next
    Redim Preserve arr(Ubound(arr)+1)
    If Err.Number<>0 Then Redim arr(0)
    On Error Goto 0
    If IsObject(ele) Then Set arr(Ubound(arr)) = ele Else arr(Ubound(arr)) = ele
End Sub
Sub pushA(arr,add)
    On Error Resume Next
    Dim ubAdd : ubAdd = Ubound(add)
    If Err.Number=0 Then
        Dim ub : ub = Ubound(arr)
        If Err.Number=0 Then
            Redim Preserve arr(ub+ubAdd+1)
        Else
            ub = -1
            Redim arr(ubAdd)
        End If

        Dim i
        For i=0 To ubAdd
            If IsObject(add(i)) Then Set arr(ub+1+i) = add(i) Else arr(ub+1+i) = add(i)
        Next
    Elseif Not IsArray(add) Then
        push arr, add
    End If
    On Error Goto 0
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
