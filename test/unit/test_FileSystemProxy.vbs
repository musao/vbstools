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

Const MY_NAME = "test_FileSystemProxy.vbs"
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
        , "ARCHIVE", 4 _
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
Sub Test_FileSystemProxy
    Dim a : Set a = new FileSystemProxy
    AssertEqualWithMessage 1, VarType(a), "VarType"
    AssertEqualWithMessage "FileSystemProxy", TypeName(a), "TypeName"
End Sub

'###################################################################################################
'FileSystemProxy.<properties>_initial
'   .actualPath
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
            dicOf( Array(  "target","actualPath"                            ,"expected",Null               , "actual", ao.actualPath) ) _
            , dicOf( Array("target","allContainers"                         ,"expected",Null               , "actual", ao.allContainers) ) _
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
Sub Test_FileProxy_properties_fs
    Dim cases
    cases=Array( _
    dicOf(  Array("Case", "1", "Definition", defShortCutFile()             )) _
    , dicOf(Array("Case", "2", "Definition", defUrlShortCutFile()          )) _
    , dicOf(Array("Case", "3", "Definition", defTextFile()                 )) _
    , dicOf(Array("Case", "4", "Definition", defFolder(Empty)              )) _
    , dicOf(Array("Case", "5", "Definition", defFolder("defShortCutFile")  )) _
    , dicOf(Array("Case", "6", "Definition", defArchive("defShortCutFile") )) _
    )
    Dim ele,path,caze,ao
    For Each ele In createData(cases)
        path = ele("Path")
        caze = ele("Case")
        Set ao = (new FileSystemProxy).of(path)

'       .actualPath
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
        assertFsProperties ao,path,caze
    Next
End Sub
Sub Test_FileProxy_properties_list
    Dim cases
'    cases=Array( _
'    dicOf(  Array("Case", "1-2-3-3", "Definition", defFolder( "defArchive(defArchive(defUrlShortCutFile))") )) _
'    , dicOf(Array("Case", "1-3-2-3", "Definition", defArchive("defFolder( defArchive(defTextFile))") )) _
'    , dicOf(Array("Case", "1-3-3-1", "Definition", defArchive("defArchive(defShortCutFile)") )) _
'    , dicOf(Array("Case", "1-3-3-2", "Definition", defArchive("defArchive(defFolder( defUrlShortCutFile))") )) _
'    , dicOf(Array("Case", "1-3-3-3", "Definition", defArchive("defArchive(defArchive(defTextFile))") )) _
'    , dicOf(Array("Case", "2-2"    , "Definition", defArchive("defTextFile,defFolder(defTextFile),defArchive(defTextFile)") )) _
'    )
    cases=Array( _
    dicOf(  Array("Case", "1-1"    , "Definition", defShortCutFile() )) _
    , dicOf(Array("Case", "1-2-3-3", "Definition", defFolder( "defArchive(defArchive(defUrlShortCutFile))") )) _
    , dicOf(Array("Case", "2-2"    , "Definition", defArchive("defTextFile,defFolder(defTextFile),defArchive(defTextFile)") )) _
    )
'All Cases
'    cases=Array( _
'    dicOf(  Array("Case", "1-1"    , "Definition", defShortCutFile() )) _
'    , dicOf(Array("Case", "1-2-1"  , "Definition", defFolder( "defUrlShortCutFile") )) _
'    , dicOf(Array("Case", "1-2-2-1", "Definition", defFolder( "defFolder( defTextFile)") )) _
'    , dicOf(Array("Case", "1-2-2-2", "Definition", defFolder( "defFolder( defFolder( defShortCutFile))") )) _
'    , dicOf(Array("Case", "1-2-2-3", "Definition", defFolder( "defFolder( defArchive(defUrlShortCutFile))") )) _
'    , dicOf(Array("Case", "1-2-3-1", "Definition", defFolder( "defArchive(defTextFile)") )) _
'    , dicOf(Array("Case", "1-2-3-2", "Definition", defFolder( "defArchive(defFolder( defShortCutFile))") )) _
'    , dicOf(Array("Case", "1-2-3-3", "Definition", defFolder( "defArchive(defArchive(defUrlShortCutFile))") )) _
'    , dicOf(Array("Case", "1-2-4"  , "Definition", defFolder( Empty) )) _
'    , dicOf(Array("Case", "1-3-1"  , "Definition", defArchive("defTextFile") )) _
'    , dicOf(Array("Case", "1-3-2-1", "Definition", defArchive("defFolder( defShortCutFile)") )) _
'    , dicOf(Array("Case", "1-3-2-2", "Definition", defArchive("defFolder( defFolder( defUrlShortCutFile))") )) _
'    , dicOf(Array("Case", "1-3-2-3", "Definition", defArchive("defFolder( defArchive(defTextFile))") )) _
'    , dicOf(Array("Case", "1-3-3-1", "Definition", defArchive("defArchive(defShortCutFile)") )) _
'    , dicOf(Array("Case", "1-3-3-2", "Definition", defArchive("defArchive(defFolder( defUrlShortCutFile))") )) _
'    , dicOf(Array("Case", "1-3-3-3", "Definition", defArchive("defArchive(defArchive(defTextFile))") )) _
'    , dicOf(Array("Case", "2-1"    , "Definition", defFolder( "defTextFile,defFolder(defTextFile),defArchive(defTextFile)") )) _
'    , dicOf(Array("Case", "2-2"    , "Definition", defArchive("defTextFile,defFolder(defTextFile),defArchive(defTextFile)") )) _
'    )
'inputbox "","",cf_toString(createData(cases))
Dim ele,path,caze,ao
    For Each ele In createData(cases)
        path = ele("Path")
        caze = ele("Case")
        Set ao = (new FileSystemProxy).of(path)

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
        assertFsEntries ao,path,caze
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
    On Error GoTo 0
End Sub
Sub Test_FileProxy_setParent_Err_NotFileSystemProxy
    On Error Resume Next
    Dim p,a
    p = WScript.ScriptFullName
    Set a = (new FileSystemProxy).of(p)
    Call a.setParent(dictionary)

    AssertEqualWithMessage "FileSystemProxy+setParent()", Err.Source, "Err.Source"
    AssertEqualWithMessage "This is not FileSystemProxy.", Err.Description, "Err.Description"
    On Error GoTo 0
End Sub
Sub Test_FileProxy_setParent_Err_NotPrentFolder
    On Error Resume Next
    Dim p,a
    p = WScript.ScriptFullName
    Set a = (new FileSystemProxy).of(p)
    Call a.setParent(a)

    AssertEqualWithMessage "FileSystemProxy+setParent()", Err.Source, "Err.Source"
    AssertEqualWithMessage "This is not a parent folder.", Err.Description, "Err.Description"
    On Error GoTo 0
End Sub

'###################################################################################################
'FileSystemProxy.setVirtualPath()
Sub Test_FileProxy_setVirtualPath
    Dim p,v,a,e
    v="C:\hoge\fuga\piyo.java"
    p = WScript.ScriptFullName
    Set a = (new FileSystemProxy).of(p)

    AssertEqualWithMessage p, a.actualPath, "before set virtualPath actualPath"
    AssertEqualWithMessage p, a.path, "before set virtualPath path"
    AssertEqualWithMessage fso.GetBaseName(p), a.baseName, "before set virtualPath baseName"
    AssertEqualWithMessage fso.GetExtensionName(p), a.extension, "before set virtualPath extension"
    AssertEqualWithMessage fso.GetFileName(p), a.name, "before set virtualPath name"

    a.setVirtualPath(v)

    AssertEqualWithMessage p, a.actualPath, "ater set virtualPath actualPath"
    AssertEqualWithMessage v, a.path, "after set virtualPath path"
    AssertEqualWithMessage fso.GetBaseName(v), a.baseName, "after set virtualPath baseName"
    AssertEqualWithMessage fso.GetExtensionName(v), a.extension, "after set virtualPath extension"
    AssertEqualWithMessage fso.GetFileName(v), a.name, "after set virtualPath name"

    a.setVirtualPath("")

    AssertEqualWithMessage p, a.actualPath, "after clear virtualPath actualPath"
    AssertEqualWithMessage p, a.path, "after clear virtualPath path"
    AssertEqualWithMessage fso.GetBaseName(p), a.baseName, "after clear virtualPath baseName"
    AssertEqualWithMessage fso.GetExtensionName(p), a.extension, "after clear virtualPath extension"
    AssertEqualWithMessage fso.GetFileName(p), a.name, "after clear virtualPath name"
End Sub
Sub Test_FileProxy_setVirtualPath_Err_Initial
    On Error Resume Next
    Dim p,v,a
    v="C:\hoge\fuga\piyo.java"
    p = WScript.ScriptFullName
    Call (new FileSystemProxy).setVirtualPath(v)

    AssertEqualWithMessage "FileSystemProxy+setVirtualPath()", Err.Source, "Err.Source"
    AssertEqualWithMessage "Please set the value before setting the virtual path.", Err.Description, "Err.Description"
    On Error GoTo 0
End Sub

'###################################################################################################
'common
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
    defFolder = defContainer(d,PeDataType("FOLDER"))
End Function
Function defArchive(d)
    defArchive = defContainer(d,PeDataType("ARCHIVE"))
End Function
Function defContainer(d,tp)
    Dim pram : pram = Array()
    If Not IsEmpty(d) Then
        Dim ele,re,func,arg
        Set re = reOf("([a-zA-Z0-9_]+)\((.+)\)", "igm")
        For Each ele In splitOuterArgs(d)
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
    defContainer = Array(tp, pram)
End Function

Function createData(cases)
    Dim ele,data
    For Each ele In cases
        pusha data, dicOf(Array("Case", ele("Case"), "Path", createDataRecursive(createCaseFolder(ele("Case")), ele("Definition"))))
    Next
    createData = data
End Function

Function createCaseFolder(caseName)
    Dim path : path = fso.BuildPath(PsPathTempFolder, "Case"&caseName&"_"&getTempName())
    fso.CreateFolder path
    createCaseFolder=path
End Function
Function createDataRecursive(targetPath,def)
    If Ubound(def)<0 Then
        Exit Function
    End If
    
    Dim tp, path
    tp = def(0)
    Select Case tp
    Case PeDataType("FOLDER"), PeDataType("ARCHIVE")
        path = createEmptyFolderAt(targetPath)
        If Ubound(def)>0 Then
            Dim ele
            For Each ele In def(1)
                createDataRecursive path, ele
            Next
        End If

        If tp=PeDataType("ARCHIVE") Then
            Dim zipPath : zipPath=getTempFilePath(fso.GetParentFolderName(path),"zip")
            Dim paths
            For Each ele in getFolderItem2(path).GetFolder.Items
                push paths, ele.path
            Next
            zip paths,zipPath
            fso.DeleteFolder path
            path = zipPath
        End If
    Case Else
        path = createSomeFileAt(tp, targetPath)
    End Select
    createDataRecursive = path
End Function
Function createEmptyFolderAt(basePath)
    Dim path : path = getTempFolderPath(basePath)
    fso.CreateFolder path
    createEmptyFolderAt=path
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

Function createShortCutFile
    createShortCutFile = createShortCutFileAt(PsPathTempFolder)
End Function
Function createUrlShortCutFile
    createUrlShortCutFile = createUrlShortCutFileAt(PsPathTempFolder)
End Function
Function createTextFile
    createTextFile = createTextFileAt(PsPathTempFolder)
End Function

'to verify the following properties
'   .actualPath
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
Sub assertFsProperties(actualObj,path,caze)
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
                dicOf(Array(  "target", "actualPath"           , "expected", path                      , "actual", .actualPath)) _
                , dicOf(Array("target", "baseName"             , "expected", fso.GetBaseName(path)     , "actual", .baseName)) _
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
            AssertEqualWithMessage expected, actual, "caseNo="&caze&"("&target&")" &", path="&path
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
            AssertEqualWithMessage expected, actual, "caseNo="&caze&"("&target&")" &", path="&path
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
Sub assertFsEntries(actualObj,path,caze)
    Dim ele, methodType, entryType, has, items, allItems, allItemsIncludingSelf, caseName
    With actualObj
        For Each ele In Array( _
                dicOf(  Array("mt", "FilesExcludingArchives", "et", PeEntryType("FILE_EXCLUDING_ARCHIVE"), "has" ,.hasFilesExcludingArchives, "items", .filesExcludingArchives, "allItems", .allFilesExcludingArchives, "allItemsIncludingSelf", .allFilesExcludingArchivesIncludingSelf)) _
                , dicOf(Array("mt", "Containers"            , "et", PeEntryType("CONTAINER")             , "has" ,.hasContainers            , "items", .containers            , "allItems", .allContainers            , "allItemsIncludingSelf", .allContainersIncludingSelf)) _
                , dicOf(Array("mt", "Entries"               , "et", PeEntryType("ENTRY")                 , "has" ,.hasEntries               , "items", .entries               , "allItems", .allEntries               , "allItemsIncludingSelf", .allEntriesIncludingSelf)) _
                )
            methodType = ele("mt")
            entryType = ele("et")
            has = ele("has")
            items = ele("items")
            allItems = ele("allItems")
            allItemsIncludingSelf = ele("allItemsIncludingSelf")
            caseName = caze&"("&methodType
            AssertEqualWithMessage expectHasEntries(entryType, path), has, caseName&",has)"
            assertFsEntriesProc entryType, path, items                , caseName&",items)"                , False, False
            assertFsEntriesProc entryType, path, allItems             , caseName&",allItems)"             , False, True
            assertFsEntriesProc entryType, path, allItemsIncludingSelf, caseName&",allItemsIncludingSelf)", True , True
    Next
    End With
End Sub

'for verify the following properties
'   .hasContainers
'   .hasEntries
'   .hasFilesExcludingArchives
Function expectHasEntries(entryType, path)
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
Sub assertFsEntriesProc(entryType, path, entries, caseName, includingSelf, recursive)
    Dim ele, dic
    Set dic = dictionary
    For Each ele In entries
        If dic.Exists(ele) Then AssertFailWithMessage "caseName="&caseName&", '"&ele&"' Entries Duplication!"
        dic.Add ele.path,False
    Next
    
    assertFsEntriesProcEachEntry entryType, path, dic, caseName, includingSelf, recursive
    
    For Each ele In dic.Keys
        If Not (dic(ele)=True) Then AssertFailWithMessage "caseName="&caseName&", '"&ele&"' Not Found !"
    Next

    AssertWithMessage True, "all ok"
    Set dic = Nothing
End Sub
Sub assertFsEntriesProcEachEntry(entryType, path, dic, caseName, includingSelf, recursive)
    If includingSelf Then existsEntry entryType, path, dic, caseName
    Dim ele
    If fso.FolderExists(path) Then
        For Each ele in fso.GetFolder(path).Files
            existsEntry entryType, ele.path, dic, caseName
            If recursive And isShellFolder(ele.path) Then
                assertFsEntriesProcEachEntryArchive entryType, ele.path, dic, caseName, includingSelf, recursive
            End If
        Next
        For Each ele in fso.GetFolder(path).SubFolders
            existsEntry entryType, ele.path, dic, caseName
            If recursive Then assertFsEntriesProcEachEntry entryType, ele.path, dic, caseName, includingSelf, recursive
        Next
    ElseIf getFolderItem2(path).IsFolder Then
        assertFsEntriesProcEachEntryArchive entryType, path, dic, caseName, includingSelf, recursive
    End If
End Sub
Sub assertFsEntriesProcEachEntryArchive(entryType, path, dic, caseName, includingSelf, recursive)
    Dim ele
    For Each ele in getFolderItem2(path).GetFolder.Items
        existsEntry entryType, ele.path, dic, caseName
        If recursive And ele.IsFolder Then
            If ele.GetFolder.Items.Count>0 Then assertFsEntriesProcEachEntry entryType, ele.path, dic, caseName, includingSelf, recursive
        End If
    Next
End Sub
Sub existsEntry(entryType, path, dic, caseName)
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
        If Not dic.Exists(path) Then AssertFailWithMessage "caseName="&caseName&","&sEntryName&" Not Exists " & path
        dic(path)=True
    End If
End Sub
Function isShellFolder(path)
    isShellFolder = False
    Dim obj
    On Error Resume Next
    Set obj = shellApp.Namespace(path)
    If Err.Number=0 Then isShellFolder = True
    On Error Goto 0
    Set obj = Nothing
End Function

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
    With fso
        Set getFolderItem2 = shellApp.Namespace(.GetParentFolderName(path)).ParseName(.GetFileName(path))
    End With
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
        & " -DestinationPath " & zpath _
        & " -CompressionLevel NoCompression "
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
Function splitOuterArgs(argString)
    Dim currentArg, depth, i, char, ret, flg
    currentArg = ""
    depth = 0
    For i = 1 To Len(argString)
        char = Mid(argString, i, 1)

        flg = False
        If char = "(" Then
            depth = depth + 1
        ElseIf char = ")" Then
            depth = depth - 1
        ElseIf char = "," And depth = 0 Then
            ' 一番外側のカンマを発見：ここまでの文字列を保存
            push ret, Trim(currentArg)
            flg = True
        End If

        currentArg = currentArg & char
        If flg Then currentArg = ""
    Next
    
    ' 最後の引数を追加
    If Len(currentArg)>0 Then push ret, Trim(currentArg)
    
    ' 配列として返す
    splitOuterArgs = ret
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
