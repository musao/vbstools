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

Const Cl_ENTRY = 0
Const Cl_FILE = 2
Const Cl_FILE_EXCLUDING_ARCHIVE = 3
Const Cl_FOLDER = 4
Const Cl_CONTAINER = 5

'###################################################################################################
'SetUp()/TearDown()
Sub SetUp()
    '実行スクリプト直下に当ファイル名で一時フォルダ作成
    PsPathTempFolder = new_Fso().BuildPath(new_Fso().GetParentFolderName(WScript.ScriptFullName), MY_NAME)
    If Not (new_Fso().FolderExists(PsPathTempFolder)) Then new_Fso().CreateFolder(PsPathTempFolder)
End Sub
Sub TearDown()
    '当テストで作成した一時フォルダを削除する
    new_Fso().DeleteFolder PsPathTempFolder
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
    dim tg,a,ao,e
    set ao = (new FileSystemProxy)

    tg = "allContainers"
    e = Null
    a = ao.allContainers
    AssertEqualWithMessage e, a, tg

    tg = "allContainersIncludingSelf"
    e = Null
    a = ao.allContainersIncludingSelf
    AssertEqualWithMessage e, a, tg

    tg = "allEntries"
    e = Null
    a = ao.allEntries
    AssertEqualWithMessage e, a, tg

    tg = "allEntriesIncludingSelf"
    e = Null
    a = ao.allEntriesIncludingSelf
    AssertEqualWithMessage e, a, tg

    tg = "allFilesExcludingArchives"
    e = Null
    a = ao.allFilesExcludingArchives
    AssertEqualWithMessage e, a, tg

    tg = "allFilesExcludingArchivesIncludingSelf"
    e = Null
    a = ao.allFilesExcludingArchivesIncludingSelf
    AssertEqualWithMessage e, a, tg

    tg = "baseName"
    e = Null
    a = ao.baseName
    AssertEqualWithMessage e, a, tg

    tg = "containers"
    e = Null
    a = ao.containers
    AssertEqualWithMessage e, a, tg

    tg = "dateLastModified"
    e = Null
    a = ao.dateLastModified
    AssertEqualWithMessage e, a, tg

    tg = "entries"
    e = Null
    a = ao.entries
    AssertEqualWithMessage e, a, tg

    tg = "extension"
    e = Null
    a = ao.extension
    AssertEqualWithMessage e, a, tg

    tg = "filesExcludingArchives"
    e = Null
    a = ao.filesExcludingArchives
    AssertEqualWithMessage e, a, tg

    tg = "hasContainers"
    e = Null
    a = ao.hasContainers
    AssertEqualWithMessage e, a, tg

    tg = "hasEntries"
    e = Null
    a = ao.hasEntries
    AssertEqualWithMessage e, a, tg

    tg = "hasFilesExcludingArchives"
    e = Null
    a = ao.hasFilesExcludingArchives
    AssertEqualWithMessage e, a, tg

    tg = "isBrowsable"
    e = Null
    a = ao.isBrowsable
    AssertEqualWithMessage e, a, tg

    tg = "isFileSystem"
    e = Null
    a = ao.isFileSystem
    AssertEqualWithMessage e, a, tg

    tg = "isFolder"
    e = Null
    a = ao.isFolder
    AssertEqualWithMessage e, a, tg

    tg = "isLink"
    e = Null
    a = ao.isLink
    AssertEqualWithMessage e, a, tg

    tg = "name"
    e = Null
    a = ao.name
    AssertEqualWithMessage e, a, tg

    tg = "parentFolder"
    e = Null
    a = ao.parentFolder
    AssertEqualWithMessage e, a, tg

    tg = "path"
    e = Null
    a = ao.path
    AssertEqualWithMessage e, a, tg

    tg = "size"
    e = Null
    a = ao.size
    AssertEqualWithMessage e, a, tg

    tg = "toString"
    e = "<FileSystemProxy>"
    a = ao.toString
    AssertEqualWithMessage e, a, tg

    tg = "type"
    e = Null
    a = ao.type
    AssertEqualWithMessage e, a, tg
End Sub
Sub Test_FileProxy_properties
    Dim data : data = createData()
    Dim cs,path,ao,obj,hasSomething
    For cs=0 To Ubound(data)
        path = data(cs)
        Set ao = new FileSystemProxy : ao.of(path)

        assertFsProperties ao,path,cs

        assertFsEntries ao,path,cs
        
'        hasSomething=expectHasEntries(path,Cl_FILE_EXCLUDING_ARCHIVE)
'        AssertEqualWithMessage hasSomething, ao.hasFilesExcludingArchives, "case="&cs&",hasFilesExcludingArchives"
'        If hasSomething Then
'            assertFsEntriesProc ao.filesExcludingArchives,path,cs&".filesExcludingArchives",False,False,Cl_FILE_EXCLUDING_ARCHIVE
'        Else
'            AssertEqualWithMessage cf_toString(Array()), cf_toString(ao.filesExcludingArchives), "case="&cs&",.filesExcludingArchives"
'        End If
'        assertFsEntriesProc ao.allFilesExcludingArchives,path,cs&".allFilesExcludingArchives",False,True,Cl_FILE_EXCLUDING_ARCHIVE
'        assertFsEntriesProc ao.allFilesExcludingArchivesIncludingSelf,path,cs&".allFilesExcludingArchivesIncludingSelf",True,True,Cl_FILE_EXCLUDING_ARCHIVE
'
'        hasSomething=expectHasEntries(path,Cl_CONTAINER)
'        AssertEqualWithMessage hasSomething, ao.hasContainers, "case="&cs&",hasContainers"
'        If hasSomething Then
'            assertFsEntriesProc ao.containers,path,cs&".containers",False,False,Cl_CONTAINER
'        Else
'            AssertEqualWithMessage cf_toString(Array()), cf_toString(ao.containers), "case="&cs&",.containers"
'        End If
'        assertFsEntriesProc ao.allContainers,path,cs&".allContainers",False,True,Cl_CONTAINER
'        assertFsEntriesProc ao.allContainersIncludingSelf,path,cs&".allContainersIncludingSelf",True,True,Cl_CONTAINER
'
'        hasSomething=expectHasEntries(path,Cl_ENTRY)
'        AssertEqualWithMessage hasSomething, ao.hasEntries, "case="&cs&",hasEntries"
'        If hasSomething Then
'            assertFsEntriesProc ao.entries,path,cs&".entries",False,False,Empty
'        Else
'            AssertEqualWithMessage cf_toString(Array()), cf_toString(ao.entries), "case="&cs&",.entries"
'       End If
'       assertFsEntriesProc ao.allEntries,path,cs&".allEntries",False,True,Empty
'       assertFsEntriesProc ao.allEntriesIncludingSelf,path,cs&".allEntriesIncludingSelf",True,True,Empty

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
    createData = caseToData(cases)
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
Function caseToData(cases)
    Dim ret,cs
    For Each cs In cases
        pushA ret,cs
    Next
    caseToData = ret
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
Function createEmptyFolderAt(basePath)
    Dim path : path = getTempFolderPath(basePath)
    fso.CreateFolder path
    createEmptyFolderAt=path
End Function
Function createFolderAt(basePath,content)
    Dim contents
    push contents,basePath
    push contents,content
    createFolderAt=createFolderAbout(contents)
End Function
Function createFolderAbout(contents)
    Dim setting : Set setting = dictionary()
    With setting
        .Add "ShortCutFile",Array(GetRef("createShortCutFileAt"),False)
        .Add "UrlShortCutFile",Array(GetRef("createUrlShortCutFileAt"),False)
        .Add "TextFile",Array(GetRef("createTextFileAt"),False)
        .Add "Folder",Array(GetRef("createFolderAbout"),True)
    End With
    Dim path : path = createEmptyFolderAt(contents(0))

    Dim ret : push ret,path
    Dim ele,ptn,pram
    For Each ele In contents(1)
        ptn=setting(ele(0))
        If ptn(1) Then
            pram=Array()
            push pram,path
            push pram,ele(1)
            pushA ret,ptn(0)(pram)
        Else
            pushA ret,ptn(0)(path)
        End If
    Next
    createFolderAbout=ret
End Function
Function defShortCutFile
    defShortCutFile = Array("ShortCutFile")
End Function
Function defUrlShortCutFile
    defUrlShortCutFile = Array("UrlShortCutFile")
End Function
Function defTextFile
    defTextFile = Array("TextFile")
End Function
Function defFolder(d)
    Dim pram : pram=Array()
    If Not IsEmpty(d) Then
        Dim i
        For Each i In Split(d,",")
            push pram,GetRef(i)()
        Next
    End If
    defFolder = Array("Folder", pram)
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
        Case Cl_CONTAINER
            If fso.FolderExists(path) Then
                ret=(fso.GetFolder(path).SubFolders.Count>0)
            ElseIf obj.IsFolder Then
                For Each ele In obj.GetFolder.Items
                    ret = ele.IsFolder
                    If ret Then Exit For
                Next
            End If
            expectHasEntries=ret
        Case Cl_FILE_EXCLUDING_ARCHIVE
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
        'Cl_ENTRY or others
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
Sub assertFsProperties(target,path,caseNo)
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

    Dim expect
    With target
        '.baseName
        expect = fso.GetBaseName(path)
        AssertEqualWithMessage expect, .baseName              , "caseNo="&caseNo&"(baseName)"             &", path="&path
        
        '.dateLastModified
        If flg Then expect = fo.DateLastModified Else expect = fi2.ModifyDate
        AssertEqualWithMessage expect, .dateLastModified      , "caseNo="&caseNo&"(dateLastModified)"     &", path="&path
        
        '.extension
        expect = fso.GetExtensionName(path)
        AssertEqualWithMessage expect, .extension             , "caseNo="&caseNo&"(extension)"            &", path="&path
        
        '.isBrowsable
        expect = fi2.IsBrowsable
        AssertEqualWithMessage expect, .isBrowsable           , "caseNo="&caseNo&"(isBrowsable)"          &", path="&path
        
        '.isFileSystem
        expect = fi2.IsFileSystem
        AssertEqualWithMessage expect, .isFileSystem          , "caseNo="&caseNo&"(isFileSystem)"         &", path="&path
        
        '.isFolder
        expect = fso.FolderExists(path)
        AssertEqualWithMessage expect, .isFolder              , "caseNo="&caseNo&"(isFolder)"             &", path="&path
        
        '.isLink
        expect = fi2.IsLink
        AssertEqualWithMessage expect, .isLink                , "caseNo="&caseNo&"(isLink)"               &", path="&path
        
        '.name
        If flg Then expect = fo.Name Else expect = fso.GetFileName(path)
        AssertEqualWithMessage expect, .name                  , "caseNo="&caseNo&"(name)"                 &", path="&path
        
        '.parentFolder
        expect = "FileSystemProxy"
        AssertEqualWithMessage expect, TypeName(.parentFolder), "caseNo="&caseNo&"(parentFolder TypeName)"&", path="&path
        If flg Then expect = fo.ParentFolder.Path Else expect = fso.GetParentFolderName(path)
        AssertEqualWithMessage expect, .parentFolder.path     , "caseNo="&caseNo&"(parentFolder path)"    &", path="&path
        
        '.path,default
        expect = path
        AssertEqualWithMessage expect, .path                  , "caseNo="&caseNo&"(path)"                 &", path="&path
        AssertEqualWithMessage expect, target                 , "caseNo="&caseNo&"(path default)"         &", path="&path
        
        '.size
        If flg Then expect = fo.Size Else expect = fi2.Size
        AssertEqualWithMessage expect, .size                  , "caseNo="&caseNo&"(size)"                 &", path="&path
        
        '.toString
        expect = "<FileSystemProxy>"&path
        AssertEqualWithMessage expect, .toString              , "caseNo="&caseNo&"(toString)"             &", path="&path
        
        '.type
        If flg Then expect = fo.Type Else expect = fi2.Type
        AssertEqualWithMessage expect, .type                  , "caseNo="&caseNo&"(type)"                 &", path="&path
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
Sub assertFsEntries(target,path,cs)
'    Dim hasSomething, a, text
'
'    For Each ele In Array(Cl_FILE_EXCLUDING_ARCHIVE, Cl_CONTAINER, Cl_ENTRY)
'        hasSomething=expectHasEntries(path,ele)
'
'        Select Case ele
'        Case Cl_FILE_EXCLUDING_ARCHIVE
'            a = target.hasFilesExcludingArchives
'            text = ",hasFilesExcludingArchives"
'        Case Cl_CONTAINER
'            a = target.hasContainers
'            text = ",hasContainers"
'        Case Cl_ENTRY
'            a = target.hasEntries
'            text = ",hasEntries"
'        End Select
'        AssertEqualWithMessage hasSomething, a, "case="&cs&text
'    Next

    Dim hasSomething
    hasSomething=expectHasEntries(path,Cl_FILE_EXCLUDING_ARCHIVE)
    AssertEqualWithMessage hasSomething, target.hasFilesExcludingArchives, "case="&cs&",hasFilesExcludingArchives"
    If hasSomething Then
        assertFsEntriesProc target.filesExcludingArchives,path,cs&".filesExcludingArchives",False,False,Cl_FILE_EXCLUDING_ARCHIVE
    Else
        AssertEqualWithMessage cf_toString(Array()), cf_toString(target.filesExcludingArchives), "case="&cs&",.filesExcludingArchives"
    End If
    assertFsEntriesProc target.allFilesExcludingArchives,path,cs&".allFilesExcludingArchives",False,True,Cl_FILE_EXCLUDING_ARCHIVE
    assertFsEntriesProc target.allFilesExcludingArchivesIncludingSelf,path,cs&".allFilesExcludingArchivesIncludingSelf",True,True,Cl_FILE_EXCLUDING_ARCHIVE

    hasSomething=expectHasEntries(path,Cl_CONTAINER)
    AssertEqualWithMessage hasSomething, target.hasContainers, "case="&cs&",hasContainers"
    If hasSomething Then
        assertFsEntriesProc target.containers,path,cs&".containers",False,False,Cl_CONTAINER
    Else
        AssertEqualWithMessage cf_toString(Array()), cf_toString(target.containers), "case="&cs&",.containers"
    End If
    assertFsEntriesProc target.allContainers,path,cs&".allContainers",False,True,Cl_CONTAINER
    assertFsEntriesProc target.allContainersIncludingSelf,path,cs&".allContainersIncludingSelf",True,True,Cl_CONTAINER
    
    hasSomething=expectHasEntries(path,Cl_ENTRY)
    AssertEqualWithMessage hasSomething, target.hasEntries, "case="&cs&",hasEntries"
    If hasSomething Then
        assertFsEntriesProc target.entries,path,cs&".entries",False,False,Empty
    Else
        AssertEqualWithMessage cf_toString(Array()), cf_toString(target.entries), "case="&cs&",.entries"
   End If
   assertFsEntriesProc target.allEntries,path,cs&".allEntries",False,True,Empty
   assertFsEntriesProc target.allEntriesIncludingSelf,path,cs&".allEntriesIncludingSelf",True,True,Empty
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
        If Not dic(ele) Then AssertFailWithMessage "caseNo="&caseNo&", '"&ele&"' Not Found !"
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
    Case Cl_FILE_EXCLUDING_ARCHIVE
        sEntryName = "filesExcludingArchives"
        If Not getFolderItem2(path).IsFolder Then flg = True
    Case Cl_CONTAINER
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
