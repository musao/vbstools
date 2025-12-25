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

Const Cl_FILE = 1
Const Cl_FOLDER = 2

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
'   .allFiles
'   .allFolders
'   .allItems
'   .baseName
'   .dateLastModified
'   .extension
'   .files
'   .folders
'   .hasFile
'   .hasFolder
'   .hasItem
'   .isBrowsable
'   .isFileSystem
'   .isFolder
'   .isLink
'   .items
'   .name
'   .parentFolder
'   .path
'   .selfAndAllFiles
'   .selfAndAllFolders
'   .selfAndAllItems
'   .size
'   .toString
'   .type
Sub Test_FileProxy_proeerties_initial
    dim tg,a,ao,e
    set ao = (new FileSystemProxy)

    tg = "allFiles"
    e = Null
    a = ao.allFiles
    AssertEqualWithMessage e, a, tg

    tg = "allFolders"
    e = Null
    a = ao.allFolders
    AssertEqualWithMessage e, a, tg

    tg = "allItems"
    e = Null
    a = ao.allItems
    AssertEqualWithMessage e, a, tg

    tg = "baseName"
    e = Null
    a = ao.baseName
    AssertEqualWithMessage e, a, tg

    tg = "dateLastModified"
    e = Null
    a = ao.dateLastModified
    AssertEqualWithMessage e, a, tg

    tg = "extension"
    e = Null
    a = ao.extension
    AssertEqualWithMessage e, a, tg

    tg = "files"
    e = Null
    a = ao.files
    AssertEqualWithMessage e, a, tg

    tg = "folders"
    e = Null
    a = ao.folders
    AssertEqualWithMessage e, a, tg

    tg = "hasFile"
    e = Null
    a = ao.hasFile
    AssertEqualWithMessage e, a, tg

    tg = "hasFolder"
    e = Null
    a = ao.hasFolder
    AssertEqualWithMessage e, a, tg

    tg = "hasItem"
    e = Null
    a = ao.hasItem
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

    tg = "items"
    e = Null
    a = ao.items
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

    tg = "selfAndAllFiles"
    e = Null
    a = ao.selfAndAllFiles
    AssertEqualWithMessage e, a, tg

    tg = "selfAndAllFolders"
    e = Null
    a = ao.selfAndAllFolders
    AssertEqualWithMessage e, a, tg

    tg = "selfAndAllItems"
    e = Null
    a = ao.selfAndAllItems
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
    Dim i,d,ao,obj,hasSomething
    For i=0 To Ubound(data)
        d = data(i)
        Set ao = new FileSystemProxy : ao.of(d)

        AssertEqualWithMessage d, ao, "i="&i&",Default"
        assertFsProperties ao,d,i
        AssertEqualWithMessage "<FileSystemProxy>"&d, ao.toString, "i="&i&",toString"
        
        Set obj = getFolderItem2(d)

        hasSomething=expectHasFile(obj)
        AssertEqualWithMessage hasSomething, ao.hasFile, "i="&i&",hasFile"
        If hasSomething Then
            assertFsItems ao.files,d,i&".files",False,False,Cl_FILE
        Else
            AssertEqualWithMessage cf_toString(Array()), cf_toString(ao.files), "i="&i&",.files"
        End If
        assertFsItems ao.allFiles,d,i&".allFiles",False,True,Cl_FILE
        assertFsItems ao.selfAndAllFiles,d,i&".selfAndAllFiles",True,True,Cl_FILE

        hasSomething=expectHasFolder(obj)
        AssertEqualWithMessage hasSomething, ao.hasFolder, "i="&i&",hasFolder"
        If hasSomething Then
            assertFsItems ao.folders,d,i&".files",False,False,Cl_FOLDER
        Else
            AssertEqualWithMessage cf_toString(Array()), cf_toString(ao.folders), "i="&i&",.folders"
        End If
        assertFsItems ao.allFolders,d,i&".allFolders",False,True,Cl_FOLDER
        assertFsItems ao.selfAndAllFolders,d,i&".selfAndAllFolders",True,True,Cl_FOLDER

        hasSomething=expectHasItem(obj)
        AssertEqualWithMessage hasSomething, ao.hasItem, "i="&i&",hasItem"
        If hasSomething Then
            assertFsItems ao.items,d,i&".items",False,False,Empty
            assertFsItems ao.allItems,d,i&".allItems",False,True,Empty
        Else
            AssertEqualWithMessage cf_toString(Array()), cf_toString(ao.items), "i="&i&",.items"
            AssertEqualWithMessage cf_toString(Array()), cf_toString(ao.allItems), "i="&i&",.allItems"
        End If
        assertFsItems ao.selfAndAllItems,d,i&".selfAndAllItems",True,True,Empty

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
    Dim i,cases
    For Each i In defs
        pusha cases,caseNormal(i)
    Next
    For Each i In extractTargetForZip(defs)
        pusha cases,caseZip(i)
    Next
    createData = caseToData(cases)
End Function
Function extractTargetForZip(defs)
    Dim i,ret : ret=Array()
    For Each i In defs
        If Not containEmpty(i) Then push ret,i
    Next
    extractTargetForZip=ret
End Function
Function containEmpty(target)
    containEmpty=True
    If isEmptyArray(target) Then Exit Function
    If IsArray(target) Then
        Dim i
        For Each i In target
            If containEmpty(i) Then Exit Function
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
    Dim ret,ele
    For Each ele In cases
        pushA ret,ele
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
Function caseNormal(d)
    caseNormal=createFolderAt(PsPathTempFolder,d)
End Function
Function caseZip(d)
    Dim flPath : flPath=createFolderAt(PsPathTempFolder,d)
    Dim i,paths
    For Each i in getFolderItem2(flPath(0)).GetFolder.Items
        push paths,i.path
    Next
    
    Dim zipPath : zipPath=getTempFilePath(PsPathTempFolder,"zip")
    zip paths,zipPath
    
    Dim ret
    For Each i In flPath
        push ret,Replace(i,flPath(0),zipPath)
    Next
    caseZip=ret
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
'   .hasFile
Function expectHasFile(obj)
    Dim ret : ret = False
    Dim path : path = obj.path
    If fso.FolderExists(path) Then
        ret=(fso.GetFolder(path).Files.Count>0)
    ElseIf obj.IsFolder Then
        Dim ele
        For Each ele In obj.GetFolder.Items
            ret = Not(ele.IsFolder)
            If ret Then Exit For
        Next
    End If
    expectHasFile=ret
End Function
'for verify the following properties
'   .hasFolder
Function expectHasFolder(obj)
    Dim ret : ret = False
    Dim path : path = obj.path
    If fso.FolderExists(path) Then
        ret=(fso.GetFolder(path).SubFolders.Count>0)
    ElseIf obj.IsFolder Then
        Dim ele
        For Each ele In obj.GetFolder.Items
            ret = ele.IsFolder
            If ret Then Exit For
        Next
    End If
    expectHasFolder=ret
End Function
'for verify the following properties
'   .hasItem
Function expectHasItem(obj)
    expectHasItem = False
    If obj.IsFolder Then expectHasItem = obj.GetFolder.Items.Count>0
End Function

'to verify the following properties
'   .baseName
'   .dateLastModified
'   .extension
'   .isBrowsable
'   .isFileSystem
'   .isFolder
'   .isLink
'   .items
'   .name
'   .parentFolder
'   .path
'   .size
'   .type
Sub assertFsProperties(target,path,comment)
    Dim obj : obj = Null
    
    If fso.FolderExists(path) Then
        Set obj=fso.GetFolder(path)
    ElseIf fso.FileExists(path) Then
        Set obj=fso.GetFile(path)
    End IF
    If Isnull(obj) Then
        With getFolderItem2(path)
            AssertEqualWithMessage .ModifyDate, target.dateLastModified, "comment="&comment&",dateLastModified"
            AssertEqualWithMessage fso.GetFileName(path), target.name, "comment="&comment&",name"
            AssertEqualWithMessage "FileSystemProxy", TypeName(target.parentFolder), "comment="&comment&",parentFolder object"
            AssertEqualWithMessage fso.GetParentFolderName(path), target.parentFolder, "comment="&comment&",parentFolder"
            AssertEqualWithMessage .Path, target.path, "comment="&comment&",path"
            AssertEqualWithMessage .Size, target.size, "comment="&comment&",size"
            AssertEqualWithMessage .Type, target.type, "comment="&comment&",type"
        End With
    Else
        With obj
            AssertEqualWithMessage .DateLastModified, target.dateLastModified, "comment="&comment&",dateLastModified"
            AssertEqualWithMessage .Name, target.name, "comment="&comment&",name"
            AssertEqualWithMessage "FileSystemProxy", TypeName(target.parentFolder), "comment="&comment&",parentFolder object"
            AssertEqualWithMessage .ParentFolder, target.parentFolder, "comment="&comment&",parentFolder"
            AssertEqualWithMessage .Path, target.path, "comment="&comment&",path"
            AssertEqualWithMessage .Size, target.size, "comment="&comment&",size"
            AssertEqualWithMessage .Type, target.type, "comment="&comment&",type"
        End With
    End If

    Set obj = getFolderItem2(path)
    With obj
        AssertEqualWithMessage .IsBrowsable, target.isBrowsable, "comment="&comment&",isBrowsable"
        AssertEqualWithMessage .IsFileSystem, target.isFileSystem, "comment="&comment&",isFileSystem"
        AssertEqualWithMessage .IsLink, target.isLink, "comment="&comment&",isLink"
    End With

    With fso
        AssertEqualWithMessage .FolderExists(path), target.isFolder, "comment="&comment&",isFolder"
        AssertEqualWithMessage .GetBaseName(path), target.baseName, "comment="&comment&",baseName"
        AssertEqualWithMessage .GetExtensionName(path), target.extension, "comment="&comment&",extension"
    End With
    
    Set obj = Nothing
End Sub
'to verify the following properties
'   .allFiles
'   .allFolders
'   .allItems
'   .files
'   .folders
'   .items
'   .selfAndAllFiles
'   .selfAndAllFolders
'   .selfAndAllItems
Sub assertFsItems(items,path,comment,self,recursive,itemType)
    Dim dic : Set dic = dictionary
    For Each i In items
        If dic.Exists(i) Then AssertFailWithMessage "comment="&comment&",items Duplication!"
        dic.Add i.path,False
    Next

    assertFsItemsEachItem path,comment,dic,self,recursive,itemType
    
    Dim i
    For Each i In dic.Keys
        If Not dic(i) Then AssertFailWithMessage "comment="&comment&", " & i & " Not Found !"
    Next

    AssertWithMessage True, "all ok"
    Set dic = Nothing
End Sub
Sub assertFsItemsEachItem(path,comment,dic,self,recursive,itemType)
    If self Then existsItem path,comment,dic,itemType
    Dim i
    If fso.FolderExists(path) Then
        For Each i in fso.GetFolder(path).Files
            existsItem i.path,comment,dic,itemType
        Next
        For Each i in fso.GetFolder(path).SubFolders
            existsItem i.path,comment,dic,itemType
            If recursive Then assertFsItemsEachItem i.path,comment,dic,self,recursive,itemType
        Next
    ElseIf getFolderItem2(path).IsFolder Then
        For Each i in getFolderItem2(path).GetFolder.Items
            existsItem i.path,comment,dic,itemType
            If recursive And i.IsFolder Then
                If i.GetFolder.Items.Count>0 Then assertFsItemsEachItem i.path,comment,dic,self,recursive,itemType
            End If
        Next
    End If
End Sub
Sub existsItem(path,comment,dic,itemType)
    Dim flg : flg = False
    Dim sItemName : sItemName = "items"
    Select Case itemType
    Case Cl_FILE
        sItemName = "files"
        If Not getFolderItem2(path).IsFolder Then flg = True
    Case Cl_FOLDER
        sItemName = "folders"
        If getFolderItem2(path).IsFolder Then flg = True
    Case Else
        flg = True
    End Select
    
    If flg Then
        If Not dic.Exists(path) Then AssertFailWithMessage "comment="&comment&","&sItemName&" Not Exists " & path
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
