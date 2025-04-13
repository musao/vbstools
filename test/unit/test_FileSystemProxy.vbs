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

Option Explicit

Const MY_NAME = "test_FileProxy.vbs"
Dim PsPathTempFolder

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
'FileSystemProxy.allItems,basename,dateLastModified,extension,hasItem,isBrowsable,isFileSystem,isFolder,isLink,items,name,parentFolder,path,size,toString,type
Sub Test_FileProxy_allItems_basename_dateLastModified_extension_hasItem_isBrowsable_isFileSystem_isFolder_isLink_items_name_parentFolder_path_size_toString_type_initial
    dim tg,a,ao,e
    set ao = (new FileSystemProxy)

    tg = "allItems"
    e = Null
    a = ao.allItems
    AssertEqualWithMessage e, a, tg

    tg = "basename"
    e = Null
    a = ao.basename
    AssertEqualWithMessage e, a, tg

    tg = "dateLastModified"
    e = Null
    a = ao.dateLastModified
    AssertEqualWithMessage e, a, tg

    tg = "extension"
    e = Null
    a = ao.extension
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
Sub Test_FileProxy_allItems_basename_dateLastModified_extension_hasItem_isBrowsable_isFileSystem_isFolder_isLink_items_name_parentFolder_path_size_toString_type
    Dim data : data = createData()
    Dim i,d,ao,obj,expectHasItem
    For i=0 To Ubound(data)
        d = data(i)
        Set ao = new FileSystemProxy : ao.of(d)

        AssertEqualWithMessage d, ao, "i="&i&",Default"
        assertFsProp ao,d,i
        AssertEqualWithMessage "<FileSystemProxy>"&d, ao.toString, "i="&i&",toString"
        
        Set obj = shellApp.Namespace(fso.GetParentFolderName(d)).Items().Item(fso.GetFileName(d))
        expectHasItem=False
        If obj.IsFolder Then expectHasItem = obj.GetFolder.Items.Count>0
        AssertEqualWithMessage expectHasItem, ao.hasItem, "i="&i&",hasItem"

        If expectHasItem Then
            assertFsItems ao.items,d,i,False
            assertFsItems ao.allItems,d,i,True
        Else
            AssertEqualWithMessage cf_toString(Array()), cf_toString(ao.items), "i="&i&",items"
            AssertEqualWithMessage cf_toString(Array()), cf_toString(ao.allItems), "i="&i&",allItems"
        End If

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
    Set ao = (new FileSystemProxy).of(createShortCutDefault())
    Call ao.of(createUrlShortCutDefault())

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
    Dim ret,i
    ret = Array( _
                WScript.ScriptFullName _
                , createShortCutDefault() _
                , createUrlShortCutDefault() _
                , createTextFileDefault() _
                , createEmptyFolderDefault() _
                )
    For Each i In createFolderDefault()
        Redim Preserve ret(Ubound(ret)+1)
        ret(Ubound(ret)) = i
    Next
    For Each i In createZipDefault()
        Redim Preserve ret(Ubound(ret)+1)
        ret(Ubound(ret)) = i
    Next
    createData = ret
End Function
Function createShortCutDefault
    createShortCutDefault = createShortCutCommon(getTempFilePath(PsPathTempFolder,"lnk"), WScript.ScriptFullName)
End Function
Function createUrlShortCutDefault
    createUrlShortCutDefault = createShortCutCommon(getTempFilePath(PsPathTempFolder,"url"), "https://www.google.com/")
End Function
Function createTextFileDefault
    Dim path : path = getTempFilePath(PsPathTempFolder,"txt")
    With fso.OpenTextFile(path, 2, True, -1)
        .Write "hoge"
        .Close
    End With
    createTextFileDefault = path
End Function
Function createEmptyFolderDefault
    Dim path
    path = getTempFolderPath(PsPathTempFolder)
    fso.CreateFolder path
    createEmptyFolderDefault=path
End Function
Function createFolderDefault
    Dim path,ret(2)
    path = getTempFolderPath(PsPathTempFolder)
    fso.CreateFolder path

    ret(0)=path
    ret(1)=createShortCutCommon(getTempFilePath(path,"lnk"), WScript.ScriptFullName)
    ret(2)=createShortCutCommon(getTempFilePath(path,"url"), "https://www.google.com/")

    createFolderDefault=ret
End Function
Function createZipDefault
    Dim fpaths(1),fold,fdef : fdef = createFolderDefault
    fold=fdef(0)
    fpaths(0)=fdef(1)
    fpaths(1)=fdef(2)

    Dim paths() : Redim paths(4)
    paths(0)=fold
    paths(1)=WScript.ScriptFullName
    paths(2)=createShortCutDefault()
    paths(3)=createUrlShortCutDefault()
    paths(4)=createTextFileDefault()
    
    Dim path : path = getTempFilePath(PsPathTempFolder,"zip")
    zip paths,path

    Dim ret() : Redim ret(7)
    ret(0)=path
    ret(1)=fso.BuildPath(path, fso.GetFileName(fold))
    ret(2)=fso.BuildPath(ret(1), fso.GetFileName(fpaths(0)))
    ret(3)=fso.BuildPath(ret(1), fso.GetFileName(fpaths(1)))
    ret(4)=fso.BuildPath(path, fso.GetFileName(paths(1)))
    ret(5)=fso.BuildPath(path, fso.GetFileName(paths(2)))
    ret(6)=fso.BuildPath(path, fso.GetFileName(paths(3)))
    ret(7)=fso.BuildPath(path, fso.GetFileName(paths(4)))

    createZipDefault = Ret
End Function
Sub assertFsProp(target,path,comment)
    Dim obj : obj = Null
    
    If fso.FolderExists(path) Then
        Set obj=fso.GetFolder(path)
    ElseIf fso.FileExists(path) Then
        Set obj=fso.GetFile(path)
    End IF
    If Isnull(obj) Then
        With shellApp.Namespace(fso.GetParentFolderName(path)).Items().Item(fso.GetFileName(path))
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

    Set obj = shellApp.Namespace(fso.GetParentFolderName(path)).Items().Item(fso.GetFileName(path))
    With obj
        AssertEqualWithMessage .IsBrowsable, target.isBrowsable, "comment="&comment&",isBrowsable"
        AssertEqualWithMessage .IsFileSystem, target.isFileSystem, "comment="&comment&",isFileSystem"
        AssertEqualWithMessage .IsFolder, target.isFolder, "comment="&comment&",isFolder"
        AssertEqualWithMessage .IsLink, target.isLink, "comment="&comment&",isLink"
    End With

    With fso
        AssertEqualWithMessage .GetBaseName(path), target.basename, "comment="&comment&",basename"
        AssertEqualWithMessage .GetExtensionName(path), target.extension, "comment="&comment&",extension"
    End With
    
    Set obj = Nothing
End Sub
Sub assertFsItems(items,path,comment,recursive)
    Dim dic
    Set dic = dictionary
    For Each i In items
        If dic.Exists(i) Then AssertFailWithMessage "comment="&comment&",items Duplication!"
        dic.Add i.path,False
    Next

    assertFsItemsEachItem path,comment,dic,recursive
    
    Dim i
    For Each i In dic.Keys
        If Not dic(i) Then AssertFailWithMessage "comment="&comment&", " & i & " Not Found !"
    Next

    AssertWithMessage True, "all ok"
    Set dic = Nothing
End Sub
Sub assertFsItemsEachItem(path,comment,dic,recursive)
    Dim i
    If fso.FolderExists(path) Then
        For Each i in fso.GetFolder(path).Files
            existsItem i,comment,dic
        Next
        For Each i in fso.GetFolder(path).SubFolders
            existsItem i,comment,dic
            If recursive Then assertFsItemsEachItem i.path,comment,dic,recursive
        Next
    Else
        For Each i in shellApp.Namespace(fso.GetParentFolderName(path)).Items().Item(fso.GetFileName(path)).GetFolder.Items
            existsItem i,comment,dic
            If recursive And i.IsFolder Then
                If i.GetFolder.Items.Count>0 Then assertFsItemsEachItem i.path,comment,dic,recursive
            End If
        Next
    End If
End Sub
Sub existsItem(target,comment,dic)
    If Not dic.Exists(target.path) Then AssertFailWithMessage "comment="&comment&",items Not Exists " & target
    dic(target.path)=True
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
    Dim path
    path = join(paths,",")

    Dim cmd : cmd = _
        "powershell -NoProfile -ExecutionPolicy Unrestricted Compress-Archive" _
        & " -Path " & path _
        & " -DestinationPath " & zpath
    Call shell.Run(cmd, 0, True)
End Sub
Function createShortCutCommon(path,target)
    With shell.CreateShortcut(path)
        .TargetPath = target
        .Save
    End With
    createShortCutCommon = path
End Function

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
