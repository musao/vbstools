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
'FileSystemProxy.basename,dateLastModified,isBrowsable,isFileSystem,isFolder,isLink,name,parentFolder,path,size,toString,type
Sub Test_FileProxy_dateLastModified_isBrowsable_isFileSystem_isFolder_isLink_name_parentFolder_path_size_toString_type_initial
    dim tg,a,ao,e
    set ao = (new FileSystemProxy)

    tg = "A.basename"
    e = Null
    a = ao.basename
    AssertEqualWithMessage e, a, tg

    tg = "B.dateLastModified"
    e = Null
    a = ao.dateLastModified
    AssertEqualWithMessage e, a, tg

    tg = "C.extension"
    e = Null
    a = ao.extension
    AssertEqualWithMessage e, a, tg

    tg = "D.isBrowsable"
    e = Null
    a = ao.isBrowsable
    AssertEqualWithMessage e, a, tg

    tg = "E.isFileSystem"
    e = Null
    a = ao.isFileSystem
    AssertEqualWithMessage e, a, tg

    tg = "F.isFolder"
    e = Null
    a = ao.isFolder
    AssertEqualWithMessage e, a, tg

    tg = "G.isLink"
    e = Null
    a = ao.isLink
    AssertEqualWithMessage e, a, tg

    tg = "H.name"
    e = Null
    a = ao.name
    AssertEqualWithMessage e, a, tg

    tg = "I.parentFolder"
    e = Null
    a = ao.parentFolder
    AssertEqualWithMessage e, a, tg

    tg = "J.path"
    e = Null
    a = ao.path
    AssertEqualWithMessage e, a, tg

    tg = "K.size"
    e = Null
    a = ao.size
    AssertEqualWithMessage e, a, tg

    tg = "L.toString"
    e = "<FileSystemProxy>"
    a = ao.toString
    AssertEqualWithMessage e, a, tg

    tg = "M.type"
    e = Null
    a = ao.type
    AssertEqualWithMessage e, a, tg
End Sub

Sub Test_FileProxy_dateLastModified_isBrowsable_isFileSystem_isFolder_isLink_name_parentFolder_path_size_toString_type
    Dim data
    data = Array( _
                WScript.ScriptFullName _
                , createShortCutDefault() _
                , createUrlShortCutDefault() _
                , createTextFileDefault() _
                , createFolderDefault() _
                , createZipDefault() _
                )

    Dim i,d,ao
    For i=0 To Ubound(data)
        d = data(i)
        Set ao = new FileSystemProxy : ao.of(d)

        AssertEqualWithMessage d, ao, "i="&i&",Default"
        AssertEqualWithMessage "<FileSystemProxy>"&d, ao.toString, "i="&i&",toString"
        assertFsItems ao,d,i
    Next
End Sub



'###################################################################################################
'common
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
Function createFolderDefault
    Dim path : path = getTempFolderPath(PsPathTempFolder)
    fso.CreateFolder path
    createFolderDefault = path
End Function
Function createZipDefault
    Dim paths(3)
    paths(0)=WScript.ScriptFullName
    paths(1)=createShortCutDefault()
    paths(2)=createUrlShortCutDefault()
    paths(3)=createTextFileDefault()
    
    Dim path : path = getTempFilePath(PsPathTempFolder,"zip")
    zip paths,path
    createZipDefault = path
End Function
Sub assertFsItems(target,path,comment)
    Dim obj
    
    If fso.FolderExists(path) Then Set obj=fso.GetFolder(path) Else Set obj=fso.GetFile(path)
    With obj
        AssertEqualWithMessage .DateLastModified, target.dateLastModified, "comment="&comment&",dateLastModified"
        AssertEqualWithMessage .Name, target.name, "comment="&comment&",name"
        AssertEqualWithMessage .ParentFolder, target.parentFolder, "comment="&comment&",parentFolder"
        AssertEqualWithMessage .Path, target.path, "comment="&comment&",path"
        AssertEqualWithMessage .Size, target.size, "comment="&comment&",size"
        AssertEqualWithMessage .Type, target.type, "comment="&comment&",type"
    End With

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

Function fso
    Set fso = CreateObject("Scripting.FileSystemObject")
End Function
Function shellApp
    Set shellApp = CreateObject("Shell.Application")
End Function
Function shell
    Set shell = CreateObject("WScript.Shell")
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
