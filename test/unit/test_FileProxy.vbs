' FileProxy.vbs: test.
' @import ../../lib/com/FileProxy.vbs
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
'FileProxy
Sub Test_FileProxy
    Dim a : Set a = new FileProxy
    AssertEqualWithMessage 1, VarType(a), "VarType"
    AssertEqualWithMessage "FileProxy", TypeName(a), "TypeName"
End Sub

'###################################################################################################
'FileProxy.of()
Sub Test_FileProxy_of
    Dim p,a
    p = WScript.ScriptFullName
    Set a = (new FileProxy).of(p)

    AssertEqualWithMessage 8, VarType(a), "VarType"
    AssertEqualWithMessage "FileProxy", TypeName(a), "TypeName"
    AssertEqualWithMessage p, a.path, "path"
End Sub
Sub Test_FileProxy_of_Err
    Dim a,d : d = vbNullString
    On Error Resume Next
    Call (new FileProxy).of(d)

    AssertEqualWithMessage "FileProxy+of()", Err.Source, "Err.Source"
    AssertMatchWithMessage "^invalid argument.*", Err.Description, "Err.Description"
End Sub
Sub Test_FileProxy_of_ErrImmutable
    On Error Resume Next
    Dim ao
    Set ao = (new FileProxy).of(makeShortCut())
    Call ao.of(makeUrlShortCut())

    AssertEqualWithMessage "FileProxy+of()", Err.Source, "Err.Source"
    AssertEqualWithMessage "Because it is an immutable variable, its value cannot be changed.", Err.Description, "Err.Description"
End Sub

'###################################################################################################
'FileProxy.dateLastModified,name,parentFolder,path,size,toString,type
Sub Test_FileProxy_dateLastModified_name_parentFolder_path_size_toString_type_initial
    dim tg,a,ao,e
    set ao = (new FileProxy)

    tg = "A.dateLastModified"
    e = Null
    a = ao.dateLastModified
    AssertEqualWithMessage e, a, tg

    tg = "B.name"
    e = Null
    a = ao.name
    AssertEqualWithMessage e, a, tg

    tg = "C.parentFolder"
    e = Null
    a = ao.parentFolder
    AssertEqualWithMessage e, a, tg

    tg = "D.path"
    e = Null
    a = ao.path
    AssertEqualWithMessage e, a, tg

    tg = "E.size"
    e = Null
    a = ao.size
    AssertEqualWithMessage e, a, tg

    tg = "F.toString"
    e = "<FileProxy>"
    a = ao.toString
    AssertEqualWithMessage e, a, tg

    tg = "G.type"
    e = Null
    a = ao.type
    AssertEqualWithMessage e, a, tg
End Sub

Sub Test_FileProxy_dateLastModified_name_parentFolder_path_size_toString_type
    Dim data
    data = Array( _
                WScript.ScriptFullName _
                , makeShortCut() _
                , makeUrlShortCut() _
                )

    Dim i,d,eo,ao
    For i=0 To Ubound(data)
        d = data(i)
        Set eo = CreateObject("Scripting.FileSystemObject").GetFile(d)
        Set ao = new FileProxy : ao.of(d)

        AssertEqualWithMessage eo, ao, "i="&i&",Default"
        AssertEqualWithMessage eo.DateLastModified, ao.DateLastModified, "i="&i&",DateLastModified"
        AssertEqualWithMessage eo.Name, ao.Name, "i="&i&",Name"
        AssertEqualWithMessage eo.ParentFolder, ao.ParentFolder, "i="&i&",ParentFolder"
        AssertEqualWithMessage eo.Path, ao.Path, "i="&i&",Path"
        AssertEqualWithMessage eo.Size, ao.Size, "i="&i&",Size"
        AssertEqualWithMessage "<FileProxy>"&ao.Path, ao.toString, "i="&i&",toString"
        AssertEqualWithMessage eo.Type, ao.Type, "i="&i&",Type"
    Next
End Sub



'###################################################################################################
'common
Function makeShortCut
    Dim path : path = CreateObject("Scripting.FileSystemObject").BuildPath(PsPathTempFolder, "test.lnk")

    With WScript.CreateObject("WScript.Shell")
        Dim shortcut
        Set shortcut = .CreateShortcut(path)
    End With
    With shortcut
        .TargetPath = WScript.ScriptFullName
        .Save
    End With
    makeShortCut = path
    Set shortcut = Nothing
End Function
Function makeUrlShortCut
    Dim path : path = CreateObject("Scripting.FileSystemObject").BuildPath(PsPathTempFolder, "test.url")

    With WScript.CreateObject("WScript.Shell")
        Dim shortcut
        Set shortcut = .CreateShortcut(path)
    End With
    With shortcut
        .TargetPath = "https://www.google.com/"
        .Save
    End With
    makeUrlShortCut = path
    Set shortcut = Nothing
End Function

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
