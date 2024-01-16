' clsAdptFile.vbs: test.
' @import ../../lib/clsAdptFile.vbs
' @import ../../lib/clsCmArray.vbs
' @import ../../lib/clsCmBroker.vbs
' @import ../../lib/clsCmBufferedReader.vbs
' @import ../../lib/clsCmBufferedWriter.vbs
' @import ../../lib/clsCmCalendar.vbs
' @import ../../lib/clsCmCharacterType.vbs
' @import ../../lib/clsCmCssGenerator.vbs
' @import ../../lib/clsCmHtmlGenerator.vbs
' @import ../../lib/clsCmReturnValue.vbs
' @import ../../lib/clsCompareExcel.vbs
' @import ../../lib/libCom.vbs

Option Explicit

Const MY_NAME = "test_clsAdptFile.vbs"
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
'clsAdptFile
Sub Test_clsAdptFile
    Dim a : Set a = new clsAdptFile
    AssertEqualWithMessage 0, VarType(a), "VarType"
    AssertEqualWithMessage "clsAdptFile", TypeName(a), "TypeName"
End Sub

'###################################################################################################
'clsAdptFile.setFileObject()
Sub Test_clsAdptFile_setFileObject
    Dim p,d
    p = WScript.ScriptFullName
    Set d = new_ShellApp().Namespace(new_Fso().GetParentFolderName(p)).Items().Item(new_Fso().GetFileName(p))

    Dim ao,a
    Set ao = new clsAdptFile
    Set a = ao.setFileObject(d)

    AssertEqualWithMessage 8, VarType(a), "VarType"
    AssertEqualWithMessage "clsAdptFile", TypeName(a), "TypeName"
    AssertEqualWithMessage p, a.path, "path"
End Sub
Sub Test_clsAdptFile_setFileObject_Err
    Dim d
    Set d = new_Dic()

    Dim ao,a
    Set ao = new clsAdptFile
    On Error Resume Next
    Set a = ao.setFileObject(d)

    AssertEqualWithMessage 438, Err.Number, "Err.Number"
    AssertEqualWithMessage "clsAdptFile.vbs:clsAdptFile+setFileObject()", Err.Source, "Err.Source"
    AssertEqualWithMessage "オブジェクトでサポートされていないプロパティまたはメソッドです。", Err.Description, "Err.Description"
End Sub

'###################################################################################################
'clsAdptFile.setFilePath()
Sub Test_clsAdptFile_setFilePath
    Dim p
    p = WScript.ScriptFullName

    Dim ao,a
    Set ao = new clsAdptFile
    Set a = ao.setFilePath(p)

    AssertEqualWithMessage 8, VarType(a), "VarType"
    AssertEqualWithMessage "clsAdptFile", TypeName(a), "TypeName"
    AssertEqualWithMessage p, a.path, "path"
End Sub
Sub Test_clsAdptFile_setFilePath_Err
    Dim d
    d = vbNullString

    Dim ao,a
    Set ao = new clsAdptFile
    On Error Resume Next
    Set a = ao.setFilePath(d)

    AssertEqualWithMessage 76, Err.Number, "Err.Number"
    AssertEqualWithMessage "clsAdptFile.vbs:clsAdptFile+setFilePath()", Err.Source, "Err.Source"
    AssertEqualWithMessage "パスが見つかりません。", Err.Description, "Err.Description"
End Sub

'###################################################################################################
'clsAdptFile.getters DateLastModified,Name,ParentFolder,Path,Size,Type
Sub Test_clsAdptFile_getters
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
        Set ao = new clsAdptFile : ao.setFilePath(d)

        AssertEqualWithMessage eo, ao, "i="&i&",Default"
        AssertEqualWithMessage eo.DateLastModified, ao.DateLastModified, "i="&i&",DateLastModified"
        AssertEqualWithMessage eo.Name, ao.Name, "i="&i&",Name"
        AssertEqualWithMessage eo.ParentFolder, ao.ParentFolder, "i="&i&",ParentFolder"
        AssertEqualWithMessage eo.Path, ao.Path, "i="&i&",Path"
        AssertEqualWithMessage eo.Size, ao.Size, "i="&i&",Size"
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
