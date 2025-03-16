' libCom.vbs: fs_* procedure test.
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

Const MY_NAME = "test_libCom_fs.vbs"
Dim PsPathTempFolder
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim adodb : Set adodb = CreateObject("ADODB.Stream")

'###################################################################################################
'SetUp()/TearDown()
Sub SetUp()
    '実行スクリプト直下に当ファイル名で一時フォルダ作成
    PsPathTempFolder = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), MY_NAME)
    If Not (fso.FolderExists(PsPathTempFolder)) Then fso.CreateFolder(PsPathTempFolder)
End Sub
Sub TearDown()
    '当テストで作成した一時フォルダを削除する
    fso.DeleteFolder PsPathTempFolder
End Sub

'###################################################################################################
'fs_getAllFiles()
Sub Test_fs_getAllFiles_OnlyFiles
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFile(True,False,"Test_fs_copyFile_Normal"))
    
End Sub


'###################################################################################################
'fs_copyFile()
Sub Test_fs_copyFile_Normal
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFile(True,False,"Test_fs_copyFile_Normal"))
    
    '実行
    Dim a : Set a = fs_copyFile(d.Item("from").Item("path"),d.Item("to").Item("path"))

    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionMergeFile(d))
End Sub
Sub Test_fs_copyFile_Normal_OverRide
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFile(True,True,"Test_fs_copyFile_Normal_OverRide"))
    
    '実行
    Dim a : Set a = fs_copyFile(d.Item("from").Item("path"),d.Item("to").Item("path"))

    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionMergeFile(d))
End Sub
Sub Test_fs_copyFile_Normal_FromFileLocked
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFile(True,False,"Test_fs_copyFile_Normal_FromFileLocked"))
    
    'fromファイルをロックする
    With lockFile(d.Item("from").Item("path"))
       '実行
        Dim a : Set a = fs_copyFile(d.Item("from").Item("path"),d.Item("to").Item("path"))

        '戻り値の検証
        Dim e
        e = True
        AssertEqualWithMessage e, a, "ret"
        e = False
        AssertEqualWithMessage e, a.isErr, "isErr"

        .Close
    End With

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionMergeFile(d))
End Sub
Sub Test_fs_copyFile_Err_FromFileNotExists
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFile(False,False,"Test_fs_copyFile_Err_FromFileNotExists"))

    '実行
    Dim a : Set a = fs_copyFile(d.Item("from").Item("path"),d.Item("to").Item("path"))
    '戻り値の検証
    Dim e
    e = False
    AssertEqualWithMessage e, a, "ret"
    e = True
    AssertEqualWithMessage e, a.isErr, "isErr"
    e = 53
    AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
    e = "ファイルが見つかりません。"
    AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionUnchange("to",d))
End Sub
Sub Test_fs_copyFile_Err_ToFileLocked
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFile(True,True,"Test_fs_copyFile_Err_ToFileLocked"))
    
    'toファイルをロックする
    With lockFile(d.Item("to").Item("path"))
       '実行
        Dim a : Set a = fs_copyFile(d.Item("from").Item("path"),d.Item("to").Item("path"))

        '戻り値の検証
        Dim e
        e = False
        AssertEqualWithMessage e, a, "ret"
        e = True
        AssertEqualWithMessage e, a.isErr, "isErr"
        e = 70
        AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
        e = "書き込みできません。"
        AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

        .Close
    End With

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionUnchange("to",d))
End Sub

'###################################################################################################
'fs_copyFolder()
Sub Test_fs_copyFolder_Normal
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFolder(True,False,"Test_fs_copyFolder_Normal"))
    
    '実行
    Dim a : Set a = fs_copyFolder(d.Item("from-folder").Item("path"),d.Item("to-folder").Item("path"))
    
    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"
    
    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionMergeFolder(d))
End Sub
Sub Test_fs_copyFolder_Normal_OverRide
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFolder(True,True,"Test_fs_copyFolder_Normal_OverRide"))

    '実行
    Dim a : Set a = fs_copyFolder(d.Item("from-folder").Item("path"),d.Item("to-folder").Item("path"))

    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"
    
    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionMergeFolder(d))
End Sub
Sub Test_fs_copyFolder_Normal_OverRideWithUnrelatedFileLocked
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFolder(True,True,"Test_fs_copyFolder_Normal_OverRideWithUnrelatedFileLocked"))
    
    '上書きしないtoフォルダのファイル（to-fla）をロックする
    With lockFile(d.Item("to-fla").Item("path"))
       '実行
        Dim a : Set a = fs_copyFolder(d.Item("from-folder").Item("path"),d.Item("to-folder").Item("path"))

        '戻り値の検証
        Dim e
        e = True
        AssertEqualWithMessage e, a, "ret"
        e = False
        AssertEqualWithMessage e, a.isErr, "isErr"

        .Close
    End With
    
    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionMergeFolder(d))
End Sub
Sub Test_fs_copyFolder_Normal_FromFileLocked
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFolder(True,True,"Test_fs_copyFolder_Normal_FromFileLocked"))
    
    'fromフォルダのファイル（from-fl1）をロックする
    With lockFile(d.Item("from-fl1").Item("path"))
       '実行
        Dim a : Set a = fs_copyFolder(d.Item("from-folder").Item("path"),d.Item("to-folder").Item("path"))

        '戻り値の検証
        Dim e
        e = True
        AssertEqualWithMessage e, a, "ret"
        e = False
        AssertEqualWithMessage e, a.isErr, "isErr"

        .Close
    End With
    
    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionMergeFolder(d))
End Sub
Sub Test_fs_copyFolder_Err_FromFileNotExists
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFolder(False,False,"Test_fs_copyFolder_Err_FromFileNotExists"))

    '実行
    Dim a : Set a = fs_copyFolder(d.Item("from-folder").Item("path"),d.Item("to-folder").Item("path"))

    '戻り値の検証
    Dim e
    e = False
    AssertEqualWithMessage e, a, "ret"
    e = True
    AssertEqualWithMessage e, a.isErr, "isErr"
    e = 76
    AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
    e = "パスが見つかりません。"
    AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"
    
    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionUnchange("to",d))
End Sub
Sub Test_fs_copyFolder_Err_ToFileLocked
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFolder(True,True,"Test_fs_copyFolder_Err_ToFileLocked"))
    
    'toフォルダの上書きするファイル（to-flb）をロックする
    Dim lockedItem : Set lockedItem = d.Item("to-flb")
    With lockFile(lockedItem.Item("path"))
       '実行
        Dim a : Set a = fs_copyFolder(d.Item("from-folder").Item("path"),d.Item("to-folder").Item("path"))

        '戻り値の検証
        Dim e
        e = False
        AssertEqualWithMessage e, a, "ret"
        e = True
        AssertEqualWithMessage e, a.isErr, "isErr"
        e = 70
        AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
        e = "書き込みできません。"
        AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

        .Close
    End With
    
    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    'toフォルダで上書きするファイル（to-flb）まではコピーする
    assertFolderItems(createExpectDefinitionMergeFolderUntilOverRideFile(d,lockedItem.Item("relativePath")))
End Sub

'###################################################################################################
'fs_copyHere()
Sub Test_fs_copyHere_Normal_file2folder
    Dim from,toto
    from = CreateFileForCopyhere("Test_fs_copyHere_Normal_file2folder")
    toto = CreateFolderForCopyhere()

    Dim a,e
    Set a = fs_copyHere(from, toto)
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"

    With fso
        e = .GetFile(from).Size
        a = .GetFile(.BuildPath(toto, .GetFileName(from))).Size
        AssertEqualWithMessage e, a, "Size"
    End With
End Sub
Sub Test_fs_copyHere_Normal_folder2folder
    Dim from,toto
    from = CreateFolderForCopyhere()
    toto = CreateFolderForCopyhere()

    Dim a,e
    Set a = fs_copyHere(from, toto)
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"

    With fso
        e = .GetFolder(from).Size
        a = .GetFolder(.BuildPath(toto, .GetFileName(from))).Size
        AssertEqualWithMessage e, a, "Size"
    End With
End Sub

Function CreateFileForCopyhere(c)
    With fso
        Dim p : p = .BuildPath(PsPathTempFolder, .GetTempName())
        Dim ts : Set ts = .OpenTextFile(p, 2, True, -1)
    End With
    With ts
        .Write c
        .Close
    End With
    CreateFileForCopyhere = p
End Function
Function CreateFolderForCopyhere()
    With fso
        Dim p : p = .BuildPath(PsPathTempFolder, .GetTempName())
        .CreateFolder p
    End With
    CreateFolderForCopyhere = p
End Function

'###################################################################################################
'fs_createFolder()
Sub Test_fs_createFolder
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForCreateFolder(False,False,"Test_fs_createFolder"))

    '実行
    Dim a : Set a = fs_createFolder(d.Item("target").Item("path"))

    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"

    'データの検証
    assertFolderItems(createExpectDefinitionCreateFolder(d))
End Sub
Sub Test_fs_createFolder_ErrExistsFile
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForCreateFolder(True,False,"Test_fs_createFolder_ErrExistsFile"))

    '実行
    Dim a : Set a = fs_createFolder(d.Item("target").Item("path"))

    '戻り値の検証
    Dim e
    e = False
    AssertEqualWithMessage e, a, "ret"
    e = True
    AssertEqualWithMessage e, a.isErr, "isErr"
    e = 58
    AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
    e = "既に同名のファイルが存在しています。"
    AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("target",d))
End Sub
Sub Test_fs_createFolder_ErrExistsFolder
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForCreateFolder(False,True,"Test_fs_createFolder_ErrExistsFolder"))

    '実行
    Dim a : Set a = fs_createFolder(d.Item("target").Item("path"))

    '戻り値の検証
    Dim e
    e = False
    AssertEqualWithMessage e, a, "ret"
    e = True
    AssertEqualWithMessage e, a.isErr, "isErr"
    e = 58
    AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
    e = "既に同名のファイルが存在しています。"
    AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("target",d))
End Sub

'###################################################################################################
'fs_deleteFile()
Sub Test_fs_deleteFile
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForDeleteFile(True,"Test_fs_deleteFile"))

    '実行
    Dim a : Set a = fs_deleteFile(d.Item("target").Item("path"))

    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"

    'データの検証
    assertFolderItems(createExpectDefinitionDisappearFile("target",d))
End Sub
Sub Test_fs_deleteFile_Err_NotExists
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForDeleteFile(False,"Test_fs_deleteFile_Err_NotExists"))

    '実行
    Dim a : Set a = fs_deleteFile(d.Item("target").Item("path"))

    '戻り値の検証
    Dim e
    e = False
    AssertEqualWithMessage e, a, "ret"
    e = True
    AssertEqualWithMessage e, a.isErr, "isErr"
    e = 53
    AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
    e = "ファイルが見つかりません。"
    AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("target",d))
End Sub
Sub Test_fs_deleteFile_Err_FileLocked
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForDeleteFile(True,"Test_fs_deleteFile_Err_FileLocked"))

    'targetファイルをロックする
    Dim target : target = d.Item("target").Item("path")
    With lockFile(target)
       '実行
        Dim a : Set a = fs_deleteFile(target)

        '戻り値の検証
        Dim e
        e = False
        AssertEqualWithMessage e, a, "ret"
        e = True
        AssertEqualWithMessage e, a.isErr, "isErr"
        e = 70
        AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
        e = "書き込みできません。"
        AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

        .Close
    End With

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("target",d))
End Sub

'###################################################################################################
'fs_deleteFolder()
Sub Test_fs_deleteFolder
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForDeleteFolder(True,"Test_fs_deleteFolder"))

    '実行
    Dim a : Set a = fs_deleteFolder(d.Item("target").Item("path"))

    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"

    'データの検証
    assertFolderItems(createExpectDefinitionDisappearFolder("target",d))
End Sub
Sub Test_fs_deleteFolder_Err_NotExists
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForDeleteFolder(False,"Test_fs_deleteFolder_Err_NotExists"))

    '実行
    Dim a : Set a = fs_deleteFolder(d.Item("target").Item("path"))

    '戻り値の検証
    Dim e
    e = False
    AssertEqualWithMessage e, a, "ret"
    e = True
    AssertEqualWithMessage e, a.isErr, "isErr"
    e = 76
    AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
    e = "パスが見つかりません。"
    AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("target",d))
End Sub
Sub Test_fs_deleteFolder_Err_FileLocked
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForDeleteFolder(True,"Test_fs_deleteFolder_Err_FileLocked"))

    'target-fileファイルをロックする
    With lockFile(d.Item("target-file").Item("path"))
       '実行
        Dim a : Set a = fs_deleteFolder(d.Item("target").Item("path"))

        '戻り値の検証
        Dim e
        e = False
        AssertEqualWithMessage e, a, "ret"
        e = True
        AssertEqualWithMessage e, a.isErr, "isErr"
        e = 70
        AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
        e = "書き込みできません。"
        AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

        .Close
    End With

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("target",d))
End Sub

'###################################################################################################
'fs_moveFile()
Sub Test_fs_moveFile_Normal
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFile(True,False,"Test_fs_moveFile_Normal"))
    
    '実行
    Dim a : Set a = fs_moveFile(d.Item("from").Item("path"),d.Item("to").Item("path"))

    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"

    'データの検証
    assertFolderItems(createExpectDefinitionDisappearFile("from",d))
    assertFolderItems(createExpectDefinitionMergeFile(d))
End Sub
Sub Test_fs_moveFile_Err_OverRide
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFile(True,True,"Test_fs_moveFile_Err_OverRide"))
    
    '実行
    Dim a : Set a = fs_moveFile(d.Item("from").Item("path"),d.Item("to").Item("path"))

    '戻り値の検証
    Dim e
    e = False
    AssertEqualWithMessage e, a, "ret"
    e = True
    AssertEqualWithMessage e, a.isErr, "isErr"
    e = 58
    AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
    e = "既に同名のファイルが存在しています。"
    AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionUnchange("to",d))
End Sub
Sub Test_fs_moveFile_Err_FromFileLocked
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFile(True,False,"Test_fs_moveFile_Err_FromFileLocked"))
    
    'fromファイルをロックする
    With lockFile(d.Item("from").Item("path"))
       '実行
        Dim a : Set a = fs_moveFile(d.Item("from").Item("path"),d.Item("to").Item("path"))

        '戻り値の検証
        Dim e
        e = False
        AssertEqualWithMessage e, a, "ret"
        e = True
        AssertEqualWithMessage e, a.isErr, "isErr"
        e = 70
        AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
        e = "書き込みできません。"
        AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

        .Close
    End With

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionUnchange("to",d))
End Sub
Sub Test_fs_moveFile_Err_FromFileNotExists
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFile(False,False,"Test_fs_moveFile_Err_FromFileNotExists"))

    '実行
    Dim a : Set a = fs_moveFile(d.Item("from").Item("path"),d.Item("to").Item("path"))
    '戻り値の検証
    Dim e
    e = False
    AssertEqualWithMessage e, a, "ret"
    e = True
    AssertEqualWithMessage e, a.isErr, "isErr"
    e = 53
    AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
    e = "ファイルが見つかりません。"
    AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionUnchange("to",d))
End Sub
Sub Test_fs_moveFile_Err_ToFileLocked
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFile(True,True,"Test_fs_moveFile_Err_ToFileLocked"))
    
    'toファイルをロックする
    With lockFile(d.Item("to").Item("path"))
       '実行
        Dim a : Set a = fs_moveFile(d.Item("from").Item("path"),d.Item("to").Item("path"))

        '戻り値の検証
        Dim e
        e = False
        AssertEqualWithMessage e, a, "ret"
        e = True
        AssertEqualWithMessage e, a.isErr, "isErr"
        e = 58
        AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
        e = "既に同名のファイルが存在しています。"
        AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

        .Close
    End With

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionUnchange("to",d))
End Sub

'###################################################################################################
'fs_moveFolder()
Sub Test_fs_moveFolder_Normal
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFolder(True,False,"Test_fs_moveFolder_Normal"))
    
    '実行
    Dim a : Set a = fs_moveFolder(d.Item("from-folder").Item("path"),d.Item("to-folder").Item("path"))

    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"
    
    'データの検証
    assertFolderItems(createExpectDefinitionDisappearFolder("from-folder",d))
    assertFolderItems(createExpectDefinitionMergeFolder(d))
End Sub
Sub Test_fs_moveFolder_Err_OverRide
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFolder(True,True,"Test_fs_moveFolder_Err_OverRide"))
    
    '実行
    Dim a : Set a = fs_moveFolder(d.Item("from-folder").Item("path"),d.Item("to-folder").Item("path"))

    '戻り値の検証
    Dim e
    e = False
    AssertEqualWithMessage e, a, "ret"
    e = True
    AssertEqualWithMessage e, a.isErr, "isErr"
    e = 58
    AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
    e = "既に同名のファイルが存在しています。"
    AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"
    
    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionUnchange("to",d))
End Sub
Sub Test_fs_moveFolder_Err_OverRideWithUnrelatedFileLocked
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFolder(True,True,"Test_fs_moveFolder_Err_OverRideWithUnrelatedFileLocked"))
    
    'to-flaをロックする
    With lockFile(d.Item("to-fla").Item("path"))
       '実行
        Dim a : Set a = fs_moveFolder(d.Item("from-folder").Item("path"),d.Item("to-folder").Item("path"))

        '戻り値の検証
        Dim e
        e = False
        AssertEqualWithMessage e, a, "ret"
        e = True
        AssertEqualWithMessage e, a.isErr, "isErr"
        e = 58
        AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
        e = "既に同名のファイルが存在しています。"
        AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

        .Close
    End With
    
    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionUnchange("to",d))
End Sub
Sub Test_fs_moveFolder_Err_FromFileLocked
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFolder(True,True,"Test_fs_moveFolder_Err_FromFileLocked"))
    
    'fromフォルダのファイル（from-fl1）をロックする
    With lockFile(d.Item("from-fl1").Item("path"))
       '実行
        Dim a : Set a = fs_moveFolder(d.Item("from-folder").Item("path"),d.Item("to-folder").Item("path"))

        '戻り値の検証
        Dim e
        e = False
        AssertEqualWithMessage e, a, "ret"
        e = True
        AssertEqualWithMessage e, a.isErr, "isErr"
        e = 70
        AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
        e = "書き込みできません。"
        AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

        .Close
    End With
    
    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionUnchange("to",d))
End Sub
Sub Test_fs_moveFolder_Err_FromFileNotExists
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFolder(False,False,"Test_fs_moveFolder_Err_FromFileNotExists"))
    
    '実行
    Dim a : Set a = fs_moveFolder(d.Item("from-folder").Item("path"),d.Item("to-folder").Item("path"))

    '戻り値の検証
    Dim e
    e = False
    AssertEqualWithMessage e, a, "ret"
    e = True
    AssertEqualWithMessage e, a.isErr, "isErr"
    e = 76
    AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
    e = "パスが見つかりません。"
    AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"
    
    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionUnchange("to",d))
End Sub
Sub Test_fs_moveFolder_Err_ToFileLocked
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFolder(True,True,"Test_fs_moveFolder_Err_ToFileLocked"))
    
    'toフォルダの上書きするファイル（to-flb）をロックする
    With lockFile(d.Item("to-flb").Item("path"))
       '実行
        Dim a : Set a = fs_moveFolder(d.Item("from-folder").Item("path"),d.Item("to-folder").Item("path"))

        '戻り値の検証
        Dim e
        e = False
        AssertEqualWithMessage e, a, "ret"
        e = True
        AssertEqualWithMessage e, a.isErr, "isErr"
        e = 58
        AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
        e = "既に同名のファイルが存在しています。"
        AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

        .Close
    End With
    
    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionUnchange("to",d))
End Sub

'###################################################################################################
'fs_readFile()
Sub Test_fs_readFile
    'データ定義と生成
    Dim cont
    cont = "lmn" & vbNewLine & "ⅢⅥⅩ" & vbNewLine & "ｱｲｳ" & vbNewLine & ChrW(12316) 'ChrW(12316)='\u301c'（波ダッシュ・波型）Sjisに変換できない文字
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile", True, cont))
    
    '実行
    Dim a : Set a = fs_readFile(d.Item("target").Item("path"))

    '戻り値の検証
    Dim e
    e = cont
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"
End Sub
Sub Test_fs_readFile_Err
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile", False, Empty))
    
    '実行
    Dim a : Set a = fs_readFile(d.Item("target").Item("path"))

    '戻り値の検証
    Dim e
    e = Empty
    AssertEqualWithMessage e, a, "ret"
    e = True
    AssertEqualWithMessage e, a.isErr, "isErr"
    e = 505
    AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
    e = "不正な参照です。"
    AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"
End Sub

'###################################################################################################
'fs_wrapInQuotes()
Sub Test_fs_wrapInQuotes
    Dim data
    data = Array( _
            Array("data", Chr(34) & "data" & Chr(34)) _
            , Array(Chr(34), Chr(34) & Chr(34)&Chr(34) & Chr(34)) _
            , Array(" ", Chr(34) & " " & Chr(34)) _
            )
    
    Dim i,d,e,a
    For i=0 To Ubound(data)
        d = data(i)(0)
        e = data(i)(1)
        a = fs_wrapInQuotes(d)
        AssertEqualWithMessage e, a, "i=" & i
    Next
End Sub

'###################################################################################################
'fs_writeFile()
Sub Test_fs_writeFile
    'データ定義と生成
    Dim cont
    cont = "abc" & vbNewLine & "あいう" & vbNewLine & "123" & vbNewLine & ChrW(12316) 'ChrW(12316)='\u301c'（波ダッシュ・波型）Sjisに変換できない文字
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile", False, Empty))
    
    '実行
    Dim p : p = d.Item("target").Item("path")
    Dim a : Set a = fs_writeFile(p, cont)

    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"
    
    'データの検証
    a = readTestFile("Unicode",p)
    e = cont
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_fs_writeFile_Rewrite
    'データ定義と生成
    Dim cont
    cont = "For" & vbNewLine & "Test_fs_writeFile_Rewrite"
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile", True, cont))
    
    '実行
    Dim p : p = d.Item("target").Item("path")
    cont = "abc" & vbNewLine & "①②③" & vbNewLine & "!#$" & vbNewLine & ChrW(12316) 'ChrW(12316)='\u301c'（波ダッシュ・波型）Sjisに変換できない文字
    Dim a : Set a = fs_writeFile(p, cont)

    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"
    
    'データの検証
    a = readTestFile("Unicode",p)
    e = cont
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_fs_writeFile_Err
    'データ定義と生成
    Dim before : before = "Test_fs_writeFile_Err" & vbNewLine & "Before"
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile", True, before))
    
    '上書きするファイル（target）をロックする
    Dim p : p = d.Item("target").Item("path")
    With lockFile(p)
        '実行
        Dim cont : cont = "Test_fs_writeFile_Err"
        Dim a : Set a = fs_writeFile(p, cont)
    
        '戻り値の検証
        Dim e
        e = False
        AssertEqualWithMessage e, a, "ret"
        e = True
        AssertEqualWithMessage e, a.isErr, "isErr"
        e = 505
        AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
        e = "不正な参照です。"
        AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

        .Close
    End With
    
    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("target",d))
    a = readTestFile("Unicode",p)
    e = before
    AssertEqualWithMessage e, a, "cont"
End Sub

'###################################################################################################
'fs_writeFileDefault()
Sub Test_fs_writeFileDefault
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile", False, Empty))
    
    '実行
    Dim p : p = d.Item("target").Item("path")
    Dim cont : cont = "Test_fs_writeFileDefault"
    Dim a : Set a = fs_writeFileDefault(p, cont)

    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"
    
    'データの検証
    Dim charset : charset = "shift-jis"
    a = readTestFile(charset,p)
    e = cont
    AssertEqualWithMessage e, a, "cont"
End Sub

'###################################################################################################
'func_FsWriteFile()
Sub Test_func_FsWriteFile_Iomode_ForWriting_Normal__Format_SystemDefault
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile", False, Empty))
    
    '実行
    Dim p : p = d.Item("target").Item("path")
    Dim create : create = True
    Dim iomode : iomode = 2          'ForWriting
    Dim format : format = -2         'TristateUseDefault
    Dim cont : cont = "Test_func_FsWriteFile_Iomode_ForWriting_Normal__Format_SystemDefault"
    Dim a : Set a = func_FsWriteFile(p, iomode, create, format, cont)

    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"
    
    'データの検証
    Dim charset : charset = "shift-jis"
    a = readTestFile(charset,p)
    e = cont
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_func_FsWriteFile_Iomode_ForWriting_Rewrite__Format_Unicode
    'データ定義と生成
    Dim before : before = "Test_func_FsWriteFile_Iomode_ForWriting_Rewrite__Format_Unicode" & vbNewLine & "Before"
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile", True, before))
    
    '実行
    Dim p : p = d.Item("target").Item("path")
    Dim create : create = True
    Dim iomode : iomode = 2          'ForWriting
    Dim format : format = -1         'TristateTrue(Unicode)
    Dim cont : cont = "Test_func_FsWriteFile_Iomode_ForWriting_Rewrite__Format_Unicode"
    Dim a : Set a = func_FsWriteFile(p, iomode, create, format, cont)

    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"
    
    'データの検証
    Dim charset : charset = "Unicode"
    a = readTestFile(charset,p)
    e = cont
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_func_FsWriteFile_Iomode_ForAppending_Normal__Format_Ascii
    'データ定義と生成
    Dim before : before = "Test_func_FsWriteFile_Iomode_ForAppending_Normal__Format_Ascii" & vbNewLine & "Before"
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile(Ascii)", True, before))
    
    '実行
    Dim p : p = d.Item("target").Item("path")
    Dim create : create = True
    Dim iomode : iomode = 8          'ForAppending
    Dim format : format = 0          'TristateFalse(Ascii)
    Dim cont : cont = "Test_func_FsWriteFile_Iomode_ForAppending_Normal__Format_Ascii"
    Dim a : Set a = func_FsWriteFile(p, iomode, create, format, cont)

    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"
    
    'データの検証
    Dim charset : charset = "shift-jis"
    a = readTestFile(charset,p)
    e = before&cont
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_func_FsWriteFile_Iomode_ForAppending_Append__Format_SystemDefault
    'データ定義と生成
    Dim before : before = "Test_func_FsWriteFile_Iomode_ForAppending_Append__Format_SystemDefault" & vbNewLine & "Before"
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile(Ascii)", True, before))
    
    '実行
    Dim p : p = d.Item("target").Item("path")
    Dim create : create = True
    Dim iomode : iomode = 8          'ForAppending
    Dim format : format = -2         'TristateUseDefault
    Dim cont : cont = "Test_func_FsWriteFile_Iomode_ForAppending_Append__Format_SystemDefault"
    Dim a : Set a = func_FsWriteFile(p, iomode, create, format, cont)

    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"
    
    'データの検証
    Dim charset : charset = "shift-jis"
    a = readTestFile(charset,p)
    e = before&cont
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_func_FsWriteFile_Create_True_Normal__Format_Unicode
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile", False, Empty))
    
    '実行
    Dim p : p = d.Item("target").Item("path")
    Dim create : create = True
    Dim iomode : iomode = 2          'ForWriting
    Dim format : format = -1         'TristateTrue(Unicode)
    Dim cont : cont = "Test_func_FsWriteFile_Create_True_Normal__Format_Unicode"
    Dim a : Set a = func_FsWriteFile(p, iomode, create, format, cont)

    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"
    
    'データの検証
    Dim charset : charset = "Unicode"
    a = readTestFile(charset,p)
    e = cont
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_func_FsWriteFile_Create_True_Rewrite__Format_Ascii
    'データ定義と生成
    Dim before : before = "Test_func_FsWriteFile_Create_True_Rewrite__Format_Ascii" & vbNewLine & "Before"
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile(Ascii)", True, before))
    
    '実行
    Dim p : p = d.Item("target").Item("path")
    Dim create : create = True
    Dim iomode : iomode = 2          'ForWriting
    Dim format : format = 0          'TristateFalse(Ascii)
    Dim cont : cont = "Test_func_FsWriteFile_Create_True_Rewrite__Format_Ascii"
    Dim a : Set a = func_FsWriteFile(p, iomode, create, format, cont)

    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"
    
    'データの検証
    Dim charset : charset = "shift-jis"
    a = readTestFile(charset,p)
    e = cont
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_func_FsWriteFile_Create_False_Err
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile", False, Empty))
    
    '実行
    Dim p : p = d.Item("target").Item("path")
    Dim create : create = False
    Dim iomode : iomode = 2          'ForWriting
    Dim format : format = -1         'TristateTrue(Unicode)
    Dim cont : cont = "Test_func_FsWriteFile_Create_False_Err"
    Dim a : Set a = func_FsWriteFile(p, iomode, create, format, cont)

    '戻り値の検証
    Dim e
    e = False
    AssertEqualWithMessage e, a, "ret"
    e = True
    AssertEqualWithMessage e, a.isErr, "isErr"
    e = 505
    AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
    e = "不正な参照です。"
    AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("target",d))
End Sub
Sub Test_func_FsWriteFile_Create_False_Rewrite__Format_Unicode
    'データ定義と生成
    Dim before : before = "Test_func_FsWriteFile_Create_False_Rewrite__Format_Unicode" & vbNewLine & "Before"
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile", True, before))
    
    '実行
    Dim p : p = d.Item("target").Item("path")
    Dim create : create = False
    Dim iomode : iomode = 2          'ForWriting
    Dim format : format = -1         'TristateTrue(Unicode)
    Dim cont : cont = "Test_func_FsWriteFile_Create_False_Rewrite__Format_Unicode"
    Dim a : Set a = func_FsWriteFile(p, iomode, create, format, cont)

    '戻り値の検証
    Dim e
    e = True
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"
    
    'データの検証
    Dim charset : charset = "Unicode"
    a = readTestFile(charset,p)
    e = cont
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_func_FsWriteFile_Err_FileLocked
    'データ定義と生成
    Dim before : before = "Test_func_FsWriteFile_Err_FileLocked" & vbNewLine & "Before"
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile(Ascii)", True, before))
    
    '上書きするファイル（target）をロックする
    Dim p : p = d.Item("target").Item("path")
    With lockFile(p)
        
        '実行
        Dim create : create = False
        Dim iomode : iomode = 2          'ForWriting
        Dim format : format = -1         'TristateTrue(Unicode)
        Dim cont : cont = "Test_func_FsWriteFile_Err_FileLocked"
        Dim a : Set a = func_FsWriteFile(p, iomode, create, format, cont)
    
        '戻り値の検証
        Dim e
        e = False
        AssertEqualWithMessage e, a, "ret"
        e = True
        AssertEqualWithMessage e, a.isErr, "isErr"
        e = 505
        AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
        e = "不正な参照です。"
        AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

        .Close
    End With
    
    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("target",d))
    Dim charset : charset = "shift-jis"
    a = readTestFile(charset,p)
    e = before
    AssertEqualWithMessage e, a, "cont"
End Sub

'###################################################################################################
'func_FsReadFile()
Sub Test_func_FsReadFile_Normal__Format_SystemDefault
    'データ定義と生成
    Dim cont : cont = "Test_func_FsReadFile_Normal__Format_SystemDefault"
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile(Ascii)", True, cont))
    
    '実行
    Dim format : format = -2         'TristateUseDefault
    Dim a : Set a = func_FsReadFile(d.Item("target").Item("path"), format)

    '戻り値の検証
    Dim e
    e = cont
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("target",d))
End Sub
Sub Test_func_FsReadFile_Normal__Format_Unicode
    'データ定義と生成
    Dim cont : cont = "Test_func_FsReadFile_Normal__Format_Unicode"
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile", True, cont))
    
    '実行
    Dim format : format = -1         'TristateTrue(Unicode)
    Dim a : Set a = func_FsReadFile(d.Item("target").Item("path"), format)

    '戻り値の検証
    Dim e
    e = cont
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("target",d))
End Sub
Sub Test_func_FsReadFile_Normal__Format_Ascii
    'データ定義と生成
    Dim cont : cont = "Test_func_FsReadFile_Normal__Format_Ascii"
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile(Ascii)", True, cont))
    
    '実行
    Dim format : format = 0          'TristateFalse(Ascii)
    Dim a : Set a = func_FsReadFile(d.Item("target").Item("path"), format)

    '戻り値の検証
    Dim e
    e = cont
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("target",d))
End Sub
Sub Test_func_FsReadFile_Err
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForReadWriteFile("textfile", False, Empty))
    
    '実行
    Dim format : format = -2         'TristateUseDefault
    Dim a : Set a = func_FsReadFile(d.Item("target").Item("path"), format)

    '戻り値の検証
    Dim e
    e = Empty
    AssertEqualWithMessage e, a, "ret"
    e = True
    AssertEqualWithMessage e, a.isErr, "isErr"
    e = 505
    AssertEqualWithMessage e, a.getErr.Item("Number"), "getErr.Item('Number')"
    e = "不正な参照です。"
    AssertEqualWithMessage e, a.getErr.Item("Description"), "getErr.Item('Description')"

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("target",d))
End Sub


'###################################################################################################
'common
Function readTestFile(c,p)
    With adodb
        .Charset = c
        .Open
        .LoadFromFile p
        readTestFile = .ReadText
        .Close
    End With
End Function
Function lockFile(p)
    Set lockFile = fso.OpenTextFile(p, 8, True, -1)
End Function
Function getTempName()
    getTempName = fso.GetTempName()
End Function
Function buildPath(pf,p)
    buildPath = fso.BuildPath(pf, p)
End Function
Function getFileName(p)
    getFileName = fso.GetFileName(p)
End Function
Sub createTextFileInUnicode(p,c)
    createTextFile p,c,-1
End Sub
Sub createTextFileInAscii(p,c)
    createTextFile p,c,0
End Sub
Sub createTextFile(p,c,f)
    With fso.OpenTextFile(p, 2, True, f)
        .Write c
        .Close
    End With
End Sub

'TestItem作成
Function createTestItems(a)
    '定義生成
    Dim i,d : Set d = CreateObject("Scripting.Dictionary")
    For Each i In a
        d.Add i(0), defineTestItem(i)
    Next
    'item作成
    For Each i In d.Items
        createTestItem i
    Next
    'info取得
    For Each i In d.Items
        getTestItemInfo i
    Next
    
    Set createTestItems = d
End Function
Function defineTestItem(a)
    Dim data : Set data = CreateObject("Scripting.Dictionary")
    With data
        .Add "type", a(1)
        .Add "parentFolder", a(2)
        .Add "relativePath", a(3)
        .Add "name", getFileName(buildPath(a(2),a(3)))
        .Add "path",  buildPath(PsPathTempFolder, buildPath(a(2),a(3)))
        .Add "isSetup", a(4)
        .Add "cont", a(5)
    End With
    Set defineTestItem = data
End Function
Sub createTestItem(i)
    With i
        If .Item("isSetup") Then
            Dim p : p = .Item("path")
            If .Item("type")="textfile" Then
                createTextFileInUnicode p, .Item("cont")
            End If
            If .Item("type")="textfile(Ascii)" Then
                createTextFileInAscii p, .Item("cont")
            End If
            If .Item("type")="folder" Then
                fso.CreateFolder p
            End If
        End If
    End With
End Sub
Sub getTestItemInfo(i)
    With i
        If .Item("isSetup") Then
            Dim p : p = .Item("path")
            If .Item("type")="textfile" Or .Item("type")="textfile(Ascii)" Then
                .Add "size", fso.GetFile(p).Size
                .Add "DateLastModified", fso.GetFile(p).DateLastModified
            End If
            If .Item("type")="folder" Then
                .Add "size", fso.GetFolder(p).Size
                .Add "DateLastModified", fso.GetFolder(p).DateLastModified
                .Add "Files.Count", fso.GetFolder(p).Files.Count
                .Add "SubFolders.Count", fso.GetFolder(p).SubFolders.Count
            End If
        End If
    End With
End Sub

'★For Test_fs_getAllFiles_*()
Function createTestItemDefinitionForGetAllFiles(f,t,n)
    Dim Folder : Folder = getTempName()
    createTestItemDefinitionForGetAllFiles = Array( _
        Array(  "folder", "folder"  , fromFolder, vbNullString , True , Empty) _
        , Array("folder", "textfile", fromFolder, getTempName(), False, Empty) _
        )
End Function

'For Test_fs_copyFile_*(),Test_fs_moveFile_*()
Function createTestItemDefinitionForFile(f,t,n)
    Dim fromFolder : fromFolder = getTempName()
    Dim toFolder : toFolder = getTempName()
    Dim ret : ret = Array( _
        Array(  "from-folder", "folder"  , fromFolder, vbNullString , True , Empty) _
        , Array("from"       , "textfile", fromFolder, getTempName(), False, Empty) _
        , Array("to-folder"  , "folder"  , toFolder  , vbNullString , True , Empty) _
        , Array("to"         , "textfile", toFolder  , getTempName(), False, Empty) _
        )
    If f Then ret(1) = _
          Array("from"       , "textfile", fromFolder, getTempName(), True, "For " & n & " as FromFile.")
    If t Then ret(3) = _
          Array("to"         , "textfile", toFolder  , getTempName(), True, "For " & n & " as ToFile.")
    createTestItemDefinitionForFile = ret
End Function
'For Test_fs_copyFolder_*(),Test_fs_moveFolder_*()
Function createTestItemDefinitionForFolder(f,t,n)
    Dim fromFolder : fromFolder = getTempName()
    Dim overRideFile : overRideFile = getTempName()
    Dim overRideFolder : overRideFolder = getTempName()
    Dim fa
    If f Then
        fa = Array( _
            Array(  "from-folder" , "folder"  , fromFolder, vbNullString                            , True , Empty) _
            , Array("from-fl1"    , "textfile", fromFolder, getTempName()                           , True , "For " & n & " as from-fl1.") _
            , Array("from-fl2"    , "textfile", fromFolder, overRideFile                            , True , "For " & n & " as from-file2.") _
            , Array("from-fd1"    , "folder"  , fromFolder, overRideFolder                          , True , Empty) _
            , Array("from-fd1-fl1", "textfile", fromFolder, buildPath(overRideFolder, getTempName()), True , "For " & n & " as from-fd1-fl1.") _
            , Array("from-fd2"    , "folder"  , fromFolder, getTempName()                           , True , Empty) _
            )
    Else
        fa = Array( _
            Array(  "from-folder" , "folder"  , fromFolder, vbNullString                            , False, Empty) _
        )
    End If

    Dim toFolder : toFolder = getTempName()
    Dim ta
    If t Then
        ta = Array( _
            Array(  "to-folder"   , "folder"  , toFolder  , vbNullString                            , True, Empty) _
            , Array("to-fla"      , "textfile", toFolder  , getTempName()                           , True, "For " & n & " as to-fl1.") _
            , Array("to-flb"      , "textfile", toFolder  , overRideFile                            , True, "For " & n & " as to-filesecond.") _
            , Array("to-fda"      , "folder"  , toFolder  , overRideFolder                          , True, Empty) _
            , Array("to-fda-fla"  , "textfile", toFolder  , buildPath(overRideFolder, getTempName()), True, "For " & n & " as to-fd1-fileone.") _
            , Array("to-fdb"      , "folder"  , toFolder  , getTempName()                           , True, Empty) _
            )
    Else
        ta = Array( _
            Array(  "to-folder"   , "folder"  , toFolder  , vbNullString                            , False, Empty) _
            )
    End If

    Dim ubfa : ubfa = Ubound(fa)
    Redim ret(ubfa)
    Dim i
    For i=0 To ubfa
        ret(i) = fa(i)
    Next
    Redim Preserve ret(ubfa+Ubound(ta)+1)
    For i=0 To Ubound(ta)
        ret(ubfa+1+i) = ta(i)
    Next
    
    createTestItemDefinitionForFolder = ret
End Function
'For Test_fs_createFolder*()
Function createTestItemDefinitionForCreateFolder(cfl,cfd,n)
    Dim tp : tp = "folder"
    If cfl Then tp = "textfile"
    Dim flg : flg = False
    If cfl Or cfd Then flg = True
    createTestItemDefinitionForCreateFolder = Array( _
        Array(  "target", tp, getTempName(), vbNullString , flg, n) _
        )
End Function
'For Test_fs_deleteFile*()
Function createTestItemDefinitionForDeleteFile(f,n)
    Dim targetFolder : targetFolder = getTempName()
    createTestItemDefinitionForDeleteFile = Array( _
        Array(  "target-folder", "folder"  , targetFolder, vbNullString , True, Empty) _
        , Array("target"       , "textfile", targetFolder, getTempName(), f   , n) _
        )
End Function
'For Test_fs_deleteFolder*()
Function createTestItemDefinitionForDeleteFolder(f,n)
    Dim rootFolder : rootFolder = getTempName()
    Dim targetFolder : targetFolder = getTempName()
    createTestItemDefinitionForDeleteFolder = Array( _
        Array(  "target-folder", "folder"  , rootFolder, vbNullString                          , True, Empty) _
        , Array("target"       , "folder"  , rootFolder, targetFolder                          , f   , Empty) _
        , Array("target-file"  , "textfile", rootFolder, buildPath(targetFolder, getTempName()), f   , n) _
        )
End Function
'For Test_fs_readFile*(),Test_fs_writeFile*()
Function createTestItemDefinitionForReadWriteFile(t,f,n)
    Dim targetFolder : targetFolder = getTempName()
    createTestItemDefinitionForReadWriteFile = Array( _
        Array(  "target-folder", "folder", targetFolder, vbNullString , True, Empty) _
        , Array("target"       , t       , targetFolder, getTempName(), f   , n) _
        )
End Function

'検証
Sub assertFolderItems(a)
    '定義生成
    Dim i,d : Set d = CreateObject("Scripting.Dictionary")
    For Each i In a
        d.Add i(0), defineAssertItem(i)
    Next
    '検証
    Dim p
    For Each i In d.Keys
        With d.Item(i)
            p = buildPath(.Item("parentFolder"), .Item("relativePath"))
            If .Item("type")="textfile" Then
                assertFile i, .Item("expect"), p
            End If
            If .Item("type")="folder" Then
                assertFolder i, .Item("expect"), p
            End If
            If .Item("type")="NotExistsFile" Then
                AssertEqualWithMessage False, fso.FileExists(p), i&"-"&.Item("type")
            End If
            If .Item("type")="NotExistsFolder" Then
                AssertEqualWithMessage False, fso.FolderExists(p), i&"-"&.Item("type")
            End If
        End With
    Next
End Sub
Function defineAssertItem(a)
    Dim data : Set data = CreateObject("Scripting.Dictionary")
    With data
        .Add "type",  a(1)
        .Add "expect",  a(2)
        .Add "parentFolder", a(3)
        .Add "relativePath", a(4)
    End With
    Set defineAssertItem = data
End Function
Sub assertFile(n,d,p)
    Dim e,a,i
    i = "name"
    If d.Exists(i) Then
        e = d.Item(i)
        a = fso.GetFile(p).Name
        AssertEqualWithMessage e, a, n&"-"&i
    End If
    i = "size"
    If d.Exists(i) Then
        e = d.Item(i)
        a = fso.GetFile(p).Size
        AssertEqualWithMessage e, a, n&"-"&i
    End If
    i = "DateLastModified"
    If d.Exists(i) Then
        e = d.Item(i)
        a = fso.GetFile(p).DateLastModified
        AssertEqualWithMessage e, a, n&"-"&i
    End If
End Sub
Sub assertFolder(n,d,p)
    Dim e,a,i
    i = "name"
    If d.Exists(i) Then
        e = d.Item(i)
        a = fso.GetFolder(p).Name
        AssertEqualWithMessage e, a, n&"-"&i
    End If
    i = "size"
    If d.Exists(i) Then
        e = d.Item(i)
        a = fso.GetFolder(p).Size
        AssertEqualWithMessage e, a, n&"-"&i
    End If
    i = "DateLastModified"
    If d.Exists(i) Then
        e = d.Item(i)
        a = fso.GetFolder(p).DateLastModified
        AssertEqualWithMessage e, a, n&"-"&i
    End If
    i = "Files.Count"
    If d.Exists(i) Then
        e = d.Item(i)
        a = fso.GetFolder(p).Files.Count
        AssertEqualWithMessage e, a, n&"-"&i
    End If
    i = "SubFolders.Count"
    If d.Exists(i) Then
        e = d.Item(i)
        a = fso.GetFolder(p).SubFolders.Count
        AssertEqualWithMessage e, a, n&"-"&i
    End If
End Sub

Function createExpectDefinitionUnchange(kd,d)
    Dim i,k,t,p : i=0
    Redim a(d.Count-1)
    For Each k In d.Keys
        If StrComp(kd,Mid(k,1,Len(kd)),vbBinaryCompare)=0 Then
            If d.Exists(kd&"-folder") Then p = d.Item(kd&"-folder").Item("path") Else p = d.Item(kd).Item("path")
            If d.Item(k).Item("isSetup") Then
                a(i) = Array(k, d.Item(k).Item("type"), d.Item(k), p, d.Item(k).Item("relativePath"))
            Else
                If StrComp(d.Item(k).Item("type"),"folder")=0 Then t="NotExistsFolder" Else t="NotExistsFile"
                a(i) = Array(k, t                     , Empty    , p, d.Item(k).Item("relativePath"))
            End If
            i = i + 1
        End If
    Next
    If i>0 Then Redim Preserve a(i-1)
    createExpectDefinitionUnchange = a
End Function
'For Test_fs_copyFile_*(),Test_fs_moveFile_*()
Function createExpectDefinitionMergeFile(d)
    Dim expToFolder : Set expToFolder = exclusionItem(d.Item("from-folder"), Array("DateLastModified")) : expToFolder.Item("name") = d.Item("to-folder").Item("name")
    Dim expTo : Set expTo = exclusionItem(d.Item("from"), Array()) : expTo.Item("name") = d.Item("to").Item("name")
    Dim ret : ret = Array( _
        Array(  "to-folder"  , "folder"       , expToFolder          , d.Item("to-folder").Item("path")  , d.Item("to-folder").Item("relativePath")) _
        , Array("to"         , "textfile"     , expTo                , d.Item("to-folder").Item("path")  , d.Item("to").Item("relativePath")) _
        )
    createExpectDefinitionMergeFile = ret
End Function
'For fs_createFolder*()
Function createExpectDefinitionCreateFolder(d)
    Dim exp : Set exp = CreateObject("Scripting.Dictionary")
    With exp
        .Add "name", d.Item("target").Item("name")
        .Add "size", 0
        .Add "Files.Countme", 0
        .Add "SubFolders.Count", 0
    End With
    createExpectDefinitionCreateFolder = Array( _
        Array(  "target", "folder", exp, d.Item("target").Item("path"), d.Item("target").Item("relativePath")) _
        )
End Function
'For Test_fs_moveFile_*(),Test_fs_deleteFile*()
Function createExpectDefinitionDisappearFile(k,d)
    Dim expFromFolder : Set expFromFolder = exclusionItem(d.Item(k&"-folder"), Array("DateLastModified"))
    With expFromFolder
        .Item("size") = 0
        .Item("Files.Count") = 0
    End With
    Dim ret : ret = Array( _
        Array(  k&"-folder", "folder"         , expFromFolder, d.Item(k&"-folder").Item("path"), d.Item(k&"-folder").Item("relativePath")) _
        , Array(k          , "NotExistsFile"  , Empty        , d.Item(k&"-folder").Item("path"), d.Item(k).Item("relativePath")) _
        )
    createExpectDefinitionDisappearFile = ret
End Function
'For Test_fs_moveFolder_*(),Test_fs_deleteFolder*()
Function createExpectDefinitionDisappearFolder(k,d)
    createExpectDefinitionDisappearFolder = Array( _
        Array(  k          , "NotExistsFolder", Empty        , d.Item(k).Item("path")          , d.Item(k).Item("relativePath")) _
        )
End Function
'For Test_fs_copyFolder_*(),Test_fs_moveFolder_*()
Function createExpectDefinitionMergeFolder(d)
    '全てのfromの情報で期待値を上書きする
    Dim f : f = createExpectDefinitionUnchange("from",d)
    createExpectDefinitionMergeFolder = createExpectDefinitionMergeFolderProc(d,f)
End Function
'For Test_fs_copyFolder_*()
Function createExpectDefinitionMergeFolderUntilOverRideFile(d,rp)
    '期待値を上書きするfromの情報は指定したrpまで
    Dim f : f = createExpectDefinitionUnchange("from",d)
    Dim i,p
    For i=0 to Ubound(f)
       p = f(i)(4)
       If StrComp(rp,p,vbBinaryCompare)=0 Then Exit For
    Next
    Dim ret
    If i=0 Then
        ret = Array()
    Else
        Redim Preserve f(i-1)
        ret = f
    End If
    createExpectDefinitionMergeFolderUntilOverRideFile = f
End Function
Function createExpectDefinitionMergeFolderProc(d,f)
    'toの期待値をベースにする
    Dim exps : exps = createExpectDefinitionUnchange("to",d)
    Dim toFp : toFp = exps(0)(3)

    'toのrelativePathとindexをrpに取得
    Dim rps : Set rps = CreateObject("Scripting.Dictionary")
    Dim i,rp,ele
    For i=0 To Ubound(exps)
        ele = exps(i)
        rp = ele(4)
        If Not rps.Exists(rp) Then rps.Add rp, i
    Next

    '引数fで期待値を上書きする
    For i=0 To Ubound(f)
        ele = f(i)
        rp = ele(4)

        'フォルダパスをtoフォルダに書き換える
        ele(3) = toFp
        If Len(rp)=0 Then
        'toフォルダ自身
            'fromフォルダの期待値の"name"をtoフォルダ名で上書きする
            ele(2).Item("name") = getFileName(toFp)
        End If

        If rps.Exists(rp) Then
        'relativePathがtoの期待値にある場合は上書き
            exps(rps.Item(rp)) = ele
        Else
        'relativePathがtoの期待値にない場合は追加
            Redim Preserve exps(Ubound(exps)+1)
            exps(Ubound(exps)) = ele
            rps.Add rp, Ubound(exps)
        End If
    Next

    'folderの期待値を書き換える
    Dim exp,j
    For i=0 To Ubound(exps)
        ele = exps(i)
        Set exp = ele(2)
        If StrComp(ele(1),"folder",vbBinaryCompare)=0 Then
        'folderの場合は期待値を修正する
            'DateLastModifiedはクリアする
            Set exp = exclusionItem(exp, Array("DateLastModified"))
            '属性の値を再集計して設定
            With createExpectDefinitionMergeFolderProcAggregate(exps,ele(4))
                For Each j In Array("size","Files.Count","SubFolders.Count")
                    exp.Item(j) = .Item(j)
                Next
            End With
            Set exps(i)(2) = exp
        End If
    Next

    createExpectDefinitionMergeFolderProc = exps
End Function
Function createExpectDefinitionMergeFolderProcAggregate(exps,rp)
    'rp以下のアイテムをexpsから取得し集計する
    Dim sz,flc,fdc : sz=0 : flc=0 : fdc=0
    Dim i,p,t,e
    For Each i In exps
        t=i(1) : Set e=i(2) : p=i(4)
        If StrComp(rp,Mid(p,1,Len(rp)),vbBinaryCompare)=0 Then
            If StrComp(rp,p,vbBinaryCompare)<>0 Then
                'サイズはフォルダ以下のファイルを全て集計する
                If StrComp("folder",t,vbBinaryCompare)<>0 Then sz=sz+e.Item("size")
                'フォルダとファイル数は直下のアイテムだけカウントする
                If InStr(1,Mid(p,Len(rp)+2,Len(p)),"\",vbBinaryCompare)=0 Then
                    If StrComp("folder",t,vbBinaryCompare)=0 Then
                        fdc = fdc + 1
                    Else
                        flc = flc + 1
                    End If
                End If
            End If
        End If
    Next
    '戻り値を返却する
    Dim ret : Set ret = CreateObject("Scripting.Dictionary")
    With ret
        .Add "size", sz
        .Add "Files.Count", flc
        .Add "SubFolders.Count", fdc
    End With
    Set createExpectDefinitionMergeFolderProcAggregate = ret
End Function
Function exclusionItem(o,e)
    Dim ret : Set ret = CreateObject("Scripting.Dictionary")
    Dim i
    For Each i In o.Keys()
        If inList(e,i)=False Then
           ret.add i, o.Item(i)
        End If
    Next
    Set exclusionItem = ret
End Function
Function inList(a,s)
    inList = False
    Dim i
    For Each i In a
        If i=s Then
            inList = True
            Exit Function
        End If
    Next
End Function

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
