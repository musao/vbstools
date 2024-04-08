' libCom.vbs: fs_* procedure test.
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

Const MY_NAME = "test_libCom_fs.vbs"
Dim PsPathTempFolder
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

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
    e = 0
    AssertEqualWithMessage e, Err.Number, "Err.Number"

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
    e = 0
    AssertEqualWithMessage e, Err.Number, "Err.Number"

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
        e = 0
        AssertEqualWithMessage e, Err.Number, "Err.Number"

        .Close
    End With

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionMergeFile(d))
End Sub
Sub Test_fs_copyFile_Err_FromFileNoExists
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFile(False,False,"Test_fs_copyFile_Err_FromFileNoExists"))

    '実行
    Dim a : Set a = fs_copyFile(d.Item("from").Item("path"),d.Item("to").Item("path"))
    '戻り値の検証
    Dim e
    e = False
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"
    e = 0
    AssertEqualWithMessage e, Err.Number, "Err.Number"

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
        e = 0
        AssertEqualWithMessage e, Err.Number, "Err.Number"

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
    e = 0
    AssertEqualWithMessage e, Err.Number, "Err.Number"
    
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
    e = 0
    AssertEqualWithMessage e, Err.Number, "Err.Number"
    
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
        e = 0
        AssertEqualWithMessage e, Err.Number, "Err.Number"

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
        e = 0
        AssertEqualWithMessage e, Err.Number, "Err.Number"

        .Close
    End With
    
    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionMergeFolder(d))
End Sub
Sub Test_fs_copyFolder_Err_FromFileNoExists
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFolder(False,False,"Test_fs_copyFolder_Err_FromFileNoExists"))

    '実行
    Dim a : Set a = fs_copyFolder(d.Item("from-folder").Item("path"),d.Item("to-folder").Item("path"))

    '戻り値の検証
    Dim e
    e = False
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"
    e = 0
    AssertEqualWithMessage e, Err.Number, "Err.Number"
    
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
        e = 0
        AssertEqualWithMessage e, Err.Number, "Err.Number"

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
    Dim p
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    assertExistsFolder p, False, "before", "createfolder", "folder"
    
    Dim ao,e
    e = True
    Set ao = fs_createFolder(p)
    AssertEqualWithMessage e, ao, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    Dim a
    e = False
    a = ao.isErr()
    AssertEqualWithMessage e, a, "isErr()"

    assertExistsFolder p, True, "after", "createfolder", "folder"
End Sub
Sub Test_fs_createFolder_ErrExistsFile
    Dim p
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))

    Dim c,d
    'ファイルを作成
    c = "UTF-8"
    d = "For" & vbNewLine & "CreateFolder Err-ExistsFile"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "createfolder", "file"
    
    Dim ao,e
    e = False
    Set ao = fs_createFolder(p)
    AssertEqualWithMessage e, ao, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    Dim a
    e = True
    a = ao.isErr()
    AssertEqualWithMessage e, a, "isErr()"

    e = 58
    a = ao.getErr().Item("Number")
    AssertEqualWithMessage e, a, "getErr().Item('Number')"

    assertExistsFile p, True, "after", "createfolder", "file"
    assertExistsFolder p, False, "after", "createfolder", "folder"
End Sub
Sub Test_fs_createFolder_ErrExistsFolder
    Dim p
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))

    Dim c,d
    'フォルダを作成
    fso.CreateFolder p
    assertExistsFolder p, True, "before", "createfolder", "folder"
    
    Dim ao,e
    e = False
    Set ao = fs_createFolder(p)
    AssertEqualWithMessage e, ao, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    Dim a
    e = False
    a = ao.isErr()
    AssertEqualWithMessage e, a, "isErr()"

    assertExistsFolder p, True, "after", "createfolder", "folder"
End Sub

'###################################################################################################
'fs_deleteFile()
Sub Test_fs_deleteFile
    Dim c,p,d
    'ファイルを作成
    c = "UTF-8"
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    d = "For" & vbNewLine & "DeleteFile Normal"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "deleteFile", "file"

    Dim e,ao
    e = True
    Set ao = fs_deleteFile(p)
    AssertEqualWithMessage e, ao, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    Dim a
    e = False
    a = ao.isErr()
    AssertEqualWithMessage e, a, "isErr()"

    assertExistsFile p, False, "after", "deleteFile", "file"
End Sub
Sub Test_fs_deleteFile_Err_NotExists
    Dim p
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "deleteFile", "file"

    Dim e,ao
    e = False
    Set ao = fs_deleteFile(p)
    AssertEqualWithMessage e, ao, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    Dim a
    e = False
    a = ao.isErr()
    AssertEqualWithMessage e, a, "isErr()"

    assertExistsFile p, False, "after", "deleteFile", "file"
End Sub
Sub Test_fs_deleteFile_Err_FileLocked
    Dim c,p,d
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "For" & vbNewLine & "DeleteFile Err FileLocked"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "deleteFile", "file"

    'ファイルをロックする
    With lockFile(p)
        Dim e,ao
        e = False
        Set ao = fs_deleteFile(p)
        
        'fs_deleteFile()がエラーになることを確認する
        AssertEqualWithMessage e, ao, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    Dim a
    e = True
    a = ao.isErr()
    AssertEqualWithMessage e, a, "isErr()"

    e = 70
    a = ao.getErr().Item("Number")
    AssertEqualWithMessage e, a, "getErr().Item('Number')"

    'ファイルが削除されていないことを確認
    assertExistsFile p, True, "after", "deleteFile", "file"
End Sub

'###################################################################################################
'fs_deleteFolder()
Sub Test_fs_deleteFolder
    Dim c,p,pf,d
    'フォルダを作成
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    fso.CreateFolder p
    'フォルダの下にファイルを作成
    c = "UTF-8"
    pf = fso.BuildPath(p, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    d = "For" & vbNewLine & "DeleteFolder Normal"
    writeTestFile c,pf,d
    assertExistsFolder p, True, "before", "deleteFolder", "folder"

    Dim e,ao
    e = True
    Set ao = fs_deleteFolder(p)
    AssertEqualWithMessage e, ao, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    Dim a
    e = False
    a = ao.isErr()
    AssertEqualWithMessage e, a, "isErr()"

    assertExistsFolder p, False, "after", "deleteFolder", "folder"
End Sub
Sub Test_fs_deleteFolder_Err_NotExists
    Dim p
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    assertExistsFolder p, False, "before", "deleteFolder", "folder"

    Dim e,ao
    e = False
    Set ao = fs_deleteFolder(p)
    AssertEqualWithMessage e, ao, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    Dim a
    e = False
    a = ao.isErr()
    AssertEqualWithMessage e, a, "isErr()"

    assertExistsFolder p, False, "after", "deleteFolder", "folder"
End Sub
Sub Test_fs_deleteFolder_Err_FileLocked
    Dim c,p,pf,d
    'フォルダを作成
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    fso.CreateFolder p
    'フォルダの下にファイルを作成
    c = "UTF-8"
    pf = fso.BuildPath(p, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    d = "For" & vbNewLine & "DeleteFolder Err FileLocked"
    writeTestFile c,pf,d

    'ファイルをロックする
    With lockFile(pf)
        Dim e,ao
        e = False
        Set ao = fs_deleteFolder(p)
        
        'fs_deleteFolder()がエラーになることを確認する
        AssertEqualWithMessage e, ao, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    Dim a
    e = True
    a = ao.isErr()
    AssertEqualWithMessage e, a, "isErr()"

    e = 70
    a = ao.getErr().Item("Number")
    AssertEqualWithMessage e, a, "getErr().Item('Number')"

    'フォルダが削除されていないことを確認
    assertExistsFolder p, True, "after", "deleteFolder", "folder"
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
    e = 0
    AssertEqualWithMessage e, Err.Number, "Err.Number"

    'データの検証
    assertFolderItems(createExpectDefinitionDisappearFromFile(d))
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
    e = 0
    AssertEqualWithMessage e, Err.Number, "Err.Number"

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
        e = 0
        AssertEqualWithMessage e, Err.Number, "Err.Number"

        .Close
    End With

    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionUnchange("to",d))
End Sub
Sub Test_fs_moveFile_Err_FromFileNoExists
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFile(False,False,"Test_fs_moveFile_Err_FromFileNoExists"))

    '実行
    Dim a : Set a = fs_moveFile(d.Item("from").Item("path"),d.Item("to").Item("path"))
    '戻り値の検証
    Dim e
    e = False
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"
    e = 0
    AssertEqualWithMessage e, Err.Number, "Err.Number"

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
        e = 0
        AssertEqualWithMessage e, Err.Number, "Err.Number"

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
    e = 0
    AssertEqualWithMessage e, Err.Number, "Err.Number"
    
    'データの検証
    assertFolderItems(createExpectDefinitionDisappearFromFolder(d))
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
    e = 0
    AssertEqualWithMessage e, Err.Number, "Err.Number"
    
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
        e = 0
        AssertEqualWithMessage e, Err.Number, "Err.Number"

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
        e = 0
        AssertEqualWithMessage e, Err.Number, "Err.Number"

        .Close
    End With
    
    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionUnchange("to",d))
End Sub
Sub Test_fs_moveFolder_Err_FromFileNoExists
    'データ定義と生成
    Dim d : Set d = createTestItems(createTestItemDefinitionForFolder(False,False,"Test_fs_moveFolder_Err_FromFileNoExists"))
    
    '実行
    Dim a : Set a = fs_moveFolder(d.Item("from-folder").Item("path"),d.Item("to-folder").Item("path"))

    '戻り値の検証
    Dim e
    e = False
    AssertEqualWithMessage e, a, "ret"
    e = False
    AssertEqualWithMessage e, a.isErr, "isErr"
    e = 0
    AssertEqualWithMessage e, Err.Number, "Err.Number"
    
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
        e = 0
        AssertEqualWithMessage e, Err.Number, "Err.Number"

        .Close
    End With
    
    'データの検証
    assertFolderItems(createExpectDefinitionUnchange("from",d))
    assertFolderItems(createExpectDefinitionUnchange("to",d))
End Sub

'###################################################################################################
'fs_readFile()
Sub Test_fs_readFile
    Dim c,p,d
    'ファイルを作成
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "lmn" & vbNewLine & "ⅢⅥⅩ" & vbNewLine & "ｱｲｳ" & vbNewLine & ChrW(12316) 'ChrW(12316)='\u301c'（波ダッシュ・波型）Sjisに変換できない文字
    writeTestFile c,p,d

    Dim e,a
    e = d
    a = fs_readFile(p)
    AssertEqualWithMessage e, a, "ret"
End Sub
Sub Test_fs_readFile_Err
    Dim p
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "readFile", "file"

    Dim e,a
    e = empty
    a = fs_readFile(p)
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"
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
    Dim p
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "writeFile", "file"

    Dim d,e,a
    d = "abc" & vbNewLine & "あいう" & vbNewLine & "123" & vbNewLine & ChrW(12316) 'ChrW(12316)='\u301c'（波ダッシュ・波型）Sjisに変換できない文字
    e = True
    a = fs_writeFile(p, d)
    AssertEqualWithMessage e, a, "ret"

    Dim c
    c = "Unicode"
    e = d
    a = readTestFile(c, p)
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_fs_writeFile_Rewrite
    Dim p,c,d
    '上書きするファイルを一旦作成
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "UTF-8"
    d = "For" & vbNewLine & "Rewrite"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "writeFile", "file"

    '上書きすることを確認
    d = "abc" & vbNewLine & "①②③" & vbNewLine & "!#$" & vbNewLine & ChrW(12316) 'ChrW(12316)='\u301c'（波ダッシュ・波型）Sjisに変換できない文字
    Dim e,a
    e = True
    a = fs_writeFile(p, d)
    AssertEqualWithMessage e, a, "ret"

    c = "Unicode"
    e = d
    a = readTestFile(c, p)
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_fs_writeFile_Err
    Dim p,c,d
    'ロックするファイルを一旦作成
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "For" & vbNewLine & "Write Error"
    writeTestFile c,p,d

    Dim de
    'ファイルをロックする
    With lockFile(p)
        de = "error" & vbNewLine & "test"
        Dim e,a
        e = False
        a = fs_writeFile(p, de)
        
        'fs_writeFile()がエラーになることを確認する
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    '上書きしていないことを確認
    e = d
    a = readTestFile(c, p)
    AssertEqualWithMessage e, a, "cont"
End Sub

'###################################################################################################
'fs_writeFileDefault()
Sub Test_fs_writeFileDefault
    Dim p
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "writeFileDefault", "file"

    Dim d,e,a
    d = "abc" & vbNewLine & "あいう" & vbNewLine & "123"
    e = True
    a = fs_writeFileDefault(p, d)
    AssertEqualWithMessage e, a, "ret"

    Dim c
    c = "shift-jis"
    e = d
    a = readTestFile(c, p)
    AssertEqualWithMessage e, a, "cont"
End Sub

'###################################################################################################
'func_FsWriteFile()
Sub Test_func_FsWriteFile_Iomode_ForWriting_Normal__Format_SystemDefault
    Dim p
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "func_FsWriteFile", "file"

    Dim d,iomode,create,f,e,a
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Iomode_ForWriting_Normal__Format_SystemDefault"
    iomode = 2     'ForWriting
    create = True
    f = -2         'TristateUseDefault
    e = True
    a = func_FsWriteFile(p, iomode, create, f, d)
    AssertEqualWithMessage e, a, "ret"

    Dim c
    c = "shift-jis"
    e = d
    a = readTestFile(c, p)
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_func_FsWriteFile_Iomode_ForWriting_Rewrite__Format_Unicode
    Dim p,c,d
    '上書きするファイルを一旦作成
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "For" & vbNewLine & "Rewrite"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "func_FsWriteFile", "file"

    Dim iomode,create,f,e,a
    '上書きすることを確認
    iomode = 2     'ForWriting
    create = True
    f = -1    'TristateTrue(Unicode)
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Iomode_ForWriting_Rewrite__Format_Unicode"
    e = True
    a = func_FsWriteFile(p, iomode, create, f, d)
    AssertEqualWithMessage e, a, "ret"

    c = "Unicode"
    e = d
    a = readTestFile(c, p)
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_func_FsWriteFile_Iomode_ForAppending_Normal__Format_Ascii
    Dim p
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "func_FsWriteFile", "file"

    Dim d,iomode,create,f,e,a
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Iomode_ForAppending_Normal__Format_Ascii"
    iomode = 8     'ForAppending
    create = True
    f = 0          'TristateFalse(Ascii)
    e = True
    a = func_FsWriteFile(p, iomode, create, f, d)
    AssertEqualWithMessage e, a, "ret"

    Dim c
    c = "shift-jis"
    e = d
    a = readTestFile(c,p)
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_func_FsWriteFile_Iomode_ForAppending_Append__Format_SystemDefault
    Dim p,c,d
    '追記するファイルを一旦作成
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "shift-jis"
    d = "For" & vbNewLine & "Append"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "func_FsWriteFile", "file"

    Dim iomode,create,f,ec,e,a
    '追記することを確認
    iomode = 8     'ForAppending
    create = True
    f = -2         'TristateUseDefault
    ec = d
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Iomode_ForAppending_Append__Format_SystemDefault"
    e = True
    a = func_FsWriteFile(p, iomode, create, f, d)
    AssertEqualWithMessage e, a, "ret"

    ec = ec & d
    a = readTestFile(c, p)
    AssertEqualWithMessage ec, a, "cont"
End Sub
Sub Test_func_FsWriteFile_Create_True_Normal__Format_Unicode
    Dim p
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "func_FsWriteFile", "file"

    Dim d,iomode,create,f,e,a
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Create_True_Normal__Format_Unicode"
    iomode = 2     'ForWriting
    create = True
    f = -1         'TristateTrue(Unicode)
    e = True
    a = func_FsWriteFile(p, iomode, create, f, d)
    AssertEqualWithMessage e, a, "ret"

    Dim c
    c = "Unicode"
    e = d
    a = readTestFile(c, p)
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_func_FsWriteFile_Create_True_Rewrite__Format_Ascii
    Dim p,c,d
    '上書きするファイルを一旦作成
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "shift-jis"
    d = "For" & vbNewLine & "Rewrite"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "func_FsWriteFile", "file"

    Dim iomode,create,f,e,a
    '上書きすることを確認
    iomode = 2     'ForWriting
    create = True
    f = 0          'TristateFalse(Ascii)
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Create_True_Rewrite__Format_Ascii"
    e = True
    a = func_FsWriteFile(p, iomode, create, f, d)
    AssertEqualWithMessage e, a, "ret"

    e = d
    a = readTestFile(c, p)
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_func_FsWriteFile_Create_False_Err
    Dim p
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "func_FsWriteFile", "file"

    Dim d,iomode,create,f,e,a
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Create_False_Err"
    iomode = 2     'ForWriting
    create = False
    f = -1         'TristateTrue(Unicode)
    e = False
    a = func_FsWriteFile(p, iomode, create, f, d)

    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"
    assertExistsFile p, False, "after", "func_FsWriteFile", "file"
End Sub
Sub Test_func_FsWriteFile_Create_False_Rewrite__Format_Unicode
    Dim p,c,d
    '上書きするファイルを一旦作成
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "For" & vbNewLine & "Rewrite"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "func_FsWriteFile", "file"

    Dim e,a,iomode,create,f
    '上書きすることを確認
    iomode = 2     'ForWriting
    create = False
    f = -1         'TristateTrue(Unicode)
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Create_False_Rewrite__Format_Unicode"
    e = True
    a = func_FsWriteFile(p, iomode, create, f, d)
    AssertEqualWithMessage e, a, "ret"

    e = d
    a = readTestFile(c, p)
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_func_FsWriteFile_Err_FileLocked
    Dim p,d,c
    'ロックするファイルを一旦作成
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
'    c = "shift-jis"
    d = "error" & vbNewLine & "FileLocked"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "func_FsWriteFile", "file"

    'ファイルをロックする
    With lockFile(p)

        Dim iomode,create,f,de,e,a
        iomode = 2     'ForWriting
        create = False
        f = 0          'TristateFalse(Ascii)
        de = "error" & vbNewLine & "test"
        e = False
        a = func_FsWriteFile(p, iomode, create, f, de)
        
        'func_FsWriteFile()がエラーになることを確認する
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    '上書きしていないことを確認
    e = d
    a = readTestFile(c, p)
    AssertEqualWithMessage e, a, "cont"
End Sub

'###################################################################################################
'func_FsReadFile()
Sub Test_func_FsReadFile_Normal__Format_SystemDefault
    Dim p
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))

    Dim d,f,c
    d = "func_FsReadFile" & vbNewLine & "のテスト" & vbNewLine & "Normal__Format_SystemDefault"
    f = -2         'TristateUseDefault
    c = "shift-jis"
    writeTestFile c,p,d

    Dim e,a
    e = d
    a = func_FsReadFile(p,f)
    AssertEqualWithMessage e, a, "ret"
End Sub
Sub Test_func_FsReadFile_Normal__Format_Unicode
    Dim p
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))

    Dim d,f,c
    d = "func_FsReadFile" & vbNewLine & "のテスト" & vbNewLine & "Normal__Format_Unicode"
    f = -1         'TristateTrue(Unicode)
    c = "Unicode"
    writeTestFile c,p,d

    Dim e,a
    e = d
    a = func_FsReadFile(p,f)
    AssertEqualWithMessage e, a, "ret"
End Sub
Sub Test_func_FsReadFile_Normal__Format_Ascii
    Dim p
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))

    Dim d,f,c
    d = "func_FsReadFile" & vbNewLine & "のテスト" & vbNewLine & "Normal__Format_Ascii"
    f = 0          'TristateFalse(Ascii)
    c = "shift-jis"
    writeTestFile c,p,d
    
    Dim e,a
    e = d
    a = func_FsReadFile(p,f)
    AssertEqualWithMessage e, a, "ret"
End Sub
Sub Test_func_FsReadFile_Err
    Dim p
    p = fso.BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "func_FsReadFile", "file"

    Dim f,e,a
    f = -2         'TristateUseDefault
    e = empty
    a = func_FsReadFile(p,f)
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"
End Sub

'###################################################################################################
'common
Sub writeTestFile(c,p,d)
    With CreateObject("ADODB.Stream")
        .Charset = c
        .Open
        .WriteText d, 0
        .SaveToFile p, 2
        .Close
    End With
End Sub
Function readTestFile(c,p)
    With CreateObject("ADODB.Stream")
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
Function getTempPath(pf)
    getTempPath = buildPath(pf, getTempName())
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
Sub createTextFile(p,c)
    With fso.OpenTextFile(p, 2, True, -1)
        .Write c
        .Close
    End With
End Sub
Sub assertExistsFile(path, expect, timestr, action, tgt)
    AssertEqualWithMessage expect, fso.FileExists(path), timestr & " " & action & " " & tgt & " exists"
End Sub
Sub assertExistsFolder(path, expect, timestr, action, tgt)
    AssertEqualWithMessage expect, fso.FolderExists(path), timestr & " " & action & " " & tgt & " exists"
End Sub
Sub assertFilesSubfoldersCount(path, expectfilecnt, expectfoldercnt, tgt)
    AssertEqualWithMessage expectfilecnt, fso.GetFolder(path).Files.Count, tgt & " folderFiles Count"
    AssertEqualWithMessage expectfoldercnt, fso.GetFolder(path).SubFolders.Count, tgt & " folderSubFolders Count"
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
                createTextFile p, .Item("cont")
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
            If .Item("type")="textfile" Then
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

Function createExpectDefinitionMergeFile(d)
    Dim expToFolder : Set expToFolder = exclusionItem(d.Item("from-folder"), Array("DateLastModified")) : expToFolder.Item("name") = d.Item("to-folder").Item("name")
    Dim expTo : Set expTo = exclusionItem(d.Item("from"), Array()) : expTo.Item("name") = d.Item("to").Item("name")
    Dim ret : ret = Array( _
        Array(  "to-folder"  , "folder"       , expToFolder          , d.Item("to-folder").Item("path")  , d.Item("to-folder").Item("relativePath")) _
        , Array("to"         , "textfile"     , expTo                , d.Item("to-folder").Item("path")  , d.Item("to").Item("relativePath")) _
        )
    createExpectDefinitionMergeFile = ret
End Function
Function createExpectDefinitionDisappearFromFile(d)
    Dim expFromFolder : Set expFromFolder = exclusionItem(d.Item("from-folder"), Array("DateLastModified"))
    With expFromFolder
        .Item("size") = 0
        .Item("Files.Count") = 0
    End With
    Dim ret : ret = Array( _
        Array(  "from-folder", "folder"       , expFromFolder        , d.Item("from-folder").Item("path"), d.Item("from-folder").Item("relativePath")) _
        , Array("from"       , "NotExistsFile", Empty                , d.Item("from-folder").Item("path"), d.Item("from").Item("relativePath")) _
        )
    createExpectDefinitionDisappearFromFile = ret
End Function
Function createExpectDefinitionDisappearFromFolder(d)
    createExpectDefinitionDisappearFromFolder = Array( _
        Array(  "from-folder" , "NotExistsFolder", Empty             , d.Item("from-folder").Item("path"), d.Item("from-fl1").Item("relativePath")) _
        )
End Function

Function createExpectDefinitionUnchange(kd,d)
    Dim i,k,t : i=0
    Redim a(d.Count-1)
    For Each k In d.Keys
        If StrComp(kd,Mid(k,1,Len(kd)),vbBinaryCompare)=0 Then
            If d.Item(k).Item("isSetup") Then
                a(i) = Array(k, d.Item(k).Item("type"), d.Item(k), d.Item(kd&"-folder").Item("path"), d.Item(k).Item("relativePath"))
            Else
                If StrComp(d.Item(k).Item("type"),"folder")=0 Then t="NotExistsFolder" Else t="NotExistsFile"
                a(i) = Array(k, t                     , Empty    , d.Item(kd&"-folder").Item("path"), d.Item(k).Item("relativePath"))
            End If
            i = i + 1
        End If
    Next
    If i>0 Then Redim Preserve a(i-1)
    createExpectDefinitionUnchange = a
End Function
Function createExpectDefinitionMergeFolder(d)
    '全てのfromの情報で期待値を上書きする
    Dim f : f = createExpectDefinitionUnchange("from",d)
    createExpectDefinitionMergeFolder = createExpectDefinitionMergeFolderProc(d,f)
End Function
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
