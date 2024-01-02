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
' @import ../../lib/clsCompareExcel.vbs
' @import ../../lib/libCom.vbs

Option Explicit

Const MY_NAME = "test_libCom_fs.vbs"
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
'fs_copyFile()
Sub Test_fs_copyFile_Normal
    com_CopyOrMoveFile_Normal(True)
End Sub
Sub Test_fs_copyFile_Normal_OverRide
    com_CopyOrMoveFile_OverRide(True)
End Sub
Sub Test_fs_copyFile_Normal_FromFileLocked
    com_CopyOrMoveFile_FromFileLocked(True)
End Sub
Sub Test_fs_copyFile_Err_FromFileNoExists
    com_CopyOrMoveFile_FromFileNoExists(True)
End Sub
Sub Test_fs_copyFile_Err_ToFileLocked
    com_CopyOrMoveFile_ToFileLocked(True)
End Sub

'###################################################################################################
'fs_copyFolder()
Sub Test_fs_copyFolder_Normal
    com_CopyOrMoveFolder_Normal(True)
End Sub
Sub Test_fs_copyFolder_Normal_OverRide
    com_CopyOrMoveFolder_OverRide(True)
End Sub
Sub Test_fs_copyFolder_Normal_OverRideWithUnrelatedFileLocked
    com_CopyOrMoveFolder_OverRideWithUnrelatedFileLocked(True)
End Sub
Sub Test_fs_copyFolder_Normal_FromFileLocked
    com_CopyOrMoveFolder_FromFileLocked(True)
End Sub
Sub Test_fs_copyFolder_Err_FromFileNoExists
    com_CopyOrMoveFolder_FromFileNoExists(True)
End Sub
Sub Test_fs_copyFolder_Err_ToFileLocked
    com_CopyOrMoveFolder_ToFileLocked(True)
End Sub

'###################################################################################################
'fs_createFolder()
Sub Test_fs_createFolder
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    assertExistsFolder p, False, "before", "createfolder", "folder"
    
    Dim a,e
    e = True
    a = fs_createFolder(p)
    AssertEqualWithMessage e, a, "ret"
    assertExistsFolder p, True, "after", "createfolder", "folder"
End Sub
Sub Test_fs_createFolder_ErrExistsFile
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))

    Dim c,d
    'ファイルを作成
    c = "UTF-8"
    d = "For" & vbNewLine & "CreateFolder Err-ExistsFile"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "createfolder", "file"
    
    Dim a,e
    e = False
    a = fs_createFolder(p)
    AssertEqualWithMessage e, a, "ret"
    assertExistsFile p, True, "after", "createfolder", "file"
    assertExistsFolder p, False, "after", "createfolder", "folder"
End Sub
Sub Test_fs_createFolder_ErrExistsFolder
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))

    Dim c,d
    'フォルダを作成
    new_Fso().CreateFolder p
    assertExistsFolder p, True, "before", "createfolder", "folder"
    
    Dim a,e
    e = False
    a = fs_createFolder(p)
    AssertEqualWithMessage e, a, "ret"
    assertExistsFolder p, True, "after", "createfolder", "folder"
End Sub

'###################################################################################################
'fs_deleteFile()
Sub Test_fs_deleteFile
    Dim c,p,d
    'ファイルを作成
    c = "UTF-8"
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    d = "For" & vbNewLine & "DeleteFile Normal"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "deleteFile", "file"

    Dim e,a
    e = True
    a = fs_deleteFile(p)
    AssertEqualWithMessage e, a, "ret"
    assertExistsFile p, False, "after", "deleteFile", "file"
End Sub
Sub Test_fs_deleteFile_Err_NotExists
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "deleteFile", "file"

    Dim e,a
    e = False
    a = fs_deleteFile(p)
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"
    assertExistsFile p, False, "after", "deleteFile", "file"
End Sub
Sub Test_fs_deleteFile_Err_FileLocked
    Dim c,p,d,f
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "For" & vbNewLine & "DeleteFile Err FileLocked"
    f = -1    'TristateTrue(Unicode)
    'ファイルを一旦作成してロックする
    With createFileAndLocked(c,p,d,f)
        Dim e,a
        e = False
        a = fs_deleteFile(p)
        
        'fs_deleteFile()がエラーになることを確認する
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    'ファイルが削除されていないことを確認
    assertExistsFile p, True, "after", "deleteFile", "file"
End Sub

'###################################################################################################
'fs_deleteFolder()
Sub Test_fs_deleteFolder
    Dim c,p,pf,d
    'フォルダを作成
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder p
    'フォルダの下にファイルを作成
    c = "UTF-8"
    pf = new_Fso().BuildPath(p, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    d = "For" & vbNewLine & "DeleteFolder Normal"
    writeTestFile c,pf,d
    assertExistsFolder p, True, "before", "deleteFolder", "folder"

    Dim e,a
    e = True
    a = fs_deleteFolder(p)
    AssertEqualWithMessage e, a, "ret"
    assertExistsFolder p, False, "after", "deleteFolder", "folder"
End Sub
Sub Test_fs_deleteFolder_Err_NotExists
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    assertExistsFolder p, False, "before", "deleteFolder", "folder"

    Dim e,a
    e = False
    a = fs_deleteFolder(p)
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"
    assertExistsFolder p, False, "after", "deleteFolder", "folder"
End Sub
Sub Test_fs_deleteFolder_Err_FileLocked
    Dim c,p,pf,d,f
    'フォルダを作成
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder p
    'フォルダの下にファイルを作成
    c = "UTF-8"
    pf = new_Fso().BuildPath(p, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    d = "For" & vbNewLine & "DeleteFolder Err FileLocked"
    f = -1    'TristateTrue(Unicode)
    'ファイルを一旦作成してロックする
    With createFileAndLocked(c,pf,d,f)
        Dim e,a
        e = False
        a = fs_deleteFolder(p)
        
        'fs_deleteFolder()がエラーになることを確認する
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    'フォルダが削除されていないことを確認
    assertExistsFolder p, True, "after", "deleteFolder", "folder"
End Sub

'###################################################################################################
'fs_moveFile()
Sub Test_fs_moveFile_Normal
    com_CopyOrMoveFile_Normal(False)
End Sub
Sub Test_fs_moveFile_Err_OverRide
    com_CopyOrMoveFile_OverRide(False)
End Sub
Sub Test_fs_moveFile_Err_FromFileLocked
    com_CopyOrMoveFile_FromFileLocked(False)
End Sub
Sub Test_fs_moveFile_Err_FromFileNoExists
    com_CopyOrMoveFile_FromFileNoExists(False)
End Sub
Sub Test_fs_moveFile_Err_ToFileLocked
    com_CopyOrMoveFile_ToFileLocked(False)
End Sub

'###################################################################################################
'fs_moveFolder()
Sub Test_fs_moveFolder_Normal
    com_CopyOrMoveFolder_Normal(False)
End Sub
Sub Test_fs_moveFolder_Err_OverRide
    com_CopyOrMoveFolder_OverRide(False)
End Sub
Sub Test_fs_moveFolder_Err_OverRideWithUnrelatedFileLocked
    com_CopyOrMoveFolder_OverRideWithUnrelatedFileLocked(False)
End Sub
Sub Test_fs_moveFolder_Err_FromFileLocked
    com_CopyOrMoveFolder_FromFileLocked(False)
End Sub
Sub Test_fs_moveFolder_Err_FromFileNoExists
    com_CopyOrMoveFolder_FromFileNoExists(False)
End Sub
Sub Test_fs_moveFolder_Err_ToFileLocked
    com_CopyOrMoveFolder_ToFileLocked(False)
End Sub

'###################################################################################################
'fs_readFile()
Sub Test_fs_readFile
    Dim c,p,d,e
    'ファイルを作成
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "lmn" & vbNewLine & "ⅢⅥⅩ" & vbNewLine & "ｱｲｳ" & vbNewLine & ChrW(12316) 'ChrW(12316)='\u301c'（波ダッシュ・波型）Sjisに変換できない文字
    e = d
    writeTestFile c,p,d

    Dim a
    a = fs_readFile(p)
    AssertEqualWithMessage e, a, "ret"
End Sub
Sub Test_fs_readFile_Err
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "readFile", "file"

    Dim e,a
    e = empty
    a = fs_readFile(p)
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"
End Sub

'###################################################################################################
'fs_writeFile()
Sub Test_fs_writeFile
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "writeFile", "file"

    Dim d,ec,ea,a
    d = "abc" & vbNewLine & "あいう" & vbNewLine & "123" & vbNewLine & ChrW(12316) 'ChrW(12316)='\u301c'（波ダッシュ・波型）Sjisに変換できない文字
    ec = d : ea = True
    a = fs_writeFile(p, d)

    Dim c,ct
    c = "Unicode"
    ct = readTestFile(c, p)
    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, ct, "cont"
End Sub
Sub Test_fs_writeFile_Rewrite
    Dim p,c,d
    '上書きするファイルを一旦作成
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "UTF-8"
    d = "For" & vbNewLine & "Rewrite"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "writeFile", "file"

    '上書きすることを確認
    d = "abc" & vbNewLine & "①②③" & vbNewLine & "!#$" & vbNewLine & ChrW(12316) 'ChrW(12316)='\u301c'（波ダッシュ・波型）Sjisに変換できない文字
    Dim a,ec,ea
    ec = d : ea = True
    a = fs_writeFile(p, d)

    Dim ct
    c = "Unicode"
    ct = readTestFile(c, p)
    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, ct, "cont"
End Sub
Sub Test_fs_writeFile_Err
    Dim p,c,d,f,ec
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "For" & vbNewLine & "Write Error"
    f = -1    'TristateTrue(Unicode)
    ec = d
    'ファイルを一旦作成してロックする
    With createFileAndLocked(c, p ,d,f)
        d = "error" & vbNewLine & "test"
        Dim ea,a
        ea = False
        a = fs_writeFile(p, d)
        
        'fs_writeFile()がエラーになることを確認する
        AssertEqualWithMessage ea, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    Dim ct
    '上書きしていないことを確認
    ct = readTestFile(c, p)
    AssertEqualWithMessage ec, ct, "cont"
End Sub

'###################################################################################################
'fs_writeFileDefault()
Sub Test_fs_writeFileDefault
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "writeFileDefault", "file"

    Dim d,ec,ea,a
    d = "abc" & vbNewLine & "あいう" & vbNewLine & "123"
    ec = d : ea = True
    a = fs_writeFileDefault(p, d)

    Dim c,ct
    c = "shift-jis"
    ct = readTestFile(c, p)
    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, ct, "cont"
End Sub

'###################################################################################################
'func_FsWriteFile()
Sub Test_func_FsWriteFile_Iomode_ForWriting_Normal__Format_SystemDefault
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "func_FsWriteFile", "file"

    Dim d,iomode,create,f,ec,ea,a
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Iomode_ForWriting_Normal__Format_SystemDefault"
    iomode = 2     'ForWriting
    create = True
    f = -2         'TristateUseDefault
    ec = d : ea = True
    a = func_FsWriteFile(p, iomode, create, f, d)

    Dim c,ct
    c = "shift-jis"
    ct = readTestFile(c, p)
    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, ct, "cont"
End Sub
Sub Test_func_FsWriteFile_Iomode_ForWriting_Rewrite__Format_Unicode
    Dim p,c,d
    '上書きするファイルを一旦作成
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "For" & vbNewLine & "Rewrite"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "func_FsWriteFile", "file"

    Dim iomode,create,f,ec,ea,a
    '上書きすることを確認
    iomode = 2     'ForWriting
    create = True
    f = -1    'TristateTrue(Unicode)
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Iomode_ForWriting_Rewrite__Format_Unicode"
    ec = d : ea = True
    a = func_FsWriteFile(p, iomode, create, f, d)

    Dim ct
    c = "Unicode"
    ct = readTestFile(c, p)
    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, ct, "cont"
End Sub
Sub Test_func_FsWriteFile_Iomode_ForAppending_Normal__Format_Ascii
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "func_FsWriteFile", "file"

    Dim d,iomode,create,f,ec,ea,a
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Iomode_ForAppending_Normal__Format_Ascii"
    iomode = 8     'ForAppending
    create = True
    f = 0          'TristateFalse(Ascii)
    ec = d : ea = True
    a = func_FsWriteFile(p, iomode, create, f, d)

    Dim c,ct
    c = "shift-jis"
    ct = readTestFile(c,p)
    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, ct, "cont"
End Sub
Sub Test_func_FsWriteFile_Iomode_ForAppending_Append__Format_SystemDefault
    Dim p,c,d
    '追記するファイルを一旦作成
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "shift-jis"
    d = "For" & vbNewLine & "Append"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "func_FsWriteFile", "file"

    Dim iomode,create,f,ec,ea,a
    '追記することを確認
    iomode = 8     'ForAppending
    create = True
    f = -2         'TristateUseDefault
    ec = d
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Iomode_ForAppending_Append__Format_SystemDefault"
    ec = ec & d : ea = True
    a = func_FsWriteFile(p, iomode, create, f, d)

    Dim ct
    ct = readTestFile(c, p)
    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, ct, "cont"
End Sub
Sub Test_func_FsWriteFile_Create_True_Normal__Format_Unicode
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "func_FsWriteFile", "file"

    Dim d,iomode,create,f,ec,ea,a
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Create_True_Normal__Format_Unicode"
    iomode = 2     'ForWriting
    create = True
    f = -1         'TristateTrue(Unicode)
    ec = d : ea = True
    a = func_FsWriteFile(p, iomode, create, f, d)

    Dim c,ct
    c = "Unicode"
    ct = readTestFile(c, p)
    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, ct, "cont"
End Sub
Sub Test_func_FsWriteFile_Create_True_Rewrite__Format_Ascii
    Dim p,c,d
    '上書きするファイルを一旦作成
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "shift-jis"
    d = "For" & vbNewLine & "Rewrite"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "func_FsWriteFile", "file"

    Dim iomode,create,f,ec,ea,a
    '上書きすることを確認
    iomode = 2     'ForWriting
    create = True
    f = 0          'TristateFalse(Ascii)
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Create_True_Rewrite__Format_Ascii"
    ec = d : ea = True
    a = func_FsWriteFile(p, iomode, create, f, d)

    Dim ct
    ct = readTestFile(c, p)
    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, ct, "cont"
End Sub
Sub Test_func_FsWriteFile_Create_False_Err
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
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
    AssertEqualWithMessage False, new_Fso().FileExists(p), "after write file exists"
End Sub
Sub Test_func_FsWriteFile_Create_False_Rewrite__Format_Unicode
    Dim p,c,d
    '上書きするファイルを一旦作成
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "For" & vbNewLine & "Rewrite"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "func_FsWriteFile", "file"

    Dim ec,ea,a,iomode,create,f
    '上書きすることを確認
    iomode = 2     'ForWriting
    create = False
    f = -1         'TristateTrue(Unicode)
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Create_False_Rewrite__Format_Unicode"
    ec = d : ea = True
    a = func_FsWriteFile(p, iomode, create, f, d)

    Dim ct
    ct = readTestFile(c, p)
    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, ct, "cont"
End Sub
Sub Test_func_FsWriteFile_Err_FileLocked
    Dim p,d,iomode,create,f,c,ec,ea,a
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    d = "error" & vbNewLine & "FileLocked"
    iomode = 2     'ForWriting
    create = False
    f = 0          'TristateFalse(Ascii)
    c = "shift-jis"
    ec = d
    'ファイルを一旦作成してロックする
    With createFileAndLocked(c, p ,d, f)
        assertExistsFile p, True, "before", "func_FsWriteFile", "file(Locked)"

        d = "error" & vbNewLine & "test"
        ea = False
        a = func_FsWriteFile(p, iomode, create, f, d)
        
        'func_FsWriteFile()がエラーになることを確認する
        AssertEqualWithMessage ea, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    '上書きしていないことを確認
    Dim ct
    ct = readTestFile(c, p)
    AssertEqualWithMessage ec, ct, "cont"
End Sub

'###################################################################################################
'func_FsReadFile()
Sub Test_func_FsReadFile_Normal__Format_SystemDefault
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))

    Dim d,f,c,e,a
    d = "func_FsReadFile" & vbNewLine & "のテスト" & vbNewLine & "Normal__Format_SystemDefault"
    f = -2         'TristateUseDefault
    c = "shift-jis"
    e = d
    writeTestFile c,p,d

    a = func_FsReadFile(p,f)
    AssertEqualWithMessage e, a, "ret"
End Sub
Sub Test_func_FsReadFile_Normal__Format_Unicode
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))

    Dim d,f,c,e,a
    d = "func_FsReadFile" & vbNewLine & "のテスト" & vbNewLine & "Normal__Format_Unicode"
    f = -1         'TristateTrue(Unicode)
    c = "Unicode"
    e = d
    writeTestFile c,p,d

    a = func_FsReadFile(p,f)
    AssertEqualWithMessage e, a, "ret"
End Sub
Sub Test_func_FsReadFile_Normal__Format_Ascii
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))

    Dim d,f,c,e,a
    d = "func_FsReadFile" & vbNewLine & "のテスト" & vbNewLine & "Normal__Format_Ascii"
    f = 0          'TristateFalse(Ascii)
    c = "shift-jis"
    e = d
    writeTestFile c,p,d
    
    a = func_FsReadFile(p,f)
    AssertEqualWithMessage e, a, "ret"
End Sub
Sub Test_func_FsReadFile_Err
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))

    Dim f,e,a
    f = -2         'TristateUseDefault
    assertExistsFile p, False, "before", "func_FsReadFile", "file"
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
Function createFileAndLocked(c,p,d,f)
    With CreateObject("ADODB.Stream")
        .Charset = c
        .Open
        .WriteText d, 0
        .SaveToFile p, 2
        .Close
    End With
    'Textstreamを作成して返却
    Set createFileAndLocked = new_Ts(p, 8, True, f)
End Function
Sub assertExistsFile(path, expect, timestr, acton, tgt)
    AssertEqualWithMessage expect, new_Fso().FileExists(path), timestr & " " & acton & " " & tgt & " exists"
End Sub
Sub assertExistsFolder(path, expect, timestr, acton, tgt)
    AssertEqualWithMessage expect, new_Fso().FolderExists(path), timestr & " " & acton & " " & tgt & " exists"
End Sub
Sub assertFilesSubfoldersCount(path, expectfilecnt, expectfoldercnt, tgt)
    AssertEqualWithMessage expectfilecnt, new_Fso().GetFolder(path).Files.Count, tgt & " folderFiles Count"
    AssertEqualWithMessage expectfoldercnt, new_Fso().GetFolder(path).SubFolders.Count, tgt & " folderSubFolders Count"
End Sub
Sub com_CopyOrMoveFile_Normal(IsCopy)
    Dim from
    'fromパスを作成
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    Dim c,df
    'fromファイルを作成
    c = "Unicode"
    df = "For" & vbNewLine & "copy/moveFile Normal"
    writeTestFile c,from,df
    assertExistsFile from, True, "before", "copy/moveFile", "fromfile"
    
    Dim toto
    'toパスを作成
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_.txt"))
    assertExistsFile toto, False, "before", "copy/moveFile", "tofile"

    Dim e,a
    '実行結果の確認
    e = True
    If IsCopy Then
        a = fs_copyFile(from,toto)
    Else
        a = fs_moveFile(from,toto)
    End If
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    'fromファイルの確認
    If isCopy Then
        assertExistsFile from, True, "after", "copy/moveFile", "fromfile"
        e = df
        a = readTestFile(c, from)
        AssertEqualWithMessage e, a, "cont"
    Else
        assertExistsFile from, False, "after", "copy/moveFile", "fromfile"
    End If

    'toファイルの確認
    assertExistsFile toto, True, "after", "copy/moveFile", "tofile"
    e = df
    a = readTestFile(c, toto)
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub com_CopyOrMoveFile_OverRide(IsCopy)
    Dim from
    'fromパスを作成
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    Dim c,df
    'fromファイルを作成
    c = "Unicode"
    df = "For" & vbNewLine & "copy/moveFile OverRide"
    writeTestFile c,from,df
    assertExistsFile from, True, "before", "copy/moveFile", "fromfile"
    
    Dim toto
    'toパスを作成
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_.txt"))
    Dim dt
    'toファイルを作成
    c = "Unicode"
    dt = "For" & vbNewLine & "copy/moveFile ToFile"
    writeTestFile c,toto,dt
    assertExistsFile toto, True, "before", "copy/moveFile", "tofile"

    'copyは正常（上書きする）moveは異常（上書きしない）になる
    Dim e,a
    '実行結果の確認
    If IsCopy Then
        e = True
        a = fs_copyFile(from,toto)
    Else
        e = False
        a = fs_moveFile(from,toto)
    End If
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    'fromファイルの確認
    If isCopy Then
        assertExistsFile from, True, "after", "copy/moveFile", "fromfile"
        e = df
        a = readTestFile(c, from)
        AssertEqualWithMessage e, a, "cont"
    Else
        assertExistsFile from, True, "after", "copy/moveFile", "fromfile"
    End If

    'toファイルの確認
    If isCopy Then
        assertExistsFile toto, True, "after", "copy/moveFile", "tofile"
        e = df
        a = readTestFile(c, toto)
        AssertEqualWithMessage e, a, "cont"
    Else
        assertExistsFile toto, True, "after", "copy/moveFile", "tofile"
        e = dt
        a = readTestFile(c, toto)
        AssertEqualWithMessage e, a, "cont"
    End If
End Sub
Sub com_CopyOrMoveFile_FromFileLocked(IsCopy)
    Dim toto
    'toパスを作成
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_.txt"))
    assertExistsFile toto, False, "before", "copy/moveFile", "tofile"

    Dim from
    'fromパスを作成
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    Dim c,df,f
    'fromファイルを作成してロックする
    c = "Unicode"
    df = "For" & vbNewLine & "copy/moveFile FromFileLocked"
    f = -1    'TristateTrue(Unicode)
    With createFileAndLocked(c,from,df,f)
        assertExistsFile from, True, "before", "copy/moveFile", "fromfile(Locked)"
        
        Dim e,a
        '実行結果の確認
        If IsCopy Then
            e = True
            a = fs_copyFile(from,toto)
        Else
            e = False
            a = fs_moveFile(from,toto)
        End If
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    'fromファイルの確認
    If isCopy Then
        assertExistsFile from, True, "after", "copy/moveFile", "fromfile"
        e = df
        a = readTestFile(c, from)
        AssertEqualWithMessage e, a, "cont"
    Else
        assertExistsFile from, True, "after", "copy/moveFile", "fromfile"
    End If

    'toファイルの確認
    If isCopy Then
        assertExistsFile toto, True, "after", "copy/moveFile", "tofile"
        e = df
        a = readTestFile(c, toto)
        AssertEqualWithMessage e, a, "cont"
    Else
        assertExistsFile toto, False, "after", "copy/moveFile", "tofile"
    End If
End Sub
Sub com_CopyOrMoveFile_FromFileNoExists(IsCopy)
    Dim from
    'fromパスを作成
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile from, False, "before", "copy/moveFile", "fromfile"
    
    Dim toto
    'toパスを作成
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_.txt"))
    assertExistsFile toto, False, "before", "copy/moveFile", "tofile"

    Dim e,a
    '実行結果の確認
    e = False
    If isCopy Then
        a = fs_copyFile(from,toto)
    Else
        a = fs_moveFile(from,toto)
    End If
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    'fromファイルの確認
    assertExistsFile from, False, "after", "copy/moveFile", "fromfile"

    'toファイルの確認
    assertExistsFile toto, False, "after", "copy/moveFile", "tofile"
End Sub
Sub com_CopyOrMoveFile_ToFileLocked(IsCopy)
    Dim from
    'fromパスを作成
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    Dim c,df
    'fromファイルを作成
    c = "Unicode"
    df = "For" & vbNewLine & "copy/moveFile OverRide"
    writeTestFile c,from,df
    assertExistsFile from, True, "before", "copy/move", "fromfile"
    
    Dim toto
    'toパスを作成
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_.txt"))
    
    Dim dt,f
    'toファイルを作成してロックする
    dt = "For" & vbNewLine & "copy/moveFile ToFileLocked"
    f = -1    'TristateTrue(Unicode)
    With createFileAndLocked(c,toto,dt,f)
        assertExistsFile toto, True, "before", "copy/move", "tofile"
        
        Dim e,a
        '実行結果の確認
        e = False
        If isCopy Then
            a = fs_copyFile(from,toto)
        Else
            a = fs_moveFile(from,toto)
        End If
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    'fromファイルの確認
    assertExistsFile from, True, "after", "copy/move", "fromfile"
    e = df
    a = readTestFile(c, from)
    AssertEqualWithMessage e, a, "cont"

    'toファイルの確認
    assertExistsFile toto, True, "after", "copy/move", "tofile"
    e = dt
    a = readTestFile(c, toto)
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub com_CopyOrMoveFolder_Normal(IsCopy)
    Dim from
    'fromフォルダを作成
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder from
    Dim c,p,ff1,ff2,ff3,df1,df2
    'fromフォルダの下にファイルとフォルダを作成
    c = "Unicode"
    ff1 = new_Now().formatAs("YYMMDD_hhmmss.000000_f1.txt")
    df1 = "For" & vbNewLine & "copy/moveFolder Normal ff1"
    p = new_Fso().BuildPath(from, ff1)
    writeTestFile c,p,df1
    ff2 = new_Now().formatAs("YYMMDD_hhmmss.000000_f2.txt")
    df2 = "For" & vbNewLine & "copy/moveFolder Normal ff2"
    p = new_Fso().BuildPath(from, ff2)
    writeTestFile c,p,df2
    ff3 = new_Now().formatAs("YYMMDD_hhmmss.000000_f3")
    p = new_Fso().BuildPath(from, ff3)
    new_Fso().CreateFolder p
    assertExistsFolder from, True, "before", "copy/moveFolder", "fromfolder"
    assertExistsFile new_Fso().BuildPath(from, ff1), True, "before", "copy/moveFolder", "fromfolder-file1"
    assertExistsFile new_Fso().BuildPath(from, ff2), True, "before", "copy/moveFolder", "fromfolder-file2"
    assertExistsFolder new_Fso().BuildPath(from, ff3), True, "before", "copy/moveFolder", "fromfolder-folder3"
    
    Dim toto
    'toパスを作成（フォルダは作成しない）
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_"))
    assertExistsFolder toto, False, "before", "copy/moveFolder", "toFolder"

    Dim e,a
    '実行結果の確認
    e = True
    If isCopy Then
        a = fs_copyFolder(from,toto)
    Else
        a = fs_moveFolder(from,toto)
    End If
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    'fromフォルダの確認
    If isCopy Then
        assertExistsFolder from, True, "after", "copy/moveFolder", "toFolder"
        assertFilesSubfoldersCount from, 2, 1, "from"
        e = df1
        a = readTestFile(c, new_Fso().BuildPath(from, ff1))
        AssertEqualWithMessage e, a, "cont file1"
        e = df2
        a = readTestFile(c, new_Fso().BuildPath(from, ff2))
        AssertEqualWithMessage e, a, "cont file2"
        assertExistsFolder new_Fso().BuildPath(from, ff3), True, "after", "copy/moveFolder", "fromfolder-folder3"
    Else
        assertExistsFolder from, False, "after", "copy/moveFolder", "toFolder"
    End If

    'toフォルダの確認
    assertExistsFolder toto, True, "after", "copy/moveFolder", "toFolder"
    assertFilesSubfoldersCount toto, 2, 1, "to"
    e = df1
    a = readTestFile(c, new_Fso().BuildPath(toto, ff1))
    AssertEqualWithMessage e, a, "cont file1"
    e = df2
    a = readTestFile(c, new_Fso().BuildPath(toto, ff2))
    AssertEqualWithMessage e, a, "cont file2"
    assertExistsFolder new_Fso().BuildPath(toto, ff3), True, "after", "copy/moveFolder", "tofolder-folder3"
End Sub
Sub com_CopyOrMoveFolder_OverRide(IsCopy)
    Dim from
    'fromフォルダを作成
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder from
    Dim c,p,ff1,ff2,ff3,df1,df2
    'fromフォルダの下にファイルとフォルダを作成
    c = "Unicode"
    ff1 = new_Now().formatAs("YYMMDD_hhmmss.000000_f1.txt")
    df1 = "For" & vbNewLine & "copy/moveFolder OverRide ff1"
    p = new_Fso().BuildPath(from, ff1)
    writeTestFile c,p,df1
    ff2 = new_Now().formatAs("YYMMDD_hhmmss.000000_f2.txt")
    df2 = "For" & vbNewLine & "copy/moveFolder OverRide ff2"
    p = new_Fso().BuildPath(from, ff2)
    writeTestFile c,p,df2
    ff3 = new_Now().formatAs("YYMMDD_hhmmss.000000_f3")
    p = new_Fso().BuildPath(from, ff3)
    new_Fso().CreateFolder p
    assertExistsFolder from, True, "before", "copy/move", "fromfolder"
    assertExistsFile new_Fso().BuildPath(from, ff1), True, "before", "copy/moveFolder", "fromfolder-file1"
    assertExistsFile new_Fso().BuildPath(from, ff2), True, "before", "copy/moveFolder", "fromfolder-file2"
    assertExistsFolder new_Fso().BuildPath(from, ff3), True, "before", "copy/moveFolder", "fromfolder-folder3"
    
    Dim toto
    'toフォルダを作成
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_"))
    new_Fso().CreateFolder toto
    Dim ft1,ft2,ft3,dt1,dt2
    'toフォルダの下にファイルとフォルダを作成
    ft1 = new_Now().formatAs("YYMMDD_hhmmss.000000_t1.txt")
    dt1 = "For" & vbNewLine & "copy/moveFolder OverRide ft1"
    p = new_Fso().BuildPath(toto, ft1)
    writeTestFile c,p,dt1
    ft2 = ff2
    dt2 = "For" & vbNewLine & "copy/moveFolder OverRide ft2"
    p = new_Fso().BuildPath(toto, ft2)
    writeTestFile c,p,dt2
    ft3 = new_Now().formatAs("YYMMDD_hhmmss.000000_t3")
    p = new_Fso().BuildPath(toto, ft3)
    new_Fso().CreateFolder p
    assertExistsFolder toto, True, "before", "copy/moveFolder", "tofolder"
    assertExistsFile new_Fso().BuildPath(toto, ft1), True, "before", "copy/moveFolder", "tofolder-file1"
    assertExistsFile new_Fso().BuildPath(toto, ft2), True, "before", "copy/moveFolder", "tofolder-file2"
    assertExistsFolder new_Fso().BuildPath(toto, ft3), True, "before", "copy/moveFolder", "tofolder-folder3"

    'copyは正常（上書きする）moveは異常（上書きしない）になる
    Dim e,a
    '実行結果の確認
    If isCopy Then
        e = True
        a = fs_copyFolder(from,toto)
    Else
        e = False
        a = fs_moveFolder(from,toto)
    End If
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    'fromフォルダの確認
    assertExistsFolder from, True, "after", "copy/moveFolder", "fromFolder"
    assertFilesSubfoldersCount from, 2, 1, "from"
    e = df1
    a = readTestFile(c, new_Fso().BuildPath(from, ff1))
    AssertEqualWithMessage e, a, "cont file1"
    e = df2
    a = readTestFile(c, new_Fso().BuildPath(from, ff2))
    AssertEqualWithMessage e, a, "cont file2"
    assertExistsFolder new_Fso().BuildPath(from, ff3), True, "after", "copy/moveFolder", "fromfolder-folder3"

    'toフォルダの確認
    If isCopy Then
        assertExistsFolder toto, True, "after", "copy/moveFolder", "toFolder"
        assertFilesSubfoldersCount toto, 3, 2, "to"
        e = dt1
        a = readTestFile(c, new_Fso().BuildPath(toto, ft1))
        AssertEqualWithMessage e, a, "cont tofolder-tofile1"
        assertExistsFolder new_Fso().BuildPath(toto, ft3), True, "after", "copy/moveFolder", "tofolder-tofolder3"
        e = df1
        a = readTestFile(c, new_Fso().BuildPath(toto, ff1))
        AssertEqualWithMessage e, a, "cont tofolder-fromfile1"
        e = df2
        a = readTestFile(c, new_Fso().BuildPath(toto, ff2))
        AssertEqualWithMessage e, a, "cont tofolder-fromfile2"
        assertExistsFolder new_Fso().BuildPath(toto, ff3), True, "after", "copy/moveFolder", "tofolder-fromfolder3"
    Else
        assertExistsFolder toto, True, "after", "copy/moveFolder", "toFolder"
        assertFilesSubfoldersCount toto, 2, 1, "to"
        e = dt1
        a = readTestFile(c, new_Fso().BuildPath(toto, ft1))
        AssertEqualWithMessage e, a, "cont tofolder-fromfile1"
        e = dt2
        a = readTestFile(c, new_Fso().BuildPath(toto, ft2))
        AssertEqualWithMessage e, a, "cont tofolder-fromfile2"
        assertExistsFolder new_Fso().BuildPath(toto, ft3), True, "after", "copy/moveFolder", "tofolder-fromfolder3"
    End If
End Sub
Sub com_CopyOrMoveFolder_OverRideWithUnrelatedFileLocked(IsCopy)
    Dim from
    'fromフォルダを作成
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder from
    Dim c,p,ff1,ff2,ff3,df1,df2
    'fromフォルダの下にファイルとフォルダを作成
    c = "Unicode"
    ff1 = new_Now().formatAs("YYMMDD_hhmmss.000000_f1.txt")
    df1 = "For" & vbNewLine & "copy/moveFolder OverRideWithUnrelatedFileLocked ff1"
    p = new_Fso().BuildPath(from, ff1)
    writeTestFile c,p,df1
    ff2 = new_Now().formatAs("YYMMDD_hhmmss.000000_f2.txt")
    df2 = "For" & vbNewLine & "copy/moveFolder OverRideWithUnrelatedFileLocked ff2"
    p = new_Fso().BuildPath(from, ff2)
    writeTestFile c,p,df2
    ff3 = new_Now().formatAs("YYMMDD_hhmmss.000000_f3")
    p = new_Fso().BuildPath(from, ff3)
    new_Fso().CreateFolder p
    assertExistsFolder from, True, "before", "copy/moveFolder", "fromfolder"
    assertExistsFile new_Fso().BuildPath(from, ff1), True, "before", "copy/moveFolder", "fromfolder-file1"
    assertExistsFile new_Fso().BuildPath(from, ff2), True, "before", "copy/moveFolder", "fromfolder-file2"
    assertExistsFolder new_Fso().BuildPath(from, ff3), True, "before", "copy/moveFolder", "fromfolder-folder3"
    
    Dim toto
    'toフォルダを作成
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_"))
    new_Fso().CreateFolder toto
    Dim ft2,ft3,dt2
    'toフォルダの下にファイルとフォルダを作成
    ft2 = ff2
    dt2 = "For" & vbNewLine & "copy/moveFolder OverRideWithUnrelatedFileLocked ft2"
    p = new_Fso().BuildPath(toto, ft2)
    writeTestFile c,p,dt2
    ft3 = new_Now().formatAs("YYMMDD_hhmmss.000000_t3")
    p = new_Fso().BuildPath(toto, ft3)
    new_Fso().CreateFolder p
    assertExistsFolder toto, True, "before", "copy/moveFolder", "tofolder"
    assertExistsFile new_Fso().BuildPath(toto, ft2), True, "before", "copy/moveFolder", "tofolder-file2"
    assertExistsFolder new_Fso().BuildPath(toto, ft3), True, "before", "copy/moveFolder", "tofolder-folder3"
    
    Dim ft1,dt1,f
    ft1 = new_Now().formatAs("YYMMDD_hhmmss.000000_t1.txt")
    dt1 = "For" & vbNewLine & "copy/moveFolder OverRideWithUnrelatedFileLocked ft1"
    p = new_Fso().BuildPath(toto, ft1)
    f = -1    'TristateTrue(Unicode)
    'toフォルダのファイル（ft1）を作成してロックする
    With createFileAndLocked(c,p,dt1,f)
        assertExistsFile p, True, "before", "copy/moveFolder", "tofolder-file1(locked)"
        
        Dim e,a
        '実行結果の確認
        If isCopy Then
            e = True
            a = fs_copyFolder(from,toto)
        Else
            e = False
            a = fs_moveFolder(from,toto)
        End If
        
        'fs_copyFolder()/fs_moveFolder()がエラーにならないことを確認する
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    'toフォルダのファイル（ft1）をロックしているが、上書きしないためcopyは正常に完了する、moveはエラーになる

    'fromフォルダの確認
    assertExistsFolder from, True, "after", "copy/moveFolder", "fromFolder"
    assertFilesSubfoldersCount from, 2, 1, "from"
    e = df1
    a = readTestFile(c, new_Fso().BuildPath(from, ff1))
    AssertEqualWithMessage e, a, "cont file1"
    e = df2
    a = readTestFile(c, new_Fso().BuildPath(from, ff2))
    AssertEqualWithMessage e, a, "cont file2"
    assertExistsFolder new_Fso().BuildPath(from, ff3), True, "after", "copy/moveFolder", "fromfolder-folder3"

    'toフォルダの確認
    If isCopy Then
        assertExistsFolder toto, True, "after", "copy/moveFolder", "toFolder"
        assertFilesSubfoldersCount toto, 3, 2, "to"
        e = dt1
        a = readTestFile(c, new_Fso().BuildPath(toto, ft1))
        AssertEqualWithMessage e, a, "cont tofolder-tofile1"
        assertExistsFolder new_Fso().BuildPath(toto, ft3), True, "after", "copy/moveFolder", "tofolder-tofolder3"
        e = df1
        a = readTestFile(c, new_Fso().BuildPath(toto, ff1))
        AssertEqualWithMessage e, a, "cont tofolder-fromfile1"
        e = df2
        a = readTestFile(c, new_Fso().BuildPath(toto, ff2))
        AssertEqualWithMessage e, a, "cont tofolder-fromfile2"
        assertExistsFolder new_Fso().BuildPath(toto, ff3), True, "after", "copy/moveFolder", "tofolder-fromfolder3"
    Else
        assertExistsFolder toto, True, "after", "copy/moveFolder", "toFolder"
        assertFilesSubfoldersCount toto, 2, 1, "to"
        e = dt1
        a = readTestFile(c, new_Fso().BuildPath(toto, ft1))
        AssertEqualWithMessage e, a, "cont tofolder-fromfile1"
        e = dt2
        a = readTestFile(c, new_Fso().BuildPath(toto, ft2))
        AssertEqualWithMessage e, a, "cont tofolder-fromfile2"
        assertExistsFolder new_Fso().BuildPath(toto, ft3), True, "after", "copy/moveFolder", "tofolder-fromfolder3"
    End If
End Sub
Sub com_CopyOrMoveFolder_FromFileLocked(IsCopy)
    Dim toto
    'toフォルダを作成
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_"))
    new_Fso().CreateFolder toto
    Dim c,p,ft1,ft2,ft3,dt1,dt2
    c = "Unicode"
    'toフォルダの下にファイルとフォルダを作成
    ft1 = new_Now().formatAs("YYMMDD_hhmmss.000000_t1.txt")
    dt1 = "For" & vbNewLine & "copy/moveFolder FromFileLocked ft1"
    p = new_Fso().BuildPath(toto, ft1)
    writeTestFile c,p,dt1
    ft2 = new_Now().formatAs("YYMMDD_hhmmss.000000_f2.txt")
    dt2 = "For" & vbNewLine & "copy/moveFolder FromFileLocked ft2"
    p = new_Fso().BuildPath(toto, ft2)
    writeTestFile c,p,dt2
    ft3 = new_Now().formatAs("YYMMDD_hhmmss.000000_t3")
    p = new_Fso().BuildPath(toto, ft3)
    new_Fso().CreateFolder p
    assertExistsFolder toto, True, "before", "copy/moveFolder", "tofolder"
    assertExistsFile new_Fso().BuildPath(toto, ft1), True, "before", "copy/moveFolder", "tofolder-file1"
    assertExistsFile new_Fso().BuildPath(toto, ft2), True, "before", "copy/moveFolder", "tofolder-file2"
    assertExistsFolder new_Fso().BuildPath(toto, ft3), True, "before", "copy/moveFolder", "tofolder-folder3"

    Dim from
    'fromフォルダを作成
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder from
    Dim ff1,ff3,df1
    'fromフォルダの下にファイルとフォルダを作成
    ff1 = new_Now().formatAs("YYMMDD_hhmmss.000000_f1.txt")
    df1 = "For" & vbNewLine & "copy/moveFolder FromFileLocked ff1"
    p = new_Fso().BuildPath(from, ff1)
    writeTestFile c,p,df1
    ff3 = new_Now().formatAs("YYMMDD_hhmmss.000000_f3")
    p = new_Fso().BuildPath(from, ff3)
    new_Fso().CreateFolder p
    assertExistsFolder from, True, "before", "copy/moveFolder", "fromfolder"
    assertExistsFile new_Fso().BuildPath(from, ff1), True, "before", "copy/moveFolder", "fromfolder-file1"
    assertExistsFolder new_Fso().BuildPath(from, ff3), True, "before", "copy/moveFolder", "fromfolder-folder3"
    
    Dim ff2,df2,f
    ff2 = ft2
    df2 = "For" & vbNewLine & "copy/moveFolder FromFileLocked ff2"
    p = new_Fso().BuildPath(from, ff2)
    f = -1    'TristateTrue(Unicode)
    'toフォルダのファイル（ff2）を作成してロックする
    With createFileAndLocked(c,p,df2,f)
        assertExistsFile p, True, "before", "copy/moveFolder", "fromfolder-file2(locked)"
        
        'copyは正常（上書きする）moveは異常（上書きしない）になる
        Dim e,a
        '実行結果の確認
        If isCopy Then
            e = True
            a = fs_copyFolder(from,toto)
        Else
            e = False
            a = fs_moveFolder(from,toto)
        End If
        
        'fs_copyFolder()/fs_moveFolder()がエラーにならないことを確認する
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    'fromフォルダのファイル（ff2）をロックしていると、copyは正常、moveがエラーになる

    'fromフォルダの確認
    assertExistsFolder from, True, "after", "copy/moveFolder", "fromFolder"
    assertFilesSubfoldersCount from, 2, 1, "from"
    e = df1
    a = readTestFile(c, new_Fso().BuildPath(from, ff1))
    AssertEqualWithMessage e, a, "cont file1"
    e = df2
    a = readTestFile(c, new_Fso().BuildPath(from, ff2))
    AssertEqualWithMessage e, a, "cont file2"
    assertExistsFolder new_Fso().BuildPath(from, ff3), True, "after", "copy/moveFolder", "fromfolder-folder3"

    'toフォルダの確認
    If isCopy Then
        assertExistsFolder toto, True, "after", "copy/moveFolder", "toFolder"
        assertFilesSubfoldersCount toto, 3, 2, "to"
        e = dt1
        a = readTestFile(c, new_Fso().BuildPath(toto, ft1))
        AssertEqualWithMessage e, a, "cont tofolder-tofile1"
        assertExistsFolder new_Fso().BuildPath(toto, ft3), True, "after", "copy/moveFolder", "tofolder-tofolder3"
        e = df1
        a = readTestFile(c, new_Fso().BuildPath(toto, ff1))
        AssertEqualWithMessage e, a, "cont tofolder-fromfile1"
        e = df2
        a = readTestFile(c, new_Fso().BuildPath(toto, ff2))
        AssertEqualWithMessage e, a, "cont tofolder-fromfile2"
        assertExistsFolder new_Fso().BuildPath(toto, ff3), True, "after", "copy/moveFolder", "tofolder-fromfolder3"
    Else
        assertExistsFolder toto, True, "after", "copy/moveFolder", "toFolder"
        assertFilesSubfoldersCount toto, 2, 1, "to"
        e = dt1
        a = readTestFile(c, new_Fso().BuildPath(toto, ft1))
        AssertEqualWithMessage e, a, "cont tofolder-fromfile1"
        e = dt2
        a = readTestFile(c, new_Fso().BuildPath(toto, ft2))
        AssertEqualWithMessage e, a, "cont tofolder-fromfile2"
        assertExistsFolder new_Fso().BuildPath(toto, ft3), True, "after", "copy/moveFolder", "tofolder-fromfolder3"
    End If
End Sub
Sub com_CopyOrMoveFolder_FromFileNoExists(IsCopy)
    Dim from
    'fromパスを作成（フォルダは作成しない）
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    assertExistsFolder from, False, "before", "copy/moveFolder", "fromfolder"
    
    Dim toto
    'toパスを作成（フォルダは作成しない）
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_"))
    assertExistsFolder toto, False, "before", "copy/moveFolder", "tofolder"

    Dim e,a
    '実行結果の確認
    e = False
    If isCopy Then
        a = fs_copyFolder(from,toto)
    Else
        a = fs_moveFolder(from,toto)
    End If
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    'fromフォルダの確認
    assertExistsFolder from, False, "after", "copy/moveFolder", "fromFolder"

    'toフォルダの確認
    assertExistsFolder toto, False, "after", "copy/moveFolder", "toFolder"
End Sub
Sub com_CopyOrMoveFolder_ToFileLocked(IsCopy)
    Dim from
    'fromフォルダを作成
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder from
    Dim c,p,ff1,ff2,ff3,df1,df2
    'fromフォルダの下にファイルとフォルダを作成
    c = "Unicode"
    ff1 = new_Now().formatAs("YYMMDD_hhmmss.000000_f1.txt")
    df1 = "For" & vbNewLine & "copy/moveFolder ToFileLocked ff1"
    p = new_Fso().BuildPath(from, ff1)
    writeTestFile c,p,df1
    ff2 = new_Now().formatAs("YYMMDD_hhmmss.000000_f2.txt")
    df2 = "For" & vbNewLine & "copy/moveFolder ToFileLocked ff2"
    p = new_Fso().BuildPath(from, ff2)
    writeTestFile c,p,df2
    ff3 = new_Now().formatAs("YYMMDD_hhmmss.000000_f3")
    p = new_Fso().BuildPath(from, ff3)
    new_Fso().CreateFolder p
    assertExistsFolder from, True, "before", "copy/moveFolder", "fromfolder"
    assertExistsFile new_Fso().BuildPath(from, ff1), True, "before", "copy/moveFolder", "fromfolder-file1"
    assertExistsFile new_Fso().BuildPath(from, ff2), True, "before", "copy/moveFolder", "fromfolder-file2"
    assertExistsFolder new_Fso().BuildPath(from, ff3), True, "before", "copy/moveFolder", "fromfolder-folder3"
    
    Dim toto
    'コピー先フォルダを作成
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_"))
    new_Fso().CreateFolder toto
    Dim ft1,ft3,dt1
    'フォルダの下にファイルとフォルダを作成
    ft1 = new_Now().formatAs("YYMMDD_hhmmss.000000_t1.txt")
    dt1 = "For" & vbNewLine & "copy/moveFolder ToFileLocked ft1"
    p = new_Fso().BuildPath(toto, ft1)
    writeTestFile c,p,dt1
    ft3 = new_Now().formatAs("YYMMDD_hhmmss.000000_t3")
    p = new_Fso().BuildPath(toto, ft3)
    new_Fso().CreateFolder p
    assertExistsFolder toto, True, "before", "copy/moveFolder", "tofolder"
    assertExistsFile new_Fso().BuildPath(toto, ft1), True, "before", "copy/moveFolder", "tofolder-file1"
    assertExistsFolder new_Fso().BuildPath(toto, ft3), True, "before", "copy/moveFolder", "tofolder-folder3"

    Dim ft2,dt2,f
    ft2 = ff2
    dt2 = "For" & vbNewLine & "copy/moveFolder ToFileLocked ft2"
    p = new_Fso().BuildPath(toto, ft2)
    f = -1    'TristateTrue(Unicode)
    'toフォルダのファイル（ft2）を作成してロックする
    With createFileAndLocked(c,p,dt2,f)
        assertExistsFile p, True, "before", "copy/moveFolder", "tofolder-file2(locked)"
        
        Dim e,a
        '実行結果の確認
        e = False
        If isCopy Then
            a = fs_copyFolder(from,toto)
        Else
            a = fs_moveFolder(from,toto)
        End If
        
        'fs_copyFolder()/fs_moveFolder()がエラーになることを確認する
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    '1つ目のファイルfn1はコピーまたは移動し、2つ目のファイルfn2のコピーまたは移動が失敗する、3つ目のフォルダfn3はコピーまたは移動しない

    'fromフォルダの確認
    assertExistsFolder from, True, "after", "copy/move", "fromFolder"
    assertFilesSubfoldersCount from, 2, 1, "from"
    e = df1
    a = readTestFile(c, new_Fso().BuildPath(from, ff1))
    AssertEqualWithMessage e, a, "cont file1"
    e = df2
    a = readTestFile(c, new_Fso().BuildPath(from, ff2))
    AssertEqualWithMessage e, a, "cont file2"
    assertExistsFolder new_Fso().BuildPath(from, ff3), True, "after", "copy/move", "fromfolder-folder3"

    'toフォルダの確認
    If isCopy Then
        assertExistsFolder toto, True, "after", "copy/move", "toFolder"
        assertFilesSubfoldersCount toto, 3, 1, "to"
        e = dt1
        a = readTestFile(c, new_Fso().BuildPath(toto, ft1))
        AssertEqualWithMessage e, a, "cont tofolder-tofile1"
        assertExistsFolder new_Fso().BuildPath(toto, ft3), True, "after", "copy/move", "tofolder-tofolder3"
        e = df1
        a = readTestFile(c, new_Fso().BuildPath(toto, ff1))
        AssertEqualWithMessage e, a, "cont tofolder-fromfile1"
        e = dt2
        a = readTestFile(c, new_Fso().BuildPath(toto, ff2))
        AssertEqualWithMessage e, a, "cont tofolder-fromfile2"
        assertExistsFolder new_Fso().BuildPath(toto, ff3), False, "after", "copy/move", "tofolder-fromfolder3"
    Else
        assertExistsFolder toto, True, "after", "copy/moveFolder", "toFolder"
        assertFilesSubfoldersCount toto, 2, 1, "to"
        e = dt1
        a = readTestFile(c, new_Fso().BuildPath(toto, ft1))
        AssertEqualWithMessage e, a, "cont tofolder-fromfile1"
        e = dt2
        a = readTestFile(c, new_Fso().BuildPath(toto, ft2))
        AssertEqualWithMessage e, a, "cont tofolder-fromfile2"
        assertExistsFolder new_Fso().BuildPath(toto, ft3), True, "after", "copy/moveFolder", "tofolder-fromfolder3"
    End If
End Sub

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
