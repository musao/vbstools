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

'###################################################################################################
'fs_copyFolder()

'###################################################################################################
'fs_createFolder()

'###################################################################################################
'fs_deleteFile()
Sub Test_fs_deleteFile
    Dim c,p,d
    'ファイルを作成
    c = "UTF-8"
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    d = "For" & vbNewLine & "DeleteFile Normal"
    writeTestFile c,p,d
    AssertEqualWithMessage True, new_Fso().FileExists(p), "before delete file exists"

    Dim e,a
    e = True
    a = fs_deleteFile(p)
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage False, new_Fso().FileExists(p), "after delete file exists"
End Sub
Sub Test_fs_deleteFile_Err_NotExists
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    AssertEqualWithMessage False, new_Fso().FileExists(p), "before delete file exists"

    Dim e,a
    e = False
    a = fs_deleteFile(p)
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage False, new_Fso().FileExists(p), "after delete file exists"
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
    AssertEqualWithMessage True, new_Fso().FileExists(p), "before delete file exists"
End Sub

'###################################################################################################
'fs_deleteFolder()
Sub Test_fs_deleteFolder
    Dim c,p,fp,d
    'フォルダを作成
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder p
    'フォルダの下にファイルを作成
    c = "UTF-8"
    fp = new_Fso().BuildPath(p, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    d = "For" & vbNewLine & "DeleteFolder Normal"
    writeTestFile c,fp,d
    AssertEqualWithMessage True, new_Fso().FolderExists(p), "before delete folder exists"

    Dim e,a
    e = True
    a = fs_deleteFolder(p)
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage False, new_Fso().FolderExists(p), "after delete folder exists"
End Sub

'###################################################################################################
'fs_moveFile()

'###################################################################################################
'fs_moveFolder()

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
    AssertEqualWithMessage False, new_Fso().FileExists(p), "before read file exists"

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
    AssertEqualWithMessage False, new_Fso().FileExists(p), "before write file exists"

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
    AssertEqualWithMessage True, new_Fso().FileExists(p), "before write file exists"

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
    AssertEqualWithMessage False, new_Fso().FileExists(p), "before write file exists"

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
    AssertEqualWithMessage False, new_Fso().FileExists(p), "before write file exists"

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
    AssertEqualWithMessage True, new_Fso().FileExists(p), "before write file exists"

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
    AssertEqualWithMessage False, new_Fso().FileExists(p), "before write file exists"

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
    AssertEqualWithMessage True, new_Fso().FileExists(p), "before write file exists"

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
    AssertEqualWithMessage False, new_Fso().FileExists(p), "before write file exists"

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
    AssertEqualWithMessage True, new_Fso().FileExists(p), "before write file exists"

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
    AssertEqualWithMessage False, new_Fso().FileExists(p), "before write file exists"

    Dim d,iomode,create,f,e,a
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Create_False_Err"
    iomode = 2     'ForWriting
    create = False
    f = -1         'TristateTrue(Unicode)
    e = False
    a = func_FsWriteFile(p, iomode, create, f, d)

    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage False, new_Fso().FileExists(p), "after write file exists"
End Sub
Sub Test_func_FsWriteFile_Create_False_Rewrite__Format_Unicode
    Dim p,c,d
    '上書きするファイルを一旦作成
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "For" & vbNewLine & "Rewrite"
    writeTestFile c,p,d
    AssertEqualWithMessage True, new_Fso().FileExists(p), "before write file exists"

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
        AssertEqualWithMessage True, new_Fso().FileExists(p), "before write file exists"

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
    AssertEqualWithMessage False, new_Fso().FileExists(p), "before read file exists"
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

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
