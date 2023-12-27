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
'fs_writeFile()
Sub Test_fs_writeFile
    Dim path,ec,ea,d,a,cont
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    AssertEqualWithMessage False, new_Fso().FileExists(path), "before write file exists"

    d = "abc" & vbNewLine & "あいう" & vbNewLine & "123" & vbNewLine & ChrW(12316) 'ChrW(12316)='\u301c'（波ダッシュ・波型）Sjisに変換できない文字
    ec = d : ea = True
    a = fs_writeFile(path, d)
    With CreateObject("ADODB.Stream")
        .Charset = "Unicode"
        .Open
        .LoadFromFile path
        cont = .ReadText
        .Close
    End With

    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, cont, "cont"
End Sub
Sub Test_fs_writeFile_Rewrite
    Dim path,ec,ea,d,a,cont
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    '適当なファイルを一旦作成
    d = "For" & vbNewLine & "Rewrite"
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .WriteText d, 0
        .SaveToFile path, 2
        .Close
    End With
    AssertEqualWithMessage True, new_Fso().FileExists(path), "before write file exists"

    '上書きすることを確認
    d = "abc" & vbNewLine & "①②③" & vbNewLine & "!#$" & vbNewLine & ChrW(12316) 'ChrW(12316)='\u301c'（波ダッシュ・波型）Sjisに変換できない文字
    ec = d : ea = True
    a = fs_writeFile(path, d)
    With CreateObject("ADODB.Stream")
        .Charset = "Unicode"
        .Open
        .LoadFromFile path
        cont = .ReadText
        .Close
    End With

    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, cont, "cont"
End Sub
Sub Test_fs_writeFile_Err
    Dim path,ec,ea,d,a,cont
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    d = "For" & vbNewLine & "Write Error"
    ec = d
    '適当なファイルを一旦作成
    With CreateObject("ADODB.Stream")
        .Charset = "Unicode"
        .Open
        .WriteText d, 0
        .SaveToFile path, 2
        .Close
    End With
    
    'TextstreamをAppendモードで作成し閉じない状態でfs_writeFile()実行エラーにする
    With new_Ts(path, 8, True, -1)
        d = "error" & vbNewLine & "test"
        ea = False
        a = fs_writeFile(path, d)
        
        AssertEqualWithMessage ea, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    '上書きしていないことを確認
    With CreateObject("ADODB.Stream")
        .Charset = "Unicode"
        .Open
        .LoadFromFile path
        cont = .ReadText
        .Close
    End With
    AssertEqualWithMessage ec, cont, "cont"
End Sub

'###################################################################################################
'fs_readFile()
Sub Test_fs_readFile
    Dim path,e,d,a
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    d = "lmn" & vbNewLine & "ⅢⅥⅩ" & vbNewLine & "ｱｲｳ" & vbNewLine & ChrW(12316) 'ChrW(12316)='\u301c'（波ダッシュ・波型）Sjisに変換できない文字
    e = d
    'ファイルを作成
    With CreateObject("ADODB.Stream")
        .Charset = "Unicode"
        .Open
        .WriteText d, 0
        .SaveToFile path, 2
        .Close
    End With
    a = fs_readFile(path)

    AssertEqualWithMessage e, a, "ret"
End Sub
Sub Test_fs_readFile_Err
    Dim path,e,a
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    AssertEqualWithMessage False, new_Fso().FileExists(path), "before read file exists"

    e = empty
    a = fs_readFile(path)
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"
End Sub

'###################################################################################################
'fs_deleteFile()
Sub Test_fs_deleteFile
    Dim path,e,d,a,ret
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    'ファイルを作成
    d = "For" & vbNewLine & "Delete Normal"
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .WriteText d, 0
        .SaveToFile path, 2
        .Close
    End With
    AssertEqualWithMessage True, new_Fso().FileExists(path), "before delete file exists"

    e = True
    a = fs_deleteFile(path)
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage False, new_Fso().FileExists(path), "after delete file exists"
End Sub
Sub Test_fs_deleteFile_Err_NotExists
    Dim path,e,d,a,ret
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    AssertEqualWithMessage False, new_Fso().FileExists(path), "before delete file exists"

    e = False
    a = fs_deleteFile(path)
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage False, new_Fso().FileExists(path), "after delete file exists"
End Sub
Sub Test_fs_deleteFile_Err_FileLocked
    Dim path,e,d,a
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    d = "For" & vbNewLine & "Delete Err FileLocked"
    '適当なファイルを一旦作成
    With CreateObject("ADODB.Stream")
        .Charset = "Unicode"
        .Open
        .WriteText d, 0
        .SaveToFile path, 2
        .Close
    End With
    
    'TextstreamをAppendモードで作成し閉じない状態でfs_deleteFile()実行エラーにする
    With new_Ts(path, 8, True, -1)
        e = False
        a = fs_deleteFile(path)
        
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    'ファイルが削除されていないことを確認
    AssertEqualWithMessage True, new_Fso().FileExists(path), "before delete file exists"
End Sub

'###################################################################################################
'func_FsWriteFile()
Sub Test_func_FsWriteFile_Iomode_ForWriting_Normal__Format_SystemDefault
    Dim path,ec,ea,d,a,cont,iomode,create,format,charset
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Iomode_ForWriting_Normal__Format_SystemDefault"
    iomode = 2     'ForWriting
    create = True
    format = -2    'TristateUseDefault
    charset = "shift-jis"
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    AssertEqualWithMessage False, new_Fso().FileExists(path), "before write file exists"
    ec = d : ea = True
    a = func_FsWriteFile(path, iomode, create, format, d)
    With CreateObject("ADODB.Stream")
        .Charset = charset
        .Open
        .LoadFromFile path
        cont = .ReadText
        .Close
    End With

    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, cont, "cont"
End Sub
Sub Test_func_FsWriteFile_Iomode_ForWriting_Rewrite__Format_Unicode
    Dim path,ec,ea,d,a,cont,iomode,create,format,charset
    iomode = 2     'ForWriting
    create = True
    format = -1    'TristateTrue(Unicode)
    charset = "Unicode"
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    '適当なファイルを一旦作成
    d = "For" & vbNewLine & "Rewrite"
    With CreateObject("ADODB.Stream")
        .Charset = charset
        .Open
        .WriteText d, 0
        .SaveToFile path, 2
        .Close
    End With
    AssertEqualWithMessage True, new_Fso().FileExists(path), "before write file exists"

    '上書きすることを確認
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Iomode_ForWriting_Rewrite__Format_Unicode"
    ec = d : ea = True
    a = func_FsWriteFile(path, iomode, create, format, d)
    With CreateObject("ADODB.Stream")
        .Charset = charset
        .Open
        .LoadFromFile path
        cont = .ReadText
        .Close
    End With
    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, cont, "cont"
End Sub
Sub Test_func_FsWriteFile_Iomode_ForAppending_Normal__Format_Ascii
    Dim path,ec,ea,d,a,cont,iomode,create,format,charset
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Iomode_ForAppending_Normal__Format_Ascii"
    iomode = 8     'ForAppending
    create = True
    format = 0     'TristateTrue(Ascii)
    charset = "shift-jis"
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    AssertEqualWithMessage False, new_Fso().FileExists(path), "before write file exists"
    ec = d : ea = True
    a = func_FsWriteFile(path, iomode, create, format, d)
    With CreateObject("ADODB.Stream")
        .Charset = charset
        .Open
        .LoadFromFile path
        cont = .ReadText
        .Close
    End With

    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, cont, "cont"
End Sub
Sub Test_func_FsWriteFile_Iomode_ForAppending_Append__Format_SystemDefault
    Dim path,ec,ea,d,a,cont,iomode,create,format,charset
    iomode = 8     'ForAppending
    create = True
    format = -2    'TristateUseDefault
    charset = "shift-jis"
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    '適当なファイルを一旦作成
    d = "For" & vbNewLine & "Append"
    ec = d
    With CreateObject("ADODB.Stream")
        .Charset = charset
        .Open
        .WriteText d, 0
        .SaveToFile path, 2
        .Close
    End With
    AssertEqualWithMessage True, new_Fso().FileExists(path), "before write file exists"

    '追記することを確認
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Iomode_ForAppending_Append__Format_SystemDefault"
    ec = ec & d : ea = True
    a = func_FsWriteFile(path, iomode, create, format, d)
    With CreateObject("ADODB.Stream")
        .Charset = charset
        .Open
        .LoadFromFile path
        cont = .ReadText
        .Close
    End With
    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, cont, "cont"
End Sub
Sub Test_func_FsWriteFile_Create_True_Normal__Format_Unicode
    Dim path,ec,ea,d,a,cont,iomode,create,format,charset
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Create_True_Normal__Format_Unicode"
    iomode = 2     'ForWriting
    create = True
    format = -1    'TristateTrue(Unicode)
    charset = "Unicode"
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    AssertEqualWithMessage False, new_Fso().FileExists(path), "before write file exists"
    ec = d : ea = True
    a = func_FsWriteFile(path, iomode, create, format, d)
    With CreateObject("ADODB.Stream")
        .Charset = charset
        .Open
        .LoadFromFile path
        cont = .ReadText
        .Close
    End With

    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, cont, "cont"
End Sub
Sub Test_func_FsWriteFile_Create_True_Rewrite__Format_Ascii
    Dim path,ec,ea,d,a,cont,iomode,create,format,charset
    iomode = 2     'ForWriting
    create = True
    format = 0     'TristateTrue(Ascii)
    charset = "shift-jis"
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    '適当なファイルを一旦作成
    d = "For" & vbNewLine & "Rewrite"
    With CreateObject("ADODB.Stream")
        .Charset = charset
        .Open
        .WriteText d, 0
        .SaveToFile path, 2
        .Close
    End With
    AssertEqualWithMessage True, new_Fso().FileExists(path), "before write file exists"

    '上書きすることを確認
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Create_True_Rewrite__Format_Ascii"
    ec = d : ea = True
    a = func_FsWriteFile(path, iomode, create, format, d)
    With CreateObject("ADODB.Stream")
        .Charset = charset
        .Open
        .LoadFromFile path
        cont = .ReadText
        .Close
    End With
    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, cont, "cont"
End Sub
Sub Test_func_FsWriteFile_Create_False_Err
    Dim path,e,d,a,cont,iomode,create,format
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Create_False_Err"
    iomode = 2     'ForWriting
    create = False
    format = -1    'TristateTrue(Unicode)
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    AssertEqualWithMessage False, new_Fso().FileExists(path), "before write file exists"
    e = False
    a = func_FsWriteFile(path, iomode, create, format, d)

    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage False, new_Fso().FileExists(path), "after write file exists"
End Sub
Sub Test_func_FsWriteFile_Create_False_Rewrite__Format_Unicode
    Dim path,ec,ea,d,a,cont,iomode,create,format,charset
    iomode = 2     'ForWriting
    create = False
    format = -1    'TristateTrue(Unicode)
    charset = "Unicode"
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    '適当なファイルを一旦作成
    d = "For" & vbNewLine & "Rewrite"
    With CreateObject("ADODB.Stream")
        .Charset = charset
        .Open
        .WriteText d, 0
        .SaveToFile path, 2
        .Close
    End With
    AssertEqualWithMessage True, new_Fso().FileExists(path), "before write file exists"

    '上書きすることを確認
    d = "func_FsWriteFile" & vbNewLine & "のテスト" & vbNewLine & "Create_False_Rewrite__Format_Unicode"
    ec = d : ea = True
    a = func_FsWriteFile(path, iomode, create, format, d)
    With CreateObject("ADODB.Stream")
        .Charset = charset
        .Open
        .LoadFromFile path
        cont = .ReadText
        .Close
    End With
    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, cont, "cont"
End Sub
Sub Test_func_FsWriteFile_Err_FileLocked
    Dim path,ec,ea,d,a,cont,iomode,create,format,charset
    iomode = 2     'ForWriting
    create = False
    format = 0     'TristateTrue(Ascii)
    charset = "shift-jis"
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    '適当なファイルを一旦作成
    d = "error" & vbNewLine & "FileLocked"
    ec = d
    With CreateObject("ADODB.Stream")
        .Charset = charset
        .Open
        .WriteText d, 0
        .SaveToFile path, 2
        .Close
    End With
    AssertEqualWithMessage True, new_Fso().FileExists(path), "before write file exists"
    
    'TextstreamをAppendモードで作成し閉じない状態でfunc_FsWriteFile()実行エラーにする
    With new_Ts(path, 8, True, format)
        d = "error" & vbNewLine & "test"
        ea = False
        a = func_FsWriteFile(path, iomode, create, format, d)
        
        AssertEqualWithMessage ea, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    '上書きしていないことを確認
    With CreateObject("ADODB.Stream")
        .Charset = charset
        .Open
        .LoadFromFile path
        cont = .ReadText
        .Close
    End With
    AssertEqualWithMessage ec, cont, "cont"
End Sub

'###################################################################################################
'func_FsReadFile()
Sub Test_func_FsReadFile_Normal__Format_SystemDefault
    Dim path,e,d,a,format,charset
    d = "func_FsReadFile" & vbNewLine & "のテスト" & vbNewLine & "Normal__Format_SystemDefault"
    format = -2    'TristateUseDefault
    charset = "shift-jis"
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    e = d
    'ファイルを作成
    With CreateObject("ADODB.Stream")
        .Charset = charset
        .Open
        .WriteText d, 0
        .SaveToFile path, 2
        .Close
    End With
    a = func_FsReadFile(path,format)

    AssertEqualWithMessage e, a, "ret"
End Sub
Sub Test_func_FsReadFile_Normal__Format_Unicode
    Dim path,e,d,a,format,charset
    d = "func_FsReadFile" & vbNewLine & "のテスト" & vbNewLine & "Normal__Format_Unicode"
    format = -1    'TristateTrue(Unicode)
    charset = "Unicode"
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    e = d
    'ファイルを作成
    With CreateObject("ADODB.Stream")
        .Charset = charset
        .Open
        .WriteText d, 0
        .SaveToFile path, 2
        .Close
    End With
    a = func_FsReadFile(path,format)

    AssertEqualWithMessage e, a, "ret"
End Sub
Sub Test_func_FsReadFile_Normal__Format_Ascii
    Dim path,e,d,a,format,charset
    d = "func_FsReadFile" & vbNewLine & "のテスト" & vbNewLine & "Normal__Format_Ascii"
    format = 0     'TristateTrue(Ascii)
    charset = "shift-jis"
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    e = d
    'ファイルを作成
    With CreateObject("ADODB.Stream")
        .Charset = charset
        .Open
        .WriteText d, 0
        .SaveToFile path, 2
        .Close
    End With
    a = func_FsReadFile(path,format)

    AssertEqualWithMessage e, a, "ret"
End Sub
Sub Test_func_FsReadFile_Err
    Dim path,e,a,format
    format = -2    'TristateUseDefault
    path = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    AssertEqualWithMessage False, new_Fso().FileExists(path), "before read file exists"

    e = empty
    a = func_FsReadFile(path,format)
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"
End Sub


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
