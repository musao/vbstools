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

'###################################################################################################
'SetUp()/TearDown()
Sub SetUp()
    '���s�X�N���v�g�����ɓ��t�@�C�����ňꎞ�t�H���_�쐬
    PsPathTempFolder = new_Fso().BuildPath(new_Fso().GetParentFolderName(WScript.ScriptFullName), MY_NAME)
    If Not (new_Fso().FolderExists(PsPathTempFolder)) Then new_Fso().CreateFolder(PsPathTempFolder)
End Sub
Sub TearDown()
    '���e�X�g�ō쐬�����ꎞ�t�H���_���폜����
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
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))

    Dim c,d
    '�t�@�C�����쐬
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
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))

    Dim c,d
    '�t�H���_���쐬
    new_Fso().CreateFolder p
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
    '�t�@�C�����쐬
    c = "UTF-8"
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
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
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
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
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "For" & vbNewLine & "DeleteFile Err FileLocked"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "deleteFile", "file"

    Dim f
    '�t�@�C�������b�N����
    f = -1    'TristateTrue(Unicode)
    With lockFile(p, f)
        Dim e,ao
        e = False
        Set ao = fs_deleteFile(p)
        
        'fs_deleteFile()���G���[�ɂȂ邱�Ƃ��m�F����
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

    '�t�@�C�����폜����Ă��Ȃ����Ƃ��m�F
    assertExistsFile p, True, "after", "deleteFile", "file"
End Sub

'###################################################################################################
'fs_deleteFolder()
Sub Test_fs_deleteFolder
    Dim c,p,pf,d
    '�t�H���_���쐬
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder p
    '�t�H���_�̉��Ƀt�@�C�����쐬
    c = "UTF-8"
    pf = new_Fso().BuildPath(p, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
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
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
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
    '�t�H���_���쐬
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder p
    '�t�H���_�̉��Ƀt�@�C�����쐬
    c = "UTF-8"
    pf = new_Fso().BuildPath(p, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    d = "For" & vbNewLine & "DeleteFolder Err FileLocked"
    writeTestFile c,pf,d

    Dim f
    '�t�@�C�������b�N����
    f = -1    'TristateTrue(Unicode)
    With lockFile(pf,f)
        Dim e,ao
        e = False
        Set ao = fs_deleteFolder(p)
        
        'fs_deleteFolder()���G���[�ɂȂ邱�Ƃ��m�F����
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

    '�t�H���_���폜����Ă��Ȃ����Ƃ��m�F
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
    Dim c,p,d
    '�t�@�C�����쐬
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "lmn" & vbNewLine & "�V�Y�]" & vbNewLine & "���" & vbNewLine & ChrW(12316) 'ChrW(12316)='\u301c'�i�g�_�b�V���E�g�^�jSjis�ɕϊ��ł��Ȃ�����
    writeTestFile c,p,d

    Dim e,a
    e = d
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

    Dim d,e,a
    d = "abc" & vbNewLine & "������" & vbNewLine & "123" & vbNewLine & ChrW(12316) 'ChrW(12316)='\u301c'�i�g�_�b�V���E�g�^�jSjis�ɕϊ��ł��Ȃ�����
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
    '�㏑������t�@�C������U�쐬
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "UTF-8"
    d = "For" & vbNewLine & "Rewrite"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "writeFile", "file"

    '�㏑�����邱�Ƃ��m�F
    d = "abc" & vbNewLine & "�@�A�B" & vbNewLine & "!#$" & vbNewLine & ChrW(12316) 'ChrW(12316)='\u301c'�i�g�_�b�V���E�g�^�jSjis�ɕϊ��ł��Ȃ�����
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
    '���b�N����t�@�C������U�쐬
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "For" & vbNewLine & "Write Error"
    writeTestFile c,p,d

    Dim de,f
    '�t�@�C�������b�N����
    f = -1    'TristateTrue(Unicode)
    With lockFile(p,f)
        de = "error" & vbNewLine & "test"
        Dim e,a
        e = False
        a = fs_writeFile(p, de)
        
        'fs_writeFile()���G���[�ɂȂ邱�Ƃ��m�F����
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    '�㏑�����Ă��Ȃ����Ƃ��m�F
    e = d
    a = readTestFile(c, p)
    AssertEqualWithMessage e, a, "cont"
End Sub

'###################################################################################################
'fs_writeFileDefault()
Sub Test_fs_writeFileDefault
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "writeFileDefault", "file"

    Dim d,e,a
    d = "abc" & vbNewLine & "������" & vbNewLine & "123"
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
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "func_FsWriteFile", "file"

    Dim d,iomode,create,f,e,a
    d = "func_FsWriteFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Iomode_ForWriting_Normal__Format_SystemDefault"
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
    '�㏑������t�@�C������U�쐬
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "For" & vbNewLine & "Rewrite"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "func_FsWriteFile", "file"

    Dim iomode,create,f,e,a
    '�㏑�����邱�Ƃ��m�F
    iomode = 2     'ForWriting
    create = True
    f = -1    'TristateTrue(Unicode)
    d = "func_FsWriteFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Iomode_ForWriting_Rewrite__Format_Unicode"
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
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "func_FsWriteFile", "file"

    Dim d,iomode,create,f,e,a
    d = "func_FsWriteFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Iomode_ForAppending_Normal__Format_Ascii"
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
    '�ǋL����t�@�C������U�쐬
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "shift-jis"
    d = "For" & vbNewLine & "Append"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "func_FsWriteFile", "file"

    Dim iomode,create,f,ec,e,a
    '�ǋL���邱�Ƃ��m�F
    iomode = 8     'ForAppending
    create = True
    f = -2         'TristateUseDefault
    ec = d
    d = "func_FsWriteFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Iomode_ForAppending_Append__Format_SystemDefault"
    e = True
    a = func_FsWriteFile(p, iomode, create, f, d)
    AssertEqualWithMessage e, a, "ret"

    ec = ec & d
    a = readTestFile(c, p)
    AssertEqualWithMessage ec, a, "cont"
End Sub
Sub Test_func_FsWriteFile_Create_True_Normal__Format_Unicode
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "func_FsWriteFile", "file"

    Dim d,iomode,create,f,e,a
    d = "func_FsWriteFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Create_True_Normal__Format_Unicode"
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
    '�㏑������t�@�C������U�쐬
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "shift-jis"
    d = "For" & vbNewLine & "Rewrite"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "func_FsWriteFile", "file"

    Dim iomode,create,f,e,a
    '�㏑�����邱�Ƃ��m�F
    iomode = 2     'ForWriting
    create = True
    f = 0          'TristateFalse(Ascii)
    d = "func_FsWriteFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Create_True_Rewrite__Format_Ascii"
    e = True
    a = func_FsWriteFile(p, iomode, create, f, d)
    AssertEqualWithMessage e, a, "ret"

    e = d
    a = readTestFile(c, p)
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_func_FsWriteFile_Create_False_Err
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile p, False, "before", "func_FsWriteFile", "file"

    Dim d,iomode,create,f,e,a
    d = "func_FsWriteFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Create_False_Err"
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
    '�㏑������t�@�C������U�쐬
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "For" & vbNewLine & "Rewrite"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "func_FsWriteFile", "file"

    Dim e,a,iomode,create,f
    '�㏑�����邱�Ƃ��m�F
    iomode = 2     'ForWriting
    create = False
    f = -1         'TristateTrue(Unicode)
    d = "func_FsWriteFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Create_False_Rewrite__Format_Unicode"
    e = True
    a = func_FsWriteFile(p, iomode, create, f, d)
    AssertEqualWithMessage e, a, "ret"

    e = d
    a = readTestFile(c, p)
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub Test_func_FsWriteFile_Err_FileLocked
    Dim p,d,c
    '���b�N����t�@�C������U�쐬
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "shift-jis"
    d = "error" & vbNewLine & "FileLocked"
    writeTestFile c,p,d
    assertExistsFile p, True, "before", "func_FsWriteFile", "file"

    Dim f
    f = 0          'TristateFalse(Ascii)
    '�t�@�C�������b�N����
    With lockFile(p,f)

        Dim iomode,create,de,e,a
        iomode = 2     'ForWriting
        create = False
        de = "error" & vbNewLine & "test"
        e = False
        a = func_FsWriteFile(p, iomode, create, f, de)
        
        'func_FsWriteFile()���G���[�ɂȂ邱�Ƃ��m�F����
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    '�㏑�����Ă��Ȃ����Ƃ��m�F
    e = d
    a = readTestFile(c, p)
    AssertEqualWithMessage e, a, "cont"
End Sub

'###################################################################################################
'func_FsReadFile()
Sub Test_func_FsReadFile_Normal__Format_SystemDefault
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))

    Dim d,f,c
    d = "func_FsReadFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Normal__Format_SystemDefault"
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
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))

    Dim d,f,c
    d = "func_FsReadFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Normal__Format_Unicode"
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
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))

    Dim d,f,c
    d = "func_FsReadFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Normal__Format_Ascii"
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
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
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
Sub com_CopyOrMoveFile_Normal(IsCopy)
    Dim from
    'from�p�X���쐬
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    Dim c,df
    'from�t�@�C�����쐬
    c = "Unicode"
    df = "For" & vbNewLine & "copy/moveFile Normal"
    writeTestFile c,from,df
    assertExistsFile from, True, "before", "copy/moveFile", "fromfile"
    
    Dim toto
    'to�p�X���쐬
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_.txt"))
    assertExistsFile toto, False, "before", "copy/moveFile", "tofile"

    Dim e,a
    '���s���ʂ̊m�F
    e = True
    If IsCopy Then
        a = fs_copyFile(from,toto)
    Else
        a = fs_moveFile(from,toto)
    End If
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    'from�t�@�C���̊m�F
    If isCopy Then
        assertExistsFile from, True, "after", "copy/moveFile", "fromfile"
        e = df
        a = readTestFile(c, from)
        AssertEqualWithMessage e, a, "cont"
    Else
        assertExistsFile from, False, "after", "copy/moveFile", "fromfile"
    End If

    'to�t�@�C���̊m�F
    assertExistsFile toto, True, "after", "copy/moveFile", "tofile"
    e = df
    a = readTestFile(c, toto)
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub com_CopyOrMoveFile_OverRide(IsCopy)
    Dim from
    'from�p�X���쐬
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    Dim c,df
    'from�t�@�C�����쐬
    c = "Unicode"
    df = "For" & vbNewLine & "copy/moveFile OverRide"
    writeTestFile c,from,df
    assertExistsFile from, True, "before", "copy/moveFile", "fromfile"
    
    Dim toto
    'to�p�X���쐬
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_.txt"))
    Dim dt
    'to�t�@�C�����쐬
    c = "Unicode"
    dt = "For" & vbNewLine & "copy/moveFile ToFile"
    writeTestFile c,toto,dt
    assertExistsFile toto, True, "before", "copy/moveFile", "tofile"

    'copy�͐���i�㏑������jmove�ُ͈�i�㏑�����Ȃ��j�ɂȂ�
    Dim e,a
    '���s���ʂ̊m�F
    If IsCopy Then
        e = True
        a = fs_copyFile(from,toto)
    Else
        e = False
        a = fs_moveFile(from,toto)
    End If
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    'from�t�@�C���̊m�F
    If isCopy Then
        assertExistsFile from, True, "after", "copy/moveFile", "fromfile"
        e = df
        a = readTestFile(c, from)
        AssertEqualWithMessage e, a, "cont"
    Else
        assertExistsFile from, True, "after", "copy/moveFile", "fromfile"
    End If

    'to�t�@�C���̊m�F
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
    Dim from
    'from�p�X���쐬
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    Dim c,df
    'from�t�@�C�����쐬
    c = "Unicode"
    df = "For" & vbNewLine & "copy/moveFile FromFileLocked"
    writeTestFile c,from,df
    assertExistsFile from, True, "before", "copy/moveFile", "fromfile"

    Dim toto
    'to�p�X���쐬
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_.txt"))
    assertExistsFile toto, False, "before", "copy/moveFile", "tofile"

    Dim f
    'from�t�@�C�������b�N����
    f = -1    'TristateTrue(Unicode)
    With lockFile(from,f)
        Dim e,a
        '���s���ʂ̊m�F
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

    'from�t�@�C���̊m�F
    If isCopy Then
        assertExistsFile from, True, "after", "copy/moveFile", "fromfile"
        e = df
        a = readTestFile(c, from)
        AssertEqualWithMessage e, a, "cont"
    Else
        assertExistsFile from, True, "after", "copy/moveFile", "fromfile"
    End If

    'to�t�@�C���̊m�F
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
    'from�p�X���쐬
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    assertExistsFile from, False, "before", "copy/moveFile", "fromfile"
    
    Dim toto
    'to�p�X���쐬
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_.txt"))
    assertExistsFile toto, False, "before", "copy/moveFile", "tofile"

    Dim e,a
    '���s���ʂ̊m�F
    e = False
    If isCopy Then
        a = fs_copyFile(from,toto)
    Else
        a = fs_moveFile(from,toto)
    End If
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    'from�t�@�C���̊m�F
    assertExistsFile from, False, "after", "copy/moveFile", "fromfile"

    'to�t�@�C���̊m�F
    assertExistsFile toto, False, "after", "copy/moveFile", "tofile"
End Sub
Sub com_CopyOrMoveFile_ToFileLocked(IsCopy)
    Dim from
    'from�p�X���쐬
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    Dim c,df
    'from�t�@�C�����쐬
    c = "Unicode"
    df = "For" & vbNewLine & "copy/moveFile fromfile ToFileLocked"
    writeTestFile c,from,df
    assertExistsFile from, True, "before", "copy/move", "fromfile"
    
    Dim toto
    'to�p�X���쐬
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_.txt"))
    Dim dt
    'to�t�@�C�����쐬
    dt = "For" & vbNewLine & "copy/moveFile ToFileLocked"
    writeTestFile c,toto,dt
    assertExistsFile toto, True, "before", "copy/moveFile", "tofile"

    Dim f
    'to�t�@�C�������b�N����
    f = -1    'TristateTrue(Unicode)
    With lockFile(toto,f)
        assertExistsFile toto, True, "before", "copy/move", "tofile"
        
        Dim e,a
        '���s���ʂ̊m�F
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

    'from�t�@�C���̊m�F
    assertExistsFile from, True, "after", "copy/move", "fromfile"
    e = df
    a = readTestFile(c, from)
    AssertEqualWithMessage e, a, "cont"

    'to�t�@�C���̊m�F
    assertExistsFile toto, True, "after", "copy/move", "tofile"
    e = dt
    a = readTestFile(c, toto)
    AssertEqualWithMessage e, a, "cont"
End Sub
Sub com_CopyOrMoveFolder_Normal(IsCopy)
    Dim from
    'from�t�H���_���쐬
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder from
    Dim c,p,ff1,ff2,ff3,df1,df2
    'from�t�H���_�̉��Ƀt�@�C���ƃt�H���_���쐬
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
    'to�p�X���쐬�i�t�H���_�͍쐬���Ȃ��j
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_"))
    assertExistsFolder toto, False, "before", "copy/moveFolder", "toFolder"

    Dim e,a
    '���s���ʂ̊m�F
    e = True
    If isCopy Then
        a = fs_copyFolder(from,toto)
    Else
        a = fs_moveFolder(from,toto)
    End If
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    'from�t�H���_�̊m�F
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

    'to�t�H���_�̊m�F
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
    'from�t�H���_���쐬
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder from
    Dim c,p,ff1,ff2,ff3,df1,df2
    'from�t�H���_�̉��Ƀt�@�C���ƃt�H���_���쐬
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
    'to�t�H���_���쐬
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_"))
    new_Fso().CreateFolder toto
    Dim ft1,ft2,ft3,dt1,dt2
    'to�t�H���_�̉��Ƀt�@�C���ƃt�H���_���쐬
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

    'copy�͐���i�㏑������jmove�ُ͈�i�㏑�����Ȃ��j�ɂȂ�
    Dim e,a
    '���s���ʂ̊m�F
    If isCopy Then
        e = True
        a = fs_copyFolder(from,toto)
    Else
        e = False
        a = fs_moveFolder(from,toto)
    End If
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    'from�t�H���_�̊m�F
    assertExistsFolder from, True, "after", "copy/moveFolder", "fromFolder"
    assertFilesSubfoldersCount from, 2, 1, "from"
    e = df1
    a = readTestFile(c, new_Fso().BuildPath(from, ff1))
    AssertEqualWithMessage e, a, "cont file1"
    e = df2
    a = readTestFile(c, new_Fso().BuildPath(from, ff2))
    AssertEqualWithMessage e, a, "cont file2"
    assertExistsFolder new_Fso().BuildPath(from, ff3), True, "after", "copy/moveFolder", "fromfolder-folder3"

    'to�t�H���_�̊m�F
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
    'from�t�H���_���쐬
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder from
    Dim c,p,ff1,ff2,ff3,df1,df2
    'from�t�H���_�̉��Ƀt�@�C���ƃt�H���_���쐬
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
    'to�t�H���_���쐬
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_"))
    new_Fso().CreateFolder toto
    Dim ft1,ft2,ft3,dt1,dt2
    'to�t�H���_�̉��Ƀt�@�C���ƃt�H���_���쐬
    ft1 = new_Now().formatAs("YYMMDD_hhmmss.000000_t1.txt")
    dt1 = "For" & vbNewLine & "copy/moveFolder OverRideWithUnrelatedFileLocked ft1"
    p = new_Fso().BuildPath(toto, ft1)
    writeTestFile c,p,dt1
    ft2 = ff2
    dt2 = "For" & vbNewLine & "copy/moveFolder OverRideWithUnrelatedFileLocked ft2"
    p = new_Fso().BuildPath(toto, ft2)
    writeTestFile c,p,dt2
    ft3 = new_Now().formatAs("YYMMDD_hhmmss.000000_t3")
    p = new_Fso().BuildPath(toto, ft3)
    new_Fso().CreateFolder p
    assertExistsFolder toto, True, "before", "copy/moveFolder", "tofolder"
    assertExistsFile new_Fso().BuildPath(toto, ft1), True, "before", "copy/moveFolder", "tofolder-file1"
    assertExistsFile new_Fso().BuildPath(toto, ft2), True, "before", "copy/moveFolder", "tofolder-file2"
    assertExistsFolder new_Fso().BuildPath(toto, ft3), True, "before", "copy/moveFolder", "tofolder-folder3"
    
    Dim f
    p = new_Fso().BuildPath(toto, ft1)
    f = -1    'TristateTrue(Unicode)
    'to�t�H���_�̃t�@�C���ift1�j�����b�N����
    With lockFile(p,f)
        Dim e,a
        '���s���ʂ̊m�F
        If isCopy Then
            e = True
            a = fs_copyFolder(from,toto)
        Else
            e = False
            a = fs_moveFolder(from,toto)
        End If
        
        'fs_copyFolder()/fs_moveFolder()���G���[�ɂȂ�Ȃ����Ƃ��m�F����
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    'to�t�H���_�̃t�@�C���ift1�j�����b�N���Ă��邪�A�㏑�����Ȃ�����copy�͐���Ɋ�������Amove�̓G���[�ɂȂ�

    'from�t�H���_�̊m�F
    assertExistsFolder from, True, "after", "copy/moveFolder", "fromFolder"
    assertFilesSubfoldersCount from, 2, 1, "from"
    e = df1
    a = readTestFile(c, new_Fso().BuildPath(from, ff1))
    AssertEqualWithMessage e, a, "cont file1"
    e = df2
    a = readTestFile(c, new_Fso().BuildPath(from, ff2))
    AssertEqualWithMessage e, a, "cont file2"
    assertExistsFolder new_Fso().BuildPath(from, ff3), True, "after", "copy/moveFolder", "fromfolder-folder3"

    'to�t�H���_�̊m�F
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
    Dim from
    'from�t�H���_���쐬
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder from
    Dim c,ff1,ff2,ff3,df1,df2
    'from�t�H���_�̉��Ƀt�@�C���ƃt�H���_���쐬
    c = "Unicode"
    ff1 = new_Now().formatAs("YYMMDD_hhmmss.000000_f1.txt")
    df1 = "For" & vbNewLine & "copy/moveFolder FromFileLocked ff1"
    p = new_Fso().BuildPath(from, ff1)
    writeTestFile c,p,df1
    ff2 = new_Now().formatAs("YYMMDD_hhmmss.000000_f2.txt")
    df2 = "For" & vbNewLine & "copy/moveFolder FromFileLocked ff2"
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
    'to�t�H���_���쐬
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_"))
    new_Fso().CreateFolder toto
    Dim p,ft1,ft2,ft3,dt1,dt2
    'to�t�H���_�̉��Ƀt�@�C���ƃt�H���_���쐬
    ft1 = new_Now().formatAs("YYMMDD_hhmmss.000000_t1.txt")
    dt1 = "For" & vbNewLine & "copy/moveFolder FromFileLocked ft1"
    p = new_Fso().BuildPath(toto, ft1)
    writeTestFile c,p,dt1
    ft2 = ff2
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

    Dim f
    p = new_Fso().BuildPath(from, ff2)
    f = -1    'TristateTrue(Unicode)
    'from�t�H���_�̃t�@�C���iff2�j�����b�N����
    With lockFile(p,f)
        'copy�͐���i�㏑������jmove�ُ͈�i�㏑�����Ȃ��j�ɂȂ�
        Dim e,a
        '���s���ʂ̊m�F
        If isCopy Then
            e = True
            a = fs_copyFolder(from,toto)
        Else
            e = False
            a = fs_moveFolder(from,toto)
        End If
        
        'fs_copyFolder()/fs_moveFolder()���G���[�ɂȂ�Ȃ����Ƃ��m�F����
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    'from�t�H���_�̃t�@�C���iff2�j�����b�N���Ă���ƁAcopy�͐���Amove���G���[�ɂȂ�

    'from�t�H���_�̊m�F
    assertExistsFolder from, True, "after", "copy/moveFolder", "fromFolder"
    assertFilesSubfoldersCount from, 2, 1, "from"
    e = df1
    a = readTestFile(c, new_Fso().BuildPath(from, ff1))
    AssertEqualWithMessage e, a, "cont file1"
    e = df2
    a = readTestFile(c, new_Fso().BuildPath(from, ff2))
    AssertEqualWithMessage e, a, "cont file2"
    assertExistsFolder new_Fso().BuildPath(from, ff3), True, "after", "copy/moveFolder", "fromfolder-folder3"

    'to�t�H���_�̊m�F
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
    'from�p�X���쐬�i�t�H���_�͍쐬���Ȃ��j
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    assertExistsFolder from, False, "before", "copy/moveFolder", "fromfolder"
    
    Dim toto
    'to�p�X���쐬�i�t�H���_�͍쐬���Ȃ��j
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_"))
    assertExistsFolder toto, False, "before", "copy/moveFolder", "tofolder"

    Dim e,a
    '���s���ʂ̊m�F
    e = False
    If isCopy Then
        a = fs_copyFolder(from,toto)
    Else
        a = fs_moveFolder(from,toto)
    End If
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"

    'from�t�H���_�̊m�F
    assertExistsFolder from, False, "after", "copy/moveFolder", "fromFolder"

    'to�t�H���_�̊m�F
    assertExistsFolder toto, False, "after", "copy/moveFolder", "toFolder"
End Sub
Sub com_CopyOrMoveFolder_ToFileLocked(IsCopy)
    Dim from
    'from�t�H���_���쐬
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder from
    Dim c,p,ff1,ff2,ff3,df1,df2
    'from�t�H���_�̉��Ƀt�@�C���ƃt�H���_���쐬
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
    '�R�s�[��t�H���_���쐬
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_"))
    new_Fso().CreateFolder toto
    Dim ft1,ft2,ft3,dt1,dt2
    '�t�H���_�̉��Ƀt�@�C���ƃt�H���_���쐬
    ft1 = new_Now().formatAs("YYMMDD_hhmmss.000000_t1.txt")
    dt1 = "For" & vbNewLine & "copy/moveFolder ToFileLocked ft1"
    p = new_Fso().BuildPath(toto, ft1)
    writeTestFile c,p,dt1
    ft2 = ff2
    dt2 = "For" & vbNewLine & "copy/moveFolder ToFileLocked ft2"
    p = new_Fso().BuildPath(toto, ft2)
    writeTestFile c,p,dt2
    ft3 = new_Now().formatAs("YYMMDD_hhmmss.000000_t3")
    p = new_Fso().BuildPath(toto, ft3)
    new_Fso().CreateFolder p
    assertExistsFolder toto, True, "before", "copy/moveFolder", "tofolder"
    assertExistsFile new_Fso().BuildPath(toto, ft1), True, "before", "copy/moveFolder", "tofolder-file1"
    assertExistsFile new_Fso().BuildPath(toto, ft2), True, "before", "copy/moveFolder", "tofolder-file2"
    assertExistsFolder new_Fso().BuildPath(toto, ft3), True, "before", "copy/moveFolder", "tofolder-folder3"

    Dim f
    p = new_Fso().BuildPath(toto, ft2)
    f = -1    'TristateTrue(Unicode)
    'to�t�H���_�̃t�@�C���ift2�j�����b�N����
    With lockFile(p,f)
        Dim e,a
        '���s���ʂ̊m�F
        e = False
        If isCopy Then
            a = fs_copyFolder(from,toto)
        Else
            a = fs_moveFolder(from,toto)
        End If
        
        'fs_copyFolder()/fs_moveFolder()���G���[�ɂȂ邱�Ƃ��m�F����
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    '1�ڂ̃t�@�C��fn1�̓R�s�[�܂��͈ړ����A2�ڂ̃t�@�C��fn2�̃R�s�[�܂��͈ړ������s����A3�ڂ̃t�H���_fn3�̓R�s�[�܂��͈ړ����Ȃ�

    'from�t�H���_�̊m�F
    assertExistsFolder from, True, "after", "copy/move", "fromFolder"
    assertFilesSubfoldersCount from, 2, 1, "from"
    e = df1
    a = readTestFile(c, new_Fso().BuildPath(from, ff1))
    AssertEqualWithMessage e, a, "cont file1"
    e = df2
    a = readTestFile(c, new_Fso().BuildPath(from, ff2))
    AssertEqualWithMessage e, a, "cont file2"
    assertExistsFolder new_Fso().BuildPath(from, ff3), True, "after", "copy/move", "fromfolder-folder3"

    'to�t�H���_�̊m�F
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
Function lockFile(p,f)
    Set lockFile = new_Ts(p, 8, True, f)
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


' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
