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
Sub Test_fs_copyFile
    Dim from
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))

    Dim c,d
    '�t�@�C�����쐬
    c = "Unicode"
    d = "For" & vbNewLine & "copyFile Normal"
    writeTestFile c,from,d
    AssertEqualWithMessage True, new_Fso().FileExists(from), "before copy fromfile exists"
    
    Dim toto
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_.txt"))
    AssertEqualWithMessage False, new_Fso().FileExists(toto), "before copy tofile exists"

    Dim ea,ec,a,ct
    ea = True : ec = d
    a = fs_copyFile(from,toto)
    ct = readTestFile(c, toto)
    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, ct, "cont"
    AssertEqualWithMessage True, new_Fso().FileExists(from), "after copy fromfile exists"
    AssertEqualWithMessage True, new_Fso().FileExists(toto), "after copy tofile exists"
End Sub
Sub Test_fs_copyFile_OverRide
    Dim from
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))

    Dim c,df
    '�t�@�C�����쐬
    c = "Unicode"
    df = "For" & vbNewLine & "copyFile OverRide"
    writeTestFile c,from,df
    AssertEqualWithMessage True, new_Fso().FileExists(from), "before copy fromfile exists"
    
    Dim toto
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_.txt"))
    
    Dim dt
    '�t�@�C�����쐬
    c = "Unicode"
    dt = "For" & vbNewLine & "copyFile ToFile"
    writeTestFile c,toto,dt
    AssertEqualWithMessage True, new_Fso().FileExists(toto), "before copy tofile exists"

    Dim ea,ec,a,ct
    ea = True : ec = df
    a = fs_copyFile(from,toto)
    ct = readTestFile(c, toto)
    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage ec, ct, "cont"
    AssertEqualWithMessage True, new_Fso().FileExists(from), "after copy fromfile exists"
    AssertEqualWithMessage True, new_Fso().FileExists(toto), "after copy tofile exists"
End Sub
Sub Test_fs_copyFile_Err_FromFileNoExists
    Dim from
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    AssertEqualWithMessage False, new_Fso().FileExists(from), "before copy fromfile exists"
    
    Dim toto
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_.txt"))
    AssertEqualWithMessage False, new_Fso().FileExists(toto), "before copy tofile exists"

    Dim e,a,ct
    e = False
    a = fs_copyFile(from,toto)
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"
    AssertEqualWithMessage False, new_Fso().FileExists(from), "after copy fromfile exists"
    AssertEqualWithMessage False, new_Fso().FileExists(toto), "after copy tofile exists"
End Sub
Sub Test_fs_copyFile_Err_ToFileLocked
    Dim from
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))

    Dim c,df
    '�t�@�C�����쐬
    c = "Unicode"
    df = "For" & vbNewLine & "copyFile OverRide"
    writeTestFile c,from,df
    AssertEqualWithMessage True, new_Fso().FileExists(from), "before copy fromfile exists"
    
    Dim toto
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_.txt"))
    
    Dim dt,f
    dt = "For" & vbNewLine & "copyFile ToFile"
    f = -1    'TristateTrue(Unicode)
    '�t�@�C������U�쐬���ă��b�N����
    With createFileAndLocked(c,toto,dt,f)
        Dim ea,a
        ea = False
        a = fs_copyFile(from,toto)
        
        'fs_copyFile()���G���[�ɂȂ邱�Ƃ��m�F����
        AssertEqualWithMessage ea, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    Dim ec,ct
    ec = dt
    ct = readTestFile(c, toto)
    AssertEqualWithMessage ec, ct, "cont"
    AssertEqualWithMessage True, new_Fso().FileExists(from), "after copy fromfile exists"
    AssertEqualWithMessage True, new_Fso().FileExists(toto), "after copy tofile exists"
End Sub

'###################################################################################################
'fs_copyFolder()
Sub Test_fs_copyFolder
    Dim from
    '�R�s�[���t�H���_���쐬
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder from
    AssertEqualWithMessage True, new_Fso().FolderExists(from), "before copy fromfolder exists"
    
    Dim c,fp,fn1,fn2,fn3,df1,df2
    '�t�H���_�̉��Ƀt�@�C���ƃt�H���_���쐬
    c = "Unicode"
    fn1 = new_Now().formatAs("YYMMDD_hhmmss.000000_f1.txt")
    df1 = "For" & vbNewLine & "copyFolder Normal fn1"
    fp = new_Fso().BuildPath(from, fn1)
    writeTestFile c,fp,df1
    fn2 = new_Now().formatAs("YYMMDD_hhmmss.000000_f2.txt")
    df2 = "For" & vbNewLine & "copyFolder Normal fn2"
    fp = new_Fso().BuildPath(from, fn2)
    writeTestFile c,fp,df2
    fn3 = new_Now().formatAs("YYMMDD_hhmmss.000000_f3")
    fp = new_Fso().BuildPath(from, fn3)
    new_Fso().CreateFolder fp
    AssertEqualWithMessage True, new_Fso().FileExists(new_Fso().BuildPath(from, fn1)), "before copy fromfolderfile1 exists"
    AssertEqualWithMessage True, new_Fso().FileExists(new_Fso().BuildPath(from, fn2)), "before copy fromfolderfile2 exists"
    AssertEqualWithMessage True, new_Fso().FolderExists(new_Fso().BuildPath(from, fn3)), "before copy fromfolderfolder3 exists"
    
    Dim toto
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_"))
    AssertEqualWithMessage False, new_Fso().FileExists(toto), "before copy tofile exists"

    Dim ea,a
    ea = True
    a = fs_copyFolder(from,toto)
    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage 2, new_Fso().GetFolder(from).Files.Count, "fromfolderFiles Count"
    AssertEqualWithMessage 1, new_Fso().GetFolder(from).SubFolders.Count, "fromfolderSubFolders Count"
    AssertEqualWithMessage 2, new_Fso().GetFolder(toto).Files.Count, "tofolderFiles Count"
    AssertEqualWithMessage 1, new_Fso().GetFolder(toto).SubFolders.Count, "tofolderSubFolders Count"

    Dim ec,ct
    ec = df1
    ct = readTestFile(c, new_Fso().BuildPath(toto, fn1))
    AssertEqualWithMessage ec, ct, "cont file1"

    ec = df2
    ct = readTestFile(c, new_Fso().BuildPath(toto, fn2))
    AssertEqualWithMessage ec, ct, "cont file2"

    ec = True
    AssertEqualWithMessage ec, new_Fso().FolderExists(new_Fso().BuildPath(toto, fn3)), "exists folder3"
End Sub
Sub Test_fs_copyFolder_OverRide
    Dim from
    '�R�s�[���t�H���_���쐬
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder from
    AssertEqualWithMessage True, new_Fso().FolderExists(from), "before copy fromfolder exists"
    
    Dim c,fp,fn1,fn2,fn3,df1,df2
    '�t�H���_�̉��Ƀt�@�C���ƃt�H���_���쐬
    c = "Unicode"
    fn1 = new_Now().formatAs("YYMMDD_hhmmss.000000_f1.txt")
    df1 = "For" & vbNewLine & "copyFolder OverRide fn1"
    fp = new_Fso().BuildPath(from, fn1)
    writeTestFile c,fp,df1
    fn2 = new_Now().formatAs("YYMMDD_hhmmss.000000_f2.txt")
    df2 = "For" & vbNewLine & "copyFolder OverRide fn2"
    fp = new_Fso().BuildPath(from, fn2)
    writeTestFile c,fp,df2
    fn3 = new_Now().formatAs("YYMMDD_hhmmss.000000_f3")
    fp = new_Fso().BuildPath(from, fn3)
    new_Fso().CreateFolder fp
    AssertEqualWithMessage True, new_Fso().FileExists(new_Fso().BuildPath(from, fn1)), "before copy fromfolderfile1 exists"
    AssertEqualWithMessage True, new_Fso().FileExists(new_Fso().BuildPath(from, fn2)), "before copy fromfolderfile2 exists"
    AssertEqualWithMessage True, new_Fso().FolderExists(new_Fso().BuildPath(from, fn3)), "before copy fromfolderfolder3 exists"
    
    Dim toto
    '�R�s�[��t�H���_���쐬
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_"))
    new_Fso().CreateFolder toto
    AssertEqualWithMessage True, new_Fso().FolderExists(toto), "before copy tofolder exists"
    
    Dim tn1,tn2,tn3,dt1,dt2
    '�t�H���_�̉��Ƀt�@�C���ƃt�H���_���쐬
    tn1 = new_Now().formatAs("YYMMDD_hhmmss.000000_t1.txt")
    dt1 = "For" & vbNewLine & "copyFolder OverRide tn1"
    fp = new_Fso().BuildPath(toto, tn1)
    writeTestFile c,fp,dt1
    tn2 = fn2
    dt2 = "For" & vbNewLine & "copyFolder OverRide tn2"
    fp = new_Fso().BuildPath(toto, tn2)
    writeTestFile c,fp,dt2
    tn3 = new_Now().formatAs("YYMMDD_hhmmss.000000_t3")
    fp = new_Fso().BuildPath(toto, tn3)
    new_Fso().CreateFolder fp
    AssertEqualWithMessage True, new_Fso().FileExists(new_Fso().BuildPath(toto, tn1)), "before copy tofolderfile1 exists"
    AssertEqualWithMessage True, new_Fso().FileExists(new_Fso().BuildPath(toto, tn2)), "before copy tofolderfile2 exists"
    AssertEqualWithMessage True, new_Fso().FolderExists(new_Fso().BuildPath(toto, tn3)), "before copy tofolderfolder3 exists"

    Dim ea,a
    ea = True
    a = fs_copyFolder(from,toto)
    AssertEqualWithMessage ea, a, "ret"
    AssertEqualWithMessage 2, new_Fso().GetFolder(from).Files.Count, "fromfolderFiles Count"
    AssertEqualWithMessage 1, new_Fso().GetFolder(from).SubFolders.Count, "fromfolderSubFolders Count"
    AssertEqualWithMessage 3, new_Fso().GetFolder(toto).Files.Count, "tofolderFiles Count"
    AssertEqualWithMessage 2, new_Fso().GetFolder(toto).SubFolders.Count, "tofolderSubFolders Count"

    Dim ec,ct
    ec = df1
    ct = readTestFile(c, new_Fso().BuildPath(from, fn1))
    AssertEqualWithMessage ec, ct, "cont fromfile1"

    ec = df2
    ct = readTestFile(c, new_Fso().BuildPath(from, fn2))
    AssertEqualWithMessage ec, ct, "cont fromfile2"

    ec = True
    AssertEqualWithMessage ec, new_Fso().FolderExists(new_Fso().BuildPath(from, fn3)), "exists from-fromfolder3"

    ec = df1
    ct = readTestFile(c, new_Fso().BuildPath(toto, fn1))
    AssertEqualWithMessage ec, ct, "cont fromfile1"

    ec = df2
    ct = readTestFile(c, new_Fso().BuildPath(toto, fn2))
    AssertEqualWithMessage ec, ct, "cont fromfile2"

    ec = True
    AssertEqualWithMessage ec, new_Fso().FolderExists(new_Fso().BuildPath(toto, fn3)), "exists to-fromfolder3"

    ec = dt1
    ct = readTestFile(c, new_Fso().BuildPath(toto, tn1))
    AssertEqualWithMessage ec, ct, "cont tofile1"

    ec = True
    AssertEqualWithMessage ec, new_Fso().FolderExists(new_Fso().BuildPath(toto, tn3)), "exists to-tofolder3"
End Sub
Sub Test_fs_copyFolder_OverRideWithUnrelatedFileLocked
    Dim from
    '�R�s�[���t�H���_���쐬
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder from
    AssertEqualWithMessage True, new_Fso().FolderExists(from), "before copy fromfolder exists"
    
    Dim c,fp,fn1,fn2,fn3,df1,df2
    '�t�H���_�̉��Ƀt�@�C���ƃt�H���_���쐬
    c = "Unicode"
    fn1 = new_Now().formatAs("YYMMDD_hhmmss.000000_f1.txt")
    df1 = "For" & vbNewLine & "copyFolder OverRideWithUnrelatedFileLocked fn1"
    fp = new_Fso().BuildPath(from, fn1)
    writeTestFile c,fp,df1
    fn2 = new_Now().formatAs("YYMMDD_hhmmss.000000_f2.txt")
    df2 = "For" & vbNewLine & "copyFolder OverRideWithUnrelatedFileLocked fn2"
    fp = new_Fso().BuildPath(from, fn2)
    writeTestFile c,fp,df2
    fn3 = new_Now().formatAs("YYMMDD_hhmmss.000000_f3")
    fp = new_Fso().BuildPath(from, fn3)
    new_Fso().CreateFolder fp
    AssertEqualWithMessage True, new_Fso().FileExists(new_Fso().BuildPath(from, fn1)), "before copy fromfolderfile1 exists"
    AssertEqualWithMessage True, new_Fso().FileExists(new_Fso().BuildPath(from, fn2)), "before copy fromfolderfile2 exists"
    AssertEqualWithMessage True, new_Fso().FolderExists(new_Fso().BuildPath(from, fn3)), "before copy fromfolderfolder3 exists"
    
    Dim toto
    '�R�s�[��t�H���_���쐬
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_"))
    new_Fso().CreateFolder toto
    AssertEqualWithMessage True, new_Fso().FolderExists(toto), "before copy tofolder exists"
    
    Dim tn2,tn3,dt2
    '�t�H���_�̉��Ƀt�@�C���ƃt�H���_���쐬
    tn2 = fn2
    dt2 = "For" & vbNewLine & "copyFolder OverRideWithUnrelatedFileLocked tn2"
    fp = new_Fso().BuildPath(toto, tn2)
    writeTestFile c,fp,dt2
    tn3 = new_Now().formatAs("YYMMDD_hhmmss.000000_t3")
    fp = new_Fso().BuildPath(toto, tn3)
    new_Fso().CreateFolder fp
    AssertEqualWithMessage True, new_Fso().FileExists(new_Fso().BuildPath(toto, tn2)), "before copy tofolderfile2 exists"
    AssertEqualWithMessage True, new_Fso().FolderExists(new_Fso().BuildPath(toto, tn3)), "before copy tofolderfolder3 exists"

    Dim tn1,dt1,f
    tn1 = new_Now().formatAs("YYMMDD_hhmmss.000000_t1.txt")
    dt1 = "For" & vbNewLine & "copyFolder OverRideWithUnrelatedFileLocked tn1"
    fp = new_Fso().BuildPath(toto, tn1)
    f = -1    'TristateTrue(Unicode)
    '�t�@�C������U�쐬���ă��b�N����
    With createFileAndLocked(c,fp,dt1,f)
        AssertEqualWithMessage True, new_Fso().FileExists(new_Fso().BuildPath(toto, tn1)), "before copy tofolderfile1 exists"
        
        Dim e,a
        e = True
        a = fs_copyFolder(from,toto)
        
        'fs_copyFolder()���G���[�ɂȂ�Ȃ����Ƃ��m�F����
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    AssertEqualWithMessage 2, new_Fso().GetFolder(from).Files.Count, "fromfolderFiles Count"
    AssertEqualWithMessage 1, new_Fso().GetFolder(from).SubFolders.Count, "fromfolderSubFolders Count"
    AssertEqualWithMessage 3, new_Fso().GetFolder(toto).Files.Count, "tofolderFiles Count"
    AssertEqualWithMessage 2, new_Fso().GetFolder(toto).SubFolders.Count, "tofolderSubFolders Count"

    '�R�s�[��t�H���_�̃t�@�C�������b�N���Ă��邪�A�㏑�����Ȃ��t�@�C���̂��߃R�s�[������Ɋ�������

    Dim ec,ct
    ec = df1
    ct = readTestFile(c, new_Fso().BuildPath(from, fn1))
    AssertEqualWithMessage ec, ct, "cont fromfile1"

    ec = df2
    ct = readTestFile(c, new_Fso().BuildPath(from, fn2))
    AssertEqualWithMessage ec, ct, "cont fromfile2"

    ec = True
    AssertEqualWithMessage ec, new_Fso().FolderExists(new_Fso().BuildPath(from, fn3)), "exists from-fromfolder3"

    ec = df1
    ct = readTestFile(c, new_Fso().BuildPath(toto, fn1))
    AssertEqualWithMessage ec, ct, "cont fromfile1"

    ec = df2
    ct = readTestFile(c, new_Fso().BuildPath(toto, fn2))
    AssertEqualWithMessage ec, ct, "cont fromfile2"

    ec = True
    AssertEqualWithMessage ec, new_Fso().FolderExists(new_Fso().BuildPath(toto, fn3)), "exists to-fromfolder3"

    ec = dt1
    ct = readTestFile(c, new_Fso().BuildPath(toto, tn1))
    AssertEqualWithMessage ec, ct, "cont tofile1"

    ec = True
    AssertEqualWithMessage ec, new_Fso().FolderExists(new_Fso().BuildPath(toto, tn3)), "exists to-tofolder3"
End Sub
Sub Test_fs_copyFolder_Err_FromFileNoExists
    Dim from
    '�R�s�[���t�H���_���쐬
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    AssertEqualWithMessage False, new_Fso().FolderExists(from), "before copy fromfolder exists"
    
    Dim toto
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_"))
    AssertEqualWithMessage False, new_Fso().FolderExists(toto), "before copy tofolder exists"

    Dim ea,a
    ea = False
    a = fs_copyFolder(from,toto)
    AssertEqualWithMessage 0, Err.Number, "Err.Number"
    AssertEqualWithMessage False, new_Fso().FolderExists(from), "after copy fromfolder exists"
    AssertEqualWithMessage False, new_Fso().FolderExists(from), "after copy tofolder exists"
End Sub
Sub Test_fs_copyFolder_ToFileLocked
    Dim from
    '�R�s�[���t�H���_���쐬
    from = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder from
    AssertEqualWithMessage True, new_Fso().FolderExists(from), "before copy fromfolder exists"
    
    Dim c,fp,fn1,fn2,fn3,df1,df2
    '�t�H���_�̉��Ƀt�@�C���ƃt�H���_���쐬
    c = "Unicode"
    fn1 = new_Now().formatAs("YYMMDD_hhmmss.000000_f1.txt")
    df1 = "For" & vbNewLine & "copyFolder ToFileLocked fn1"
    fp = new_Fso().BuildPath(from, fn1)
    writeTestFile c,fp,df1
    fn2 = new_Now().formatAs("YYMMDD_hhmmss.000000_f2.txt")
    df2 = "For" & vbNewLine & "copyFolder ToFileLocked fn2"
    fp = new_Fso().BuildPath(from, fn2)
    writeTestFile c,fp,df2
    fn3 = new_Now().formatAs("YYMMDD_hhmmss.000000_f3")
    fp = new_Fso().BuildPath(from, fn3)
    new_Fso().CreateFolder fp
    AssertEqualWithMessage True, new_Fso().FileExists(new_Fso().BuildPath(from, fn1)), "before copy fromfolderfile1 exists"
    AssertEqualWithMessage True, new_Fso().FileExists(new_Fso().BuildPath(from, fn2)), "before copy fromfolderfile2 exists"
    AssertEqualWithMessage True, new_Fso().FolderExists(new_Fso().BuildPath(from, fn3)), "before copy fromfolderfolder3 exists"
    
    Dim toto
    '�R�s�[��t�H���_���쐬
    toto = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000_"))
    new_Fso().CreateFolder toto
    AssertEqualWithMessage True, new_Fso().FolderExists(toto), "before copy tofolder exists"
    
    Dim tn1,tn3,dt1
    '�t�H���_�̉��Ƀt�@�C���ƃt�H���_���쐬
    tn1 = new_Now().formatAs("YYMMDD_hhmmss.000000_t1.txt")
    dt1 = "For" & vbNewLine & "copyFolder ToFileLocked tn1"
    fp = new_Fso().BuildPath(toto, tn1)
    writeTestFile c,fp,dt1
    tn3 = new_Now().formatAs("YYMMDD_hhmmss.000000_t3")
    fp = new_Fso().BuildPath(toto, tn3)
    new_Fso().CreateFolder fp
    AssertEqualWithMessage True, new_Fso().FileExists(new_Fso().BuildPath(toto, tn1)), "before copy tofolderfile1 exists"
    AssertEqualWithMessage True, new_Fso().FolderExists(new_Fso().BuildPath(toto, tn3)), "before copy tofolderfolder3 exists"

    Dim tn2,dt2,f
    tn2 = fn2
    dt2 = "For" & vbNewLine & "copyFolder ToFileLocked tn2"
    fp = new_Fso().BuildPath(toto, tn2)
    f = -1    'TristateTrue(Unicode)
    '�t�@�C������U�쐬���ă��b�N����
    With createFileAndLocked(c,fp,dt2,f)
        AssertEqualWithMessage True, new_Fso().FileExists(new_Fso().BuildPath(toto, tn2)), "before copy tofolderfile2 exists"
        
        Dim e,a
        e = False
        a = fs_copyFolder(from,toto)
        
        'fs_copyFolder()���G���[�ɂȂ邱�Ƃ��m�F����
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    AssertEqualWithMessage 2, new_Fso().GetFolder(from).Files.Count, "fromfolderFiles Count"
    AssertEqualWithMessage 1, new_Fso().GetFolder(from).SubFolders.Count, "fromfolderSubFolders Count"
    AssertEqualWithMessage 3, new_Fso().GetFolder(toto).Files.Count, "tofolderFiles Count"
    AssertEqualWithMessage 1, new_Fso().GetFolder(toto).SubFolders.Count, "tofolderSubFolders Count"

    '1�ڂ̃t�@�C��fn1�̓R�s�[�܂��͈ړ����A2�ڂ̃t�@�C��fn2�̃R�s�[�܂��͈ړ������s����A3�ڂ̃t�H���_fn3�̓R�s�[�܂��͈ړ����Ȃ�

    Dim ec,ct
    ec = df1
    ct = readTestFile(c, new_Fso().BuildPath(from, fn1))
    AssertEqualWithMessage ec, ct, "cont fromfile1"

    ec = df2
    ct = readTestFile(c, new_Fso().BuildPath(from, fn2))
    AssertEqualWithMessage ec, ct, "cont fromfile2"

    ec = True
    AssertEqualWithMessage ec, new_Fso().FolderExists(new_Fso().BuildPath(from, fn3)), "exists from-fromfolder3"

    ec = df1
    ct = readTestFile(c, new_Fso().BuildPath(toto, fn1))
    AssertEqualWithMessage ec, ct, "cont fromfile1"

    ec = dt1
    ct = readTestFile(c, new_Fso().BuildPath(toto, tn1))
    AssertEqualWithMessage ec, ct, "cont tofile1"

    ec = dt2
    ct = readTestFile(c, new_Fso().BuildPath(toto, tn2))
    AssertEqualWithMessage ec, ct, "cont tofile1"

    ec = True
    AssertEqualWithMessage ec, new_Fso().FolderExists(new_Fso().BuildPath(toto, tn3)), "exists to-tofolder3"
End Sub

'###################################################################################################
'fs_createFolder()
Sub Test_fs_createFolder
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    AssertEqualWithMessage False, new_Fso().FolderExists(p), "before create folder exists"
    
    Dim a,e
    e = True
    a = fs_createFolder(p)
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage True, new_Fso().FolderExists(p), "after create folder exists"
End Sub
Sub Test_fs_createFolder_ErrExistsFile
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))

    Dim c,d
    '�t�@�C�����쐬
    c = "UTF-8"
    d = "For" & vbNewLine & "CreateFolder Err-ExistsFile"
    writeTestFile c,p,d
    AssertEqualWithMessage True, new_Fso().FileExists(p), "before create folder file exists"
    
    Dim a,e
    e = False
    a = fs_createFolder(p)
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage True, new_Fso().FileExists(p), "after create folder file exists"
    AssertEqualWithMessage False, new_Fso().FolderExists(p), "after create folder folder exists"
End Sub
Sub Test_fs_createFolder_ErrExistsFile
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))

    Dim c,d
    '�t�H���_���쐬
    new_Fso().CreateFolder p
    AssertEqualWithMessage True, new_Fso().FolderExists(p), "before create folder folder exists"
    
    Dim a,e
    e = False
    a = fs_createFolder(p)
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage True, new_Fso().FolderExists(p), "after create folder folder exists"
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
    AssertEqualWithMessage 0, Err.Number, "Err.Number"
    AssertEqualWithMessage False, new_Fso().FileExists(p), "after delete file exists"
End Sub
Sub Test_fs_deleteFile_Err_FileLocked
    Dim c,p,d,f
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "For" & vbNewLine & "DeleteFile Err FileLocked"
    f = -1    'TristateTrue(Unicode)
    '�t�@�C������U�쐬���ă��b�N����
    With createFileAndLocked(c,p,d,f)
        Dim e,a
        e = False
        a = fs_deleteFile(p)
        
        'fs_deleteFile()���G���[�ɂȂ邱�Ƃ��m�F����
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    '�t�@�C�����폜����Ă��Ȃ����Ƃ��m�F
    AssertEqualWithMessage True, new_Fso().FileExists(p), "before delete file exists"
End Sub

'###################################################################################################
'fs_deleteFolder()
Sub Test_fs_deleteFolder
    Dim c,p,fp,d
    '�t�H���_���쐬
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder p
    '�t�H���_�̉��Ƀt�@�C�����쐬
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
Sub Test_fs_deleteFolder_Err_NotExists
    Dim p
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    AssertEqualWithMessage False, new_Fso().FolderExists(p), "before delete folder exists"

    Dim e,a
    e = False
    a = fs_deleteFolder(p)
    AssertEqualWithMessage e, a, "ret"
    AssertEqualWithMessage 0, Err.Number, "Err.Number"
    AssertEqualWithMessage False, new_Fso().FolderExists(p), "after delete folder exists"
End Sub
Sub Test_fs_deleteFolder_Err_FileLocked
    Dim c,p,fp,d,f
    '�t�H���_���쐬
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000"))
    new_Fso().CreateFolder p
    '�t�H���_�̉��Ƀt�@�C�����쐬
    c = "UTF-8"
    fp = new_Fso().BuildPath(p, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    d = "For" & vbNewLine & "DeleteFolder Err FileLocked"
    f = -1    'TristateTrue(Unicode)
    '�t�@�C������U�쐬���ă��b�N����
    With createFileAndLocked(c,fp,d,f)
        Dim e,a
        e = False
        a = fs_deleteFolder(p)
        
        'fs_deleteFolder()���G���[�ɂȂ邱�Ƃ��m�F����
        AssertEqualWithMessage e, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    '�t�H���_���폜����Ă��Ȃ����Ƃ��m�F
    AssertEqualWithMessage True, new_Fso().FolderExists(p), "after delete folder exists"
End Sub

'###################################################################################################
'fs_moveFile()

'###################################################################################################
'fs_moveFolder()

'###################################################################################################
'fs_readFile()
Sub Test_fs_readFile
    Dim c,p,d,e
    '�t�@�C�����쐬
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "lmn" & vbNewLine & "�V�Y�]" & vbNewLine & "���" & vbNewLine & ChrW(12316) 'ChrW(12316)='\u301c'�i�g�_�b�V���E�g�^�jSjis�ɕϊ��ł��Ȃ�����
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
    d = "abc" & vbNewLine & "������" & vbNewLine & "123" & vbNewLine & ChrW(12316) 'ChrW(12316)='\u301c'�i�g�_�b�V���E�g�^�jSjis�ɕϊ��ł��Ȃ�����
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
    '�㏑������t�@�C������U�쐬
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "UTF-8"
    d = "For" & vbNewLine & "Rewrite"
    writeTestFile c,p,d
    AssertEqualWithMessage True, new_Fso().FileExists(p), "before write file exists"

    '�㏑�����邱�Ƃ��m�F
    d = "abc" & vbNewLine & "�@�A�B" & vbNewLine & "!#$" & vbNewLine & ChrW(12316) 'ChrW(12316)='\u301c'�i�g�_�b�V���E�g�^�jSjis�ɕϊ��ł��Ȃ�����
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
    '�t�@�C������U�쐬���ă��b�N����
    With createFileAndLocked(c, p ,d,f)
        d = "error" & vbNewLine & "test"
        Dim ea,a
        ea = False
        a = fs_writeFile(p, d)
        
        'fs_writeFile()���G���[�ɂȂ邱�Ƃ��m�F����
        AssertEqualWithMessage ea, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    Dim ct
    '�㏑�����Ă��Ȃ����Ƃ��m�F
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
    d = "abc" & vbNewLine & "������" & vbNewLine & "123"
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
    d = "func_FsWriteFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Iomode_ForWriting_Normal__Format_SystemDefault"
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
    '�㏑������t�@�C������U�쐬
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "For" & vbNewLine & "Rewrite"
    writeTestFile c,p,d
    AssertEqualWithMessage True, new_Fso().FileExists(p), "before write file exists"

    Dim iomode,create,f,ec,ea,a
    '�㏑�����邱�Ƃ��m�F
    iomode = 2     'ForWriting
    create = True
    f = -1    'TristateTrue(Unicode)
    d = "func_FsWriteFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Iomode_ForWriting_Rewrite__Format_Unicode"
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
    d = "func_FsWriteFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Iomode_ForAppending_Normal__Format_Ascii"
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
    '�ǋL����t�@�C������U�쐬
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "shift-jis"
    d = "For" & vbNewLine & "Append"
    writeTestFile c,p,d
    AssertEqualWithMessage True, new_Fso().FileExists(p), "before write file exists"

    Dim iomode,create,f,ec,ea,a
    '�ǋL���邱�Ƃ��m�F
    iomode = 8     'ForAppending
    create = True
    f = -2         'TristateUseDefault
    ec = d
    d = "func_FsWriteFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Iomode_ForAppending_Append__Format_SystemDefault"
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
    d = "func_FsWriteFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Create_True_Normal__Format_Unicode"
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
    '�㏑������t�@�C������U�쐬
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "shift-jis"
    d = "For" & vbNewLine & "Rewrite"
    writeTestFile c,p,d
    AssertEqualWithMessage True, new_Fso().FileExists(p), "before write file exists"

    Dim iomode,create,f,ec,ea,a
    '�㏑�����邱�Ƃ��m�F
    iomode = 2     'ForWriting
    create = True
    f = 0          'TristateFalse(Ascii)
    d = "func_FsWriteFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Create_True_Rewrite__Format_Ascii"
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
    d = "func_FsWriteFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Create_False_Err"
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
    '�㏑������t�@�C������U�쐬
    p = new_Fso().BuildPath(PsPathTempFolder, new_Now().formatAs("YYMMDD_hhmmss.000000.txt"))
    c = "Unicode"
    d = "For" & vbNewLine & "Rewrite"
    writeTestFile c,p,d
    AssertEqualWithMessage True, new_Fso().FileExists(p), "before write file exists"

    Dim ec,ea,a,iomode,create,f
    '�㏑�����邱�Ƃ��m�F
    iomode = 2     'ForWriting
    create = False
    f = -1         'TristateTrue(Unicode)
    d = "func_FsWriteFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Create_False_Rewrite__Format_Unicode"
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
    '�t�@�C������U�쐬���ă��b�N����
    With createFileAndLocked(c, p ,d, f)
        AssertEqualWithMessage True, new_Fso().FileExists(p), "before write file exists"

        d = "error" & vbNewLine & "test"
        ea = False
        a = func_FsWriteFile(p, iomode, create, f, d)
        
        'func_FsWriteFile()���G���[�ɂȂ邱�Ƃ��m�F����
        AssertEqualWithMessage ea, a, "ret"
        AssertEqualWithMessage 0, Err.Number, "Err.Number"

        .Close
    End With

    '�㏑�����Ă��Ȃ����Ƃ��m�F
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
    d = "func_FsReadFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Normal__Format_SystemDefault"
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
    d = "func_FsReadFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Normal__Format_Unicode"
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
    d = "func_FsReadFile" & vbNewLine & "�̃e�X�g" & vbNewLine & "Normal__Format_Ascii"
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
    'Textstream���쐬���ĕԋp
    Set createFileAndLocked = new_Ts(p, 8, True, f)
End Function

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
