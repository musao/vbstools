Option Explicit

'�萔
Private Const Cs_FOLDER_INCLUDE = "include"
Private Const Cs_FOLDER_TEMP = "tmp"

'Include�p�֐���`
Sub sub_Include( _
    byVal asIncludeFileName _
    )
    With CreateObject("Scripting.FileSystemObject")
        Dim sParentFolderName : sParentFolderName = .GetParentFolderName(WScript.ScriptFullName)
        Dim sIncludeFilePath
        sIncludeFilePath = .BuildPath(sParentFolderName, Cs_FOLDER_INCLUDE)
        sIncludeFilePath = .BuildPath(sIncludeFilePath, asIncludeFileName)
        ExecuteGlobal .OpenTextfile(sIncludeFilePath).ReadAll
    End With
End Sub
'Include
Call sub_Include("VbsBasicLibCommon.vbs")
Call sub_Include("clsCompareExcel.vbs")
Call sub_Include("clsCmCalendar.vbs")
Call sub_Include("clsCmBufferedWriter.vbs")
Call sub_Include("clsCmArray.vbs")


Dim sPath, oTs, sText, oFunc
Set oFunc = new_Func( _
                   "ts => {" _
                           & "dim s:" _
                           & "With ts:" _
                               & "s=""Line = ""&.Line&vbNewLine" _
                                   & "&""Column = ""&.Column&vbNewLine" _
                                   & "&""AtEndOfLine = ""&.AtEndOfLine&vbNewLine" _
                                   & "&""AtEndOfStream = ""&.AtEndOfStream:" _
                           & "End With:" _
                           & "return s" _
                           & "}" _
                )


sPath = func_CM_FsGetPrivateFilePath("tmp", "sample.txt")   '�`\tmp\sample.txt
Set oTs = func_CM_FsOpenTextFile(sPath, 1, False, -2)

''Read
'Msgbox oFunc(oTs)          '1,1,False,False
'sText = oTs.Read(10)
'Inputbox "","",sText       '���͏\�ꌎ�Ƃ��Ă���
'Msgbox oFunc(oTs)          '1,11,False,False

''Read2
'Msgbox oFunc(oTs)          '1,1,False,False
'sText = oTs.Read(23605)
'Inputbox "","",sText       '���͏\�ꌎ�Ƃ��Ă����`
'Msgbox oFunc(oTs)          '147,23,True,False
'sText = oTs.Read(1)
'Inputbox "","",sText       '
'Msgbox oFunc(oTs)          '148,1,True,True
'sText = oTs.Read(1)        '�G���[
'Msgbox oFunc(oTs)          '-

''Read3
'Msgbox oFunc(oTs)          '1,1,False,False
'sText = oTs.Read(23606)
'Inputbox "","",sText       '���͏\�ꌎ�Ƃ��Ă����`
'Msgbox oFunc(oTs)          '148,1,True,True
'sText = oTs.Read(1)        '�G���[
'Msgbox oFunc(oTs)          '-

''Read4
'Msgbox oFunc(oTs)          '1,1,False,False
'sText = oTs.Read(23607)
'Inputbox "","",sText       '���͏\�ꌎ�Ƃ��Ă����`
'msgbox Len(sText)          '23606
'Msgbox oFunc(oTs)          '148,1,True,True
'sText = oTs.Read(1)        '�G���[
'Msgbox oFunc(oTs)          '-


''ReadLine1
'Msgbox oFunc(oTs)          '1,1,False,False
'sText = oTs.ReadLine()
'Inputbox "","",sText
''���͏\�ꌎ�Ƃ��Ă����̌��R�S�Ƃ������̂̂��߂�p�������B�������Ď��������h�S�͂����̔��W�ł��傽�قǂ��]��Ă���܂��ɂ͕������܂��܂����āA�܂��ɂ͂��悽�ł����낽�B���ɂ߂����ł����̂͂Ƃ��Ă��ق��ɂ܂�łł��łȂ���B�قƂ�ǎO���ɂ��b���㗬�����b�ɂ��܂��������̕K�킢�����\�ɂƂ����̘b�܂��ł��ł������A�킪�\�������Ȃ������M�K�n���ނ��āA��X����̕��֌x�����Ă̎����܂�ł�����Ƃ��̂ł���������ӌ��������悤�Ƃǂ��������W�œZ�߂Ȃ��ȂA���₵������قǔ��΂��͂܂邾�ƂȂ�ł̂����邽���B
'Msgbox oFunc(oTs)          '2,1,False,False

''ReadLine2
'Msgbox oFunc(oTs)          '1,1,False,False
'sText = oTs.Read(23605)
'Inputbox "","",sText       '���͏\�ꌎ�Ƃ��Ă����`
'Msgbox oFunc(oTs)          '147,23,True,False
'sText = oTs.ReadLine()
'Inputbox "","",sText       '
'Msgbox oFunc(oTs)          '148,1,True,True
'sText = oTs.ReadLine()     '�G���[

''ReadLine3
sPath = func_CM_FsGetPrivateFilePath("tmp", "sample2.txt")   '�`\tmp\sample2.txt
Set oTs = func_CM_FsOpenTextFile(sPath, 1, False, -2)

''ReadLine3-1
'Msgbox oFunc(oTs)          '1,1,False,False
'sText = oTs.Read(1)
'Inputbox "","",sText       '��  [Pointer=1]
'Msgbox oFunc(oTs)          '1,2,False,False
'sText = oTs.ReadLine()
'Inputbox "","",sText       '��  [Pointer=2]
'Msgbox oFunc(oTs)          '2,1,False,False
'sText = oTs.ReadLine()
'Inputbox "","",sText       '��  [Pointer=5]
'Msgbox oFunc(oTs)          '3,1,False,False
'oTs.SkipLine
'Msgbox oFunc(oTs)          '4,1,False,False
'oTs.SkipLine
'Msgbox oFunc(oTs)          '5,1,True,False
'oTs.SkipLine
'Msgbox oFunc(oTs)          '6,1,False,False
'sText = oTs.ReadLine()
'Inputbox "","",sText       'ef  [Pointer=19]
'Msgbox oFunc(oTs)          '6,3,True,True

'ReadLine3-2
Msgbox oFunc(oTs)          '1,1,False,False
sText = oTs.Read(1)
Inputbox "","",sText       '��  [Pointer=1]
Msgbox Len(sText)          '1
Msgbox oFunc(oTs)          '1,2,False,False
sText = oTs.Read(1)
Inputbox "","",sText       '��  [Pointer=2]
Msgbox Len(sText)          '1
Msgbox oFunc(oTs)          '1,3,True,False
sText = oTs.Read(3)
Inputbox "","",sText       '<���s>��  [Pointer=5]
Msgbox Len(sText)          '3
Msgbox oFunc(oTs)          '2,2,True,False
sText = oTs.Read(14)
Inputbox "","",sText       '<���s>����<���s>abc<���s>d<���s>  [Pointer=19]
Msgbox Len(sText)          '14
Msgbox oFunc(oTs)          '6,1,False,False
sText = oTs.Read(100)
Inputbox "","",sText       'ef  [Pointer=21]
Msgbox Len(sText)          '2
Msgbox oFunc(oTs)          '6,3,True,True


''ReadAll
'Msgbox oFunc(oTs)          '1,1,False,False
'sText = oTs.ReadAll()
'Inputbox "","",sText
'msgbox Len(sText)          '23606
'Msgbox oFunc(oTs)          '148,1,True,True
'sText = oTs.ReadAll()      '�G���[
'Msgbox oFunc(oTs)          '-

''Skip
'Msgbox oFunc(oTs)          '1,1,False,False
'oTs.Skip(10)
'Msgbox oFunc(oTs)          '1,11,False,False
'sText = oTs.Read(10)
'Inputbox "","",sText       '�̌��R�S�Ƃ������̂�
'Msgbox oFunc(oTs)          '1,21,False,False
'oTs.Skip(230)
'Msgbox oFunc(oTs)          '1,251,False,False
'sText = oTs.Read(10)
'Inputbox "","",sText       '���邽���B<���s>������
'Msgbox oFunc(oTs)          '2,4,False,False

''Skip2
'Msgbox oFunc(oTs)          '1,1,False,False
'oTs.Skip(254)
'Msgbox oFunc(oTs)          '1,255,False,False
'sText = oTs.Read(1)
'Inputbox "","",sText       '�B
'Msgbox "Len(sText) = " & Len(sText) & vbNewLine & "Asc(sText) = " & Asc(sText) & vbNewLine & "Asc(""�B"") = " & Asc("�B")   '1,-32446,-32446
'Msgbox oFunc(oTs)          '1,256,True,False
'sText = oTs.Read(1)
'Inputbox "","",sText       '
'Msgbox "Len(sText) = " & Len(sText) & vbNewLine & "Asc(sText) = " & Asc(sText) & vbNewLine & "Asc(vbCr) = " & Asc(vbCr)   '1,13,13
'Msgbox oFunc(oTs)          '1,257,True,False
'sText = oTs.Read(1)
'Inputbox "","",sText       '
'Msgbox "Len(sText) = " & Len(sText) & vbNewLine & "Asc(sText) = " & Asc(sText) & vbNewLine & "Asc(vbLf) = " & Asc(vbLf)   '1,10,10
'Msgbox oFunc(oTs)          '2,1,False,False
'sText = oTs.Read(1)
'Inputbox "","",sText       '��
'Msgbox "Len(sText) = " & Len(sText) & vbNewLine & "Asc(sText) = " & Asc(sText) & vbNewLine & "Asc(""��"") = " & Asc("��")   '1,-32075,-32075
'Msgbox oFunc(oTs)          '2,2,False,False


''SkipLine
'Msgbox oFunc(oTs)          '1,1,False,False
'oTs.SkipLine
'Msgbox oFunc(oTs)          '2,1,False,False
'sText = oTs.Read(10)
'Inputbox "","",sText       '��������������������
'Msgbox oFunc(oTs)          '2,11,False,False
'oTs.SkipLine
'Msgbox oFunc(oTs)          '3,1,False,False
'sText = oTs.Read(10)
'Inputbox "","",sText       '�����������݂Ɍ�����
'Msgbox oFunc(oTs)          '3,11,False,False

wscript.quit
