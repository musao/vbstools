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


Dim sPath : sPath = func_CM_FsGetPrivateFilePath("tmp", "sample.txt")   '�`\tmp\sample.txt
Dim oTs : Set oTs = func_CM_FsOpenTextFile(sPath, 1, False, -2)
Dim sText
Dim oFunc
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

''Read
'Msgbox oFunc(oTs)          '1,1,False,False
'sText = oTs.Read(10)
'Inputbox "","",sText       '���͏\�ꌎ�Ƃ��Ă���
'Msgbox oFunc(oTs)          '1,11,False,False


''ReadLine
'Msgbox oFunc(oTs)          '1,1,False,False
'sText = oTs.ReadLine()
'Inputbox "","",sText
''���͏\�ꌎ�Ƃ��Ă����̌��R�S�Ƃ������̂̂��߂�p�������B�������Ď��������h�S�͂����̔��W�ł��傽�قǂ��]��Ă���܂��ɂ͕������܂��܂����āA�܂��ɂ͂��悽�ł����낽�B���ɂ߂����ł����̂͂Ƃ��Ă��ق��ɂ܂�łł��łȂ���B�قƂ�ǎO���ɂ��b���㗬�����b�ɂ��܂��������̕K�킢�����\�ɂƂ����̘b�܂��ł��ł������A�킪�\�������Ȃ������M�K�n���ނ��āA��X����̕��֌x�����Ă̎����܂�ł�����Ƃ��̂ł���������ӌ��������悤�Ƃǂ��������W�œZ�߂Ȃ��ȂA���₵������قǔ��΂��͂܂邾�ƂȂ�ł̂����邽���B
'Msgbox oFunc(oTs)          '2,1,False,False

''ReadAll
'Msgbox oFunc(oTs)          '1,1,False,False
'sText = oTs.ReadAll()
'Inputbox "","",sText
'msgbox Len(sText)
'Msgbox oFunc(oTs)          '148,1,True,True

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
