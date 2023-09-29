Option Explicit

'定数
Private Const Cs_FOLDER_INCLUDE = "include"
Private Const Cs_FOLDER_TEMP = "tmp"

'Include用関数定義
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


Dim sPath : sPath = func_CM_FsGetPrivateFilePath("tmp", "sample.txt")   '～\tmp\sample.txt
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
'Inputbox "","",sText       '私は十一月とうていそ
'Msgbox oFunc(oTs)          '1,11,False,False


''ReadLine
'Msgbox oFunc(oTs)          '1,1,False,False
'sText = oTs.ReadLine()
'Inputbox "","",sText
''私は十一月とうていその欠乏心というもののためを用いうた。けっして事実が尊敬心はついその発展でしょたほどを云わておりませには附随きまっますだて、まだにはしよたですだろた。国にめがけですものはとうていほかにまるでですでなけれ。ほとんど三宅さんにお話し上流そう話にしまし自分その必竟いつか乱暴にという肝話ますですですたが、わが十月もあなたか自信尻馬が釣って、大森さんの方へ警視総監の私をまるでお相違としのでそれ個性をご意見をもっようとどうかご発展で纏めないなつつ、いやしくもよほど反対がはまるだとならでのが困るたう。
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
'Inputbox "","",sText       'の欠乏心というものの
'Msgbox oFunc(oTs)          '1,21,False,False
'oTs.Skip(230)
'Msgbox oFunc(oTs)          '1,251,False,False
'sText = oTs.Read(10)
'Inputbox "","",sText       '困るたう。<改行>しかし
'Msgbox oFunc(oTs)          '2,4,False,False

''Skip2
'Msgbox oFunc(oTs)          '1,1,False,False
'oTs.Skip(254)
'Msgbox oFunc(oTs)          '1,255,False,False
'sText = oTs.Read(1)
'Inputbox "","",sText       '。
'Msgbox "Len(sText) = " & Len(sText) & vbNewLine & "Asc(sText) = " & Asc(sText) & vbNewLine & "Asc(""。"") = " & Asc("。")   '1,-32446,-32446
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
'Inputbox "","",sText       'し
'Msgbox "Len(sText) = " & Len(sText) & vbNewLine & "Asc(sText) = " & Asc(sText) & vbNewLine & "Asc(""し"") = " & Asc("し")   '1,-32075,-32075
'Msgbox oFunc(oTs)          '2,2,False,False


''SkipLine
'Msgbox oFunc(oTs)          '1,1,False,False
'oTs.SkipLine
'Msgbox oFunc(oTs)          '2,1,False,False
'sText = oTs.Read(10)
'Inputbox "","",sText       'しかししかしご国から
'Msgbox oFunc(oTs)          '2,11,False,False
'oTs.SkipLine
'Msgbox oFunc(oTs)          '3,1,False,False
'sText = oTs.Read(10)
'Inputbox "","",sText       'しかしご建設に向いて
'Msgbox oFunc(oTs)          '3,11,False,False

wscript.quit
