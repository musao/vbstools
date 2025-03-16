' libCom.vbs: util_* procedure test.
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

Const MY_NAME = "test_libCom_util.vbs"
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
'util_escapeForPs()


'###################################################################################################
'util_getIpAddress()
Sub Test_util_getIpAddress
    Const RE_IP4 = "^(([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])$"
    Const RE_IP6 = "^(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))$"
    
    dim a
    a = util_getIpAddress()

    dim i
    For i=0 To Ubound(a)
        AssertWithMessage Len(a(i).Item("Caption"))>0, "i="&i&":Caption"
        AssertMatchWithMessage RE_IP4, a(i).Item("Ip").Item("V4"), "i="&i&":IpV4"
        AssertMatchWithMessage RE_IP6, a(i).Item("Ip").Item("V6"), "i="&i&":IpV6"
    Next
End Sub

'###################################################################################################
'util_randStr()
Sub Test_util_randStr
    dim d,a,s
    s = 1000
    With new_Char()
        d = .charList(.typeHalfWidthNumbers)
    End With
    
    a = util_randStr(d,s)
    Dim i,j,t,flg : i=1
    Do While i<Len(a)
        t = Mid(a,i,1)
        flg = false
        For Each j In d
            If cf_isSame(t,j) Then
                flg=True
                Exit For
            End If
        Next
        If Not flg Then
            AssertFailWithMessage "util_randStr(d,s)="&a&", s="&s&", i="&i
        End If
        i=i+1
    Loop
    Assert True
End Sub

'###################################################################################################
'util_isZipWithPassword()
Sub Test_util_isZipWithPassword_WithPassword
    Dim e,d,a
    e = True
    d = makeDummyZipFile(1)

    a = util_isZipWithPassword(d)
    AssertEqualWithMessage e, a, "util_isZipWithPassword()"
End Sub
Sub Test_util_isZipWithPassword_WithNoPassword
    Dim e,d,a
    e = False
    d = makeDummyZipFile(0)

    a = util_isZipWithPassword(d)
    AssertEqualWithMessage e, a, "util_isZipWithPassword()"
End Sub
Sub Test_util_isZipWithPassword_NotZipFile
    Dim e,d,a
    e = False
    d = WScript.ScriptFullName

    a = util_isZipWithPassword(d)
    AssertEqualWithMessage e, a, "util_isZipWithPassword()"
End Sub
Sub Test_util_isZipWithPassword_NotExisistsFile
    Dim e,d,a
    e = False
    d = new_Fso().BuildPath(PsPathTempFolder, new_Fso().GetTempName())

    a = util_isZipWithPassword(d)
    AssertEqualWithMessage e, a, "util_isZipWithPassword()"
End Sub
Sub Test_util_isZipWithPassword_NotPath
    Dim e,d,a
    e = False
    d = vbNullString

    a = util_isZipWithPassword(d)
    AssertEqualWithMessage e, a, "util_isZipWithPassword()"
End Sub




'###################################################################################################
'common
Function makeDummyZipFile(c)
    Dim p : p = new_Fso().BuildPath(PsPathTempFolder, new_Fso().GetTempName())
    Dim ar(6),i
    For i=0 To 5
        ar(i)=CByte(AscW("Z"))
    Next
    ar(6)=CByte(c)

    With CreateObject("ADODB.Stream")
        .Type = 2       'adTypeText
        .Charset = "UTF-16BE"
        .Open
        
        For i = 0 To UBound(ar) - 1 Step 2
            .WriteText ChrW(ar(i) * &H100 + ar(i + 1))
        Next
        If ((UBound(ar) And 1) = 0) Then
            .WriteText ChrW(ar(UBound(ar)) * &H100)
            .Position = .Position - 1
            .SetEOS
        End If

        .SaveToFile p
        .Close
    End With
    makeDummyZipFile = p
End Function

' Local Variables:
' mode: Visual-Basic
' indent-tabs-mode: nil
' End:
