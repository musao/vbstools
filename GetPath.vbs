'***************************************************************************************************
'FILENAME                    : GetPath.vbs
'Overview                    : 引数のファイルパスをクリップボードにコピーする
'Detailed Description        : Sendtoからファイルパスを取得するのに使用する
'Argument
'     PATH1,2...             : ファイルのパス1,2,...
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/08/10         Y.Fujii                  First edition
'***************************************************************************************************
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


'メイン関数実行
Call Main()
Wscript.Quit


'***************************************************************************************************
'Processing Order            : First
'Function/Sub Name           : Main()
'Overview                    : メイン関数
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2016/08/10         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    Dim sLineFeedCode : sLineFeedCode = vbCrLf
    
    '引数を改行で連結する
    Dim sOutput : sOutput = ""
    Dim sItem
    For Each sItem In Wscript.Arguments
        If Not(Len(sItem)) Then
            If (Len(sOutput)) Then sOutput = sOutput & sLineFeedCode
            sOutput = sOutput & sItem
        End If
    Next
    
    '一時ファイルのパスを作成
    Dim sParentFolderPath : sParentFolderPath = func_CM_FsGetParentFolderPath(WScript.ScriptFullName)
    Dim sFolderPath : sFolderPath = func_CM_FsBuildPath(sParentFolderPath, Cs_FOLDER_TEMP)
    If Not(func_CM_FsFolderExists(sFolderPath)) Then func_CM_FsCreateFolder(sFolderPath)
    Dim sTempFilePaths : sTempFilePaths = func_CM_FsBuildPath(sFolderPath, func_CM_FsGetTempFileName())
    
    '一時ファイルに連結した引数を出力
    Call sub_CM_FsWriteFile(sTempFilePaths, sOutput)
    
    'クリップボードに一時ファイルの内容を出力
    Call CreateObject("Wscript.Shell").Run("cmd /c clip <""" & sTempFilePaths & """", 0, True)
    
    '一時ファイルを削除
    Call func_CM_FsDeleteFile(sTempFilePaths)
    
End Sub
