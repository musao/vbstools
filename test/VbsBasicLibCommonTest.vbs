'***************************************************************************************************
'FILENAME                    : libComTest.vbs
'Overview                    : 共通関数ライブラリのテスト
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/11         Y.Fujii                  First edition
'***************************************************************************************************
Option Explicit

'定数
Private Const Cs_FOLDER_LIB = "lib"
Private Const Cs_UTLIB_FILE = "VbsUtLib.vbs"
Private Const Cs_UTAST_FILE = "clsUtAssistant.vbs"
Private Const Cs_TEST_FILE = "libCom.vbs"

With CreateObject("Scripting.FileSystemObject")
    '単体テスト用ライブラリ読み込み
    Dim sIncludeFolderPath : sIncludeFolderPath = .BuildPath(.GetParentFolderName(WScript.ScriptFullName), Cs_FOLDER_LIB)
    ExecuteGlobal .OpenTextfile(.BuildPath(sIncludeFolderPath, Cs_UTLIB_FILE)).ReadAll
    ExecuteGlobal .OpenTextfile(.BuildPath(sIncludeFolderPath, Cs_UTAST_FILE)).ReadAll
    '単体テスト対象ソース読み込み
    Dim sParentFolderPath : sParentFolderPath = .GetParentFolderName(.GetParentFolderName(WScript.ScriptFullName))
    sIncludeFolderPath = .BuildPath(sParentFolderPath, Cs_FOLDER_LIB)
    ExecuteGlobal .OpenTextfile(.BuildPath(sIncludeFolderPath, Cs_TEST_FILE)).ReadAll
End With

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
'2022/10/11         Y.Fujii                  First edition
'***************************************************************************************************
Sub Main()
    Dim oUtAssistant : Set oUtAssistant = New clsUtAssistant
    
    'func_CM_FsDeleteFile()のテスト
    Call func_CM_FsDeleteFileTest(oUtAssistant)
    'func_CM_FsGetParentFolderPath()のテスト
    Call func_CM_FsGetParentFolderPathTest(oUtAssistant)
    
    '結果出力
    Call sub_UtResultOutput(oUtAssistant)
    
    Set oUtAssistant = Nothing
    
End Sub

'***************************************************************************************************
'Processing Order            : Last
'Function/Sub Name           : sub_OutputReport()
'Overview                    : 結果出力
'Detailed Description        : 工事中
'Argument
'     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub sub_OutputReport( _
    byRef aoUtAssistant _
    )
    Call sub_UtWriteFile(func_UtGetThisLogFilePath(), aoUtAssistant.OutputReportInTsvFormat())
End Sub


'***************************************************************************************************
'Processing Order            : 1
'Function/Sub Name           : func_CM_FsDeleteFileTest()
'Overview                    : func_CM_FsDeleteFile()のテスト
'Detailed Description        : 工事中
'Argument
'     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/11         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub func_CM_FsDeleteFileTest( _
    byRef aoUtAssistant _
    )
    
    '1-1 削除成功
    Call aoUtAssistant.Run("func_CM_FsDeleteFileTestSuccess")
    '1-2 削除失敗
    Call aoUtAssistant.Run("func_CM_FsDeleteFileTestFailure")
    
End Sub

'***************************************************************************************************
'Processing Order            : 1-1
'Function/Sub Name           : func_CM_FsDeleteFileTestSuccess()
'Overview                    : func_CM_FsDeleteFile()のテスト
'Detailed Description        : 削除成功の場合
'Argument
'     なし
'Return Value
'     結果 True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/11         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsDeleteFileTestSuccess( _
    )
    func_CM_FsDeleteFileTestSuccess = False
    
    '一時ファイルのフルパスを取得
    Dim sPath : sPath = func_UtGetThisTempFilePath()
    
    With CreateObject("Scripting.FileSystemObject")
        'ファイルを作成
        Call .CreateTextFile(sPath)
        
        'ファイルができていることを確認
        If Not(.FileExists(sPath)) Then Exit Function
        
        'テスト対象実行
        Dim boResult : boResult = func_CM_FsDeleteFile(sPath)
        
        '戻り値を確認
        If Not(boResult) Then Exit Function
        
        'ファイルが削除できていたら成功
        func_CM_FsDeleteFileTestSuccess = Not(.FileExists(sPath))
    End With
    
End Function

'***************************************************************************************************
'Processing Order            : 1-2
'Function/Sub Name           : func_CM_FsDeleteFileTestFailure()
'Overview                    : func_CM_FsDeleteFile()のテスト
'Detailed Description        : 削除失敗の場合
'Argument
'     なし
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsDeleteFileTestFailure( _
    )
    func_CM_FsDeleteFileTestFailure = False
    
    '一時ファイルのフルパスを取得
    Dim sPath : sPath = func_UtGetThisTempFilePath()
    
    With CreateObject("Scripting.FileSystemObject")
        
        'ファイルがないことを確認
        If .FileExists(sPath) Then Exit Function
        
        'テスト対象実行
        Dim boResult : boResult = func_CM_FsDeleteFile(sPath)
        
        'ファイルがないので失敗したら成功
        func_CM_FsDeleteFileTestFailure = Not(boResult)
    End With
    
End Function

'***************************************************************************************************
'Processing Order            : 2
'Function/Sub Name           : func_CM_FsGetParentFolderPathTest()
'Overview                    : func_CM_FsGetParentFolderPath()のテスト
'Detailed Description        : 工事中
'Argument
'     aoUtAssistant          : 単体テスト用アシスタントクラスのインスタンス
'Return Value
'     なし
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Sub func_CM_FsGetParentFolderPathTest( _
    byRef aoUtAssistant _
    )
    
    '2-1 正常
    Call aoUtAssistant.Run("func_CM_FsGetParentFolderPathTestNormal")
'    '1-2 削除失敗
'    Call aoUtAssistant.Run("func_CM_FsDeleteFileTestFailure")
    
End Sub

'***************************************************************************************************
'Processing Order            : 2-1
'Function/Sub Name           : func_CM_FsGetParentFolderPathTestNormal()
'Overview                    : func_CM_FsGetParentFolderPath()のテスト
'Detailed Description        : 工事中
'Argument
'     なし
'Return Value
'     結果 True,Flase
'---------------------------------------------------------------------------------------------------
'Histroy
'Date               Name                     Reason for Changes
'----------         ----------------------   -------------------------------------------------------
'2022/10/12         Y.Fujii                  First edition
'***************************************************************************************************
Private Function func_CM_FsGetParentFolderPathTestNormal( _
    )
    func_CM_FsGetParentFolderPathTestNormal = False
    
    Dim oParams : Set oParams = new_Dic()
    With oParams               '入力値、期待値
        Call .Add("c:\a\b", "c:\a")
        Call .Add("C:\A\", "C:\")
        Call .Add("C:\a", "C:\")
        Call .Add("c:\", "")
        Call .Add("C:", "")
        
        Dim sKey
        For Each sKey In .Keys
            If StrComp(.Item(sKey), func_CM_FsGetParentFolderPath(sKey)) Then Exit Function
        Next
    End With
    
    func_CM_FsGetParentFolderPathTestNormal = True
    
    Set oParams = Nothing
    
End Function
