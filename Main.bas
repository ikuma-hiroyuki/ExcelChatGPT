Attribute VB_Name = "Main"
'@Folder "ChatGPT"
Option Explicit

'@EntryPoint
Public Sub ExecuteChatGPT()
    '選択範囲に対してChatGPTを実行する
    Dim rg As Range
    For Each rg In Selection
        Dim content1 As String, content2 As String, output As String, role As String
        content1 = Cells(1, rg.Column).Value
        content2 = Cells(rg.Row, "C")
        output = Cells(rg.Row, "B")
        role = Cells(rg.Row, "A")
        rg.Value = ChatGPT(content1, content2, output, role)
    Next
End Sub

Public Function ChatGPT(ByVal userContent1 As String, Optional ByVal userContent2 As String, _
                        Optional ByVal output As String, Optional ByVal role As String) As String
Attribute ChatGPT.VB_Description = "ChatGPTに問い合わせます。"
Attribute ChatGPT.VB_ProcData.VB_Invoke_Func = " \n14"
'----------------------------------------------------------------------------------------------------
' Procedure : ChatGPT
' DateTime  : 2023/04/14
' Author    : ikuma
' Purpose   : 与えられたユーザーコンテンツを使用して、OpenAIのChatGPT APIにアクセスし、生成された応答を返します。
'           : オプションで、追加のユーザーコンテンツ、出力形式、ロールを指定できます。
' Input     : ByVal userContent1 As String          - 最初のユーザーコンテンツ
'           : Optional ByVal userContent2 As String - 二番目のユーザーコンテンツ（オプション）
'           : Optional ByVal output As String       - 出力形式（オプション）
'           : Optional ByVal role As String         - ロール（オプション）
' Output    : String - GPTが生成した応答
'----------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    'APIの設定
    Dim settings As GptSettings: Set settings = New GptSettings
    
    'リクエストのロールと内容の設定
    Dim requestMessage As ChatRequest: Set requestMessage = New ChatRequest
    With requestMessage
        '基本設定
        Set .SetSettings = settings
        'ロールの設定
        If role <> vbNullString Then .AddMessage("system") = role
        .AddMessage("user") = userContent1
        If userContent2 <> vbNullString Then .AddMessage("user") = userContent2
        If output <> vbNullString Then .AddMessage("system") = "出力形式は 「" & output & "」とすること。"
    End With
    
    'HTTP リクエスト
    Dim http As MSXML2.XMLHTTP60: Set http = New MSXML2.XMLHTTP60
    With http
        'リクエスト送信先の設定
        .Open "POST", settings.ApiURL
        
        'ヘッダーの設定
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & settings.ApiKEY
        
        'リクエストの送信
        Dim json As String: json = requestMessage.GenerateJson
        .send json
        
        'レスポンスの取得
        Dim parsed As Scripting.Dictionary: Set parsed = JsonConverter.ParseJson(.responseText)
        
        'GPTからの回答を出力
        If InStr(.responseText, """error"": {") > 0 Then
            ChatGPT = parsed("error")("message")
        Else
            ChatGPT = parsed("choices")(1)("message")("content")
        End If
    End With
    
    DoEvents '途中経過を出力するために DoEvents
    Exit Function

ErrHandler:
    MsgBox Err.Description, vbCritical, "エラー"
End Function


