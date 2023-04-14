VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChatRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ChatGPT"
Option Explicit

Private mainDict As Scripting.Dictionary '大本の辞書
Private message As Scripting.Dictionary  'JSON の messages 内に格納するための message １行用の辞書
Private messagesArray As Variant         'JSON の messages 用配列。この中に message を格納する。

Private Sub Class_Initialize()
    Set mainDict = New Scripting.Dictionary
    messagesArray = Array()
End Sub

'@Ignore WriteOnlyProperty
Public Property Set SetSettings(ByVal settings As GptSettings)
    '基本設定
    mainDict("model") = settings.Model
    mainDict("max_tokens") = settings.MaxTokens
    mainDict("temperature") = settings.Temperature
End Property

'@Ignore WriteOnlyProperty
Public Property Let AddMessage(ByVal roleValue As String, ByVal contentValue As String)
    Set message = New Scripting.Dictionary
    message("role") = roleValue
    message("content") = contentValue
    ReDim Preserve messagesArray(UBound(messagesArray) + 1)
    Set messagesArray(UBound(messagesArray)) = message
End Property

Public Function GenerateJson() As String
    '最終的に以下の形式のJSONを生成する｡
    '{
    '  "model": "gpt-3.5-turbo",
    '  "messages": [
    '    { "role": "system", "content": "hogehoge" },
    '    { "role": "user", "content": "fugafuga" },
    '    { "role": "system", "content": "barbar" }
    '  ],
    '  "max_tokens": 1000,
    '  "temperature": 0.8
    '}
    
    '最後にmessages 用配列に格納
    mainDict("messages") = messagesArray
    
    'JSONに変換
    Dim json As String: json = JsonConverter.ConvertToJson(mainDict)
    
    GenerateJson = json
End Function
