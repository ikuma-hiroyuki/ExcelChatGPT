VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GptSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ChatGPT"
Option Explicit

Private ws_ As Worksheet

Private Sub Class_Initialize()
    Set ws_ = ThisWorkbook.Worksheets("設定")
End Sub

Public Property Get ApiURL() As String
    ApiURL = "https://api.openai.com/v1/chat/completions"
End Property

Public Property Get ApiKEY() As String
    Dim key As String: key = ws_.Range("B2").Value
    If key = vbNullString Then
        Err.Raise 999, Description:="設定シートにAPIキーが設定されていません。"
    End If
    ApiKEY = key
End Property

Public Property Get Model() As String
    Dim api_model As String: api_model = ws_.Range("B3").Value
    If api_model = vbNullString Then
        Model = "gpt-3.5-turbo"
    Else
        Model = api_model
    End If
End Property

Public Property Get MaxTokens() As Long
    Dim token As String: token = ws_.Range("B4").Value
    If token = vbNullString Then
        MaxTokens = 100
    Else
        MaxTokens = token
    End If
End Property

Public Property Get Temperature() As Double
    Dim temp As Double: temp = ws_.Range("B5").Value
    If temp = 0 Then
        Temperature = 1
    Else
        Temperature = temp
    End If
End Property

