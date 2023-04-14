Attribute VB_Name = "Main"
'@Folder "ChatGPT"
Option Explicit

'@EntryPoint
Public Sub ExecuteChatGPT()
    '�I��͈͂ɑ΂���ChatGPT�����s����
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
Attribute ChatGPT.VB_Description = "ChatGPT�ɖ₢���킹�܂��B"
Attribute ChatGPT.VB_ProcData.VB_Invoke_Func = " \n14"
'----------------------------------------------------------------------------------------------------
' Procedure : ChatGPT
' DateTime  : 2023/04/14
' Author    : ikuma
' Purpose   : �^����ꂽ���[�U�[�R���e���c���g�p���āAOpenAI��ChatGPT API�ɃA�N�Z�X���A�������ꂽ������Ԃ��܂��B
'           : �I�v�V�����ŁA�ǉ��̃��[�U�[�R���e���c�A�o�͌`���A���[�����w��ł��܂��B
' Input     : ByVal userContent1 As String          - �ŏ��̃��[�U�[�R���e���c
'           : Optional ByVal userContent2 As String - ��Ԗڂ̃��[�U�[�R���e���c�i�I�v�V�����j
'           : Optional ByVal output As String       - �o�͌`���i�I�v�V�����j
'           : Optional ByVal role As String         - ���[���i�I�v�V�����j
' Output    : String - GPT��������������
'----------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    'API�̐ݒ�
    Dim settings As GptSettings: Set settings = New GptSettings
    
    '���N�G�X�g�̃��[���Ɠ��e�̐ݒ�
    Dim requestMessage As ChatRequest: Set requestMessage = New ChatRequest
    With requestMessage
        '��{�ݒ�
        Set .SetSettings = settings
        '���[���̐ݒ�
        If role <> vbNullString Then .AddMessage("system") = role
        .AddMessage("user") = userContent1
        If userContent2 <> vbNullString Then .AddMessage("user") = userContent2
        If output <> vbNullString Then .AddMessage("system") = "�o�͌`���� �u" & output & "�v�Ƃ��邱�ƁB"
    End With
    
    'HTTP ���N�G�X�g
    Dim http As MSXML2.XMLHTTP60: Set http = New MSXML2.XMLHTTP60
    With http
        '���N�G�X�g���M��̐ݒ�
        .Open "POST", settings.ApiURL
        
        '�w�b�_�[�̐ݒ�
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & settings.ApiKEY
        
        '���N�G�X�g�̑��M
        Dim json As String: json = requestMessage.GenerateJson
        .send json
        
        '���X�|���X�̎擾
        Dim parsed As Scripting.Dictionary: Set parsed = JsonConverter.ParseJson(.responseText)
        
        'GPT����̉񓚂��o��
        If InStr(.responseText, """error"": {") > 0 Then
            ChatGPT = parsed("error")("message")
        Else
            ChatGPT = parsed("choices")(1)("message")("content")
        End If
    End With
    
    DoEvents '�r���o�߂��o�͂��邽�߂� DoEvents
    Exit Function

ErrHandler:
    MsgBox Err.Description, vbCritical, "�G���["
End Function


