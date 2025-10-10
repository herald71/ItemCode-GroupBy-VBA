Attribute VB_Name = "Module35"
' ----------------------------------------------------------
' ���α׷��� : 1688 �߱��� ���� �ڵ��Է� + 1688 �ٷΰ��� ��ũ ����
' �ۼ�����   : 2025-07-07
' ����       : v1.3
' ����       : �Է¿��� �α� �˻�� 1688���� ���� �˻��� ���Ǵ� �߱��� Ű����� ������� �ڵ� �Է��ϰ�,
'              ����� ������ ���� 1688 ���������� �����۸�ũ ����
' ----------------------------------------------------------

Option Explicit

Private Const API_URL As String = "https://api.openai.com/v1/chat/completions"
Private Const MODEL_NAME As String = "gpt-4o"

Sub �߱���Ű�������()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' ����ڷκ��� �Է¿��� ����� ���ĺ� �Է� �ޱ� (�⺻�� B/H)
    Dim inputCol As String
    Dim outputCol As String

    inputCol = InputBox("������ �ѱ� Ű���尡 �ִ� ���� �Է��ϼ��� (��: B)", "�Է¿� ����", "B")
    If inputCol = "" Then Exit Sub

    outputCol = InputBox("�߱��� Ű���带 ������ ��� ���� �Է��ϼ��� (��: H)", "����� ����", "H")
    If outputCol = "" Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, inputCol).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox inputCol & "���� �α� �˻�� �����ϴ�.", vbExclamation
        Exit Sub
    End If

    Dim i As Long
    Dim cntTotal As Long: cntTotal = lastRow - 1
    Application.ScreenUpdating = False
    ws.Range(outputCol & "2:" & outputCol & lastRow).Value = "ó�� ��..."

    ' ����� ������(������ ��) ��� (AA � ����)
    Dim rightCol As String
    rightCol = NextColLetter(outputCol)

    ' ���(����) �߰�
    ws.Cells(1, outputCol).Value = "�߱���"
    ws.Cells(1, rightCol).Value = "1688 �ٷΰ���"

    For i = 2 To lastRow
        Dim korKeyword As String, zhKeyword As String
        korKeyword = Trim(ws.Cells(i, inputCol).Value)
        If korKeyword <> "" Then
            zhKeyword = Get1688Chinese(korKeyword)
            ws.Cells(i, outputCol).Value = zhKeyword

            ' ----- [������ ���� 1688 ���� �����۸�ũ ����]
            If zhKeyword <> "" And Left(zhKeyword, 1) <> "A" Then ' ����/������ ����
                ws.Hyperlinks.Add Anchor:=ws.Cells(i, rightCol), _
                    Address:="https://www.1688.com/", _
                    TextToDisplay:="1688 �ٷΰ���"
            Else
                ws.Cells(i, rightCol).Value = ""
            End If
        Else
            ws.Cells(i, outputCol).Value = ""
            ws.Cells(i, rightCol).Value = ""
        End If

        Application.StatusBar = "����: " & (i - 1) & "/" & cntTotal & " (" & _
             Format((i - 1) / cntTotal, "0%") & ")"
        DoEvents
    Next i

    ws.Columns(outputCol).AutoFit
    ws.Columns(rightCol).AutoFit
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "�߱��� ���� �� 1688 �ٷΰ��� ��ũ(����������) ������ �Ϸ�Ǿ����ϴ�!", vbInformation
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "���� �߻�: " & Err.Description, vbCritical
End Sub

' 1688�� �߱��� �˻��� ���� �Լ� (1���� ��ȯ)
Private Function Get1688Chinese(korKeyword As String) As String
    Dim prompt As String
    prompt = "�Ʒ��� �ѱ� Ű���带 1688.com���� ������ ���� ���̴� �߱��� �˻��� 1���� ������ �ּ���. " & _
             "�����ϸ� �ܾ� �״�θ�, ���� ���� �߱��� �˻�� ��ȯ�ϼ���. " & _
             "Ű����: " & korKeyword

    Get1688Chinese = CallGPTAPI(prompt)
End Function

' OpenAI API ȣ�� (���信�� ù �ٸ� ��ȯ)
Private Function CallGPTAPI(prompt As String) As String
    On Error GoTo ErrorHandler

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    Dim requestBody As String
    requestBody = "{" & _
                  """model"": """ & MODEL_NAME & """," & _
                  """messages"": [" & _
                  "{""role"": ""system"", ""content"": ""You are a helpful assistant that translates Korean e-commerce search keywords into real Chinese search keywords for 1688.com. Reply with only the main Chinese keyword, no extra explanations or examples.""}, " & _
                  "{""role"": ""user"", ""content"": """ & JsonEscape(prompt) & """}" & _
                  "]," & _
                  """temperature"": 0.3" & _
                  "}"

    With http
        .Open "POST", API_URL, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & GetAPIKey()
        .Send requestBody

        If .Status <> 200 Then
            CallGPTAPI = "API ����: " & .Status
            Exit Function
        End If

        CallGPTAPI = ExtractContent(.responseText)
    End With
    Exit Function

ErrorHandler:
    CallGPTAPI = "���� ����"
End Function

' API Key 반환 (실제 키 값은 별도 관리)
Private Function GetAPIKey() As String
    ' TODO: 실제 사용 시 본인의 OpenAI API 키를 입력하세요
    GetAPIKey = "YOUR_OPENAI_API_KEY_HERE"
End Function

' JsonEscape
Private Function JsonEscape(text As String) As String
    Dim result As String
    result = Replace(text, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    JsonEscape = result
End Function

' ���信�� content ���� (ù �ٸ� ��ȯ)
Private Function ExtractContent(json As String) As String
    On Error GoTo ErrorHandler

    Dim contentStart As Long
    contentStart = InStr(json, """content"": """)
    If contentStart = 0 Then ExtractContent = "���� ����": Exit Function

    contentStart = contentStart + Len("""content"": """)
    Dim contentEnd As Long
    contentEnd = InStr(contentStart, json, """")
    If contentEnd = 0 Then ExtractContent = "���� ����": Exit Function

    Dim content As String
    content = Mid(json, contentStart, contentEnd - contentStart)
    content = Replace(content, "\n", vbNewLine)
    content = Replace(content, "\""", """")
    content = Replace(content, "\r", "")
    content = Replace(content, "\t", "")
    content = Replace(content, "\u", "")

    content = Trim(Split(content, vbNewLine)(0))
    ExtractContent = content
    Exit Function

ErrorHandler:
    ExtractContent = "���� �Ľ� ����"
End Function

' ���� �� ���ĺ� �� ĭ ���� (Z->AA ����)
Private Function NextColLetter(colLetter As String) As String
    Dim n As Long
    n = Range(colLetter & "1").column + 1
    NextColLetter = Split(Cells(1, n).Address(True, False), "$")(0)
End Function






