Attribute VB_Name = "Module7"
Sub ���λ�ǰ������Ʈ_�ڵ�ȭ()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim img As Picture
    Dim imageUrl As String
    Dim deliveryType As String
    Dim i As Long
    Dim rocketCount As Long
    Dim sellerRocketCount As Long
    Dim normalDeliveryCount As Long
    Dim outputRow As Long

    ' Ȱ�� ��Ʈ ����
    Set ws = ActiveSheet

    ' ������ �� ã�� (H�� ����)
    lastRow = ws.Cells(ws.Rows.count, "H").End(xlUp).Row

    ' F���� �����Ͱ� ���� �� ����
    For i = lastRow To 2 Step -1
        If ws.Cells(i, "F").Value = "" Then
            ws.Rows(i).Delete
        End If
    Next i

    ' ������ �� �ٽ� ��� (���� �����Ǿ��� ������ �ٽ� ����ؾ� ��)
    lastRow = ws.Cells(ws.Rows.count, "H").End(xlUp).Row

    ' 1. H���� ��ǰURL�� K���� �����۸�ũ�� ����
    For Each cell In ws.Range("H2:H" & lastRow)
        If cell.Value Like "http*://*" Then
            ws.Hyperlinks.Add Anchor:=ws.Cells(cell.Row, "K"), Address:=cell.Value, TextToDisplay:="�ٷΰ���"
        End If
    Next cell

    ' 2. I���� �̹���URL�� �����Ͽ� L���� �̹��� ����
    For Each cell In ws.Range("I2:I" & lastRow)
        imageUrl = cell.Value
        If imageUrl <> "" Then
            ' �̹��� ����
            On Error Resume Next
            Set img = ws.Pictures.Insert(imageUrl)
            If Not img Is Nothing Then
                With img
                    ' �� ���̸� �� �� ũ�� �����Ͽ� �̹����� �� ���̵���
                    ws.Cells(cell.Row, "L").RowHeight = 50
                    .Top = ws.Cells(cell.Row, "L").Top
                    .Left = ws.Cells(cell.Row, "L").Left
                    .Width = ws.Cells(cell.Row, "L").Width
                    .Height = ws.Cells(cell.Row, "L").Height
                    
                    ' ������ ���� �̹��� ũ�� ���� (�ʺ� ��������)
                    Dim ratio As Double
                    ratio = .Width / .Height
                    If .Width > ws.Cells(cell.Row, "L").Width Then
                        .Width = ws.Cells(cell.Row, "L").Width
                        .Height = .Width / ratio
                    End If
                    
                    ' �̹����� �� �߾ӿ� ��ġ�ϵ��� ����
                    .Top = ws.Cells(cell.Row, "L").Top + (ws.Cells(cell.Row, "L").Height - .Height) / 2
                    .Left = ws.Cells(cell.Row, "L").Left + (ws.Cells(cell.Row, "L").Width - .Width) / 2
                End With
            End If
            On Error GoTo 0
        End If
    Next cell

    ' 3. J���� ������¿� ���� �� ���� ���� �� ���� ����
    rocketCount = 0
    sellerRocketCount = 0
    normalDeliveryCount = 0

    For Each cell In ws.Range("J2:J" & lastRow)
        deliveryType = cell.Value
        Select Case deliveryType
            Case "�Ǹ��ڷ���"
                cell.Interior.Color = RGB(255, 200, 150) ' ���� ��������
                sellerRocketCount = sellerRocketCount + 1
            Case "���Ϲ��"
                cell.Interior.Color = RGB(150, 200, 255) ' ���� �Ķ���
                rocketCount = rocketCount + 1
            Case "�Ϲݹ��"
                cell.Interior.Color = RGB(255, 150, 150) ' ���� ������
                normalDeliveryCount = normalDeliveryCount + 1
        End Select
    Next cell

    ' 4. A��, H��, I�� �����
    ws.Columns("A:A").Hidden = True
    ws.Columns("H:I").Hidden = True

    ' 5. �׵θ� �� ���� ����
    With ws.Range("A1:L" & lastRow)
        .Borders.LineStyle = xlContinuous ' �׵θ� ����
        .HorizontalAlignment = xlCenter ' �ؽ�Ʈ ��� ����
        .VerticalAlignment = xlCenter ' ���� ��� ����
    End With

    ' C���� ���� ���� ����, C1�� ��� ����
    ws.Range("C2:C" & lastRow).HorizontalAlignment = xlLeft
    ws.Range("C1").HorizontalAlignment = xlCenter

    ' 6. K1 ���� L1 ���� "�ٷΰ���"�� "�̹���" �ؽ�Ʈ �߰�
    ws.Range("K1").Value = "�ٷΰ���"
    ws.Range("L1").Value = "�̹���"

    ' 7. ��� ���� ������ C�� ������ �� �Ʒ� �ι�°�� ���
    outputRow = lastRow + 2
    ws.Cells(outputRow, "C").Value = "������� ���: "
    ws.Cells(outputRow + 1, "C").Value = "���Ϲ��: " & rocketCount & "��"
    ws.Cells(outputRow + 2, "C").Value = "�Ǹ��ڷ���: " & sellerRocketCount & "��"
    ws.Cells(outputRow + 3, "C").Value = "�Ϲݹ��: " & normalDeliveryCount & "��"

    ' �� ��° ��� ���� - ���ݴ� �м�

    Dim priceRange As Range
    Dim qtyRange As Range

    ' ������ �� ã�� (F�� ����)
    lastRow = ws.Cells(ws.Rows.count, "F").End(xlUp).Row

    ' �ǸŰ��� �ǸŰ��� ���� ����
    Set priceRange = ws.Range("F2:F" & lastRow) ' F���� �ǸŰ�
    Set qtyRange = ws.Range("G2:G" & lastRow)   ' G���� �ǸŰ���

    ' ���ݴ� ���� ���� - ����� �Է�
    Dim interval As Double
    interval = Application.InputBox("���ݴ� ������ �Է��ϼ���:", "���ݴ� ���� �Է�", Type:=1)
    
    ' �Է°��� 0 ������ ��� ����
    If interval <= 0 Then
        MsgBox "�ùٸ� ���ݴ� ������ �Է��ϼ���."
        Exit Sub
    End If

    ' ���ݴ뺰 �ǸŰ����� ������ Dictionary ����
    Dim priceGroups As Object
    Set priceGroups = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        Dim price As Double
        Dim qty As Double
        price = ws.Cells(i, "F").Value
        qty = ws.Cells(i, "G").Value

        ' ��ȿ�� �������� Ȯ��
        If IsNumeric(price) And IsNumeric(qty) Then
            ' ���ݴ� ���
            Dim groupStart As Double
            groupStart = Int(price / interval) * interval
            groupKey = Format(groupStart, "#,##0") & " - " & Format(groupStart + interval - 1, "#,##0") ' ��ǥ �߰�

            ' Dictionary�� ����
            If priceGroups.exists(groupKey) Then
                priceGroups(groupKey) = priceGroups(groupKey) + qty
            Else
                priceGroups.Add groupKey, qty
            End If
        End If
    Next i

    ' ���� ���� �ȸ� ���ݴ� ã��
    Dim maxQty As Double
    Dim maxGroup As String
    Dim secondMaxQty As Double
    Dim secondMaxGroup As String
    Dim thirdMaxQty As Double
    Dim thirdMaxGroup As String
    maxQty = 0
    secondMaxQty = 0
    thirdMaxQty = 0

    For Each key In priceGroups.Keys
        If priceGroups(key) > maxQty Then
            ' ���� �ִ밪���� �� �ܰ辿 �б�
            thirdMaxQty = secondMaxQty
            thirdMaxGroup = secondMaxGroup
            secondMaxQty = maxQty
            secondMaxGroup = maxGroup
            maxQty = priceGroups(key)
            maxGroup = key
        ElseIf priceGroups(key) > secondMaxQty Then
            ' ���� �� ��° �ִ밪�� �� ��°�� �̵�
            thirdMaxQty = secondMaxQty
            thirdMaxGroup = secondMaxGroup
            secondMaxQty = priceGroups(key)
            secondMaxGroup = key
        ElseIf priceGroups(key) > thirdMaxQty Then
            ' �� ��° �ִ밪 ����
            thirdMaxQty = priceGroups(key)
            thirdMaxGroup = key
        End If
    Next key

    ' ��� �޽��� ����
    Dim message As String
    message = ""
    If maxGroup <> "" Then
        message = "���� ���� �ȸ� ���ݴ�� " & maxGroup & "�̸�, �� �ǸŰ����� " & Format(maxQty, "#,##0") & "�Դϴ�." & vbCrLf
        If secondMaxGroup <> "" Then
            message = message & "�� ��°�� ���� �ȸ� ���ݴ�� " & secondMaxGroup & "�̸�, �� �ǸŰ����� " & Format(secondMaxQty, "#,##0") & "�Դϴ�." & vbCrLf
        End If
        If thirdMaxGroup <> "" Then
            message = message & "�� ��°�� ���� �ȸ� ���ݴ�� " & thirdMaxGroup & "�̸�, �� �ǸŰ����� " & Format(thirdMaxQty, "#,##0") & "�Դϴ�."
        End If
        MsgBox message
    Else
        MsgBox "�����Ͱ� ���ų� ��ȿ�� ���� �����ϴ�."
        Exit Sub
    End If

    ' ������� C���� �Ϲݹ�� ��谪 �Ʒ��� ���
    outputRow = outputRow + 5 ' ��� ��� ���� 5�� �Ʒ��� ��� ���
    ws.Cells(outputRow, "C").Value = "���ݴ� �м� ���:"
    ws.Cells(outputRow + 1, "C").Value = "���� ���� �ȸ� ���ݴ�� " & maxGroup & "�̸�, �� �ǸŰ����� " & Format(maxQty, "#,##0") & "�Դϴ�."

    If secondMaxGroup <> "" Then
        ws.Cells(outputRow + 2, "C").Value = "�� ��°�� ���� �ȸ� ���ݴ�� " & secondMaxGroup & "�̸�, �� �ǸŰ����� " & Format(secondMaxQty, "#,##0") & "�Դϴ�."
    End If

    If thirdMaxGroup <> "" Then
        ws.Cells(outputRow + 3, "C").Value = "�� ��°�� ���� �ȸ� ���ݴ�� " & thirdMaxGroup & "�̸�, �� �ǸŰ����� " & Format(thirdMaxQty, "#,##0") & "�Դϴ�."
    End If

End Sub


