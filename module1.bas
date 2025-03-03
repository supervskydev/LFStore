Attribute VB_Name = "module1"
Option Explicit
'module 1



Public Sub setInfoFromCode(ByVal ws As Worksheet) '��������˳���벹ȫ��Ʒ��Ϣ
    Dim j As Long
    
    Dim k As Long
    Dim rg As Range
    For Each rg In ws.Range("A2:M21").CurrentRegion.Rows
        If rg.Cells(1, 1) <> "" And rg.Cells(1, 1) <> "��Ʒ����" Then
            j = module1.getCodeNum(rg.Cells(1, 1))
            If j <> 0 Then
                For k = 2 To 6
                   rg.Cells(1, k) = Sheet4.Cells(j, k)
                Next k
                If ws.Range("A1").CurrentRegion.Columns.Count <= 6 Then
                Else
                    If rg.Cells(1, 7) = "" Then
                        rg.Cells(1, 7) = 1
                    End If
                    If rg.Cells(1, 8) = "" Then
                        rg.Cells(1, 8) = Sheet4.Cells(j, 7)
                    End If
                    If rg.Cells(1, 9) = "" Then
                        rg.Cells(1, 9) = Sheet4.Cells(j, 8)
                    End If
                End If
            Else
                Debug.Print "����ⲻ���ڴ˱���" & rg.Cells(1, 1)
            End If
        Else
        End If
    Next rg
End Sub

Public Sub setInfoFromBarcode(ByVal ws As Worksheet)
    Dim j As Long
    Dim k As Long
    Dim rg As Range
    For Each rg In ws.Range("A2:M21").CurrentRegion.Rows
        If rg.Cells(1, 2) <> "" And rg.Cells(1, 1) <> "��Ʒ����" Then
            j = module1.getBarcodeNum(rg.Cells(1, 2))
            If j <> 0 Then
                For k = 1 To 6
                   rg.Cells(1, k) = Sheet4.Cells(j, k)
                Next k
                If rg.Cells(1, 7) = "" Then
                    rg.Cells(1, 7) = 1
                End If
                If rg.Cells(1, 8) = "" Then
                    rg.Cells(1, 8) = Sheet4.Cells(j, 7)
                End If
                If rg.Cells(1, 9) = "" Then
                    rg.Cells(1, 9) = Sheet4.Cells(j, 8)
                End If
            Else
                Debug.Print "��Ʒ���в����ڸ������� " & rg.Cells(1, 2)
                rg.Cells(1, 1) = Sheet4.UsedRange.Rows.Count '�����������±���⣬�Զ���ȫδʹ�õ���С����
            End If
        Else
        End If
    Next rg
End Sub

Public Sub autoSetInfo(ByVal ws As Worksheet)
    Call setInfoFromBarcode(ws)
    Call setInfoFromCode(ws)
End Sub


Public Function getStockNum(ByVal code As Long) '�ṩ��Ʒ��˳���롿������˳�����ڡ��������кţ��鲻������0
    Dim j As Long
    Dim find As Boolean
    find = False
    For j = 2 To Sheet2.UsedRange.Rows.Count
        If code = Sheet2.Cells(j, 1) Then
            find = True
            Exit For
        End If
    Next j
    If find Then
        getStockNum = j
    Else
        getStockNum = 0
    End If
End Function

Public Function getStockCounts(ByVal code As Long) '�ṩ��Ʒ��˳���롿�����ҡ������������ҵ������������鲻������0
    If getStockNum(code) > 0 Then
        getStockCounts = Sheet2.Cells(getStockNum(code), 7)
    Else
        getStockCounts = 0
    End If
End Function

Public Function getCodeNum(ByVal code As Long) '�ṩ��Ʒ��˳���롿������˳�����ڡ���������кţ��鲻������0
    Dim j As Long
    Dim find As Boolean
    find = False
    For j = 2 To Sheet4.UsedRange.Rows.Count
        If code = Sheet4.Cells(j, 1) Then
            find = True
            Exit For
        End If
    Next j
    If find Then
        getCodeNum = j
    Else
        getCodeNum = 0
    End If
End Function
Public Function getBarcodeNum(ByVal barcode As String) '�ṩ��Ʒ�������롿������˳�����ڡ�����⡿���кţ��鲻������0
    Dim j As Long
    Dim find As Boolean
    find = False
    For j = 2 To Sheet4.UsedRange.Rows.Count
        If barcode = Sheet4.Cells(j, 2) Then
            find = True
            Exit For
        End If
    Next j
    If find Then
        getBarcodeNum = j
    Else
        getBarcodeNum = 0
    End If
End Function

Public Sub reset(ByVal s As String) 'ɾ��������ڶ��п�ʼ��������
    Dim row As Long
    Dim ws As Worksheet
    Set ws = Worksheets(s)
    row = ws.UsedRange.Rows.Count + 1 '����+1 ����row = 1 ��ɾ�����б�ͷ
    ws.Activate '���н��select������Ч������
    ws.Range("a2:" & "a" & row).EntireRow.Select
    Selection.Delete
End Sub



Public Function getShoppingTimesFromCode(ByVal code As Long) '�ṩ��Ʒ��˳���롿��������Ʒ����������鲻������0
    Dim i As Long
    i = getStockNum(code)
    If i <> 0 Then
        getShoppingTimesFromCode = Sheet2.Cells(i, 13)
    Else
        getShoppingTimesFromCode = 0
    End If
'    Debug.Print Sheet2.Cells(1, 3) & "���������" & getShoppingTimesFromCode
End Function


Public Sub newCodeRecord(ByVal rg As Range) ' ���ݲ���rg��Ϊ���������������¼����rg��Ϣ
    Dim newRow As Long
    Dim rg1 As Range
    Dim i%
    newRow = Sheet4.UsedRange.Rows.Count + 1 '׷�������ˮ���к�
    Set rg1 = Sheet4.Range("A" & newRow & ":H" & newRow)
    rg1.HorizontalAlignment = xlCenter
    For i = 1 To 6
         rg1.Cells(1, i) = rg.Cells(1, i)
    Next i
    If rg.Cells(1, 8) <> "" Then
        rg1.Cells(1, 7) = rg.Cells(1, 8)
    End If
    
     If rg.Cells(1, 9) <> "" Then
        
        rg1.Cells(1, 8) = rg.Cells(1, 9)
    End If
End Sub

Public Function totalStockValue() '��ȡʵʱ�����
   Dim i As Long
   i = Sheet3.UsedRange.Rows.Count
   If i > 1 Then
        totalStockValue = Sheet3.Range("o" & i)
    Else
        totalStockValue = 0
    End If
End Function

Public Function totalSales() '��ȡʵʱ�����
   Dim i As Long: Dim j As Long
   j = Sheet3.UsedRange.Rows.Count
   If j > 1 Then
        For i = 2 To j
            With Sheet3
                If .Range("G" & i) = "���۳���" Then
                    totalSales = totalSales + .Range("I" & i) * .Range("Q" & i)
                End If
            End With
        Next i
    Else
        totalSales = 0
   End If
End Function

Public Function totalSalesProfits() '��ȡʵʱ����ë��
   Dim i As Long: Dim j As Long
   j = Sheet3.UsedRange.Rows.Count
   If j > 1 Then
        For i = 2 To j
            With Sheet3
                If .Range("G" & i) = "���۳���" Then
                    totalSalesProfits = totalSalesProfits + .Range("M" & i)
                End If
            End With
        Next i
    Else
        totalSalesProfits = 0
   End If
End Function


Public Function totalSalesProfitRate() '��ȡʵʱ����ë����
   If totalSales <> 0 Then
        totalSalesProfitRate = totalSalesProfits / totalSales
   Else
        totalSalesProfitRate = 0
   End If
End Function

Public Function salesProfitRateToday() '��ȡʵʱ����ë����
   If salesToday <> 0 Then
        salesProfitRateToday = salesProfitsToday / salesToday
   Else
        salesProfitRateToday = 0
   End If
End Function


Public Function salesToday()
    Dim i As Long: Dim row As Long
    row = Sheet3.UsedRange.Rows.Count
    salesToday = 0
    With Sheet3
         For i = row To 2 Step -1
            If .Cells(i, 7) = "���۳���" And DateValue(Now()) = DateValue(.Cells(i, 6)) Then
                salesToday = salesToday + .Range("I" & i) * .Range("Q" & i)
            End If
        Next i
    End With
End Function

Public Function salesProfitsToday()
    Dim i As Long: Dim row As Long
    row = Sheet3.UsedRange.Rows.Count
    salesProfitsToday = 0
    With Sheet3
         For i = row To 2 Step -1
            If .Cells(i, 7) = "���۳���" And DateValue(Now()) = DateValue(.Cells(i, 6)) Then
                salesProfitsToday = salesProfitsToday + .Range("L" & i)
            End If
        Next i
    End With
End Function

Sub ������7_����() 'ѡ�񸶿ʽ
    Dim rg As Range
    Set rg = Range("C29:C32")
    Dim n As Integer
    n = Range("G25") '���ʽ��Ӧ���� 1 to 4
    rg.Cells(n, 1) = rg.Cells(n, 1) + Range("I25")
    Set rg = Nothing
    Range("I25") = 0
End Sub
