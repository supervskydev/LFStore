Attribute VB_Name = "module1"
Option Explicit
'module 1



Public Sub setInfoFromCode(ByVal ws As Worksheet) '根据输入顺序码补全商品信息
    Dim j As Long
    
    Dim k As Long
    Dim rg As Range
    For Each rg In ws.Range("A2:M21").CurrentRegion.Rows
        If rg.Cells(1, 1) <> "" And rg.Cells(1, 1) <> "商品编码" Then
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
                Debug.Print "编码库不存在此编码" & rg.Cells(1, 1)
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
        If rg.Cells(1, 2) <> "" And rg.Cells(1, 1) <> "商品编码" Then
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
                Debug.Print "商品库中不存在该条形码 " & rg.Cells(1, 2)
                rg.Cells(1, 1) = Sheet4.UsedRange.Rows.Count '可能是入库更新编码库，自动补全未使用的最小编码
            End If
        Else
        End If
    Next rg
End Sub

Public Sub autoSetInfo(ByVal ws As Worksheet)
    Call setInfoFromBarcode(ws)
    Call setInfoFromCode(ws)
End Sub


Public Function getStockNum(ByVal code As Long) '提供商品【顺序码】，查找顺序码在【库存表】的行号，查不到返回0
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

Public Function getStockCounts(ByVal code As Long) '提供商品【顺序码】，查找【库存表】数量，找到返回数量，查不到返回0
    If getStockNum(code) > 0 Then
        getStockCounts = Sheet2.Cells(getStockNum(code), 7)
    Else
        getStockCounts = 0
    End If
End Function

Public Function getCodeNum(ByVal code As Long) '提供商品【顺序码】，查找顺序码在【编码表】的行号，查不到返回0
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
Public Function getBarcodeNum(ByVal barcode As String) '提供商品【条形码】，查找顺序码在【编码库】的行号，查不到返回0
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

Public Sub reset(ByVal s As String) '删除工作表第二行开始的所有行
    Dim row As Long
    Dim ws As Worksheet
    Set ws = Worksheets(s)
    row = ws.UsedRange.Rows.Count + 1 '必须+1 否则当row = 1 会删除首行表头
    ws.Activate '此行解决select方法无效的问题
    ws.Range("a2:" & "a" & row).EntireRow.Select
    Selection.Delete
End Sub



Public Function getShoppingTimesFromCode(ByVal code As Long) '提供商品【顺序码】，查找商品购买次数，查不到返回0
    Dim i As Long
    i = getStockNum(code)
    If i <> 0 Then
        getShoppingTimesFromCode = Sheet2.Cells(i, 13)
    Else
        getShoppingTimesFromCode = 0
    End If
'    Debug.Print Sheet2.Cells(1, 3) & "购买次数：" & getShoppingTimesFromCode
End Function


Public Sub newCodeRecord(ByVal rg As Range) ' 传递参数rg作为参数，新增编码记录保存rg信息
    Dim newRow As Long
    Dim rg1 As Range
    Dim i%
    newRow = Sheet4.UsedRange.Rows.Count + 1 '追加入库流水的行号
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

Public Function totalStockValue() '获取实时库存金额
   Dim i As Long
   i = Sheet3.UsedRange.Rows.Count
   If i > 1 Then
        totalStockValue = Sheet3.Range("o" & i)
    Else
        totalStockValue = 0
    End If
End Function

Public Function totalSales() '获取实时库存金额
   Dim i As Long: Dim j As Long
   j = Sheet3.UsedRange.Rows.Count
   If j > 1 Then
        For i = 2 To j
            With Sheet3
                If .Range("G" & i) = "销售出库" Then
                    totalSales = totalSales + .Range("I" & i) * .Range("Q" & i)
                End If
            End With
        Next i
    Else
        totalSales = 0
   End If
End Function

Public Function totalSalesProfits() '获取实时销售毛利
   Dim i As Long: Dim j As Long
   j = Sheet3.UsedRange.Rows.Count
   If j > 1 Then
        For i = 2 To j
            With Sheet3
                If .Range("G" & i) = "销售出库" Then
                    totalSalesProfits = totalSalesProfits + .Range("M" & i)
                End If
            End With
        Next i
    Else
        totalSalesProfits = 0
   End If
End Function


Public Function totalSalesProfitRate() '获取实时销售毛利率
   If totalSales <> 0 Then
        totalSalesProfitRate = totalSalesProfits / totalSales
   Else
        totalSalesProfitRate = 0
   End If
End Function

Public Function salesProfitRateToday() '获取实时销售毛利率
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
            If .Cells(i, 7) = "销售出库" And DateValue(Now()) = DateValue(.Cells(i, 6)) Then
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
            If .Cells(i, 7) = "销售出库" And DateValue(Now()) = DateValue(.Cells(i, 6)) Then
                salesProfitsToday = salesProfitsToday + .Range("L" & i)
            End If
        Next i
    End With
End Function

Sub 下拉框7_更改() '选择付款方式
    Dim rg As Range
    Set rg = Range("C29:C32")
    Dim n As Integer
    n = Range("G25") '付款方式对应数字 1 to 4
    rg.Cells(n, 1) = rg.Cells(n, 1) + Range("I25")
    Set rg = Nothing
    Range("I25") = 0
End Sub
