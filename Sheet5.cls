VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'sheet5 出入库

Private Sub cmdClear_Click() '清空输入数据
    Range("A2:M21").ClearContents
    Range("A2:M21").Interior.Color = vbWhite
    Range("L2:L21") = 0
    Range("G23,I23,I25").ClearContents
End Sub


Private Sub cmdOut_Click() '出库
    Range("A2:M23").Interior.Color = vbWhite
        Dim rg As Range
        Set rg = Range("A2:M23")
        '提交库存处理要先检查库存是否够用，如果不够用提示库存数量，并对品种的数量标红色
        If isDuplicate Then
             MsgBox "存在重复数据！请修改后重新更新！"
             summerizeData
        ElseIf isStockEnough Then
            Range("M23") = Range("M23") + 1  '每成功提交一次销售出库，客户人次+1
            Call Sheet2.getInput("out", rg)
            Range("I25") = Range("G23") '作用？
            module1.下拉框7_更改  '触发运行下拉框代码, 计算销售额，分4种收款方式展示
            updateStatistics
        Else
            Debug.Print "cmdOut_Click：库存不足，提示修改后重新提交"
        End If
'    Else
'        '取消
'    End If
End Sub

Private Sub cmdReset_Click()
    '清空库存表 清空流水账 清空输入表
    Dim reset As Integer
    reset = MsgBox("此操作会清空库存表和流水账，而且当前输入也会被清空，请确定是否继续？", vbOKCancel, "恢复出厂提醒")
    If reset = 1 Then
        module1.reset ("库存")
        module1.reset ("流水账")
        Sheet5.Range("A2:M21,C29:C32,C35:C38").ClearContents
        Sheet5.Range("M23") = 0
    End If
    Dim ws As Worksheet
    Set ws = Worksheets("出入库")
    ws.Activate '界面回到出入库
End Sub

Private Sub cmdIn_Click() '入库按钮被点击
    Dim i As Long: Dim rg As Range: Dim r As Range
    Range("A2:M21").Interior.Color = vbWhite
    Set rg = Range("A2:M21")
    For Each r In rg.Rows
        If r.Cells(1, 1) <> "" Then
            i = module1.getCodeNum(r.Cells(1, 1))
            If i = 0 Then '编码库无记录，需要插入新的编码记录
                Call module1.newCodeRecord(r) '调用带参数的过程记得添加call关键字
            Else
                '编码库有记录, 单元格选择变化时自动补全商品信息。
                '如果进货价和零售价非空，价格更新到编码库
                Sheet4.Cells(i, 7) = r.Cells(1, 8)
                Sheet4.Cells(i, 8) = r.Cells(1, 9)
                updateStatistics
            End If
        End If
    Next r
    Call Sheet2.getInput("in", rg)
End Sub



Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Call module1.autoSetInfo(Sheets(ActiveSheet.name))
    getSaleDetail
    Range("G23") = salesVolumn
    Range("I23") = purchaseCosts
    getSaleDetail
    updateStatistics
End Sub

Private Sub updateStatistics() '获取今日销售数据
    Range("C38") = module1.totalStockValue
    Range("C35") = module1.salesToday
    Range("C36") = module1.salesProfitsToday
    Range("C37") = module1.salesProfitRateToday
End Sub

Private Sub getSaleDetail() '计算表格分项金额
    Dim i As Integer
    For i = 2 To 21
        If Cells(i, 7) <> "" And Cells(i, 9) <> "" Then
            Cells(i, 12) = Cells(i, 7) * Cells(i, 9)
            Cells(i, 13) = module1.getStockCounts(Cells(i, 1))
        End If
    Next i
End Sub


Public Function isStockEnough()
'    Debug.Print "检查库存是否充足"
    Dim rg As Range
    Dim j As Long
    Dim s As String
    isStockEnough = True
    For Each rg In Range("A2:M21").Rows
        If rg.Cells(1, 1) <> "" Then
            j = module1.getStockNum(rg.Cells(1, 1))
            If j = 0 Then
                rg.Cells(1, 7).Interior.Color = vbRed
                s = s & rg.Cells(1, 3) & "无库存记录!请检查后修改！" & Chr(10)
                rg.Cells(1, 7) = 0
                isStockEnough = False
            Else
                Dim n As Double
                n = Sheet2.Cells(j, 7)
                If n - rg.Cells(1, 7) >= 0 Then
                    If rg.Cells(1, 7) = 0 Then
                        rg.Cells(1, 7).Interior.Color = vbRed
                        MsgBox rg.Cells(1, 3) & "操作数量不能为0！"
                        isStockEnough = False
                    Else
                        s = s & rg.Cells(1, 3) & "库存足够!" & Chr(10)
                    End If
                Else
                    rg.Cells(1, 7).Interior.Color = vbRed
                    rg.Cells(1, 7) = n
                    s = s & rg.Cells(1, 3) & "库存不足!实际库存只有" & n & Sheet2.Cells(2, 5) & Chr(10)
                    isStockEnough = False
                End If 'n
            End If ' j
        End If 'rg
    Next rg
    If isStockEnough Then
        MsgBox s & "销售出库成功！"
    Else
        MsgBox s & "已自动修改为库存数量，请检查后重新提交！这次操作不会记录到流水账日志！"
    End If
End Function

Public Function isDuplicate() '检查是否有重复品种
    Range("A2:M21").Interior.Color = vbWhite
    Dim dict As New Scripting.Dictionary
    Dim rg As Range
    Dim i As Long
    dict.RemoveAll
    isDuplicate = False
    For Each rg In Range("A2:M21").Rows
        If rg.Cells(1, 1) <> "" Then
            i = rg.Cells(1, 1)
            If Not dict.Exists(i) Then
                dict.Add i, rg.Cells(1, 10)
            Else
                isDuplicate = True
                Range(rg.Cells(1, 1), rg.Cells(1, 13)).Interior.Color = vbRed
            End If
        End If
    Next rg
    Set dict = Nothing
End Function

Private Sub summerizeData()
'    Debug.Print "汇总重复数据"
    Dim rg As Range: Dim i As Long: Dim j As Integer: Dim max_row As Integer
    max_row = Sheet5.Range("A2:M21").CurrentRegion.Columns.Count
    For i = 2 To 21
        If Range("A" & i).Interior.Color = vbRed And Range("G" & i) <> 0 Then
            For j = 2 To i
                If Cells(i, 1) = Cells(j, 1) And Not Range("A" & j).Interior.Color = vbRed Then
                    Cells(j, 7) = Cells(j, 7) + Cells(i, 7)
                   Range(Cells(i, 1), Cells(i, max_row)).Interior.Color = vbWhite
                   Range(Cells(i, 1), Cells(i, max_row)).ClearContents
                   Cells(i, 12) = 0 '出现空行要补数据，避免currentRegion出错
                   Exit For
                End If
            Next j
        ElseIf Range("A" & i).Interior.Color = vbRed Then
            Range(Cells(i, 1), Cells(i, max_row)).Interior.Color = vbWhite
            Range(Cells(i, 1), Cells(i, max_row)).ClearContents
            Cells(i, 12) = 0 '出现空行要补数据，避免currentRegion出错
        Else
            
        End If
    Next i
End Sub

Private Function salesVolumn() '自动更新销售额
    Dim i As Integer
    For i = 2 To 21
        If Cells(i, 1) <> "" Then
            salesVolumn = salesVolumn + Cells(i, 12)
        End If
    Next i
End Function

Private Function purchaseCosts() '自动更新进货费用
    Dim i As Integer
    For i = 2 To 21
        If Cells(i, 1) <> "" Then
            purchaseCosts = purchaseCosts + Cells(i, 7) * Cells(i, 8)
        End If
    Next i
End Function



