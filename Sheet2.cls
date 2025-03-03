VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'sheet2 库存


Public Sub getInput(ByVal str As String, ByVal rg As Range) '获取出入库数据
        Dim r As Range
        Dim s As New stockClass '库存对象声明
        For Each r In rg.Rows
            If r.Cells(1, 1) <> "" Then
                '开始初始化库存对象
                s.code = r.Cells(1, 1)
                s.barcode = r.Cells(1, 2)
                s.name = r.Cells(1, 3)
                s.spec = r.Cells(1, 4)
                s.unit = r.Cells(1, 5)
                s.clasf = r.Cells(1, 6)
                s.num = r.Cells(1, 7)
                s.costPrice = r.Cells(1, 8)
                s.salePrice = r.Cells(1, 9)
                s.visitor = Sheet5.Range("M23")
                '结束初始化库存对象
                If r.Cells(1, 10) <> "" And r.Cells(1, 10) <> 0 Then ' r.Cells(1, 10) <> 0 识别是库存转换时无库存记录的商品
                    s.prodDate = DateValue(r.Cells(1, 10))
                Else
                    s.prodDate = DateValue(Now())  '无填入生产日期，自动设为入库日期
                End If
                If r.Cells(1, 11) <> "" Then
                    s.expiryDays = r.Cells(1, 11)
                Else
                    s.expiryDays = 180  '无填入有效期，自动设置6个月，参考牛奶有效期6个月
                End If
                Call updateAndSave(str, s) '保存对象数据到库存表
                Call writeLog(str, s) '写入日志流水账
            Else
                '空行忽略
            End If
        Next r
        If str = "in" Then
            MsgBox "入库成功！"
        Else
    '       MsgBox "出库成功！"
        End If
        Set s = Nothing
End Sub



Private Sub updateAndSave(ByVal str As String, ByVal s As stockClass) '操作数据在库存更新并保存
    Dim n As Long
    n = module1.getStockNum(s.code) '以商品编码为主键，查找库存表出现商品编码的行数
    If n > 0 Then
        '有库存记录，更新数量
        If str = "in" Then
            Cells(n, 7) = Cells(n, 7) + s.num '入库库存增加
            Cells(n, 12) = Cells(n, 12) + s.totalValue '价值总量增加
        Else
            Cells(n, 7) = Cells(n, 7) - s.num '出库库存减少
             Cells(n, 12) = Cells(n, 12) - s.totalValue '价值总量减少
             Cells(n, 13) = Cells(n, 13) + 1 '销售次数更新
        End If
        '开始初始化库存对象 前6项不变
        Cells(n, 8) = s.costPrice
        Cells(n, 9) = s.salePrice
        Cells(n, 10) = s.profit
        Cells(n, 11) = s.profitRate
        Cells(n, 14) = Now()
        Cells(n, 15) = DateValue(s.prodDate)
        Cells(n, 16) = s.expiryDays
        Cells(n, 17) = DateValue(s.expiryDate)
        '结束初始化库存对象
    Else ' n = 0
        '入库无库存记录，新建记录并填写数据 出库已拦截此类情况，所以无需操作
        '开始填写
        If str = "in" Then
            n = Sheet2.UsedRange.Rows.Count + 1
            Cells(n, 1) = s.code
            Cells(n, 2) = s.barcode
            Cells(n, 3) = s.name
            Cells(n, 4) = s.spec
            Cells(n, 5) = s.unit
            Cells(n, 6) = s.clasf
            Cells(n, 7) = s.num
            Cells(n, 8) = s.costPrice
            Cells(n, 9) = s.salePrice
            Cells(n, 10) = s.profit
            Cells(n, 11) = s.profitRate
            Cells(n, 12) = s.totalValue
            Cells(n, 13) = 0
            Cells(n, 14) = Now()
            Cells(n, 15) = DateValue(s.prodDate)
            Cells(n, 16) = s.expiryDays
            Cells(n, 17) = DateValue(s.expiryDate)
        Else
            '出库库无库存记录，点击出库时已经查询过库存，所以不会出现
        End If
        '结束填写
    End If
End Sub

'商品编码 1
'商品名称 2
'规格 3
'单位 4
'分类 5
'操作时间 6
'操作类型 7
'操作前库存 8
'操作数量 9
'操作后库存 10
'操作金额 11
'销售利润 12
'历史总利润 13
'客户ID 14
'库存总价值 15
'进货价 16
'销售价 17
'单位毛利 18
'毛利率 19
'销售次数 20
'生产日期 21
'保质期 22
'过期 23


Private Sub writeLog(ByVal op As String, ByVal s As stockClass) '写入日志流水账
     Dim n As Long
    n = Sheet3.UsedRange.Rows.Count + 1 '获取新空行
    '开始填写
    With Sheet3
            .Cells(n, 1) = s.code
            .Cells(n, 2) = s.name
            .Cells(n, 3) = s.spec
            .Cells(n, 4) = s.unit
            .Cells(n, 5) = s.clasf
            .Cells(n, 6) = Now()
            .Cells(n, 9) = s.num
            .Cells(n, 10) = module1.getStockCounts(s.code)
            .Cells(n, 11) = s.totalValue
            .Cells(n, 14) = s.visitor
            .Cells(n, 16) = s.costPrice
            .Cells(n, 17) = s.salePrice
            .Cells(n, 18) = s.profit
            .Cells(n, 19) = s.profitRate
            .Cells(n, 20) = module1.getShoppingTimesFromCode(s.code)
            .Cells(n, 21) = s.prodDate
            .Cells(n, 22) = s.expiryDays
            .Cells(n, 23) = s.expiryDate
            .Cells(n, 24) = s.salePrice * s.num
            If op = "in" Then
                .Cells(n, 7) = "进货入库"
                 .Cells(n, 12) = 0 '进货，利润不变
                 .Cells(n, 8) = module1.getStockCounts(s.code) - s.num  '入库 操作前库存 = 操作后库存 - 操作数量
                If n = 2 Then
                    .Cells(n, 15) = s.totalValue '首行入库，库存总价值 = 入库金额
                    .Cells(n, 13) = 0 '首行入库, 历史总利润 = 0
                Else
                    .Cells(n, 15) = s.totalValue + .Cells(n - 1, 15) '非首行入库，库存总价值 = 入库金额 + 库存金额
                    .Cells(n, 13) = .Cells(n - 1, 13) '非首行入库, 历史总利润不变
                End If
            Else 'op = out
                .Cells(n, 7) = "销售出库"
                 .Cells(n, 12) = s.profit * s.num
                 .Cells(n, 8) = module1.getStockCounts(s.code) + s.num  '出库库 操作前库存 = 操作后库存 + 操作数量
                If n = 2 Then
                    '首行流水日志不会是销售出库，前端会检测库存
                    .Cells(n, 13) = s.profit * s.num '首行销售, 历史总利润 = 销售利润
                Else
                    .Cells(n, 13) = s.profit * s.num + .Cells(n - 1, 13) '非首行销售, 历史总利润 = 销售利润 + 历史总利润
                    .Cells(n, 15) = .Cells(n - 1, 15) - s.totalValue '非首行出库，库存总价值 = 库存金额 - 出库金额
                End If
            End If
    End With
    '结束填写
End Sub
