VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'sheet6 库存转换

Private Sub cmdChange1_Click()
    Call exchange(2, 3)
End Sub

Private Sub cmdChange2_Click()
    Call exchange(3, 2)
End Sub

Private Sub exchange(ByVal r2 As Integer, ByVal r3 As Integer) ' 2 行数据之间进行数据转换, r2, r3 是行号
    Dim i As Long: Dim j As Long
    Dim n As Double
    Dim a As Double: Dim b As Double
    a = Cells(r2, 7): b = Cells(r3, 7)
    If a * b <> 0 Then '转换比值不能为零
        If Cells(r2, 6) = "香烟" Then
            If Cells(r2, 5) = "条" Then '条 的编码比 包的编码大1
                a = 1: b = 10
                Cells(r3, 1) = Cells(r2, 1) - 1
            Else
                a = 10: b = 1
                Cells(r3, 1) = Cells(r2, 1) + 1
            End If
        Else
            If a <= b Then
                b = b / a: a = 1
            Else
                a = a / b: b = 1
            End If
            Debug.Print "NOT 香烟"
        End If
        Cells(r2, 7) = a: Cells(r3, 7) = b
'        Debug.Print "a=" & a & " b=" & b
    Else
        MsgBox "'转换比值不能为零!"
    End If
    If Cells(r2, 1) <> "" And Cells(r3, 1) <> "" Then
       i = module1.getStockNum(Cells(r2, 1))
       j = module1.getStockNum(Cells(r3, 1))
         With Sheet2
          If Cells(r2, 10) >= a Then
               If i > 0 Then
                  If j > 0 Then
                      .Cells(i, 7) = .Cells(i, 7) - a '第二行商品库存更新
                      .Cells(i, 12) = .Cells(i, 7) * .Cells(i, 8) '第二行商品库存总量更新
                      .Cells(j, 7) = .Cells(j, 7) + b '第三行商品库存更新
                      .Cells(j, 12) = .Cells(j, 7) * .Cells(j, 8) '第三行商品库存总量更新
                  Else
                      Cells(r3, 7) = 0 '库存无记录，入库数量为0
                       Call Sheet2.getInput("in", Range("A" & r3 & ":J" & r3)) '增加库存记录，但是库存数量是0
                  End If 'j
               Else
                  MsgBox "用来转换单位的商品，库存查不到入库记录！请入库后再转换！" 'i = 0
               End If 'i
          Else
              If i = 0 Then
                MsgBox "用来转换的商品无库存记录，无法转换，请入库后再试！"
              Else
                 MsgBox .Cells(i, 3) & "库存为" & .Cells(i, 7) & .Cells(i, 5) & "库存不足，无法转换！"
              End If
          End If
           Call Worksheet_SelectionChange(Range("G2"))
        End With
    Else
        MsgBox "商品编码为空无法转换！请补全商品编码再试试！"
    End If ' Cells(r2, 1)
End Sub

Private Sub cmdClear_Click()
    Range("A2:J3").ClearContents
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Call module1.autoSetInfo(Sheets(ActiveSheet.name))
    exchangeCigerate
    Call module1.autoSetInfo(Sheets(ActiveSheet.name))
    If Cells(2, 1) <> "" Then
        Cells(2, 10) = module1.getStockCounts(Cells(2, 1))
    End If
    If Cells(3, 1) <> "" Then
        Cells(3, 10) = module1.getStockCounts(Cells(3, 1))
    End If
End Sub

Private Sub exchangeCigerate()
     Dim a As Double: Dim b As Double
    a = Cells(2, 7): b = Cells(3, 7)
    If Cells(2, 6) = "香烟" Then
        If Cells(2, 5) = "条" Then '条 的编码比 包的编码大1
            a = 1: b = 10
            Cells(3, 1) = Cells(2, 1) - 1
        Else
            a = 10: b = 1
            Cells(3, 1) = Cells(2, 1) + 1
        End If
'        Debug.Print "香烟"
    Else
        '非香烟处理代码
    End If
    Cells(2, 7) = a: Cells(3, 7) = b
    
End Sub
