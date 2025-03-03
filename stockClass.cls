VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stockClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'库存类
Private code_ As Long '商品编码
Private barcode_ As String '条形码
Private name_    As String      '商品名称
Private num_ As Double         '库存数量
Private unit_    As String       '单位
Private spec_    As String       '规格
Private clasf_   As String       '分类
Private costPrice_   As Double       '进货价
Private salePrice_   As Double       '销售价
Private checkInDate_ As Date         '入库日期
Private shelve_  As String       '货架
Private prodDate_ As Date         '生产日期
Private expiryDays_  As Integer      '保质天数
Private visitor_ As Long
'soldTimes 销量不能初始化或直接计算得出，所以不能列为类属性或方法


'方法列表
Private Sub Class_Initialize() '初始化
    checkInDate_ = Now()
    shelve_ = "X"
    visitor_ = Sheet5.Range("M23")
End Sub

Function profit() '单票毛利
    If salePrice_ >= 0 And costPrice_ >= 0 Then
        profit = salePrice_ - costPrice_
    Else
        Debug.Print "NOT salePrice_ >= 0 And costPrice_ >= 0"
    End If
End Function

Function profitRate() '毛利率
    If salePrice_ <> 0 Then
        profitRate = profit / salePrice_
    Else
        Debug.Print "salePrice is 0!"
        profitRate = 0
    End If
End Function

Function totalValue() '商品价值总量
    If costPrice_ >= 0 And num_ >= 0 Then
        totalValue = costPrice_ * num_
    Else
        Debug.Print "NOT costPrice_ >= 0 And num_ >= 0!"
    End If
End Function



Function expiryDate() '过期日期
    expiryDate = DateAdd("d", expiryDays_, prodDate_)
End Function


' 属性 getter
Property Get code() As Long
    code = code_
End Property

Property Get barcode() As String
    barcode = barcode_
End Property

Property Get name() As String
    name = name_
End Property

Property Get num() As Double
    num = num_
End Property

Property Get unit() As String
    unit = unit_
End Property

Property Get spec() As String
    spec = spec_
End Property

Property Get clasf() As String
    clasf = clasf_
End Property

Property Get costPrice() As Double
    costPrice = costPrice_
End Property

Property Get salePrice() As Double
    salePrice = salePrice_
End Property

Property Get checkInDate() As Date
    checkInDate = checkInDate_
End Property

Property Get shelve() As String
    shelve = shelve_
End Property

Property Get prodDate() As Date
    prodDate = prodDate_
End Property

Property Get expiryDays() As Integer
    expiryDays = expiryDays_
End Property

Property Get visitor() As Long
    visitor = visitor_
End Property


'setter
Property Let name(v As String)
    name_ = v
End Property

Property Let num(v As Double)
    num_ = v
End Property

Property Let unit(v As String)
    unit_ = v
End Property

Property Let spec(v As String)
    spec_ = v
End Property

Property Let clasf(v As String)
    clasf_ = v
End Property

Property Let costPrice(v As Double)
    costPrice_ = v
End Property

Property Let salePrice(v As Double)
    salePrice_ = v
End Property

Property Let checkInDate(v As Date)
    checkInDate_ = v
End Property

Property Let shelve(v As String)
    shelve_ = v
End Property

Property Let prodDate(v As Date)
    prodDate_ = v
End Property

Property Let expiryDays(v As Integer)
    expiryDays_ = v
End Property
Property Let code(v As Long)
    code_ = v
End Property

Property Let barcode(v As String)
    barcode_ = v
End Property

Property Let visitor(v As Long)
    visitor_ = v
End Property
