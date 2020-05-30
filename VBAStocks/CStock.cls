VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pTicker As String
Private pYOpenPrice As Double
Private pYClosePrice As Double
Private pYVolume As Double
Private pDOpenPrice As Double
Private pDClosePrice As Double
Private pDVolume As Long
Private pStockDate As Long



'Ticker Property
Public Property Get Ticker() As String
    Ticker = pTicker
End Property

Public Property Let Ticker(Value As String)
    pTicker = Value
End Property

'YOpenPrice Property
Public Property Get YOpenPrice() As Double
    YOpenPrice = pYOpenPrice
End Property

Public Property Let YOpenPrice(Value As Double)
    pYOpenPrice = Value
End Property

'YClosePrice Property
Public Property Get YClosePrice() As Double
    YClosePrice = pYClosePrice
End Property

Public Property Let YClosePrice(Value As Double)
    pYClosePrice = Value
End Property

'YPriceChange Property
Public Property Get YPriceChange() As Double
    YPriceChange = pYClosePrice - pYOpenPrice
End Property

'YPercentChange Property
Public Property Get YPercentChange() As Double
    If pYOpenPrice <> 0 Then
        YPercentChange = (pYClosePrice - pYOpenPrice) / pYOpenPrice
    ElseIf (pYClosePrice = 0 And pYClosePrice = 0) Then
        YPercentChange = 0
    ElseIf pYClosePrice > 0 Then
        On Error Resume Next
            YPercentChange = pYClosePrice / 0        '  1.#INF
    End If
        
End Property


'YVolume Property
Public Property Get YVolume() As Double
    YVolume = pYVolume
End Property

Public Property Let YVolume(Value As Double)
    pYVolume = Value
End Property

'DOpenPrice Property
Public Property Get DOpenPrice() As Double
    DOpenPrice = pDOpenPrice
End Property

Public Property Let DOpenPrice(Value As Double)
    pDOpenPrice = Value
End Property

'DClosePrice Property
Public Property Get DClosePrice() As Double
    DClosePrice = pDClosePrice
End Property

Public Property Let DClosePrice(Value As Double)
    pDClosePrice = Value
End Property

'DVolume Property
Public Property Get DVolume() As Long
    DVolume = pDVolume
End Property

Public Property Let DVolume(Value As Long)
    pDVolume = Value
End Property

'StockDate Property
Public Property Get StockDate() As Long
    StockDate = pStockDate
End Property

Public Property Let StockDate(Value As Long)
    pStockDate = Value
End Property

