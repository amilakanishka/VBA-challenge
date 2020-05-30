VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGreatest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pTicker As String
Private pParam As Double

'Ticker Property
Public Property Get Ticker() As String
    Ticker = pTicker
End Property

Public Property Let Ticker(Value As String)
    pTicker = Value
End Property

'Param Property
Public Property Get Param() As Double
    Param = pParam
End Property

Public Property Let Param(Value As Double)
    pParam = Value
End Property

