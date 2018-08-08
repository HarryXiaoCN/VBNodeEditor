Attribute VB_Name = "N_Type"
Public Type 三维坐标
    X As Single
    Y As Single
    Z As Single
    Target As Long
End Type
Public Type 段
    Title As String
    Content() As String
    Start As Long
    End As Long
End Type
Public Type 单元
    A As Boolean '存在性
    Strat As Boolean
    Order As Long
    Title As String
    Content() As String
    X As Single
    Y As Single
    Z As Long
End Type
Public Type 联系
    A As Boolean
    Target As Long
    Source As Long
End Type
