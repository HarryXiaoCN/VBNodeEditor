Attribute VB_Name = "N_Type"
Public Type 三维坐标
    X As Single
    Y As Single
    Z As Single
    Target As Long
    Aim As Long
End Type
Public Type 段
    Title As String
    Content() As String
    Start As Long
    End As Long
End Type
Public Type 单元
    a As Boolean '存在性
    Strat As Boolean
    Order As Long
    Color As Long
    Title As String
    Content() As String
    ContentOne As String
    SourceSum As Long
    TargetSum As Long
    X As Single
    Y As Single
    Z As Long
End Type
Public Type 联系
    a As Boolean
    Target As Long
    Source As Long
End Type
