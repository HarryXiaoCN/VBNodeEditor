Attribute VB_Name = "N_VarSet"
Option Explicit
Public VBbasPath, VBbasName, VBbasFunctionName  As String
Public ZSum, ZNow, EditLock, EditMod As Long
Public ScalingRate As Single
Public XMB() As New Home
Public MousePlace As 三维坐标
Public Package(1 To 999) As 段
Public NewNode As 单元
Public Const Pi = 3.14159265358979

Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A

Global lpPrevWndProcA As Long

Public bMouseFlag As Boolean '鼠标事件激活标志

Public Sub HookMouse(ByVal hwnd As Long)
lpPrevWndProcA = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHookMouse(ByVal hwnd As Long)
SetWindowLong hwnd, GWL_WNDPROC, lpPrevWndProcA
End Sub

Private Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case uMsg
Case WM_MOUSEWHEEL '滚动
Dim wzDelta, wKeys As Integer
'wzDelta传递滚轮滚动的快慢，该值小于零表示滚轮向后滚动（朝用户方向），
'大于零表示滚轮向前滚动（朝显示器方向）
wzDelta = HIWORD(wParam)
'wKeys指出是否有CTRL=8、SHIFT=4、鼠标键(左=2、中=16、右=2、附加)按下，允许复合
wKeys = LOWORD(wParam)
'--------------------------------------------------
If wzDelta < 0 Then '朝用户方向
    ScalingRate = ScalingRate + 0.1
    If ScalingRate > 3 Then ScalingRate = 3
Else '朝显示器方向
    ScalingRate = ScalingRate - 0.1
    If ScalingRate < 0.1 Then ScalingRate = 0.1
End If
'--------------------------------------------------
Case Else
WindowProc = CallWindowProc(lpPrevWndProcA, hw, uMsg, wParam, lParam)
End Select
End Function

Private Function HIWORD(LongIn As Long) As Integer
HIWORD = (LongIn And &HFFFF0000) \ &H10000 '取出32位值的高16位
End Function
Private Function LOWORD(LongIn As Long) As Integer
LOWORD = LongIn And &HFFFF& '取出32位值的低16位
End Function
