Attribute VB_Name = "N_VarSet"
Option Explicit
Public VBbasPath, VBbasName, VBbasFunctionName  As String
Public ZSum, ZNow, EditLock, EditMod As Long
Public ScalingRate As Single
Public XMB() As New Home
Public MousePlace As ��ά����
Public Package(1 To 999) As ��
Public NewNode As ��Ԫ
Public Const Pi = 3.14159265358979

Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A

Global lpPrevWndProcA As Long

Public bMouseFlag As Boolean '����¼������־

Public Sub HookMouse(ByVal hwnd As Long)
lpPrevWndProcA = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHookMouse(ByVal hwnd As Long)
SetWindowLong hwnd, GWL_WNDPROC, lpPrevWndProcA
End Sub

Private Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case uMsg
Case WM_MOUSEWHEEL '����
Dim wzDelta, wKeys As Integer
'wzDelta���ݹ��ֹ����Ŀ�������ֵС�����ʾ���������������û����򣩣�
'�������ʾ������ǰ����������ʾ������
wzDelta = HIWORD(wParam)
'wKeysָ���Ƿ���CTRL=8��SHIFT=4������(��=2����=16����=2������)���£�������
wKeys = LOWORD(wParam)
'--------------------------------------------------
If wzDelta < 0 Then '���û�����
    ScalingRate = ScalingRate + 0.1
    If ScalingRate > 3 Then ScalingRate = 3
Else '����ʾ������
    ScalingRate = ScalingRate - 0.1
    If ScalingRate < 0.1 Then ScalingRate = 0.1
End If
'--------------------------------------------------
Case Else
WindowProc = CallWindowProc(lpPrevWndProcA, hw, uMsg, wParam, lParam)
End Select
End Function

Private Function HIWORD(LongIn As Long) As Integer
HIWORD = (LongIn And &HFFFF0000) \ &H10000 'ȡ��32λֵ�ĸ�16λ
End Function
Private Function LOWORD(LongIn As Long) As Integer
LOWORD = LongIn And &HFFFF& 'ȡ��32λֵ�ĵ�16λ
End Function
