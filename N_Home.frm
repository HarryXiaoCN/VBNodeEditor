VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Home 
   AutoRedraw      =   -1  'True
   Caption         =   "Home"
   ClientHeight    =   6960
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   13815
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   13200
      Top             =   6480
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu 菜单 
      Caption         =   "菜单"
      Begin VB.Menu 新建 
         Caption         =   "新建"
      End
      Begin VB.Menu 打开 
         Caption         =   "打开"
      End
      Begin VB.Menu 保存 
         Caption         =   "保存"
      End
      Begin VB.Menu 另存为 
         Caption         =   "另存为"
      End
      Begin VB.Menu 退出 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private 节点() As 单元
Private 连接(1 To 999) As 联系
Private MeID, LSum As Long
Private NodeReDimLock As Boolean
Private Sub Form_Activate()
Dim STemp '获得窗体的Z坐标
STemp = Split(Me.Caption, " - ")
MeID = Val(STemp(0))
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
If NodeReDimLock = False Then Exit Sub
For i = 1 To UBound(节点) - 1
    If 节点(i).A = True Then
        If X + 100 > 节点(i).X And X - 100 < 节点(i).X _
        And Y + 100 > 节点(i).Y And Y - 100 < 节点(i).Y Then
            MousePlace.Target = i: Exit Sub
        End If
    End If
Next
MousePlace.Target = 0
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePlace.Target = 0
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With MousePlace
    .X = X
    .Y = Y
    .Z = MeID
End With
End Sub
'Private Sub Form_Resize()
'Me.Scale (0, 0)-(Home.Width, Home.Height)
'End Sub
Private Sub Timer1_Timer()
Dim i As Long
Me.Cls
If NodeReDimLock = False Then Exit Sub
Me.Print UBound(节点)
For i = 1 To UBound(节点) - 1
    If 节点(i).A = True Then
        Me.CurrentX = 节点(i).X
        Me.CurrentY = 节点(i).Y
        Me.Print 节点(i).Title
        If MousePlace.X + 100 > 节点(i).X And MousePlace.X - 100 < 节点(i).X _
        And MousePlace.Y + 100 > 节点(i).Y And MousePlace.Y - 100 < 节点(i).Y Then
            Me.Circle (节点(i).X, 节点(i).Y), 100, RGB(255, 0, 0)
            Me.CurrentX = 0: Me.CurrentY = 200: Me.Print i
        Else
            Me.Circle (节点(i).X, 节点(i).Y), 100, RGB(0, 191, 255)
        End If
    End If
Next
For i = 1 To LSum
    If 连接(i).A = True Then
        Me.Line (节点(连接(i).Source).X, 节点(连接(i).Source).Y)-(节点(连接(i).Target).X, 节点(连接(i).Target).Y), RGB(126, 126, 126)
    End If
Next
If MousePlace.Target <> 0 And MousePlace.Z = MeID Then
    节点(MousePlace.Target).X = MousePlace.X
    节点(MousePlace.Target).Y = MousePlace.Y
End If
End Sub
Private Sub 打开_Click()
' 设置“CancelError”为 True
CommonDialog1.CancelError = True
'On Error GoTo ErrHandler
' 设置标志
CommonDialog1.Flags = cdlOFNHideReadOnly
' 设置过滤器
CommonDialog1.Filter = "VBBas Files" & _
"(*.bas)|*.bas|All Files (*.*)|*.*"
' 指定缺省的过滤器
CommonDialog1.FilterIndex = 1
' 显示“打开”对话框
CommonDialog1.ShowOpen
' 显示选定文件的名字
VBbasName = CommonDialog1.FileTitle
VBbasPath = CommonDialog1.FileName
'---------------VBbasToN-----------
Dim StrLine() As String: Dim LineSum, fSUM, i, c, j As Long
Dim SNTemp As Long
Open VBbasPath For Input As #1
    Do Until EOF(1)
        LineSum = LineSum + 1
        ReDim Preserve StrLine(LineSum)
        Line Input #1, StrLine(LineSum)
        '---------------VBbasToN-----------
        If LineSum = 1 Then
            SNTemp = InStr(1, StrLine(LineSum), """") + 1
            VBbasFunctionName = Mid(StrLine(LineSum), SNTemp, InStrRev(StrLine(LineSum), """", Len(StrLine(LineSum))) - SNTemp)
            Me.Caption = Me.Caption & " - " & VBbasFunctionName
        Else
            Dim STemp
            STemp = Split(StrLine(LineSum), " ")
            If InStr(1, StrLine(LineSum), "Public Function ") = 1 _
            Or InStr(1, StrLine(LineSum), "Public Sub ") = 1 _
            Or InStr(1, StrLine(LineSum), "Private Function ") = 1 _
            Or InStr(1, StrLine(LineSum), "Private Sub ") = 1 Then
                fSUM = fSUM + 1: Package(fSUM).Start = LineSum
                Package(fSUM).Title = Mid(STemp(2), 1, InStr(1, STemp(2), "(") - 1)
            End If
            If InStr(1, StrLine(LineSum), "End Function") = 1 _
            Or InStr(1, StrLine(LineSum), "End Sub") = 1 Then
                Package(fSUM).End = LineSum
                ReDim Package(fSUM).Content(1 To LineSum - Package(fSUM).Start)
                j = 0
                For c = Package(fSUM).Start + 1 To LineSum
                    j = j + 1
                    Package(fSUM).Content(j) = StrLine(c)
                Next
            End If
        End If
        '----------------End---------------
    Loop
Close #1
ReDim Preserve 节点(fSUM)
For i = 1 To fSUM - 1
    节点(i).A = True
    节点(i).Title = Package(i).Title
    节点(i).X = Me.Width / fSUM * i
    节点(i).Y = Me.Height / fSUM * i
    ReDim 节点(i).Content(1 To UBound(Package(i).Content))
    For c = 1 To UBound(Package(i).Content)
        节点(i).Content(c) = Package(i).Content(c)
    Next
Next
LSum = 0
For i = 1 To fSUM - 1
        For c = 1 To i - 1
            For j = 1 To UBound(节点(c).Content)
                If InStr(1, 节点(c).Content(j), " " & 节点(i).Title & " ") <> 0 _
                Or InStr(1, 节点(c).Content(j), " " & 节点(i).Title & "(") <> 0 Then
                    LSum = LSum + 1
                    连接(LSum).A = True: 连接(LSum).Source = i: 连接(LSum).Target = c
                End If
            Next
        Next
        For c = i + 1 To fSUM - 1
            For j = 1 To UBound(节点(c).Content)
                If InStr(1, 节点(c).Content(j), " " & 节点(i).Title & " ") <> 0 _
                Or InStr(1, 节点(c).Content(j), " " & 节点(i).Title & "(") <> 0 Then
                    LSum = LSum + 1
                    连接(LSum).A = True: 连接(LSum).Source = i: 连接(LSum).Target = c
                End If
            Next
        Next
Next
NodeReDimLock = True
'----------------End---------------
Exit Sub
ErrHandler:
' 用户按了“取消”按钮
End Sub
Private Sub 退出_Click()
End
End Sub
