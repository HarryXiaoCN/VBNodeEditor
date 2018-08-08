VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Home 
   AutoRedraw      =   -1  'True
   Caption         =   "Home"
   ClientHeight    =   6960
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13815
   FillColor       =   &H00C0C000&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
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
Private MeID, MeName, LSum As Long
Private MouseDownPosition(2) As Single
Private NodeReDimLock, MouseDownLock  As Boolean
Private Sub Form_Activate()
Dim STemp '获得窗体的Z坐标
On Error GoTo Er
STemp = Split(Me.Caption, " - ")
MeID = Val(STemp(0))
'MeName = STemp(1)
Er:
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
If NodeReDimLock = False Then Exit Sub
For i = 1 To UBound(节点) - 1
    If 节点(i).a = True Then
        If X + 100 > 节点(i).X And X - 100 < 节点(i).X _
        And Y + 100 > 节点(i).Y And Y - 100 < 节点(i).Y Then
            MousePlace.Target = i: Exit Sub
        End If
    End If
Next
MousePlace.Target = 0: MouseDownPosition(0) = X: MouseDownPosition(1) = Y: MouseDownLock = True
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePlace.Target = 0
MouseDownLock = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With MousePlace
    .X = X
    .Y = Y
    .Z = MeID
End With
End Sub
Private Sub Form_Resize()
Me.Scale (-Me.Width / 2, Me.Height / 2)-(Me.Width / 2, -Me.Height / 2)
End Sub
Private Sub Timer1_Timer()
Dim i As Long
Me.Cls
If NodeReDimLock = False Then Exit Sub
For i = 1 To LSum
    If 连接(i).a = True Then
        Me.Line (节点(连接(i).Source).X, 节点(连接(i).Source).Y)-((节点(连接(i).Target).X + 节点(连接(i).Source).X) / 2, (节点(连接(i).Target).Y + 节点(连接(i).Source).Y) / 2), RGB(255, 0, 0)
        Me.Line ((节点(连接(i).Target).X + 节点(连接(i).Source).X) / 2, (节点(连接(i).Target).Y + 节点(连接(i).Source).Y) / 2)-(节点(连接(i).Target).X, 节点(连接(i).Target).Y), RGB(0, 0, 255)
    End If
Next
For i = 1 To UBound(节点) - 1
    If 节点(i).a = True Then
        If MousePlace.X + 120 > 节点(i).X And MousePlace.X - 120 < 节点(i).X _
        And MousePlace.Y + 120 > 节点(i).Y And MousePlace.Y - 120 < 节点(i).Y Then
            Me.FillColor = RGB(255, 0, 0)
            Me.Circle (节点(i).X, 节点(i).Y), 150, RGB(255, 0, 0)
            Me.FillColor = &HC0C000
            Me.CurrentX = -Me.Width / 2: Me.CurrentY = Me.Height / 2: Me.Print i; 节点(i).SourceSum; 节点(i).TargetSum
        Else
            Me.Circle (节点(i).X, 节点(i).Y), 100, RGB(0, 191, 255)
        End If
        Me.CurrentX = 节点(i).X
        Me.CurrentY = 节点(i).Y
        Me.Print 节点(i).Title
    End If
Next
If MousePlace.Target <> 0 And MousePlace.Z = MeID Then
    节点(MousePlace.Target).X = MousePlace.X
    节点(MousePlace.Target).Y = MousePlace.Y
End If
If MouseDownLock = True Then
    For i = 1 To UBound(节点) - 1
        If 节点(i).a = True Then
            节点(i).X = 节点(i).X + MousePlace.X - MouseDownPosition(0)
            节点(i).Y = 节点(i).Y + MousePlace.Y - MouseDownPosition(1)
        End If
    Next
    MouseDownPosition(0) = MousePlace.X
    MouseDownPosition(1) = MousePlace.Y
End If
End Sub
Private Sub 打开_Click()
Dim Crosswise, Lengthways As Long
CommonDialog1.CancelError = True
'On Error GoTo ErrHandler
' 设置标志
CommonDialog1.Flags = cdlOFNHideReadOnly
' 设置过滤器
CommonDialog1.Filter = "VBBas Files" & _
"(*.bas)|*.bas|All Files (*.*)|*.*"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen
VBbasName = CommonDialog1.FileTitle
VBbasPath = CommonDialog1.FileName
'---------------VBbasToN-----------
Dim StrLine() As String: Dim LineSum, fSUM, i, c, j, k As Long
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
            
            Me.Caption = MeID & " - " & VBbasName
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
    节点(i).a = True
    节点(i).Title = Package(i).Title
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
                    For k = 1 To LSum '过滤已经有连接的情况
                        If (连接(k).Source = i And 连接(LSum).Target = c) _
                        Or (连接(k).Source = c And 连接(LSum).Target = i) Then GoTo CF1
                    Next
                    LSum = LSum + 1
                    连接(LSum).a = True: 连接(LSum).Source = i: 连接(LSum).Target = c
CF1:
                End If
            Next
        Next
        For c = i + 1 To fSUM - 1
            For j = 1 To UBound(节点(c).Content)
                If InStr(1, 节点(c).Content(j), " " & 节点(i).Title & " ") <> 0 _
                Or InStr(1, 节点(c).Content(j), " " & 节点(i).Title & "(") <> 0 Then
                    For k = 1 To LSum
                        If (连接(k).Source = i And 连接(LSum).Target = c) _
                        Or (连接(k).Source = c And 连接(LSum).Target = i) Then GoTo CF2
                    Next
                    LSum = LSum + 1
                    连接(LSum).a = True: 连接(LSum).Source = i: 连接(LSum).Target = c
                End If
CF2:
            Next
        Next
Next
Dim Max(2, 2) As Long
For i = 1 To LSum '统计最多连接，好置于中心
    If 连接(i).a = True Then
        节点(连接(i).Source).SourceSum = 节点(连接(i).Source).SourceSum + 1
        节点(连接(i).Target).TargetSum = 节点(连接(i).Target).TargetSum + 1
        If 节点(连接(i).Source).SourceSum > Max(0, 0) Then Max(0, 0) = 节点(连接(i).Source).SourceSum: Max(0, 1) = 连接(i).Source
        If 节点(连接(i).Target).TargetSum > Max(1, 0) Then Max(1, 0) = 节点(连接(i).Target).TargetSum: Max(1, 1) = 连接(i).Target
    End If
Next
If Max(0, 0) > 0 Or Max(1, 0) > 0 Then '检查有没有连接好看画不画圆
    Dim Angle, CX As Single: Dim TarSumTemp As Long: Dim VisLock() As Boolean
    ReDim VisLock(1 To fSUM)
    If Max(0, 0) > Max(1, 0) Then '检查哪种画比较好看
        节点(Max(0, 1)).X = 0: 节点(Max(0, 1)).Y = 0
        Angle = 2 * Pi / Max(0, 0)
        VisLock(Max(0, 1)) = True
        For i = 1 To LSum
            If 连接(i).a = True And 连接(i).Source = Max(0, 1) Then
                CX = Angle * TarSumTemp
                 VisLock(连接(i).Target) = True
                If CX > Pi / 2 Then
                    CX = Pi - CX
                    节点(连接(i).Target).X = 2000 * -Cos(CX)
                Else
                    节点(连接(i).Target).X = 2000 * Cos(CX)
                End If
                    节点(连接(i).Target).Y = 2000 * Sin(CX)
                TarSumTemp = TarSumTemp + 1
            End If
        Next
        Angle = 2 * Pi / (fSUM - 2 - Max(0, 0))
    Else '检查哪种画比较好看的转折
        节点(Max(1, 1)).X = 0: 节点(Max(1, 1)).Y = 0
        Angle = 2 * Pi / Max(1, 0)
        VisLock(Max(1, 1)) = True:
        For i = 1 To LSum
            If 连接(i).a = True And 连接(i).Target = Max(1, 1) Then
                CX = Angle * TarSumTemp
                VisLock(连接(i).Source) = True
                If CX > Pi / 2 Then
                    CX = Pi - CX
                    节点(连接(i).Source).X = 2000 * -Cos(CX)
                Else
                    节点(连接(i).Source).X = 2000 * Cos(CX)
                End If
                    节点(连接(i).Source).Y = 2000 * Sin(CX)
                TarSumTemp = TarSumTemp + 1
            End If
        Next
        Angle = 2 * Pi / (fSUM - 2 - Max(1, 0))
    End If
    '这里到下一个注释是把最多连接不直接相关的节点暂时安置一下
    TarSumTemp = 0
    For i = 1 To fSUM - 1
        If VisLock(i) = False Then
            CX = Angle * TarSumTemp
            If CX > Pi / 2 Then
                CX = Pi - CX
                节点(i).X = 5000 * -Cos(CX)
            Else
                节点(i).X = 5000 * Cos(CX)
            End If
                节点(i).Y = 5000 * Sin(CX)
            TarSumTemp = TarSumTemp + 1
        End If
    Next
Else '检查有没有连接好看画不画圆转折
    '下面到最外围的End If是没有互相连接的函数阵列安置
    Lengthways = 1
    For i = 1 To fSUM - 1
        Crosswise = Crosswise + 1
        If Crosswise > Sqr(fSUM) Then
            Crosswise = 1: Lengthways = Lengthways + 1
        End If
         If Lengthways Mod 2 = 0 Then
            节点(i).X = -Me.Width / (Int(Sqr(fSUM)) + 1) * (Crosswise + 0.2) + Me.Width / 2
        Else
            节点(i).X = -Me.Width / (Int(Sqr(fSUM)) + 1) * (Crosswise - 0.2) + Me.Width / 2
        End If
        If Crosswise Mod 2 = 0 Then
            节点(i).Y = -Me.Height / (Int(Sqr(fSUM)) + 1) * (Lengthways + 0.2) + Me.Height / 2
        Else
            节点(i).Y = -Me.Height / (Int(Sqr(fSUM)) + 1) * (Lengthways - 0.2) + Me.Height / 2
        End If
    Next
End If
NodeReDimLock = True
'----------------End---------------
Exit Sub
ErrHandler:
' 用户按了“取消”按钮
End Sub
Private Sub 退出_Click()
End
End Sub
