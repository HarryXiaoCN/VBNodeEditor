VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Home 
   AutoRedraw      =   -1  'True
   Caption         =   "Home"
   ClientHeight    =   6960
   ClientLeft      =   225
   ClientTop       =   855
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
         Shortcut        =   ^N
      End
      Begin VB.Menu 打开 
         Caption         =   "打开"
         Shortcut        =   ^O
      End
      Begin VB.Menu 保存 
         Caption         =   "保存"
         Shortcut        =   ^S
      End
      Begin VB.Menu 另存为 
         Caption         =   "另存为"
      End
      Begin VB.Menu 退出 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu 视图 
      Caption         =   "视图"
      Begin VB.Menu 节点布局 
         Caption         =   "节点布局"
         Begin VB.Menu Vis 
            Caption         =   "自动布局"
            Index           =   0
         End
         Begin VB.Menu Vis 
            Caption         =   "被引用中心"
            Index           =   1
         End
         Begin VB.Menu Vis 
            Caption         =   "调用中心"
            Index           =   2
         End
         Begin VB.Menu Vis 
            Caption         =   "阵列布局"
            Index           =   3
         End
      End
      Begin VB.Menu 显示 
         Caption         =   "显示"
         Begin VB.Menu 显示节点标题 
            Caption         =   "显示节点标题"
            Shortcut        =   ^T
         End
         Begin VB.Menu 显示节点连接 
            Caption         =   "显示节点连接"
            Shortcut        =   ^L
         End
         Begin VB.Menu 隐藏时指针显示 
            Caption         =   "隐藏时指针显示"
            Begin VB.Menu 引用次数指向 
               Caption         =   "引用次数指向"
               Shortcut        =   ^Y
            End
            Begin VB.Menu 调用次数指向 
               Caption         =   "调用次数指向"
               Shortcut        =   ^D
            End
         End
      End
   End
   Begin VB.Menu 功能 
      Caption         =   "功能"
      Begin VB.Menu 节点搜索 
         Caption         =   "节点搜索"
         Shortcut        =   ^F
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
Private NodeReDimLock, MouseDownLock, FindLock As Boolean
Private Sub Form_Activate()
Dim STemp '获得窗体的Z坐标
On Error GoTo Er
STemp = Split(Me.Caption, " - ")
MeID = Val(STemp(0))
'MeName = STemp(1)
Er:
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long
Select Case KeyCode
    Case 27 'ESC
        If FindLock = True Then
            For i = 1 To UBound(节点)
                If 节点(i).Color = 2 Then 节点(i).Color = 0
            Next
            FindLock = False
        End If
End Select
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
If NodeReDimLock = False Then Exit Sub
For i = 1 To UBound(节点) - 1
    If 节点(i).a = True Then
        If X + 100 > 节点(i).X And X - 100 < 节点(i).X _
        And Y + 100 > 节点(i).Y And Y - 100 < 节点(i).Y Then
            MousePlace.Target = i: If 节点(i).Color = 0 Then 节点(i).Color = 1
            If 节点(i).Color = 2 Then 节点(i).Color = 3
            Exit Sub
        End If
    End If
Next
MousePlace.Target = 0: MouseDownPosition(0) = X: MouseDownPosition(1) = Y: MouseDownLock = True
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If NodeReDimLock = False Then Exit Sub
If 节点(MousePlace.Target).Color = 1 Then 节点(MousePlace.Target).Color = 0
If 节点(MousePlace.Target).Color = 3 Then 节点(MousePlace.Target).Color = 2
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

Private Sub Form_Unload(Cancel As Integer)
If MeID = 0 Then End
End Sub

Private Sub Timer1_Timer()
Dim i, c As Long
Me.Cls
If NodeReDimLock = False Then Exit Sub
For i = 1 To LSum
    If 连接(i).a = True Then
        If 显示节点连接.Checked = True Then
            Me.Line (节点(连接(i).Source).X, 节点(连接(i).Source).Y)-((节点(连接(i).Target).X + 节点(连接(i).Source).X) / 2, (节点(连接(i).Target).Y + 节点(连接(i).Source).Y) / 2), RGB(255, 0, 0)
            Me.Line ((节点(连接(i).Target).X + 节点(连接(i).Source).X) / 2, (节点(连接(i).Target).Y + 节点(连接(i).Source).Y) / 2)-(节点(连接(i).Target).X, 节点(连接(i).Target).Y), RGB(0, 0, 255)
        Else
            If 调用次数指向.Checked = True And 连接(i).Source = MousePlace.Aim Then
                Me.Line (节点(连接(i).Source).X, 节点(连接(i).Source).Y)-((节点(连接(i).Target).X + 节点(连接(i).Source).X) / 2, (节点(连接(i).Target).Y + 节点(连接(i).Source).Y) / 2), RGB(255, 0, 0)
                Me.Line ((节点(连接(i).Target).X + 节点(连接(i).Source).X) / 2, (节点(连接(i).Target).Y + 节点(连接(i).Source).Y) / 2)-(节点(连接(i).Target).X, 节点(连接(i).Target).Y), RGB(0, 0, 255)
            End If
            If 引用次数指向.Checked = True And 连接(i).Target = MousePlace.Aim Then
                Me.Line (节点(连接(i).Source).X, 节点(连接(i).Source).Y)-((节点(连接(i).Target).X + 节点(连接(i).Source).X) / 2, (节点(连接(i).Target).Y + 节点(连接(i).Source).Y) / 2), RGB(255, 0, 0)
                Me.Line ((节点(连接(i).Target).X + 节点(连接(i).Source).X) / 2, (节点(连接(i).Target).Y + 节点(连接(i).Source).Y) / 2)-(节点(连接(i).Target).X, 节点(连接(i).Target).Y), RGB(0, 0, 255)
            End If
        End If
    End If
Next
For i = 1 To UBound(节点) - 1
    If 节点(i).a = True Then
        Select Case 节点(i).Color
            Case 0
                Me.Circle (节点(i).X, 节点(i).Y), 100, RGB(0, 191, 255)
            Case 1
                Me.FillColor = RGB(255, 0, 0)
                Me.Circle (节点(i).X, 节点(i).Y), 150, RGB(255, 0, 0)
                Me.FillColor = &HC0C000
            Case 2
                Me.FillColor = RGB(34, 139, 34)
                Me.Circle (节点(i).X, 节点(i).Y), 120, RGB(34, 139, 34)
                Me.FillColor = &HC0C000
                If 显示节点标题.Checked = False Then
                    Me.CurrentX = 节点(i).X
                    Me.CurrentY = 节点(i).Y
                    Me.Print 节点(i).Title
                End If
            Case 3
                Me.FillColor = RGB(128, 0, 128)
                Me.Circle (节点(i).X, 节点(i).Y), 150, RGB(128, 0, 128)
                Me.FillColor = &HC0C000
        End Select
        If MousePlace.X + 120 > 节点(i).X And MousePlace.X - 120 < 节点(i).X _
        And MousePlace.Y + 120 > 节点(i).Y And MousePlace.Y - 120 < 节点(i).Y Then
            If 节点(i).Color = 0 Then 节点(i).Color = 1: MousePlace.Aim = i
            If 节点(i).Color = 2 Then 节点(i).Color = 3
            If 显示节点标题.Checked = False Then
                Me.CurrentX = 节点(i).X
                Me.CurrentY = 节点(i).Y
                Me.Print 节点(i).Title
            End If
            Me.CurrentX = -Me.Width / 2: Me.CurrentY = Me.Height / 2: Me.Print i; 节点(i).SourceSum; 节点(i).TargetSum
        Else
            If MousePlace.Target <> i And 节点(i).Color = 1 Then 节点(i).Color = 0
            If MousePlace.Target <> i And 节点(i).Color = 3 Then 节点(i).Color = 2
        End If
        If 显示节点标题.Checked = True Then
            Me.CurrentX = 节点(i).X
            Me.CurrentY = 节点(i).Y
            Me.Print 节点(i).Title
        End If
    End If
Next
If 调用次数指向.Checked = True Or 引用次数指向.Checked = True Then
    For i = 1 To LSum
        If 连接(i).a = True Then
            If 调用次数指向.Checked = True And 连接(i).Source = MousePlace.Aim Then
                Me.CurrentX = 节点(连接(i).Target).X
                Me.CurrentY = 节点(连接(i).Target).Y
                Me.Print 节点(连接(i).Target).Title
            End If
            If 引用次数指向.Checked = True And 连接(i).Target = MousePlace.Aim Then
                Me.CurrentX = 节点(连接(i).Source).X
                Me.CurrentY = 节点(连接(i).Source).Y
                Me.Print 节点(连接(i).Source).Title
            End If
        End If
    Next
End If
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
Private Sub Vis_Click(Index As Integer)
Dim i As Long
If Vis(Index).Checked = False Then
    For i = 0 To 3
        Vis(i).Checked = False
    Next
    Vis(Index).Checked = True
Else
    For i = 0 To 3
        Vis(i).Checked = False
    Next
    Vis(0).Checked = True
End If
End Sub

Private Sub 打开_Click()
Dim Crosswise, Lengthways As Long
新建_Click
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
Dim StrLine() As String: Dim LineSum, fSUM, i, c, j, k, Max(2, 2) As Long
Dim Angle, CX As Single: Dim TarSumTemp As Long: Dim VisLock() As Boolean
Dim SNTemp, 圆阵() As Long
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
If Vis(0).Checked = True Then
    For i = 1 To LSum '统计最多连接，好置于中心
        If 连接(i).a = True Then
            节点(连接(i).Source).SourceSum = 节点(连接(i).Source).SourceSum + 1
            节点(连接(i).Target).TargetSum = 节点(连接(i).Target).TargetSum + 1
            If 节点(连接(i).Source).SourceSum > Max(0, 0) Then Max(0, 0) = 节点(连接(i).Source).SourceSum: Max(0, 1) = 连接(i).Source
            If 节点(连接(i).Target).TargetSum > Max(1, 0) Then Max(1, 0) = 节点(连接(i).Target).TargetSum: Max(1, 1) = 连接(i).Target
        End If
    Next
    If Max(0, 0) > 0 Or Max(1, 0) > 0 Then '检查有没有连接好看画不画圆
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
End If
If Vis(1).Checked = True Then
    For i = 1 To LSum '统计最多连接，好置于中心
        If 连接(i).a = True Then
            节点(连接(i).Source).SourceSum = 节点(连接(i).Source).SourceSum + 1
            节点(连接(i).Target).TargetSum = 节点(连接(i).Target).TargetSum + 1
            If 节点(连接(i).Source).SourceSum > Max(0, 0) Then Max(0, 0) = 节点(连接(i).Source).SourceSum: Max(0, 1) = 连接(i).Source
            If 节点(连接(i).Target).TargetSum > Max(1, 0) Then Max(1, 0) = 节点(连接(i).Target).TargetSum: Max(1, 1) = 连接(i).Target
        End If
    Next
    If Max(0, 0) > 0 Then '检查有没有连接好看画不画圆
        ReDim VisLock(1 To fSUM)
        节点(Max(0, 1)).X = 0: 节点(Max(0, 1)).Y = 0
        Angle = 2 * Pi / Max(0, 0)
        VisLock(Max(0, 1)) = True:
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
        GoTo 阵列
    End If
End If
If Vis(2).Checked = True Then
    For i = 1 To LSum '统计最多连接，好置于中心
        If 连接(i).a = True Then
            节点(连接(i).Source).SourceSum = 节点(连接(i).Source).SourceSum + 1
            节点(连接(i).Target).TargetSum = 节点(连接(i).Target).TargetSum + 1
            If 节点(连接(i).Source).SourceSum > Max(0, 0) Then Max(0, 0) = 节点(连接(i).Source).SourceSum: Max(0, 1) = 连接(i).Source
            If 节点(连接(i).Target).TargetSum > Max(1, 0) Then Max(1, 0) = 节点(连接(i).Target).TargetSum: Max(1, 1) = 连接(i).Target
        End If
    Next
    If Max(1, 0) > 0 Then '检查有没有连接好看画不画圆
        ReDim VisLock(1 To fSUM)
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
        GoTo 阵列
    End If
End If
If Vis(3).Checked = True Then
阵列:
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

Private Sub 节点搜索_Click()
Dim FindStr As String: Dim i As Long
On Error GoTo Er
FindLock = True
FindStr = InputBox("请输入您要搜索的内容")
For i = 1 To UBound(节点) - 1
    If 节点(i).a = True And InStr(1, 节点(i).Title, FindStr) <> 0 Then
        节点(i).Color = 2
    End If
Next
Er:
End Sub

Private Sub 调用次数指向_Click()
If 调用次数指向.Checked = True Then 调用次数指向.Checked = False Else 调用次数指向.Checked = True
End Sub

Private Sub 退出_Click()
End
End Sub

Private Sub 显示节点标题_Click()
If 显示节点标题.Checked = True Then
    隐藏时指针显示.Enabled = True: 显示节点标题.Checked = False
Else
    If 显示节点标题.Checked = True Then 隐藏时指针显示.Enabled = False
    显示节点标题.Checked = True
End If
End Sub

Private Sub 显示节点连接_Click()
If 显示节点连接.Checked = True Then
    隐藏时指针显示.Enabled = True: 显示节点连接.Checked = False
Else
    If 显示节点标题.Checked = True Then 隐藏时指针显示.Enabled = False
    显示节点连接.Checked = True
End If
End Sub

Private Sub 新建_Click()
NodeReDimLock = False
Erase 节点
Erase 连接
End Sub

Private Sub 引用次数指向_Click()
If 引用次数指向.Checked = True Then 引用次数指向.Checked = False Else 引用次数指向.Checked = True
End Sub
