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
      Name            =   "΢���ź�"
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
   StartUpPosition =   3  '����ȱʡ
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
   Begin VB.Menu �˵� 
      Caption         =   "�˵�"
      Begin VB.Menu �½� 
         Caption         =   "�½�"
         Shortcut        =   ^N
      End
      Begin VB.Menu �� 
         Caption         =   "��"
         Shortcut        =   ^O
      End
      Begin VB.Menu ���� 
         Caption         =   "����"
         Shortcut        =   ^S
      End
      Begin VB.Menu ���Ϊ 
         Caption         =   "���Ϊ"
      End
      Begin VB.Menu �˳� 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu ��ͼ 
      Caption         =   "��ͼ"
      Begin VB.Menu �ڵ㲼�� 
         Caption         =   "�ڵ㲼��"
         Begin VB.Menu Vis 
            Caption         =   "�Զ�����"
            Index           =   0
         End
         Begin VB.Menu Vis 
            Caption         =   "����������"
            Index           =   1
         End
         Begin VB.Menu Vis 
            Caption         =   "��������"
            Index           =   2
         End
         Begin VB.Menu Vis 
            Caption         =   "���в���"
            Index           =   3
         End
      End
      Begin VB.Menu ��ʾ 
         Caption         =   "��ʾ"
         Begin VB.Menu ��ʾ�ڵ���� 
            Caption         =   "��ʾ�ڵ����"
            Shortcut        =   ^T
         End
         Begin VB.Menu ��ʾ�ڵ����� 
            Caption         =   "��ʾ�ڵ�����"
            Shortcut        =   ^L
         End
         Begin VB.Menu ����ʱָ����ʾ 
            Caption         =   "����ʱָ����ʾ"
            Begin VB.Menu ���ô���ָ�� 
               Caption         =   "���ô���ָ��"
               Shortcut        =   ^Y
            End
            Begin VB.Menu ���ô���ָ�� 
               Caption         =   "���ô���ָ��"
               Shortcut        =   ^D
            End
         End
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu �ڵ����� 
         Caption         =   "�ڵ�����"
         Shortcut        =   ^F
      End
   End
End
Attribute VB_Name = "Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private �ڵ�() As ��Ԫ
Private ����(1 To 999) As ��ϵ
Private MeID, MeName, LSum As Long
Private MouseDownPosition(2) As Single
Private NodeReDimLock, MouseDownLock, FindLock As Boolean
Private Sub Form_Activate()
Dim STemp '��ô����Z����
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
            For i = 1 To UBound(�ڵ�)
                If �ڵ�(i).Color = 2 Then �ڵ�(i).Color = 0
            Next
            FindLock = False
        End If
End Select
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
If NodeReDimLock = False Then Exit Sub
For i = 1 To UBound(�ڵ�) - 1
    If �ڵ�(i).a = True Then
        If X + 100 > �ڵ�(i).X And X - 100 < �ڵ�(i).X _
        And Y + 100 > �ڵ�(i).Y And Y - 100 < �ڵ�(i).Y Then
            MousePlace.Target = i: If �ڵ�(i).Color = 0 Then �ڵ�(i).Color = 1
            If �ڵ�(i).Color = 2 Then �ڵ�(i).Color = 3
            Exit Sub
        End If
    End If
Next
MousePlace.Target = 0: MouseDownPosition(0) = X: MouseDownPosition(1) = Y: MouseDownLock = True
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If NodeReDimLock = False Then Exit Sub
If �ڵ�(MousePlace.Target).Color = 1 Then �ڵ�(MousePlace.Target).Color = 0
If �ڵ�(MousePlace.Target).Color = 3 Then �ڵ�(MousePlace.Target).Color = 2
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
    If ����(i).a = True Then
        If ��ʾ�ڵ�����.Checked = True Then
            Me.Line (�ڵ�(����(i).Source).X, �ڵ�(����(i).Source).Y)-((�ڵ�(����(i).Target).X + �ڵ�(����(i).Source).X) / 2, (�ڵ�(����(i).Target).Y + �ڵ�(����(i).Source).Y) / 2), RGB(255, 0, 0)
            Me.Line ((�ڵ�(����(i).Target).X + �ڵ�(����(i).Source).X) / 2, (�ڵ�(����(i).Target).Y + �ڵ�(����(i).Source).Y) / 2)-(�ڵ�(����(i).Target).X, �ڵ�(����(i).Target).Y), RGB(0, 0, 255)
        Else
            If ���ô���ָ��.Checked = True And ����(i).Source = MousePlace.Aim Then
                Me.Line (�ڵ�(����(i).Source).X, �ڵ�(����(i).Source).Y)-((�ڵ�(����(i).Target).X + �ڵ�(����(i).Source).X) / 2, (�ڵ�(����(i).Target).Y + �ڵ�(����(i).Source).Y) / 2), RGB(255, 0, 0)
                Me.Line ((�ڵ�(����(i).Target).X + �ڵ�(����(i).Source).X) / 2, (�ڵ�(����(i).Target).Y + �ڵ�(����(i).Source).Y) / 2)-(�ڵ�(����(i).Target).X, �ڵ�(����(i).Target).Y), RGB(0, 0, 255)
            End If
            If ���ô���ָ��.Checked = True And ����(i).Target = MousePlace.Aim Then
                Me.Line (�ڵ�(����(i).Source).X, �ڵ�(����(i).Source).Y)-((�ڵ�(����(i).Target).X + �ڵ�(����(i).Source).X) / 2, (�ڵ�(����(i).Target).Y + �ڵ�(����(i).Source).Y) / 2), RGB(255, 0, 0)
                Me.Line ((�ڵ�(����(i).Target).X + �ڵ�(����(i).Source).X) / 2, (�ڵ�(����(i).Target).Y + �ڵ�(����(i).Source).Y) / 2)-(�ڵ�(����(i).Target).X, �ڵ�(����(i).Target).Y), RGB(0, 0, 255)
            End If
        End If
    End If
Next
For i = 1 To UBound(�ڵ�) - 1
    If �ڵ�(i).a = True Then
        Select Case �ڵ�(i).Color
            Case 0
                Me.Circle (�ڵ�(i).X, �ڵ�(i).Y), 100, RGB(0, 191, 255)
            Case 1
                Me.FillColor = RGB(255, 0, 0)
                Me.Circle (�ڵ�(i).X, �ڵ�(i).Y), 150, RGB(255, 0, 0)
                Me.FillColor = &HC0C000
            Case 2
                Me.FillColor = RGB(34, 139, 34)
                Me.Circle (�ڵ�(i).X, �ڵ�(i).Y), 120, RGB(34, 139, 34)
                Me.FillColor = &HC0C000
                If ��ʾ�ڵ����.Checked = False Then
                    Me.CurrentX = �ڵ�(i).X
                    Me.CurrentY = �ڵ�(i).Y
                    Me.Print �ڵ�(i).Title
                End If
            Case 3
                Me.FillColor = RGB(128, 0, 128)
                Me.Circle (�ڵ�(i).X, �ڵ�(i).Y), 150, RGB(128, 0, 128)
                Me.FillColor = &HC0C000
        End Select
        If MousePlace.X + 120 > �ڵ�(i).X And MousePlace.X - 120 < �ڵ�(i).X _
        And MousePlace.Y + 120 > �ڵ�(i).Y And MousePlace.Y - 120 < �ڵ�(i).Y Then
            If �ڵ�(i).Color = 0 Then �ڵ�(i).Color = 1: MousePlace.Aim = i
            If �ڵ�(i).Color = 2 Then �ڵ�(i).Color = 3
            If ��ʾ�ڵ����.Checked = False Then
                Me.CurrentX = �ڵ�(i).X
                Me.CurrentY = �ڵ�(i).Y
                Me.Print �ڵ�(i).Title
            End If
            Me.CurrentX = -Me.Width / 2: Me.CurrentY = Me.Height / 2: Me.Print i; �ڵ�(i).SourceSum; �ڵ�(i).TargetSum
        Else
            If MousePlace.Target <> i And �ڵ�(i).Color = 1 Then �ڵ�(i).Color = 0
            If MousePlace.Target <> i And �ڵ�(i).Color = 3 Then �ڵ�(i).Color = 2
        End If
        If ��ʾ�ڵ����.Checked = True Then
            Me.CurrentX = �ڵ�(i).X
            Me.CurrentY = �ڵ�(i).Y
            Me.Print �ڵ�(i).Title
        End If
    End If
Next
If ���ô���ָ��.Checked = True Or ���ô���ָ��.Checked = True Then
    For i = 1 To LSum
        If ����(i).a = True Then
            If ���ô���ָ��.Checked = True And ����(i).Source = MousePlace.Aim Then
                Me.CurrentX = �ڵ�(����(i).Target).X
                Me.CurrentY = �ڵ�(����(i).Target).Y
                Me.Print �ڵ�(����(i).Target).Title
            End If
            If ���ô���ָ��.Checked = True And ����(i).Target = MousePlace.Aim Then
                Me.CurrentX = �ڵ�(����(i).Source).X
                Me.CurrentY = �ڵ�(����(i).Source).Y
                Me.Print �ڵ�(����(i).Source).Title
            End If
        End If
    Next
End If
If MousePlace.Target <> 0 And MousePlace.Z = MeID Then
    �ڵ�(MousePlace.Target).X = MousePlace.X
    �ڵ�(MousePlace.Target).Y = MousePlace.Y
End If
If MouseDownLock = True Then
    For i = 1 To UBound(�ڵ�) - 1
        If �ڵ�(i).a = True Then
            �ڵ�(i).X = �ڵ�(i).X + MousePlace.X - MouseDownPosition(0)
            �ڵ�(i).Y = �ڵ�(i).Y + MousePlace.Y - MouseDownPosition(1)
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

Private Sub ��_Click()
Dim Crosswise, Lengthways As Long
�½�_Click
CommonDialog1.CancelError = True
'On Error GoTo ErrHandler
' ���ñ�־
CommonDialog1.Flags = cdlOFNHideReadOnly
' ���ù�����
CommonDialog1.Filter = "VBBas Files" & _
"(*.bas)|*.bas|All Files (*.*)|*.*"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen
VBbasName = CommonDialog1.FileTitle
VBbasPath = CommonDialog1.FileName
'---------------VBbasToN-----------
Dim StrLine() As String: Dim LineSum, fSUM, i, c, j, k, Max(2, 2) As Long
Dim Angle, CX As Single: Dim TarSumTemp As Long: Dim VisLock() As Boolean
Dim SNTemp, Բ��() As Long
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
ReDim Preserve �ڵ�(fSUM)
For i = 1 To fSUM - 1
    �ڵ�(i).a = True
    �ڵ�(i).Title = Package(i).Title
    ReDim �ڵ�(i).Content(1 To UBound(Package(i).Content))
    For c = 1 To UBound(Package(i).Content)
        �ڵ�(i).Content(c) = Package(i).Content(c)
    Next
Next
LSum = 0
For i = 1 To fSUM - 1
        For c = 1 To i - 1
            For j = 1 To UBound(�ڵ�(c).Content)
                If InStr(1, �ڵ�(c).Content(j), " " & �ڵ�(i).Title & " ") <> 0 _
                Or InStr(1, �ڵ�(c).Content(j), " " & �ڵ�(i).Title & "(") <> 0 Then
                    For k = 1 To LSum '�����Ѿ������ӵ����
                        If (����(k).Source = i And ����(LSum).Target = c) _
                        Or (����(k).Source = c And ����(LSum).Target = i) Then GoTo CF1
                    Next
                    LSum = LSum + 1
                    ����(LSum).a = True: ����(LSum).Source = i: ����(LSum).Target = c
CF1:
                End If
            Next
        Next
        For c = i + 1 To fSUM - 1
            For j = 1 To UBound(�ڵ�(c).Content)
                If InStr(1, �ڵ�(c).Content(j), " " & �ڵ�(i).Title & " ") <> 0 _
                Or InStr(1, �ڵ�(c).Content(j), " " & �ڵ�(i).Title & "(") <> 0 Then
                    For k = 1 To LSum
                        If (����(k).Source = i And ����(LSum).Target = c) _
                        Or (����(k).Source = c And ����(LSum).Target = i) Then GoTo CF2
                    Next
                    LSum = LSum + 1
                    ����(LSum).a = True: ����(LSum).Source = i: ����(LSum).Target = c
                End If
CF2:
            Next
        Next
Next
If Vis(0).Checked = True Then
    For i = 1 To LSum 'ͳ��������ӣ�����������
        If ����(i).a = True Then
            �ڵ�(����(i).Source).SourceSum = �ڵ�(����(i).Source).SourceSum + 1
            �ڵ�(����(i).Target).TargetSum = �ڵ�(����(i).Target).TargetSum + 1
            If �ڵ�(����(i).Source).SourceSum > Max(0, 0) Then Max(0, 0) = �ڵ�(����(i).Source).SourceSum: Max(0, 1) = ����(i).Source
            If �ڵ�(����(i).Target).TargetSum > Max(1, 0) Then Max(1, 0) = �ڵ�(����(i).Target).TargetSum: Max(1, 1) = ����(i).Target
        End If
    Next
    If Max(0, 0) > 0 Or Max(1, 0) > 0 Then '�����û�����Ӻÿ�������Բ
        ReDim VisLock(1 To fSUM)
        If Max(0, 0) > Max(1, 0) Then '������ֻ��ȽϺÿ�
            �ڵ�(Max(0, 1)).X = 0: �ڵ�(Max(0, 1)).Y = 0
            Angle = 2 * Pi / Max(0, 0)
            VisLock(Max(0, 1)) = True
            For i = 1 To LSum
                If ����(i).a = True And ����(i).Source = Max(0, 1) Then
                    CX = Angle * TarSumTemp
                     VisLock(����(i).Target) = True
                    If CX > Pi / 2 Then
                        CX = Pi - CX
                        �ڵ�(����(i).Target).X = 2000 * -Cos(CX)
                    Else
                        �ڵ�(����(i).Target).X = 2000 * Cos(CX)
                    End If
                        �ڵ�(����(i).Target).Y = 2000 * Sin(CX)
                    TarSumTemp = TarSumTemp + 1
                End If
            Next
            Angle = 2 * Pi / (fSUM - 2 - Max(0, 0))
        Else '������ֻ��ȽϺÿ���ת��
            �ڵ�(Max(1, 1)).X = 0: �ڵ�(Max(1, 1)).Y = 0
            Angle = 2 * Pi / Max(1, 0)
            VisLock(Max(1, 1)) = True:
            For i = 1 To LSum
                If ����(i).a = True And ����(i).Target = Max(1, 1) Then
                    CX = Angle * TarSumTemp
                    VisLock(����(i).Source) = True
                    If CX > Pi / 2 Then
                        CX = Pi - CX
                        �ڵ�(����(i).Source).X = 2000 * -Cos(CX)
                    Else
                        �ڵ�(����(i).Source).X = 2000 * Cos(CX)
                    End If
                        �ڵ�(����(i).Source).Y = 2000 * Sin(CX)
                    TarSumTemp = TarSumTemp + 1
                End If
            Next
            Angle = 2 * Pi / (fSUM - 2 - Max(1, 0))
        End If
        '���ﵽ��һ��ע���ǰ�������Ӳ�ֱ����صĽڵ���ʱ����һ��
        TarSumTemp = 0
        For i = 1 To fSUM - 1
            If VisLock(i) = False Then
                CX = Angle * TarSumTemp
                If CX > Pi / 2 Then
                    CX = Pi - CX
                    �ڵ�(i).X = 5000 * -Cos(CX)
                Else
                    �ڵ�(i).X = 5000 * Cos(CX)
                End If
                    �ڵ�(i).Y = 5000 * Sin(CX)
                TarSumTemp = TarSumTemp + 1
            End If
        Next
    Else '�����û�����Ӻÿ�������Բת��
        '���浽����Χ��End If��û�л������ӵĺ������а���
        Lengthways = 1
        For i = 1 To fSUM - 1
            Crosswise = Crosswise + 1
            If Crosswise > Sqr(fSUM) Then
                Crosswise = 1: Lengthways = Lengthways + 1
            End If
             If Lengthways Mod 2 = 0 Then
                �ڵ�(i).X = -Me.Width / (Int(Sqr(fSUM)) + 1) * (Crosswise + 0.2) + Me.Width / 2
            Else
                �ڵ�(i).X = -Me.Width / (Int(Sqr(fSUM)) + 1) * (Crosswise - 0.2) + Me.Width / 2
            End If
            If Crosswise Mod 2 = 0 Then
                �ڵ�(i).Y = -Me.Height / (Int(Sqr(fSUM)) + 1) * (Lengthways + 0.2) + Me.Height / 2
            Else
                �ڵ�(i).Y = -Me.Height / (Int(Sqr(fSUM)) + 1) * (Lengthways - 0.2) + Me.Height / 2
            End If
        Next
    End If
End If
If Vis(1).Checked = True Then
    For i = 1 To LSum 'ͳ��������ӣ�����������
        If ����(i).a = True Then
            �ڵ�(����(i).Source).SourceSum = �ڵ�(����(i).Source).SourceSum + 1
            �ڵ�(����(i).Target).TargetSum = �ڵ�(����(i).Target).TargetSum + 1
            If �ڵ�(����(i).Source).SourceSum > Max(0, 0) Then Max(0, 0) = �ڵ�(����(i).Source).SourceSum: Max(0, 1) = ����(i).Source
            If �ڵ�(����(i).Target).TargetSum > Max(1, 0) Then Max(1, 0) = �ڵ�(����(i).Target).TargetSum: Max(1, 1) = ����(i).Target
        End If
    Next
    If Max(0, 0) > 0 Then '�����û�����Ӻÿ�������Բ
        ReDim VisLock(1 To fSUM)
        �ڵ�(Max(0, 1)).X = 0: �ڵ�(Max(0, 1)).Y = 0
        Angle = 2 * Pi / Max(0, 0)
        VisLock(Max(0, 1)) = True:
        For i = 1 To LSum
            If ����(i).a = True And ����(i).Source = Max(0, 1) Then
                CX = Angle * TarSumTemp
                VisLock(����(i).Target) = True
                If CX > Pi / 2 Then
                    CX = Pi - CX
                    �ڵ�(����(i).Target).X = 2000 * -Cos(CX)
                Else
                    �ڵ�(����(i).Target).X = 2000 * Cos(CX)
                End If
                    �ڵ�(����(i).Target).Y = 2000 * Sin(CX)
                TarSumTemp = TarSumTemp + 1
            End If
        Next
        Angle = 2 * Pi / (fSUM - 2 - Max(0, 0))
        '���ﵽ��һ��ע���ǰ�������Ӳ�ֱ����صĽڵ���ʱ����һ��
        TarSumTemp = 0
        For i = 1 To fSUM - 1
            If VisLock(i) = False Then
                CX = Angle * TarSumTemp
                If CX > Pi / 2 Then
                    CX = Pi - CX
                    �ڵ�(i).X = 5000 * -Cos(CX)
                Else
                    �ڵ�(i).X = 5000 * Cos(CX)
                End If
                    �ڵ�(i).Y = 5000 * Sin(CX)
                TarSumTemp = TarSumTemp + 1
            End If
        Next
    Else '�����û�����Ӻÿ�������Բת��
        '���浽����Χ��End If��û�л������ӵĺ������а���
        GoTo ����
    End If
End If
If Vis(2).Checked = True Then
    For i = 1 To LSum 'ͳ��������ӣ�����������
        If ����(i).a = True Then
            �ڵ�(����(i).Source).SourceSum = �ڵ�(����(i).Source).SourceSum + 1
            �ڵ�(����(i).Target).TargetSum = �ڵ�(����(i).Target).TargetSum + 1
            If �ڵ�(����(i).Source).SourceSum > Max(0, 0) Then Max(0, 0) = �ڵ�(����(i).Source).SourceSum: Max(0, 1) = ����(i).Source
            If �ڵ�(����(i).Target).TargetSum > Max(1, 0) Then Max(1, 0) = �ڵ�(����(i).Target).TargetSum: Max(1, 1) = ����(i).Target
        End If
    Next
    If Max(1, 0) > 0 Then '�����û�����Ӻÿ�������Բ
        ReDim VisLock(1 To fSUM)
        �ڵ�(Max(1, 1)).X = 0: �ڵ�(Max(1, 1)).Y = 0
        Angle = 2 * Pi / Max(1, 0)
        VisLock(Max(1, 1)) = True:
        For i = 1 To LSum
            If ����(i).a = True And ����(i).Target = Max(1, 1) Then
                CX = Angle * TarSumTemp
                VisLock(����(i).Source) = True
                If CX > Pi / 2 Then
                    CX = Pi - CX
                    �ڵ�(����(i).Source).X = 2000 * -Cos(CX)
                Else
                    �ڵ�(����(i).Source).X = 2000 * Cos(CX)
                End If
                    �ڵ�(����(i).Source).Y = 2000 * Sin(CX)
                TarSumTemp = TarSumTemp + 1
            End If
        Next
        Angle = 2 * Pi / (fSUM - 2 - Max(1, 0))
        '���ﵽ��һ��ע���ǰ�������Ӳ�ֱ����صĽڵ���ʱ����һ��
        TarSumTemp = 0
        For i = 1 To fSUM - 1
            If VisLock(i) = False Then
                CX = Angle * TarSumTemp
                If CX > Pi / 2 Then
                    CX = Pi - CX
                    �ڵ�(i).X = 5000 * -Cos(CX)
                Else
                    �ڵ�(i).X = 5000 * Cos(CX)
                End If
                    �ڵ�(i).Y = 5000 * Sin(CX)
                TarSumTemp = TarSumTemp + 1
            End If
        Next
    Else '�����û�����Ӻÿ�������Բת��
        '���浽����Χ��End If��û�л������ӵĺ������а���
        GoTo ����
    End If
End If
If Vis(3).Checked = True Then
����:
    Lengthways = 1
    For i = 1 To fSUM - 1
        Crosswise = Crosswise + 1
        If Crosswise > Sqr(fSUM) Then
            Crosswise = 1: Lengthways = Lengthways + 1
        End If
         If Lengthways Mod 2 = 0 Then
            �ڵ�(i).X = -Me.Width / (Int(Sqr(fSUM)) + 1) * (Crosswise + 0.2) + Me.Width / 2
        Else
            �ڵ�(i).X = -Me.Width / (Int(Sqr(fSUM)) + 1) * (Crosswise - 0.2) + Me.Width / 2
        End If
        If Crosswise Mod 2 = 0 Then
            �ڵ�(i).Y = -Me.Height / (Int(Sqr(fSUM)) + 1) * (Lengthways + 0.2) + Me.Height / 2
        Else
            �ڵ�(i).Y = -Me.Height / (Int(Sqr(fSUM)) + 1) * (Lengthways - 0.2) + Me.Height / 2
        End If
    Next
End If
NodeReDimLock = True
'----------------End---------------
Exit Sub
ErrHandler:
' �û����ˡ�ȡ������ť
End Sub

Private Sub �ڵ�����_Click()
Dim FindStr As String: Dim i As Long
On Error GoTo Er
FindLock = True
FindStr = InputBox("��������Ҫ����������")
For i = 1 To UBound(�ڵ�) - 1
    If �ڵ�(i).a = True And InStr(1, �ڵ�(i).Title, FindStr) <> 0 Then
        �ڵ�(i).Color = 2
    End If
Next
Er:
End Sub

Private Sub ���ô���ָ��_Click()
If ���ô���ָ��.Checked = True Then ���ô���ָ��.Checked = False Else ���ô���ָ��.Checked = True
End Sub

Private Sub �˳�_Click()
End
End Sub

Private Sub ��ʾ�ڵ����_Click()
If ��ʾ�ڵ����.Checked = True Then
    ����ʱָ����ʾ.Enabled = True: ��ʾ�ڵ����.Checked = False
Else
    If ��ʾ�ڵ����.Checked = True Then ����ʱָ����ʾ.Enabled = False
    ��ʾ�ڵ����.Checked = True
End If
End Sub

Private Sub ��ʾ�ڵ�����_Click()
If ��ʾ�ڵ�����.Checked = True Then
    ����ʱָ����ʾ.Enabled = True: ��ʾ�ڵ�����.Checked = False
Else
    If ��ʾ�ڵ����.Checked = True Then ����ʱָ����ʾ.Enabled = False
    ��ʾ�ڵ�����.Checked = True
End If
End Sub

Private Sub �½�_Click()
NodeReDimLock = False
Erase �ڵ�
Erase ����
End Sub

Private Sub ���ô���ָ��_Click()
If ���ô���ָ��.Checked = True Then ���ô���ָ��.Checked = False Else ���ô���ָ��.Checked = True
End Sub
