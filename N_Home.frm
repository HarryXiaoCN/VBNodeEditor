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
      End
      Begin VB.Menu �� 
         Caption         =   "��"
      End
      Begin VB.Menu ���� 
         Caption         =   "����"
      End
      Begin VB.Menu ���Ϊ 
         Caption         =   "���Ϊ"
      End
      Begin VB.Menu �˳� 
         Caption         =   "�˳�"
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
Private MeID, LSum As Long
Private NodeReDimLock As Boolean
Private Sub Form_Activate()
Dim STemp '��ô����Z����
STemp = Split(Me.Caption, " - ")
MeID = Val(STemp(0))
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
If NodeReDimLock = False Then Exit Sub
For i = 1 To UBound(�ڵ�) - 1
    If �ڵ�(i).A = True Then
        If X + 100 > �ڵ�(i).X And X - 100 < �ڵ�(i).X _
        And Y + 100 > �ڵ�(i).Y And Y - 100 < �ڵ�(i).Y Then
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
Me.Print UBound(�ڵ�)
For i = 1 To UBound(�ڵ�) - 1
    If �ڵ�(i).A = True Then
        Me.CurrentX = �ڵ�(i).X
        Me.CurrentY = �ڵ�(i).Y
        Me.Print �ڵ�(i).Title
        If MousePlace.X + 100 > �ڵ�(i).X And MousePlace.X - 100 < �ڵ�(i).X _
        And MousePlace.Y + 100 > �ڵ�(i).Y And MousePlace.Y - 100 < �ڵ�(i).Y Then
            Me.Circle (�ڵ�(i).X, �ڵ�(i).Y), 100, RGB(255, 0, 0)
            Me.CurrentX = 0: Me.CurrentY = 200: Me.Print i
        Else
            Me.Circle (�ڵ�(i).X, �ڵ�(i).Y), 100, RGB(0, 191, 255)
        End If
    End If
Next
For i = 1 To LSum
    If ����(i).A = True Then
        Me.Line (�ڵ�(����(i).Source).X, �ڵ�(����(i).Source).Y)-(�ڵ�(����(i).Target).X, �ڵ�(����(i).Target).Y), RGB(126, 126, 126)
    End If
Next
If MousePlace.Target <> 0 And MousePlace.Z = MeID Then
    �ڵ�(MousePlace.Target).X = MousePlace.X
    �ڵ�(MousePlace.Target).Y = MousePlace.Y
End If
End Sub
Private Sub ��_Click()
' ���á�CancelError��Ϊ True
CommonDialog1.CancelError = True
'On Error GoTo ErrHandler
' ���ñ�־
CommonDialog1.Flags = cdlOFNHideReadOnly
' ���ù�����
CommonDialog1.Filter = "VBBas Files" & _
"(*.bas)|*.bas|All Files (*.*)|*.*"
' ָ��ȱʡ�Ĺ�����
CommonDialog1.FilterIndex = 1
' ��ʾ���򿪡��Ի���
CommonDialog1.ShowOpen
' ��ʾѡ���ļ�������
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
ReDim Preserve �ڵ�(fSUM)
For i = 1 To fSUM - 1
    �ڵ�(i).A = True
    �ڵ�(i).Title = Package(i).Title
    �ڵ�(i).X = Me.Width / fSUM * i
    �ڵ�(i).Y = Me.Height / fSUM * i
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
                    LSum = LSum + 1
                    ����(LSum).A = True: ����(LSum).Source = i: ����(LSum).Target = c
                End If
            Next
        Next
        For c = i + 1 To fSUM - 1
            For j = 1 To UBound(�ڵ�(c).Content)
                If InStr(1, �ڵ�(c).Content(j), " " & �ڵ�(i).Title & " ") <> 0 _
                Or InStr(1, �ڵ�(c).Content(j), " " & �ڵ�(i).Title & "(") <> 0 Then
                    LSum = LSum + 1
                    ����(LSum).A = True: ����(LSum).Source = i: ����(LSum).Target = c
                End If
            Next
        Next
Next
NodeReDimLock = True
'----------------End---------------
Exit Sub
ErrHandler:
' �û����ˡ�ȡ������ť
End Sub
Private Sub �˳�_Click()
End
End Sub
