Attribute VB_Name = "N_Function"
Sub Main()
VBbasName = "NewVB.bas"
VBbasPath = App.Path & "\" & VBbasName
NewFileLoad
End Sub
Public Function NewFileLoad()
Dim i As Long
ReDim XMB(10) As New Home
For i = 0 To 9
    If i = 0 Then XMB(i).Show Else Load XMB(i)
    XMB(i).Scale (-XMB(i).Width / 2, XMB(i).Height / 2)-(XMB(i).Width / 2, -XMB(i).Height / 2)
    XMB(i).Vis(0).Checked = True
    XMB(i).��ʾ�ڵ����.Checked = True
    XMB(i).��ʾ�ڵ�����.Checked = True
    XMB(i).����ʱָ����ʾ.Enabled = False
    HomeCapationVisFilePath XMB(i), i
Next
End Function
Public Function HomeCapationVisFilePath(ByRef FormName, ByRef Z As Long)
FormName.Caption = Z & " - " & VBbasName
End Function
