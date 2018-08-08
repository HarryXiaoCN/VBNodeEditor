Attribute VB_Name = "N_Function"
Sub Main()
VBbasName = "NewVB.bas"
VBbasPath = App.Path & "\" & VBbasName
NewFileLoad
End Sub
Public Function NewFileLoad()
Dim i As Long
ReDim Main(10) As New Home
For i = 0 To ZSum
    Main(i).Show
    Main(i).Scale (0, 0)-(Main(i).Width, Main(i).Height)
    HomeCapationVisFilePath Main(i), i
Next
End Function
Public Function HomeCapationVisFilePath(ByRef FormName, ByRef Z As Long)
FormName.Caption = Z & " - " & VBbasName
End Function
