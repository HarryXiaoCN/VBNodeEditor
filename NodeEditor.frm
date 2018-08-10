VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form NodeEditor 
   Caption         =   "Editor"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7485
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   7485
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   11033
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"NodeEditor.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "È·¶¨"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   7800
      Width           =   7215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Content"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   7335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "NodeEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
With NewNode
    .Title = Text1.Text
    .ContentOne = RichTextBox1.Text
End With
Unload Me
End Sub

Private Sub Form_Resize()
If Me.Height < 4000 Then Me.Height = 4000: Me.Enabled = False: Me.Enabled = True
On Error GoTo Er
For i = 0 To 1
    With Label1(i)
        .Width = Me.Width - 450
    End With
Next
With Text1
    .Width = Me.Width - 450
End With
With RichTextBox1
    .Height = Me.Height - 2565
    .Width = Me.Width - 450
End With
With Command1
    .Width = Me.Width - 450
    .Top = RichTextBox1.Height + 1545
End With
Er:
End Sub

Private Sub Form_Unload(Cancel As Integer)
XMB(EditLock).Enabled = True
End Sub

