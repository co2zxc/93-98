VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   2775
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "輸入："
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
If Text1.Text = "" Then Exit Sub
Label2.Caption = 1 & vbCrLf & 11 & vbCrLf
n = Text1.Text
str1 = "11"
If n < 2 Or n > 10 Then Label2.Caption = "": Exit Sub
For i = 2 To Val(n)
a = Left(str1, 1)
c = 0
For i1 = 1 To Len(str1)
g = Mid(str1, i1, 1)
If g = a Then
b = g
c = c + 1
Else
a = g
ouput = ouput & c & b
c = 1
b = g
End If
Next
ouput = ouput & c & b
str1 = ouput
ouput = ""
Label2.Caption = Label2.Caption & str1 & vbCrLf
Next


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Label2.Caption = "": Exit Sub
End Sub
