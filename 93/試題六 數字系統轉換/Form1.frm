VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "數字系統轉換"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "計算"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "輸出"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "數值"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "基底"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As Double
r = Text1.Text
e = 0
For i = 1 To Len(Text2.Text) '取小數點前
If Mid(Text2.Text, i, 1) = "." Then
e = 1
Exit For
Else
input1 = input1 & Mid(Text2.Text, i, 1)
End If
Next
If e = 1 Then input2 = Right(Text2.Text, Len(Text2.Text) - Len(input1) - 1) '取小數點後


If r < 2 Or r > 16 Then Text3.Text = "基底輸入錯誤": Exit Sub '進制範圍2~16

If r >= 2 And r < 10 Then '檢查錯誤
For i = 1 To Len(input1)
If Mid(input1, i, 1) >= r Then MsgBox "數值輸入，請重新輸入", , "錯誤": Exit Sub
Next
End If

If r >= 10 And r <= 16 Then '檢查10進制以上錯誤
For i = 1 To Len(input1)
num = Mid(input1, i, 1)
If num > "F" Then MsgBox "數值輸入，請重新輸入", , "錯誤": Exit Sub

If num = "A" Then
num = 10
ElseIf num = "B" Then
num = 11
ElseIf num = "C" Then
num = 12
ElseIf num = "D" Then
num = 13
ElseIf num = "E" Then
num = 14
ElseIf num = "F" Then
num = 15
End If

If Val(num) >= r Then MsgBox "數值輸入，請重新輸入", , "錯誤": Exit Sub
Next
End If

For i = 0 To Len(input1) - 1 '計算加權後的10進位數值
num = Mid(input1, Len(input1) - i, 1)
If num = "A" Then
num = 10
ElseIf num = "B" Then
num = 11
ElseIf num = "C" Then
num = 12
ElseIf num = "D" Then
num = 13
ElseIf num = "E" Then
num = 14
ElseIf num = "F" Then
num = 15
End If
Sum = Sum + Val(num) * r ^ i
Next

a = -1
For i = 1 To Len(input2) '小數點加權計算
sum2 = sum2 + Val(Mid(input2, i, 1)) * r ^ a
a = a - 1
Next

Text3.Text = Sum + sum2


End Sub
