VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "垃圾處理費計算"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   7335
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "垃圾處理費計算"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

input1 = InputBox("請輸入水費度數及電費度數", "輸入", 0)

a = Split(input1)


If a(0) < 50 Then e1 = 1
If a(0) >= 50 And a(0) <= 100 Then e1 = 2
If a(0) > 100 Then e1 = 3

If a(1) < 100 Then e2 = 1
If a(1) >= 100 And a(1) <= 200 Then e2 = 2
If a(1) > 200 Then e2 = 3

If e1 = 1 And e2 = 1 Then
Sum = (Val(a(0)) + Val(a(1))) / 2 * 0.6
ElseIf e1 = 1 And e2 = 2 Or e1 = 2 And e2 = 1 Then
Sum = (Val(a(0)) + Val(a(1))) / 2 * 0.8
ElseIf e1 = 3 And e2 = 3 Then
Sum = (Val(a(0)) + Val(a(1))) / 2 * 1.4
ElseIf e1 = 3 And e2 = 2 Or e1 = 2 And e2 = 3 Then
Sum = (Val(a(0)) + Val(a(1))) / 2 * 1.2
Else
Sum = (Val(a(0)) + Val(a(1))) / 2
End If

Label1.Caption = "計算後的垃圾處理費用:" & Sum * 5




End Sub
