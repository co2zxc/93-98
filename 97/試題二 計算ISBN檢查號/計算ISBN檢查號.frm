VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "計算ISBN檢查號"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   6930
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   3000
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   2280
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "計算ISNBN檢查號"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "ISBN-13"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "ISBN-10"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "請輸入9位數字"
      Height          =   495
      Left            =   240
      TabIndex        =   1
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
a = Text1.Text


For i = 1 To Len(a)
str1 = Mid(a, i, 1)
If str1 >= 0 And str1 <= 9 And str1 <> "-" Then
str2 = str2 & str1
End If
Next

If Len(str2) = 9 Then

For i1 = 10 To 2 Step -1
num = Mid(str2, 11 - i1, 1) * i1
Sum = Sum + num
Next

m = Sum Mod 11
n = 11 - m
If n = 10 Then
chk = "X"
ElseIf n = 11 Then
chk = "0"
Else
chk = n
End If

'957857358
'957442355
If Mid(str2, 4, 3) = "157" Or Mid(str2, 4, 3) = "204" Or Mid(str2, 4, 3) = "421" Or Mid(str2, 4, 3) = "442" Then
Text2.Text = Mid(str2, 1, 3) & "-" & Mid(str2, 4, 3) & "-" & Mid(str2, 7, 3) & "-" & chk
ElseIf Mid(str2, 4, 4) = "7198" Or Mid(str2, 4, 4) = "7323" Or Mid(str2, 4, 4) = "8573" Then
Text2.Text = Mid(str2, 1, 3) & "-" & Mid(str2, 4, 4) & "-" & Mid(str2, 8, 2) & "-" & chk
Else
MsgBox "請重新輸入", , "輸入錯誤"
End If

str3 = "978" & str2
x = 3
Sum = 0
For i2 = 1 To 12
If x = 3 Then
x = 1
num = Mid(str3, i2, 1) * x
ElseIf x = 1 Then
x = 3
num = Mid(str3, i2, 1) * x

End If
Sum = Sum + num
Next
m = Sum Mod 10
If m = 0 Then
chk = "0"
Else
chk = 10 - m
End If

If Mid(str2, 4, 3) = "157" Or Mid(str2, 4, 3) = "204" Or Mid(str2, 4, 3) = "421" Or Mid(str2, 4, 3) = "442" Then
Text3.Text = "978" & "-" & Mid(str2, 1, 3) & "-" & Mid(str2, 4, 3) & "-" & Mid(str2, 7, 3) & "-" & chk
ElseIf Mid(str2, 4, 4) = "7198" Or Mid(str2, 4, 4) = "7323" Or Mid(str2, 4, 4) = "8573" Then
Text3.Text = "978" & "-" & Mid(str2, 1, 3) & "-" & Mid(str2, 4, 4) & "-" & Mid(str2, 8, 2) & "-" & chk
End If


Else

MsgBox "輸入錯誤", , "輸入錯誤 """
End If

End Sub

