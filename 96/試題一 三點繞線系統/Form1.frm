VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "三點繞線系統"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   427.875
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   9420
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text4 
      Height          =   1935
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   240
      ScaleHeight     =   700
      ScaleMode       =   0  '使用者自訂
      ScaleWidth      =   700
      TabIndex        =   4
      Top             =   840
      Width           =   4935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "X-Routing"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  '置中對齊
      Caption         =   "Saving:"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "  on-45 Length:"
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Non-45 Length:"
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "X-Routing for any three points"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x(3), x1(3), y(3), y1(3), x2(3), a
Private Sub Command1_Click()

Picture1.Cls
Randomize

For i = 1 To 3 '隨機產生三點
x(i) = Int(Rnd * 650)
y(i) = Int(Rnd * 650)
x1(i) = x(i)
y1(i) = y(i)
Next

For i = 1 To 3 '畫點
Picture1.DrawWidth = 3
Picture1.Circle (x(i), y(i)), 2
Next

For i = 1 To 3 '排列小→大
For i1 = i To 3
If y(i) > y(i1) Then
a = y(i): b = x(i)
y(i) = y(i1): x(i) = x(i1)
y(i1) = a: x(i1) = b
End If
Next
Next

For i = 1 To 3
x2(i) = x(i)
Next

For i = 1 To 3
For i1 = i To 3
If x1(i) > x1(i1) Then
a = x1(i)
x1(i) = x1(i1)
x1(i1) = a
End If
Next
Next

Picture1.DrawWidth = 1 '劃線
Picture1.Line (x(1), y(1))-(x(1), y(2))
Picture1.Line (x(3), y(3))-(x(3), y(2))
Picture1.Line (x(1), y(2))-(x(2), y(2))
Picture1.Line (x(2), y(2))-(x(3), y(2))
Text1.Text = ""
Text2.Text = ""
Text1.Text = Abs(y(2) - y(1)) + Abs(y(3) - y(2)) + Abs(x1(3) - x1(1))
Text4.Text = ""
Text3.Text = ""
For i = 1 To 3
Text4.Text = Text4.Text & x(i) & "," & y(i) & vbCrLf
Next
End Sub

Private Sub Command2_Click()
s1 = Abs(y(2) - y(1))
s2 = Abs(y(3) - y(2))
ans = 0

For i = 1 To 3
Picture1.DrawWidth = 3
Picture1.Circle (x(i), y(i)), 2, RGB(255, 0, 0)
Next

e1 = 0
e2 = 0

Picture1.DrawWidth = 1
If Abs(x(2) - x(1)) > s1 And x(1) < x(2) Then '判斷為哪一種圖形
    Picture1.Line (x(1), y(1))-(x(1) + s1, y(2)), RGB(255, 0, 0)
    Picture1.Line (x(2), y(2))-(x(1) + s1, y(2)), RGB(255, 0, 0)
    x(1) = Abs(x(2) - x(1) - s1) 'x(1)為轉為45度後與x(2)的距離
    e1 = 1
ElseIf Abs(x(2) - x(1)) > s1 And x(1) > x(2) Then
    Picture1.Line (x(1), y(1))-(x(1) - s1, y(2)), RGB(255, 0, 0)
    Picture1.Line (x(2), y(2))-(x(1) - s1, y(2)), RGB(255, 0, 0)
    x(1) = Abs(x(1) - s1 - x(2))
    e1 = 1
Else
    Picture1.Line (x(1), y(1))-(x(1), y(2)), RGB(255, 0, 0)
    Picture1.Line (x(1), y(2))-(x(2), y(2)), RGB(255, 0, 0)
    x(1) = Abs(x(1) - x(2))
    e2 = 0
End If


If Abs(x(3) - x(2)) > s2 And x(3) < x(2) Then '判斷為哪一種圖形
    Picture1.Line (x(3), y(3))-(x(3) + s2, y(2)), RGB(255, 0, 0)
    Picture1.Line (x(2), y(2))-(x(3) + s2, y(2)), RGB(255, 0, 0)
    x(3) = Abs(x(2) - x(3) - s2) 'x(3)為轉為45度後與x(2)的距離
    e2 = 1
ElseIf Abs(x(2) - x(3)) > s2 And x(3) > x(2) Then
    Picture1.Line (x(3), y(3))-(x(3) - s2, y(2)), RGB(255, 0, 0)
    Picture1.Line (x(2), y(2))-(x(3) - s2, y(2)), RGB(255, 0, 0)
    x(3) = Abs(x(3) - s2 - x(2))
    e2 = 1
Else
    Picture1.Line (x(3), y(3))-(x(3), y(2)), RGB(255, 0, 0)
    Picture1.Line (x(3), y(2))-(x(2), y(2)), RGB(255, 0, 0)
    x(3) = Abs(x(3) - x(2))
    e2 = 0
End If

'以下為4種可能情況
If e1 = 0 And e2 = 0 Then Text2.Text = ""

  If e1 = 1 And e2 = 0 Then
    If x(1) > x(3) Then
      ans = (2 * s1 ^ 2) ^ 0.5 + (y(3) - y(2)) + x(1)
    Else
      ans = (2 * s1 ^ 2) ^ 0.5 + (y(3) - y(2)) + x(3)
  End If
End If

If e1 = 0 And e2 = 1 Then
  If x(3) > x(1) Then
    ans = (2 * s2 ^ 2) ^ 0.5 + (y(2) - y(1)) + x(3)
  Else
    ans = (2 * s2 ^ 2) ^ 0.5 + (y(2) - y(1)) + x(1)
  End If
End If


If e1 = 1 And e2 = 1 Then
  If x(1) > x(3) Then
   ans = (2 * s1 ^ 2) ^ 0.5 + (2 * (s2 ^ 2)) ^ 0.5 + x(1)
  Else
   ans = (2 * s1 ^ 2) ^ 0.5 + (2 * (s2 ^ 2)) ^ 0.5 + x(3)
  End If
End If

If Text1.Text > ans Then '判斷轉換後總長度是否小於原長度
  Text2.Text = Int(ans)
  Text3.Text = Format((Text1.Text - Text2.Text) / Text1.Text * 100, "##,##") & "%"
Else '如沒有小於 則顯示原圖形
  Text2.Text = ""
  Text3.Text = ""
  Picture1.Cls

For i = 1 To 3
  Picture1.DrawWidth = 3
  Picture1.Circle (x2(i), y(i)), 2, RGB(255, 0, 0)
Next
  Picture1.DrawWidth = 1
  Picture1.Line (x2(1), y(1))-(x2(1), y(2)), RGB(255, 0, 0)
  Picture1.Line (x2(3), y(3))-(x2(3), y(2)), RGB(255, 0, 0)
  Picture1.Line (x2(1), y(2))-(x2(2), y(2)), RGB(255, 0, 0)
  Picture1.Line (x2(2), y(2))-(x2(3), y(2)), RGB(255, 0, 0)
  Text2.Text = 0
  Text3.Text = Format((Text1.Text - Text2.Text) / Text1.Text * 100, "##,##") & "%"
End If

End Sub


Private Sub Command3_Click()
End
End Sub


