VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "同音代換加密法"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   6990
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   2520
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "加密"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Text            =   "abcabcabcabc"
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "密文"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "明文"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'A~J陣列放入數值
a = Array("", "09", "12", "33", "47", "53", "67", "78", "92")
b = Array("", "48", "81")
c = Array("", 13, 41, 62)
d = Array("", "01", "03", 45, 79)
e = Array("", 14, 16, 24, 44, 46, 55, 57, 64, 74, 82, 87, 98)
f = Array("", 10, 31)
g = Array("", "06", 25)
h = Array("", 23, 39, 50, 56, 65, 68)
i = Array("", 32, 70, 73, 83, 88, 93)
j = Array("", 15)

input1 = Text1.Text
Text2.Text = ""
For i = 1 To Len(input1)
str1 = Mid(input1, i, 1)
Select Case str1
Case "a"
asum = asum + 1: If asum > 8 Then asum = 1
Text2.Text = Text2.Text & " " & a(asum)
Case "b"
bsum = bsum + 1: If bsum > 2 Then bsum = 1
Text2.Text = Text2.Text & " " & b(bsum)
Case "c"
csum = csum + 1: If csum > 3 Then csum = 1
Text2.Text = Text2.Text & " " & c(csum)
Case "d"
dsum = dsum + 1: If dsum > 4 Then dsum = 1
Text2.Text = Text2.Text & " " & d(dsum)
Case "e"
esum = esum + 1: If esum > 12 Then esum = 1
Text2.Text = Text2.Text & " " & e(esum)
Case "f"
fsum = fsum + 1: If fsum > 2 Then fsum = 1
Text2.Text = Text2.Text & " " & f(fsum)
Case "g"
gsum = gsum + 1: If gsum > 2 Then gsum = 1
Text2.Text = Text2.Text & " " & g(gsum)
Case "h"
hsum = hsum + 1: If hsum > 6 Then hsum = 1
Text2.Text = Text2.Text & " " & h(hsum)
Case "i"
isum = isum + 1: If isum > 6 Then isum = 1
Text2.Text = Text2.Text & " " & i(isum)
Case "j"
jsum = jsum + 1: If jsum > 1 Then jsum = 1
Text2.Text = Text2.Text & " " & j(jsum)
End Select

Next





End Sub

