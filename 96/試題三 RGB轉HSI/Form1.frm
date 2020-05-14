VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   4950
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "HSI轉RGB"
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox text3 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox text2 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox text1 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RGB轉HSI"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "I"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "S"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "H"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "B"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "G"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "R"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   480
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function ACos(x As Double) As Double '反三角函數COS
pi = 3.14159265358979
If x = 1 Then
ACos = 0
ElseIf x = -1 Then
ACos = pi
Else
ACos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End If
End Function

Private Sub Command1_Click()

r = Val(text1.Text)
g = Val(text2.Text)
b = Val(text3.Text)

pi = 3.14159265358979
S = r + g + b
r1 = r / S
g1 = g / S
b1 = b / S

a1 = 0.5 * ((r1 - g1) + (r1 - b1))
a2 = ((r1 - g1) ^ 2 + (r1 - b1) * (g1 - b1)) ^ 0.5

If b <= g Then
H = ACos(0.5 * ((r1 - g1) + (r1 - b1)) / ((r1 - g1) ^ 2 + (r1 - b1) * (g1 - b1)) ^ 0.5)

Else
H = 2 * pi - ACos(0.5 * ((r1 - g1) + (r1 - b1)) / ((r1 - g1) ^ 2 + (r1 - b1) * (g1 - b1)) ^ 0.5)
End If


S = 1 - 3 * min(r1, g1, b1)
I = (r + g + b) / (3 * 255)

H = H * 180 / pi
S = S * 255
I = I * 255

Text4.Text = H
Text5.Text = S
Text6.Text = I
End Sub

Function min(ByVal r As Double, ByVal g As Double, ByVal b As Double) As Double
min = r
If min > g Then min = g
If min > b Then min = b
End Function

Private Sub Command2_Click()

H = Val(Text4.Text)
S = Val(Text5.Text)
I = Val(Text6.Text)

pi = 3.14159265358979

h1 = H * pi / 180
s1 = S / 255
i1 = I / 255

If h1 < 2 * pi / 3 Then
x = i1 * (1 - s1)
y = i1 * (1 + s1 * Cos(h1) / Cos(pi / 3 - h1))
z = 3 * i1 - (x + y)

b = x
r = y
g = z

ElseIf h1 >= 2 * pi / 3 And h1 < 4 * pi / 3 Then
h1 = h1 - 2 * pi / 3
x = i1 * (1 - s1)
y = i1 * (1 + s1 * Cos(h1) / Cos(pi / 3 - h1))
z = 3 * i1 - (x + y)
r = x
g = y
b = z

ElseIf h1 >= 4 * pi / 3 And h1 < 2 * pi Then
h1 = h1 - 4 * pi / 3
x = i1 * (1 - s1)
y = i1 * (1 + s1 * Cos(h1) / Cos(pi / 3 - h1))
z = 3 * i1 - (x + y)
g = x
b = y
r = z
End If

text1.Text = Format(r * 255, "0")
text2.Text = Format(g * 255, "0")
text3.Text = Format(b * 255, "0")

End Sub
