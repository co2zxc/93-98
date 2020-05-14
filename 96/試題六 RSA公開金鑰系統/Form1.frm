VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "RSA公開金鑰系統"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   6870
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "解密"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   4080
      Width           =   5055
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Text            =   "68269"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Text            =   "9907"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   2280
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "加密"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Text            =   "68269"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Text            =   "8315"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "明文"
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "密文"
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "秘密金鑰"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "密文"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "明文"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "公開金鑰"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
input1 = Text3.Text
e = Val(Text1.Text)
n = Val(Text2.Text)

Do While e > 0
  b = e Mod 2 & b
  e = e \ 2
Loop
For i = 1 To Len(input1)
  If Asc(Mid(input1, i, 1)) > 0 Then
    Call aa(Asc(Mid(input1, i, 1)), b, n, ouput)
  Else
    Call aa(Asc(Mid(input1, i, 1)) + 65536, b, n, ouput) '中文字ASC < 0 所以要+65536 因為文中字佔2Byte
  End If
Next
Text4.Text = ouput
End Sub

Sub aa(a, b, c, ouput)
s = 1
For i = 1 To Len(b)
    s = s * s
    Do While s > c
      s = s - c
    Loop
  If Mid(b, i, 1) = 1 Then
    s = a * s
    Do While s > c
      s = s - c
    Loop
    End If
Next
If Len(s) < 5 Then s = 0 & s
ouput = ouput & s
End Sub

Private Sub Command2_Click()
input1 = Text7.Text
d = Val(Text5.Text)
n = Val(Text6.Text)

Do While d > 0
  b = d Mod 2 & b
  d = d \ 2
Loop

c = Len(input1) / 5

For i = 0 To c - 1
  Call bb(Val(Mid(input1, 1 + 5 * i, 5)), b, n, ouput)
Next

Text8.Text = ouput
End Sub
Sub bb(a, b, c, ouput)
s = 1
For i = 1 To Len(b)
    s = s * s
    Do While s > c
      s = s - c
    Loop
  If Mid(b, i, 1) = 1 Then
    s = a * s
    Do While s > c
      s = s - c
    Loop
    End If
Next
ouput = ouput & Chr(s)
End Sub
