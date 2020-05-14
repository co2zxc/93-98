VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "連線延遲"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   4020
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "計算"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "輸出"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "輸入"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
input1 = Text1.Text
a = Split(input1)

w = a(0)
r = a(1)
c = Val(Left(a(2), Len(a(2)) - 1)) * 10 ^ -12

If w >= 0 And r < 200 Or c < 0.4 * 10 ^ -12 Then Text2.Text = "輸入錯誤 請重新輸入": Exit Sub

Ro = 350
Cl = 0.2 * 10 ^ -12
TB = 350 * 10 ^ -12
RB = 350
CB = 0.04 * 10 ^ -12

If w = 0 Then
Td = (Ro + r) * (c + Cl)
ElseIf w > 0 Then
Td = (Ro + r / 2) * (c / 2 + w * CB) + TB + (RB / w + r / 2) * (c / 2 + Cl)
Else
End If

Td = Td * 10 ^ 12 & "ps"

Text2.Text = Td



End Sub

