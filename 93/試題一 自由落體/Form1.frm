VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "自由落體"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   3750
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "計算"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "輸出"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "輸入"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   1335
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
ouput = a(0)

Do While a(0) > 0
h = Int(a(0) / 2 - a(1) / 5)
If h < 0 Then h = 0
a(0) = h
e = e + 1
ouput = ouput & " " & h
Loop
e = e - 1
If e < 0 Then e = 0
Text2.Text = ouput & " " & e

End Sub

