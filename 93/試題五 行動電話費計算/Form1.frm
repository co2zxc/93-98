VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "行動電話月租費計算"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   3405
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "計算"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "輸出"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "輸入"
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
input1 = Text1.Text
a = Split(input1)
a(1) = Val(a(1))

Select Case a(0)
Case "1"
ouput = 600 + a(1) * 5
Case "2"
ouput = 200 + a(1) * 9
If a(1) * 9 >= 200 Then ouput = ouput - 400
Case "3"
If a(1) >= 5 Then
ouput = 66 + (a(1) - 5) * 12
Else
ouput = 66
End If

If (a(1) - 5) * 12 > 66 Then ouput = ouput - 132

End Select

Text2.Text = ouput

End Sub

