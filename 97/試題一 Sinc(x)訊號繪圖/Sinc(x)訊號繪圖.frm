VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sinc(x)訊號繪圖"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   8445
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "Repolt"
      Height          =   495
      Left            =   5760
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   360
      ScaleHeight     =   2595
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   360
      Width           =   7695
   End
   Begin VB.Label Label2 
      Caption         =   "Max"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Min"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   3480
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
b = Text2.Text
If a < -29 Or a > 29 Or b < -29 Or b > 29 Then
MsgBox "請輸入-30~30之間", , "輸入錯誤"
Else
Picture1.Cls

Picture1.Scale (-30, 12)-(30, -3)
Picture1.Line (0, 10)-(0, -2)
Picture1.Line (-30, 0)-(30, 0)

For i = -2 To 10 Step 0.4 '畫Y
Picture1.Line (0, i)-(0.2, i)
Next

For i1 = a To b Step 0.4 '畫X
Picture1.Line (i1, 0)-(i1, 0.2)
Next

For i2 = -2 To 10 Step 2 'Y軸數字
Picture1.CurrentX = -1.3
Picture1.CurrentY = i2 + 0.2
Picture1.Print i2 / 10
Next

For i3 = a To b Step 5
Picture1.CurrentX = i3 - 0.6
Picture1.CurrentY = -0.3
Picture1.Print i3
Next


For i4 = a To b Step 0.01
If i4 = 0 Then
Picture1.PSet (0, 10), RGB(255, 0, 0)
Else
Picture1.PSet (i4, 10 * Sin(i4) / i4), RGB(255, 0, 0)
End If
Next
End If
End Sub


