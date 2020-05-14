VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "畫Sin 及Cos 函數圖形"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   8115
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command3 
      Caption         =   "結束"
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "畫出"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "函數圖形"
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   4680
      Width           =   2655
      Begin VB.OptionButton Option2 
         Caption         =   "Cos圖形"
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sin圖形"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   240
      ScaleHeight     =   10
      ScaleMode       =   0  '使用者自訂
      ScaleWidth      =   10
      TabIndex        =   0
      Top             =   240
      Width           =   7455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

pi = 3.14
Picture1.Scale (0, 10)-(2 * pi, -10)
Picture1.Line (0, 0)-(10, 0)

If Option1.Value = True Then
For i = 0 To 2 * pi Step 0.001
Picture1.PSet (i, Sin(i) * 5)
Next
End If

If Option2.Value = True Then
For i = 0 To 2 * pi Step 0.001
Picture1.PSet (i, Cos(i) * 5)
Next
End If



End Sub

Private Sub Command2_Click()
Picture1.Cls
End Sub

Private Sub Command3_Click()
End

End Sub
