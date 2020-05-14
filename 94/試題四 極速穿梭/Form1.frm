VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "極速穿梭"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   2895
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "計算"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   1695
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "輸出"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "輸入"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
input1 = Split(Text1.Text) '超簡單不做說明~"~
n = UBound(input1)
Text2.Text = ""

For i = 0 To n - 1 Step 3
ouput = Val(input1(i)) * Val(Mid(input1(i + 1), 1, 1)) / Val(Mid(input1(i + 1), 3, 1)) / Val(input1(i + 2))
Text2.Text = Text2.Text & Format(ouput, 0) & vbCrLf
Next



End Sub
