VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "試題四"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "計算"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   1695
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
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
Dim input1 As Long
input1 = Text1.Text
e = 0
a = 0
For i = 1 To Sqr(input1)
For i1 = 1 To Sqr(input1)
If i ^ 2 + i1 ^ 2 = input1 Then
e = e + 1
ouput = ouput & vbCrLf & e & " " & i & " " & i1
a = 1
End If
Next
Next

If a = 1 Then
Text2.Text = "count X Y" & ouput & vbCrLf & "There are " & e & " possible answers."
Else
Text2.Text = "count X Y" & vbCrLf & "Sorry,No answer found."
End If
End Sub

