VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "璸衡の玻ネ借计计"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   6525
   StartUpPosition =   3  '╰参箇砞
   Begin VB.CommandButton Command1 
      Caption         =   "玻ネ借计计"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label5 
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "程借计"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "借计计Τ"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "叫块计"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a(3)
n = Text1.Text
c = 1
For i = 2 To n
  If IsPrime(i) Then
  Sum = Sum + 1
  End If
Next

num = 1
For i = n To 2 Step -1
    If IsPrime(i) Then
        a(num) = i
        num = num + 1
    End If
If num = 4 Then Exit For
Next

Label4.Caption = Sum
Label5.Caption = a(1) & " " & a(2) & " " & a(3)

End Sub

Function IsPrime(ByVal n As Long) As Boolean
If n = 2 Or n = 3 Then IsPrime = True: Exit Function
    For i = 2 To Sqr(n)
        If n Mod i = 0 Then IsPrime = False: Exit Function
    Next
IsPrime = True

End Function
