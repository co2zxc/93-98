VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��Ƥ���"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   6600
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.TextBox Text2 
      Alignment       =   2  '�m�����
      Height          =   2775
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�@���s"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "�п�J�@���N"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sum As Integer, e As Integer
Private Sub Command1_Click()
n = Val(Text1.Text)
If Text1.Text = "***" And e = 2 Then
e = 0
Text1.Text = ""
Label2.Caption = ""
Text1.ForeColor = RGB(0, 0, 0)
sum = 0
ElseIf n < 1 Or n > 10 Then '�P�_��J���~
sum = sum + 1
Label2.Caption = "��J���~"
e = 1
End If

If sum > 3 Then '���~�W�L�T���Y��ܿ��~�T��
Label2.Caption = "��J�W�L3��"
Text1.Text = "???"
Text1.ForeColor = RGB(255, 0, 0)
e = 2
End If

If e = 0 Then
Text2.Text = ""
For i = n To 1 Step -1
Text2.Text = Text2.Text & i & " "
For i1 = n - i To 1 Step -1
Text2.Text = Text2.Text & 1 & " "
Next
Text2.Text = Text2.Text & vbCrLf
Next
End If

End Sub
