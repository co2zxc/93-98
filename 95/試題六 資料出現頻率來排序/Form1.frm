VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�̸�ƥX�{�W�v�ӱƧ�"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   10335
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Text            =   "Conventional approaches for encoding technique"
      Top             =   960
      Width           =   9375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��J"
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   9375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim word(), sum(), s()
Dim ishave As Boolean
input1 = Text1.Text
ReDim word(Len(input1)), sum(Len(input1)), s(Len(input1))


For i = 1 To Len(input1)
ishave = False
For i1 = 1 To a
If word(i1) = Mid(input1, i, 1) Then
sum(i1) = sum(i1) + 1 '�p�����ƫh����r�X�{����+1,���X�^��
ishave = True
Exit For
End If
Next

If ishave = False Then
a = a + 1 '�����ƫh�N�s�r�����U�Ӱ}�C
word(a) = Mid(input1, i, 1)
sum(a) = 1
End If
Next

For i = 1 To a '�ѥX�{���Ƥj��p�ƦC
For i1 = i To a
If sum(i) < sum(i1) Then
tmp = sum(i): tmp1 = word(i)
sum(i) = sum(i1): word(i) = word(i1)
sum(i1) = tmp: word(i1) = tmp1
End If
Next
Next

For i = 1 To a
Label1.Caption = Label1.Caption & word(i) & "=" & sum(i) & ";  "
Next
End Sub

