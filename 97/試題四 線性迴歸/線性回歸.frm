VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�u�ʦ^�k"
   ClientHeight    =   8460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   11085
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   6720
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   4800
      ScaleHeight     =   5475
      ScaleWidth      =   5955
      TabIndex        =   1
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label Label2 
      Caption         =   "�п�J����`��"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Height          =   5535
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim x(10) As Single
Dim y(10) As Single
a = Text1.Text

If a >= 2 And a <= 10 Then

For i = 1 To a
1:
x(i) = InputBox("�п�J��" & i & "�Ix�y��", "��J�C�I�y��")
If x(i) < 1 Or x(i) > 8 Then MsgBox "�п�J1~8����", , "��J���~": GoTo 1
2:
y(i) = InputBox("�п�J��" & i & "�Iy�y��", "��J�C�I�y��")
If y(i) < 1 Or y(i) > 8 Then MsgBox "�п�J1~8����", , "��J���~": GoTo 2
Next


For i = 1 To a
sumx = sumx + x(i)
sumy = sumy + y(i)
sumxx = sumxx + x(i) ^ 2
sumxy = sumxy + x(i) * y(i)
Next

avgx = sumx / a
avgy = sumy / a
m = Format((sumxy - sumx * avgy) / (sumxx - sumx * avgx), "0.000")
b = Format(avgy - m * avgx, "0.000")

Picture1.Scale (0, 9)-(9, 0)

Picture1.DrawStyle = 0
Picture1.Line (1, 1)-(8, 1)
Picture1.Line (1, 1)-(1, 8)
Picture1.Line (1, 8)-(8, 8)
Picture1.Line (8, 1)-(8, 8)

For i = 2 To 7
Picture1.DrawStyle = 2
Picture1.Line (i, 1)-(i, 8)
Picture1.Line (1, i)-(8, i)
Next

For i = 1 To a
Picture1.DrawStyle = 0
Picture1.Circle (x(i), y(i)), 0.1, RGB(255, 0, 0)
Next

For i = 1 To 8
Picture1.CurrentX = i
Picture1.CurrentY = 0.5
Picture1.Print i
Picture1.CurrentX = 0.7
Picture1.CurrentY = i
Picture1.Print i
Next

For i = 1 To a - 1
Picture1.Line (x(i), y(i))-(x(i + 1), y(i + 1)), RGB(255, 0, 0)
Next


Label1.Caption = Label1.Caption & "�u�ʦ^�k(Linear Regression)" & vbCrLf & "�Q�γ̤p�����k�ӱ���@���I" & vbCrLf & "�п�J����I�`��" & a
For i = 1 To a
Label1.Caption = Label1.Caption & vbCrLf & "�п�J�C�@�I��ƪ�x,y�y��[x y] : " & "[" & x(i) & " " & y(i) & "]"
Next
Label1.Caption = Label1.Caption & vbCrLf & "�̤p����Ƚu���^�k�Y�� "
Label1.Caption = Label1.Caption & vbCrLf & "�ײv(m)     = " & m
Label1.Caption = Label1.Caption & vbCrLf & "�`�Z(b)     = " & b
Label1.Caption = Label1.Caption & vbCrLf & "�`����I��     = " & a

Else

MsgBox "�п�J2~10����", , "��J���~"
End If

End Sub

