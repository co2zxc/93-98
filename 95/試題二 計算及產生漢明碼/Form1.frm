VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�p��β��ͺ~���X"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   5895
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.TextBox Text2 
      Alignment       =   1  '�a�k���
      Height          =   270
      Left            =   2880
      TabIndex        =   3
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���ͺ~���X"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�a�k���
      Height          =   270
      Left            =   2880
      TabIndex        =   1
      Text            =   "1101101011"
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "�t���~���X���T��"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "���ǻ����T��"
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim input1 As String, ouput As String
input1 = Text1.Text
k = 1
n = Len(input1)

If Len(input1) > 11 Then Text2.Text = "���ǻ��T�������פ��W�L11�줸": Exit Sub

For i = 1 To Len(input1)
If Val(Mid(input1, i, 1)) > 1 Then Text2.Text = "���ǻ����T������0��1": Exit Sub
Next


Do Until 2 ^ k >= n + k + 1 '�D�oK���ˬd�줸
  k = k + 1
Loop
ouput = input1

For i = 1 To k '�N�ˬd�줸���J�T��
  If i = 1 Or i = 2 Then
    ouput = ouput & "A"
  Else
    a = 2 ^ (i - 1)
    s1 = Mid(ouput, 1, Len(ouput) - a + 1)
    s2 = Mid(ouput, Len(ouput) - a + 2)
    ouput = s1 & "A" & s2
  End If
Next

xornum = 0
For i = 1 To Len(ouput) '�N�줸��1����XOR�B��
  If Mid(ouput, i, 1) = "1" Then
    xornum = xornum Xor Len(ouput) + 1 - i
  End If
Next

Do While xornum > 0 '�ˬd�X��G�i��
  If xornum Mod 2 = 0 Then
    str1 = 0 & str1
  Else
    str1 = 1 & str1
  End If
  xornum = xornum \ 2
Loop

xornum = str1

Do While Len(xornum) < k '�p���ץ��F�ˬd�줸 �h��0
xornum = "0" & xornum
Loop

i1 = k
i2 = 1
For i = 1 To Len(ouput)
  If Mid(ouput, i, 1) = "A" Then
    a = 2 ^ (i1 - 1)
    s1 = Mid(ouput, 1, Len(ouput) - a)
    s2 = Mid(ouput, Len(ouput) - a + 2)
    ouput = s1 & Mid(xornum, i2, 1) & s2
    i2 = i2 + 1
    i1 = i1 - 1
  End If
Next

Text2.Text = ouput

End Sub

