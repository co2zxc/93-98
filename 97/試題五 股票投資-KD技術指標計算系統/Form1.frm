VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "�Ѳ����-KD�޳N���аO��t��"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   7890
   StartUpPosition =   3  '�t�ιw�]��
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Left            =   4200
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "��8��d��"
      Height          =   495
      Left            =   4800
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "��8��k��"
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "��X�ɮ׸��|"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "��J�ɮ׸��|"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim str1 As String
Dim str2 As String
Dim str3 As String
Dim num(12) As String
Dim max(12) As String '���C��̰���
Dim min(12) As String '���C��̧C��
Dim ���L(12) As String
Dim k(12) As String 'K��
Dim d(12) As String 'D��
Dim rsv(12) As Single
k(8) = Text3.Text  '��8��KD��
d(8) = Text4.Text
file = Text1.Text

Open file For Input As #1
Input #1, str1, str2, str3: Close #1 'Ū���̰� �̧C ���L��

For i = 1 To 12
max(i) = Format(Mid(str1, 6 * i - 5, 5), "00.00")
min(i) = Format(Mid(str2, 6 * i - 5, 5), "00.00")
���L(i) = Format(Mid(str3, 6 * i - 5, 5), "00.00")
Next


For z = 1 To 9 '��X�̰���
For z1 = 1 To 9 - z
If max(z1) > max(z1 + 1) Then
tmp = max(z1): max(z1) = max(z1 + 1): max(z1 + 1) = tmp
End If
Next
Next


For x = 1 To 9 '��X�̧C��
For X1 = 1 To 9 - x
If min(X1) < min(X1 + 1) Then
tmp1 = min(X1): min(X1) = min(X1 + 1): min(X1 + 1) = tmp1
End If
Next
Next



rsv(9) = Format((���L(9) - min(9)) / (max(9) - min(9)) * 100, "00.00")



'
For i = 1 To 12
max(i) = Format(Mid(str1, 6 * i - 5, 5), "00.00")
min(i) = Format(Mid(str2, 6 * i - 5, 5), "00.00")
���L(i) = Format(Mid(str3, 6 * i - 5, 5), "00.00")
Next



For z = 1 To 10 '��X�̰���
For z1 = 1 To 10 - z
If max(z1) > max(z1 + 1) Then
tmp = max(z1): max(z1) = max(z1 + 1): max(z1 + 1) = tmp
End If
Next
Next


For x = 1 To 10 '��X�̧C��
For X1 = 1 To 10 - x
If min(X1) < min(X1 + 1) Then
tmp1 = min(X1): min(X1) = min(X1 + 1): min(X1 + 1) = tmp1
End If
Next
Next

'

rsv(10) = Format((���L(10) - min(10)) / (max(10) - min(10)) * 100, "00.00")

For i = 1 To 12
max(i) = Format(Mid(str1, 6 * i - 5, 5), "00.00")
min(i) = Format(Mid(str2, 6 * i - 5, 5), "00.00")
���L(i) = Format(Mid(str3, 6 * i - 5, 5), "00.00")
Next



For z = 1 To 11 '��X�̰���
For z1 = 1 To 11 - z
If max(z1) > max(z1 + 1) Then
tmp = max(z1): max(z1) = max(z1 + 1): max(z1 + 1) = tmp
End If
Next
Next


For x = 1 To 11 '��X�̧C��
For X1 = 1 To 11 - x
If min(X1) < min(X1 + 1) Then
tmp1 = min(X1): min(X1) = min(X1 + 1): min(X1 + 1) = tmp1
End If
Next
Next



rsv(11) = Format((���L(11) - min(11)) / (max(11) - min(11)) * 100, "00.00")

'
For i = 1 To 12
max(i) = Format(Mid(str1, 6 * i - 5, 5), "00.00")
min(i) = Format(Mid(str2, 6 * i - 5, 5), "00.00")
���L(i) = Format(Mid(str3, 6 * i - 5, 5), "00.00")
Next



For z = 1 To 12 '��X�̰���
For z1 = 1 To 12 - z
If max(z1) > max(z1 + 1) Then
tmp = max(z1): max(z1) = max(z1 + 1): max(z1 + 1) = tmp
End If
Next
Next


For x = 1 To 12 '��X�̧C��
For X1 = 1 To 12 - x
If min(X1) < min(X1 + 1) Then
tmp1 = min(X1): min(X1) = min(X1 + 1): min(X1 + 1) = tmp1
End If
Next
Next



rsv(12) = Format((���L(12) - min(12)) / (max(12) - min(12)) * 100, "00.00")

For j = 9 To 12 '�p��C��KD��
k(j) = Format(2 / 3 * k(j - 1) + 1 / 3 * rsv(j), "00.00")
d(j) = Format(2 / 3 * d(j - 1) + 1 / 3 * k(j), "00.00")
Next

Open Text2.Text For Output As #2
Write #2, k(8), k(9), k(10), k(11), k(12)
Write #2, d(8), d(9), d(10), d(11), d(12)
Close #2


End Sub

Private Sub Command2_Click()
End
End Sub



Private Sub Command3_Click()

CommonDialog1.ShowOpen
file = CommonDialog1.FileName
Dim b() As Byte
Dim c() As Byte


ReDim b(FileLen(file))
ReDim c(FileLen(Text2.Text))
Open file For Binary As #1
Get #1, , b
Close #1


Open Text2.Text For Binary As #2
Get #2, , c
Close #2

If StrConv(b, 64) = StrConv(c, 64) Then
MsgBox "��Ƥ�勵�T", , "�祿"
Else
MsgBox "��Ƥ����~", , "�祿"
End If


End Sub
