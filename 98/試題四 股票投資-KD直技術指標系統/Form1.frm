VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   13155
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   15
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   14
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   13
      Top             =   4560
      Width           =   10215
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   12
      Top             =   3840
      Width           =   10215
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   3120
      Width           =   10215
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   2400
      Width           =   10215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   1800
      Width           =   10215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label Label7 
      Caption         =   "�w����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "�h�ŭ�"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "20�馨���"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "5�饭����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "��X�ɸ��|�W��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "��J�ɸ��|�W��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As String
Dim str(35) As String '�C�Ѧ��L��
Dim b() As Byte
Dim c As String
Dim str1 As String '���L��
Dim str2 As String '5�饭����
Dim str3 As String '5�饭���ȸ��
Dim str4 As String   '20�饭����
Dim str5 As String
Dim str6 As String '�C��h�ŭ�
Dim str7 As String '�h�ŭ�
Dim str8 As String '�C��w����
Dim str9 As String '�w����
a = Text1.Text
ReDim b(FileLen(a))
Open a For Binary As #1
Get #1, , b()


For i = 1 To 34 'Ū��20��H����
str(i) = Format(Mid(StrConv(b, 64), 5 * i - 4, 5), "00.00")
If str(i) = "" Then str(i) = "00.00"
If i >= 20 Then
str1 = str1 & str(i) & " " '20��H����
str8 = Format(1 / 3 * (4 * Val(str(i - 4)) - Val(str(i - 19))), "00.00")
str9 = str9 & str8 & " "
If i <= 30 Then
str2 = (Val(str(i)) + Val(str(i - 1)) + Val(str(i - 2)) + Val(str(i - 3)) + Val(str(i - 4))) / 5 '�p��C5�饭����
'�H�U�p��20������
str4 = Format((Val(str(i)) + Val(str(i - 1)) + Val(str(i - 2)) + Val(str(i - 3)) + Val(str(i - 4)) + Val(str(i - 5)) + Val(str(i - 6)) + Val(str(i - 7)) + Val(str(i - 8)) + Val(str(i - 9)) + Val(str(i - 10)) + Val(str(i - 11)) + Val(str(i - 12)) + Val(str(i - 13)) + Val(str(i - 14)) + Val(str(i - 15)) + Val(str(i - 16)) + Val(str(i - 17)) + Val(str(i - 18)) + Val(str(i - 19))) / 20, "00.00")
str6 = Format(Val(str2) - Val(str4), "00.00")

str7 = str7 & str6 & " "
str3 = str3 & str2 & " " '5�饭���ȸ��
str5 = str5 & str4 & " " '20�饭����
End If
End If
Next i


Text3.Text = str1
Text4.Text = str3
Text5.Text = str5
Text6.Text = str7
Text7.Text = str9
Close #1
'�d��.txt

c = Text2.Text
Open c For Output As #2
Write #2, Text3.Text
Write #2, Text4.Text
Write #2, Text5.Text
Write #2, Text6.Text
Write #2, Text7.Text

Close #2
End Sub

Private Sub Command2_Click()
End
End Sub

