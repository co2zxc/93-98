VERSION 5.00
Begin VB.Form x 
   Caption         =   "�I�M���X�����F����"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14175
   LinkTopic       =   "Form1"
   ScaleHeight     =   5642.458
   ScaleMode       =   0  '�ϥΪ̦ۭq
   ScaleWidth      =   14175
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.TextBox y 
      Height          =   270
      Index           =   8
      Left            =   2400
      TabIndex        =   38
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox x 
      Height          =   270
      Index           =   8
      Left            =   1560
      TabIndex        =   37
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '�S���ؽu
      FillStyle       =   0  '���
      Height          =   4620
      Left            =   8280
      ScaleHeight     =   5833.333
      ScaleMode       =   0  '�ϥΪ̦ۭq
      ScaleWidth      =   5668.693
      TabIndex        =   35
      Top             =   3480
      Visible         =   0   'False
      Width           =   5595
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�e�X�I�����G"
      Height          =   495
      Left            =   8280
      TabIndex        =   34
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   9600
      TabIndex        =   32
      Top             =   2280
      Width           =   3735
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   9600
      TabIndex        =   31
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   9600
      TabIndex        =   30
      Top             =   1080
      Width           =   3735
   End
   Begin VB.CommandButton avg 
      Caption         =   "�D�����Z��"
      Height          =   495
      Left            =   8280
      TabIndex        =   29
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton min 
      Caption         =   "�D�̤p�Z��"
      Height          =   495
      Left            =   8280
      TabIndex        =   28
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton max 
      Caption         =   "�D�̤j�Z��"
      Height          =   495
      Left            =   8280
      TabIndex        =   27
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox y 
      Height          =   270
      Index           =   0
      Left            =   6600
      TabIndex        =   26
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox x 
      Height          =   270
      Index           =   0
      Left            =   5640
      TabIndex        =   25
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox y 
      Height          =   270
      Index           =   7
      Left            =   2400
      TabIndex        =   24
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox y 
      Height          =   270
      Index           =   6
      Left            =   2400
      TabIndex        =   23
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox y 
      Height          =   270
      Index           =   5
      Left            =   2400
      TabIndex        =   22
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox y 
      Height          =   270
      Index           =   4
      Left            =   2400
      TabIndex        =   21
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox y 
      Height          =   270
      Index           =   3
      Left            =   2400
      TabIndex        =   20
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox y 
      Height          =   270
      Index           =   2
      Left            =   2400
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox y 
      Height          =   270
      Index           =   1
      Left            =   2400
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox x 
      Height          =   270
      Index           =   7
      Left            =   1560
      TabIndex        =   17
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox x 
      Height          =   270
      Index           =   6
      Left            =   1560
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox x 
      Height          =   270
      Index           =   5
      Left            =   1560
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox x 
      Height          =   270
      Index           =   4
      Left            =   1560
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox x 
      Height          =   270
      Index           =   3
      Left            =   1560
      TabIndex        =   13
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox x 
      Height          =   270
      Index           =   2
      Left            =   1560
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox x 
      Height          =   270
      Index           =   1
      Left            =   1560
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
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
      Left            =   5760
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "��x8�I�y��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   36
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "�Z��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   33
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "��x7�I�y��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "��x6�I�y��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "��x5�I�y��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "��x4�I�y��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "��x3�I�y��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "��x2�I�y��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "��x1�I�y��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "�п�J�����I���y�� : x�My��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "�п�J6���I�y�� : x�My��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "�п�J���XC�����X���I ( �̦h8���I )�G"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x1(10) As Single
Dim y1(10) As Single
Dim tmp As String
Dim tmp1 As Double
Dim min1(8) As Double
Dim a As Integer



Private Sub avg_Click() '�����Z��

For z = 0 To a
If (x(z).Text < 0 Or x(z).Text > 6) Or (y(z).Text < 1 Or y(z).Text > 6) Then
MsgBox "���XC����J�I�ƿ��~�A���ιw�]��8���I�A�ӹ��X�I�������C�ЦA��J�@��!"
Exit Sub
Else
End If
Next


Dim avg As Double
For i = 1 To a
min1(i) = ((x(0).Text - x(i).Text) ^ 2 + (y(0).Text - y(i).Text) ^ 2) ^ 0.5
Next
For i1 = 1 To a
avg = avg + min1(i1)
Next
Text6.Text = avg / a

End Sub

Private Sub Command1_Click()
Picture1.Cls
Picture1.Visible = True
Picture1.Scale (-2, 8)-(9, -1)
Picture1.Circle (x(0).Text, y(0).Text), 0.1, RGB(0, 0, 0)

For i1 = 1 To a '�e�I
Picture1.Circle (x(i1).Text, y(i1).Text), 0.1, RGB(255, 0, 255)
Next

For i = 0 To 6 Step 0.5
Picture1.Line (i, 0)-(i, 6) '�eX
Picture1.Line (0, i)-(6, i) '�eY
Next

For i2 = 0 To 6
Picture1.CurrentX = i2 - 0.1 ' �eX�b���
Picture1.CurrentY = -0.2
Picture1.Print i2
Picture1.CurrentX = -0.4 '�eY�b���
Picture1.CurrentY = i2 + 0.1
Picture1.Print i2
Next

For i3 = 0 To a '�C�@�I�y��
Picture1.CurrentX = x(i3).Text + 0.1
Picture1.CurrentY = y(i3).Text + 0.4
Picture1.Print "x" & i3
Next

End Sub

Private Sub max_Click() '�̤j�Z��
For z = 0 To a
If (x(z).Text < 0 Or x(z).Text > 6) Or (y(z).Text < 1 Or y(z).Text > 6) Then
MsgBox "���XC����J�I�ƿ��~�A���ιw�]��8���I�A�ӹ��X�I�������C�ЦA��J�@��!"
Exit Sub
Else
End If
Next

a = Text1.Text
Text4.Text = ""
For i = 1 To a
min1(i) = ((x(0).Text - x(i).Text) ^ 2 + (y(0).Text - y(i).Text) ^ 2) ^ 0.5
Next
For i1 = 1 To a
For i2 = 1 To a - 1
If min1(i2) > min1(i2 + 1) Then
tmp1 = min1(i2)
min1(i2) = min1(i2 + 1)
min1(i2 + 1) = tmp1
End If
Next
Next
Text4.Text = min1(a)

End Sub



Private Sub min_Click() '�̵u�Z��
For z = 0 To a
If (x(z).Text < 0 Or x(z).Text > 6) Or (y(z).Text < 1 Or y(z).Text > 6) Then
MsgBox "���XC����J�I�ƿ��~�A���ιw�]��8���I�A�ӹ��X�I�������C�ЦA��J�@��!"
Exit Sub
Else
End If
Next

a = Text1.Text
Text5.Text = ""
For i = 1 To a
min1(i) = ((x(0).Text - x(i).Text) ^ 2 + (y(0).Text - y(i).Text) ^ 2) ^ 0.5
Next
For i1 = 1 To a
For i2 = 1 To a - 1
If min1(i2) > min1(i2 + 1) Then
tmp1 = min1(i2)
min1(i2) = min1(i2 + 1)
min1(i2 + 1) = tmp1
End If
Next
Next
Text5.Text = min1(1)

End Sub



Private Sub Text1_Change()
'��ܿ�J�y��
For z = 0 To a
x(z).Text = ""
y(z).Text = ""
Next

Text4.Text = ""
Text5.Text = ""
Text6.Text = ""


a = Val(Text1.Text)
If Text1.Text <> "" And a > 0 And a < 9 Then
Label2.Visible = True
Label2.Caption = "�п�J" & Str(a) & "���I�y�� : x�My��"
Label3.Visible = True
x(0).Visible = True
y(0).Visible = True
For i = 1 To a
Label4(i).Visible = True
x(i).Visible = True
y(i).Visible = True
Next i
Else

Label2.Visible = False
Label3.Visible = False
x(0).Visible = False
y(0).Visible = False
For i1 = 1 To 8
Label4(i1).Visible = False
x(i1).Visible = False
y(i1).Visible = False
Next i1
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 49 Or KeyAscii > 56 Then KeyAscii = 0

End Sub


