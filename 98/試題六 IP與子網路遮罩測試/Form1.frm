VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   5310
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   3720
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   2535
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   4695
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "Form1.frx":0000
      Left            =   3840
      List            =   "Form1.frx":0064
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form1.frx":00DF
      Left            =   360
      List            =   "Form1.frx":00EC
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Mask"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Net"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Open App.Path & "/test1.txt" For Input As #1
Dim a(4), b(4)

Call bb(b(), Combo1.Text)

Text1.Text = ""

    Do While Not EOF(1)
        Line Input #1, input1
        check = Left(input1, 1)
        input1 = Right(input1, Len(input1) - 2)
        
        For i = 1 To 4
            a(i) = ""
        Next
        
        mask1 = mask(Combo2.List(Combo2.ListIndex)) '�N�B�n�ର�G�i��
        Call aa(a(), input1) '�N��JIP��J�}�C
        c = "": d = ""

        If Left(check, 1) = "C" And (Val(a(1)) < 191 Or Val(a(1)) > 223) Then '�P�_��JIP�O�_���T
            If Len(a(1)) = 3 Then
            input1 = b(1) & Mid(input1, 4, Len(input1) - 3)
            a(1) = b(1)
            ElseIf Len(a(1)) = 2 Then
            input1 = b(1) & Mid(input1, 3, Len(input1) - 2)
            a(1) = b(1)
            End If
        ElseIf Left(check, 1) = "B" And (Val(a(1)) > 191 Or Val(a(1)) < 128) Then
            If Len(a(1)) = 3 Then
                input1 = b(1) & Mid(input1, 4, Len(input1) - 3)
                a(1) = b(1)
            ElseIf Len(a(1)) = 2 Then
                input1 = b(1) & Mid(input1, 3, Len(input1) - 2)
                a(1) = b(1)
            End If
        End If
        
        For i = 1 To 4
            c = c & add(bin(a(i))) '�ܼ�C����JIP�ର�G�i��
            d = d & add(bin(b(i))) '�ܼ�D���Ѧ�IP�ର�G�i��
        Next

        ouput = "": ouput1 = ""

        For i = 1 To Len(mask1)
            If Mid(mask1, i, 1) = 1 And Mid(c, i, 1) = 1 Then '�N�Ѧ�IP�B�B�n��AND�B��
                ouput = ouput & 1
            Else
                ouput = ouput & 0
            End If
            If Mid(mask1, i, 1) = 1 And Mid(d, i, 1) = 1 Then '�NŪ��IP�B�B�n��AND�B��
                ouput1 = ouput1 & 1
            Else
                ouput1 = ouput1 & 0
            End If
        Next
        '���G�ۦP�h��XIP
        If ouput = ouput1 Then Text1.Text = Text1.Text & "IP:" & input1 & vbCrLf

    Loop

Close #1
End Sub

Private Sub Form_Load()
Combo1.ListIndex = 0
Combo2.ListIndex = 23
End Sub


Function mask(a As Integer) As String '�N�B�n�ର�G�i��
For i = 1 To a
mask = mask & 1
Next
For i = 1 To 32 - a
mask = mask & 0
Next
End Function

Sub aa(a(), ByVal b As String) '�NŪ��IP��J��J�}�C
e = 1
For i = 1 To Len(b)
If IsNumeric(Mid(b, i, 1)) Then
a(e) = a(e) & Mid(b, i, 1)
Else
If e = 4 Then Exit For
e = e + 1
End If
Next
End Sub

Sub bb(a(), ByVal b As String) '�N�Ѧ�IP��J�}�C
e = 1
For i = 1 To Len(b)
If IsNumeric(Mid(b, i, 1)) Then
a(e) = a(e) & Mid(b, i, 1)
Else
If e = 4 Then Exit For
e = e + 1
End If
Next
End Sub

Function bin(ByVal a As Integer) As String '�NIP�ର�G�i��
Do While a > 0
bin = a Mod 2 & bin
a = a \ 2
Loop
End Function

Function add(a As String) As String '�NIP�ɨ�32�줸
a = Trim(a)
add = a
For i = 1 To 8 - Len(a)
add = 0 & add
Next
End Function


