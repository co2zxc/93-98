VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "計算次方及餘數"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   7800
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "計算"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "餘數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "c"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   720
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As Long, b As Long, c As Long
a1 = Text1.Text
b1 = Text2.Text
c1 = Text3.Text
e = 0

a = Val(a1)
b = Val(b1)
c = Val(c1)
Call test(a1, e) '檢察輸入是否為數字
If e = 1 Then Exit Sub
Call test(b1, e)
If e = 1 Then Exit Sub
Call test(c1, e)
If e = 1 Then Exit Sub



While b > 0 '將指數轉為二進制
g = b Mod 2 & g
b = b \ 2
Wend

s = a
For i = 2 To Len(g)
s = s ^ 2
While s > c '取餘數
s = s - c
Wend

If Mid(g, i, 1) = 1 Then
s = s * a
While s > c '取餘數
s = s - c
Wend
End If
Next

Text4.Text = s


End Sub

Sub test(a, e)

For i = 1 To Len(a)
If Asc((Mid(a, i, 1))) < 48 Or Asc(Mid(a, i, 1)) > 57 Then
MsgBox "請輸入正確數字", , "錯誤"
e = 1
End If
Next

End Sub
