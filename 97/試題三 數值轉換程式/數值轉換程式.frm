VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "數值轉換程式"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   6870
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convert"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   2280
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "24-bit Binary"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Real number"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "Number System Conversion"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = Text1.Text
b = Abs(a) - Int(Abs(a))

If a >= 0 Then

While a > 0
m = Int(Abs(a)) Mod 2
If m = 1 Then
str1 = 1 & str1
Else
str1 = 0 & str1
End If
a = Int(Abs(a)) \ 2
Wend

While b > 0
m = b * 2
If m >= 1 Then
str2 = "1" & str2
b = m - 1
Else
str2 = "0" & str2
b = m
End If
Wend

If Len(str1) > 15 Or Len(str2) > 9 Then
Text2.Text = "overflow"
Else
Text2.Text = Format(str1 & "." & str2, "0000000000000000.000000000")
End If


Else

While a < 0
m = Int(Abs(a)) Mod 2
If m = 1 Then
str1 = 1 & str1
Else
str1 = 0 & str1
End If
a = Fix(a) \ 2
Wend

While b > 0
m = b * 2
If m >= 1 Then
str2 = "1" & str2
b = m - 1
Else
str2 = "0" & str2
b = m
End If
Wend

If Len(str1) > 15 Or Len(str2) > 9 Then
Text2.Text = "overflow"
Else
Text2.Text = "1" & Format(str1 & "." & str2, "000000000000000.000000000")
End If
End If

End Sub

Private Sub Command2_Click()
End
End Sub

