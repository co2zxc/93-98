VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "環佇列運作系統"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   8325
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  '平面
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2160
      Width           =   3735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Message:"
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
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Circular Queue"
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
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim e As Integer, fr As Integer, re As Integer, now As Integer
Private Sub Command1_Click()

For i = 0 To now '先判斷是否只剩下一格
If Text1(i).Text = "" Then
s = s + 1
End If
Next

If s = 1 Then '如只剩下一格即增加記憶體位置
Last = now + 6
For i = now + 1 To Last
Load Text1(i)
With Text1(i)
.Left = 360 + i * 495
.Top = 960
.Visible = True
.Text = ""
End With
Next

For i1 = 0 To now - 1 '整理原有資料
If Text1(now - i1).Text = "" Then Exit For
Text1(Last - i1).Text = Text1(now - i1).Text
Next

For i = now - i1 To Last - i1
Text1(i).Text = ""
Next
If re > now Then re = 0
s = 6 '數字隨意給 為了能執行下面程式
now = Last
End If

If s > 1 Then '增加資料到記憶體位置
If re > now Then re = 0
c = Int(Rnd * 999) + 1
Text1(re).Text = c
re = re + 1
Text2.Text = "Added    " & c
End If

End Sub

Private Sub Command2_Click()

For i = 0 To now - 1 '找到佇列中FRONT端
If Text1(i).Text = "" And Text1(i + 1) <> "" Then
fr = i + 1
Exit For
Else
fr = 0
End If
Next

If fr > now Then fr = 0
If Text1(fr).Text = "" Then
Text2.Text = "Queue is empty"
Else
c = Text1(fr).Text
Text1(fr).Text = ""
fr = fr + 1
Text2.Text = "Removed   " & c
End If

End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Activate()
now = 5
Randomize
For i = 1 To 5
Load Text1(i)
With Text1(i)
.Left = 360 + i * 495
.Top = 960
.Visible = True
End With
Next

a = Int(Rnd * 6)
b = Int(Rnd * 999) + 1
b1 = Int(Rnd * 999) + 1

If a < 5 Then
Text1(a).Text = b
Text1(a + 1).Text = b1
e = 0
Else
Text1(a).Text = b
Text1(0).Text = b1
e = 1
End If

If e = 0 Then re = a + 2: fr = a
If e = 1 Then re = 1: fr = a



End Sub

