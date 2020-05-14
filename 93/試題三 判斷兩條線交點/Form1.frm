VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "試題三"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   3555
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "輸出"
      Height          =   495
      Left            =   1200
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   1200
      TabIndex        =   9
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Text            =   "4 1"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Text            =   "3 3"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Text            =   "1 1"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "2 4"
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "輸出格式:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "D(X4,Y4)="
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "C(X3,Y3)="
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "B(X2,Y2)="
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "A(X1,Y1)="
      Height          =   255
      Left            =   480
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
a1 = Split(Text1.Text)
a2 = Split(Text2.Text)
a3 = Split(Text3.Text)
a4 = Split(Text4.Text)
Dim x1 As Integer, x2 As Integer, x3 As Integer, x4 As Integer, y1 As Integer, y2 As Integer, y3 As Integer, y4 As Integer

x1 = a1(0)
y1 = a1(1)

x2 = a2(0)
y2 = a2(1)

x3 = a3(0)
y3 = a3(1)

x4 = a4(0)
y4 = a4(1)

If x1 - x2 = 0 Then
    e = 1
Else
    m1 = (y1 - y2) / (x1 - x2)
End If
    
If x3 - x4 = 0 Then
    e = 2
Else
    m2 = (y3 - y4) / (x3 - x4)
End If
    
c1 = y1 - m1 * x1
c2 = y3 - m2 * x3

s = -m1 - (-m2)
sx = c1 - c2
sy = c2 - c1

If s <> 0 Then
x = sx / s
y = sy / s
End If

If e = 1 Then
    If (x3 = x1 Or x4 = x1) And (y1 <= y3 And y3 <= y2) Or (y2 <= y3 And y3 <= y1) Or (y1 <= y4 And y4 <= y2) Or (y2 <= y4 And y4 <= y1) Then Text5.Text = "線有相交"
ElseIf e = 2 Then
    If (x1 = x3 Or x2 = x3) And (y3 <= y1 And y1 <= y4) Or (y4 <= y1 And y1 <= y3) Or (y3 <= y2 And y2 <= y4) Or (y4 <= y2 And y2 <= y3) Then Text5.Text = "線有相交"
    Else
        If ((x1 <= x And x <= x2) Or (x2 <= x And x <= x1)) And ((x3 <= x And x <= x4) Or (x4 <= x And x <= x3)) Then
            Text5.Text = "線有相交"
        Else
            Text5.Text = "線無相交"
        End If
End If
End Sub

