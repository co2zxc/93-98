VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "虏璸衡诀"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   7095
   StartUpPosition =   3  '╰参箇砞
   Begin VB.CommandButton Command4 
      Caption         =   "."
      Height          =   495
      Left            =   3000
      TabIndex        =   19
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "="
      Height          =   495
      Index           =   6
      Left            =   4320
      TabIndex        =   18
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "/"
      Height          =   495
      Index           =   5
      Left            =   5640
      TabIndex        =   17
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   495
      Index           =   4
      Left            =   4320
      TabIndex        =   16
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "X"
      Height          =   495
      Index           =   3
      Left            =   5640
      TabIndex        =   15
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+"
      Height          =   495
      Index           =   2
      Left            =   4320
      TabIndex        =   14
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ac"
      Height          =   495
      Index           =   1
      Left            =   5640
      TabIndex        =   13
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Log"
      Height          =   495
      Index           =   0
      Left            =   4320
      TabIndex        =   12
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+/-"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      Height          =   495
      Index           =   9
      Left            =   3000
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      Height          =   495
      Index           =   8
      Left            =   1680
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      Height          =   495
      Index           =   7
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      Height          =   495
      Index           =   6
      Left            =   3000
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      Height          =   495
      Index           =   5
      Left            =   1680
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      Height          =   495
      Index           =   4
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      Height          =   495
      Index           =   3
      Left            =   3000
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      Height          =   495
      Index           =   2
      Left            =   1680
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '綼癸霍
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Dim a As Integer
Dim b As String
Dim c As String
Dim d As String
Private Sub Command1_Click(Index As Integer)
If x = 1 Then Text1.Text = ""
Select Case Index
Case 0
a = 0
Case 1
a = 1
Case 2
a = 2
Case 3
a = 3
Case 4
a = 4
Case 5
a = 5
Case 6
a = 6
Case 7
a = 7
Case 8
a = 8
Case 9
a = 9
End Select
x = 0
Text1.Text = Text1.Text & a

End Sub

Private Sub Command2_Click()
If Val(Text1.Text) > 0 Then
Text1.Text = "-" & Text1.Text
Else
Text1.Text = Abs(Val(Text1.Text))
End If
End Sub

Private Sub Command3_Click(Index As Integer)

Select Case Index
Case 0
If Text1.Text = 0 Then
Text1.Text = 1
Else
Text1.Text = Log(Text1.Text) / Log(10)
End If
Case 1
Text1.Text = "0"
Case 2
b = "+": c = Text1.Text
Case 3
b = "*": c = Text1.Text
Case 4
b = "-": c = Text1.Text
Case 5
b = "/": c = Text1.Text
Case 6

If b = "+" Then
Text1.Text = mAdd(c, Text1.Text)
ElseIf b = "-" Then
Text1.Text = mSubt(c, Text1.Text)
ElseIf b = "*" Then
Text1.Text = mMult(c, Text1.Text)
ElseIf b = "/" Then
Text1.Text = mDiv(c, Text1.Text)
End If
End Select
x = 1

End Sub

Function mAdd(ByVal strX As String, ByVal strY As String) As String '计猭
Dim x1, y1, z1(), a1, b1
b1 = IIf(Len(strX) >= Len(strY), Len(strX) + 1, Len(strY) + 1)
x1 = StrReverse(strX): y1 = StrReverse(strY)
ReDim z1(b1)
mAdd = ""
For a1 = 1 To b1 - 1
z1(a1) = z1(a1) + Val(Mid(x1, a1, 1)) + Val(Mid(y1, a1, 1))
z1(a1 + 1) = z1(a1) \ 10
z1(a1) = z1(a1) Mod 10
mAdd = Trim(Str(z1(a1))) & mAdd
Next
If z1(b1) <> 0 Then
mAdd = Trim(Str(z1(b1))) & mAdd
End If
End Function


Function mSubt(ByVal strX As String, ByVal strY As String) As String  '计搭猭
Dim x, y, z(), a, b
b = Len(strX): ReDim z(b)
x = StrReverse(strX): y = StrReverse(strY)
mSubt = ""
For a = 1 To b
z(a) = z(a) + Val(Mid(x, a, 1)) - Val(Mid(y, a, 1))
If z(a) < 0 Then
z(a + 1) = z(a + 1) - 1
z(a) = z(a) + 10
End If
mSubt = Trim(Str(z(a))) & mSubt
Next
For a = 1 To Len(mSubt)
b = a
If Mid(mSubt, a, 1) <> "0" Then Exit For
Next
mSubt = Mid(mSubt, b, Len(mSubt) - b + 1)
End Function


Function mMult(ByVal strX As String, ByVal strY As String) As String '计猭
Dim x(), y(), z()
Dim a, b, d, i, j, k
ReDim x(Len(strX)): ReDim y(Len(strY))
b = 1
For a = UBound(x) To 1 Step -1
x(b) = Val(Mid(strX, a, 1))
b = b + 1
Next
b = 1
For a = UBound(y) To 1 Step -1
y(b) = Val(Mid(strY, a, 1))
b = b + 1
Next
ReDim z(Len(strX) + Len(strY))
For a = 1 To UBound(x)
i = a
For b = 1 To UBound(y)
d = z(i) + x(a) * y(b)
z(i) = (d Mod 10)
z(i + 1) = z(i + 1) + (d \ 10)
i = i + 1
Next
Next
mMult = ""
k = i + 100
If k > UBound(z) Then k = UBound(z)
For a = k To 1 Step -1
b = a
If z(a) <> 0 Then Exit For
Next
For a = 1 To b
mMult = Trim(Str(z(a))) + mMult
Next
 End Function


Function mDiv(ByVal strX As String, ByVal strY As String) As String '计埃猭
Dim x, y, i, j, m, n
mDiv = "0": x = strX: y = strY
Do Until (x = y) Or (isBig(y, x))
m = Len(x) - Len(y)
i = y & String(m, "0")
If m > 0 Then
j = y & String(m - 1, "0")
End If
If isBig(x, i) Then
x = mSubt(x, i)
mDiv = mAdd(mDiv, mPow("10", Trim(Str(m))))
Else
x = mSubt(x, j)
mDiv = mAdd(mDiv, mPow("10", Trim(Str(m - 1))))
End If
Loop
If x = y Then
mDiv = mAdd(mDiv, 1)
End If
 End Function
 
Function mPow(ByVal strX As String, ByVal strY As String) As String '计Ωよ
Dim a As String
mPow = "1"
If strY = "0" Then
Exit Function
End If
Do Until (a = strY) Or isBig(a, strY)
mPow = mMult(mPow, strX)
a = mAdd(a, "1")
Loop
End Function


Function isBig(ByVal strA As String, ByVal strB As String) As Boolean '计ゑ
isBig = False
If strA = strB Then Exit Function
Dim a, b, c
If Len(strB) > Len(strA) Then Exit Function
If Len(strB) < Len(strA) Then
isBig = True
Exit Function
End If
For a = 1 To Len(strA)
b = Val(Mid(strA, a, 1))
c = Val(Mid(strB, a, 1))
If b < c Then
isBig = False
Exit Function
End If
If b > c Then
isBig = True
Exit Function
End If
Next
isBig = True
 End Function
 
Private Sub Command4_Click()
Text1.Text = Text1.Text & "."
End Sub

Private Sub Form_Load()
Text1.Text = "0"
x = 1
End Sub
