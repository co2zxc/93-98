VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "�H�Ѱ����k"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   5520
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.TextBox Text1 
      Height          =   4575
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  '�������b
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim e, x, x1, x2, x3, y, y1, y2, y3, chkx, chky As Integer

Private Sub Form_Activate()
e = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If e = 0 Then '�Ĥ@����J�����y�лP��ê��
  x = Val(Chr(KeyAscii)): e = 1
ElseIf e = 1 Then
  y = Val(Chr(KeyAscii)): e = 2
ElseIf e = 2 Then
  x1 = Val(Chr(KeyAscii)): e = 3
ElseIf e = 3 Then
  y1 = Val(Chr(KeyAscii)): e = 4
ElseIf e = 4 Then
  x2 = Val(Chr(KeyAscii)): e = 5
ElseIf e = 5 Then
  y2 = Val(Chr(KeyAscii)): e = 6
ElseIf e = 6 Then
  x3 = Val(Chr(KeyAscii)): e = 7
ElseIf e = 7 Then
  y3 = Val(Chr(KeyAscii)): e = 8
ElseIf e = 8 Then
  Text1.Locked = True
  a = Val(Chr(KeyAscii))
  If a = 9 Then Text1.Text = Text1.Text & vbCrLf & "��J���ʼƦr��G" & a & "(�������{��)": Exit Sub
  Call cover(x, y, x1, y1): Call cover(x, y, x2, y2): Call cover(x, y, x3, y3)
  
  If out(x, y, a) = True Then
    Text1.Text = Text1.Text & vbCrLf & "��J���ʼƦr��G" & a
    Text1.Text = Text1.Text & vbCrLf & "�����ʦܷs��m�G" & x & " " & y & "(�]�W�X�ѽL�~�ӫO���쪬)"
  ElseIf bump(x, y, x1, y1, a) = True Or bump(x, y, x2, y2, a) = True Or bump(x, y, x3, y3, a) = True Then
    Text1.Text = Text1.Text & vbCrLf & "��J���ʼƦr��G" & a
    Text1.Text = Text1.Text & vbCrLf & "�����ʦܷs��m�G" & x & " " & y & "(�]�����}�ӫO���쪬)"
  Else
      Call tomove(x, y, a)
      Text1.Text = Text1.Text & vbCrLf & "��J���ʼƦr��G" & a
      Text1.Text = Text1.Text & vbCrLf & "�����ʦܷs��m�G" & x & " " & y
  End If
End If



End Sub

Sub tomove(x, y, a) '���ʨ��������m
Select Case a
Case 0:
y = y + 2: x = x + 1
Case 1:
y = y + 2: x = x - 1
Case 2:
y = y + 1: x = x - 2
Case 3:
y = y - 1: x = x - 2
Case 4:
y = y - 2: x = x - 1
Case 5:
y = y - 2: x = x + 1
Case 6:
y = y - 1: x = x + 2
Case 7:
y = y + 1: x = x + 2
End Select
End Sub

Function bump(ByVal x, ByVal y, ByVal x1, ByVal y1, a) As Boolean '�P�_�O�_�䰨�}�A�Q�ΰ������V�P�_
bump = False
Select Case a
Case 0:
  If y1 = y + 1 And x1 = x Then
    bump = True
  Else
    y = y + 2: x = x + 1
  End If
Case 1:
  If y1 = y + 1 And x1 = x Then
    bump = True
  Else
    y = y + 2: x = x - 1
  End If
Case 2:
  If x1 = x - 1 And y1 = y Then
    bump = True
  Else
    y = y + 1: x = x - 2
  End If
Case 3:
  If x1 = x - 1 And y1 = y Then
    bump = True
  Else
    y = y - 1: x = x - 2
  End If
Case 4:
  If y1 = y - 1 And x1 = x Then
    bump = True
  Else
    y = y - 2: x = x - 1
  End If
Case 5:
  If y1 = y - 1 And x1 = x Then
    bump = True
  Else
    y = y - 2: x = x + 1
  End If
Case 6:
  If x1 = x + 1 And y1 = y Then
    bump = True
  Else
    y = y - 1: x = x + 2
  End If
Case 7:
  If x1 = x + 1 And y1 = y Then
    bump = True
  Else
    y = y + 1: x = x + 2
  End If
End Select
End Function

Function out(ByVal x, ByVal y, a) As Boolean '�P�_�O�_���X�ѽL
out = False
Select Case a
Case 0:
y = y + 2: x = x + 1
Case 1:
y = y + 2: x = x - 1
Case 2:
y = y + 1: x = x - 2
Case 3:
y = y - 1: x = x - 2
Case 4:
y = y - 2: x = x - 1
Case 5:
y = y - 2: x = x + 1
Case 6:
y = y - 1: x = x + 2
Case 7:
y = y + 1: x = x + 2
End Select
If x > 8 Or y > 8 Or x < 1 Or y < 1 Then out = True
End Function

Sub cover(x, y, x1, y1) '�P�_�O�_����ê���W
If x = x1 And y = y1 Then
x1 = 0: y1 = 0
End If
End Sub
