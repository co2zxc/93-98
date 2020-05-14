VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "資料匯流排傳輸系統"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   9870
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
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
      Index           =   4
      Left            =   7800
      TabIndex        =   15
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Index           =   3
      Left            =   840
      TabIndex        =   14
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Index           =   2
      Left            =   7800
      TabIndex        =   13
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Index           =   1
      Left            =   840
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   11
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Tansmit"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   10
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Random Set"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   9
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Line Line17 
      X1              =   5160
      X2              =   5400
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line16 
      X1              =   4800
      X2              =   5400
      Y1              =   3120
      Y2              =   2880
   End
   Begin VB.Line Line15 
      X1              =   4200
      X2              =   4440
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line14 
      X1              =   4200
      X2              =   4800
      Y1              =   2880
      Y2              =   3120
   End
   Begin VB.Line Line13 
      X1              =   5160
      X2              =   5400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line12 
      X1              =   4800
      X2              =   5400
      Y1              =   1800
      Y2              =   2040
   End
   Begin VB.Line Line11 
      X1              =   5160
      X2              =   5160
      Y1              =   2040
      Y2              =   2880
   End
   Begin VB.Line Line10 
      X1              =   4200
      X2              =   4440
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line9 
      X1              =   4200
      X2              =   4800
      Y1              =   2040
      Y2              =   1800
   End
   Begin VB.Line Line8 
      X1              =   4440
      X2              =   4440
      Y1              =   2040
      Y2              =   2880
   End
   Begin VB.Line Line7 
      Index           =   1
      X1              =   5640
      X2              =   6120
      Y1              =   3000
      Y2              =   3480
   End
   Begin VB.Line Line5 
      Index           =   7
      X1              =   5640
      X2              =   5640
      Y1              =   3720
      Y2              =   3960
   End
   Begin VB.Line Line5 
      Index           =   6
      X1              =   5640
      X2              =   5640
      Y1              =   3000
      Y2              =   3240
   End
   Begin VB.Line Line5 
      Index           =   5
      X1              =   3840
      X2              =   3840
      Y1              =   3720
      Y2              =   3960
   End
   Begin VB.Line Line6 
      Index           =   1
      X1              =   5640
      X2              =   6120
      Y1              =   3960
      Y2              =   3480
   End
   Begin VB.Line Line5 
      Index           =   4
      X1              =   3840
      X2              =   3840
      Y1              =   3000
      Y2              =   3240
   End
   Begin VB.Line Line4 
      Index           =   1
      X1              =   3840
      X2              =   5640
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   3480
      X2              =   3840
      Y1              =   3480
      Y2              =   3960
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   3480
      X2              =   3840
      Y1              =   3480
      Y2              =   3000
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   3840
      X2              =   5640
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line7 
      Index           =   0
      X1              =   5640
      X2              =   6120
      Y1              =   960
      Y2              =   1440
   End
   Begin VB.Line Line5 
      Index           =   3
      X1              =   5640
      X2              =   5640
      Y1              =   1680
      Y2              =   1920
   End
   Begin VB.Line Line5 
      Index           =   2
      X1              =   5640
      X2              =   5640
      Y1              =   960
      Y2              =   1200
   End
   Begin VB.Line Line5 
      Index           =   1
      X1              =   3840
      X2              =   3840
      Y1              =   1680
      Y2              =   1920
   End
   Begin VB.Line Line6 
      Index           =   0
      X1              =   5640
      X2              =   6120
      Y1              =   1920
      Y2              =   1440
   End
   Begin VB.Line Line5 
      Index           =   0
      X1              =   3840
      X2              =   3840
      Y1              =   960
      Y2              =   1200
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   3840
      X2              =   5640
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   3480
      X2              =   3840
      Y1              =   1440
      Y2              =   1920
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   3480
      X2              =   3840
      Y1              =   1440
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   3840
      X2              =   5640
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label9 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label8 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label7 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   4
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   2
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Data Bus Transmission System"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)

Select Case Index '

Case 1:
  If Command1(1).Caption = "Ih" Then
    '當為IH則變成LD
    Command1(1).Caption = "Ld"
    '當為LD則先看其他是否都不是EN才變成EN
  ElseIf Command1(1).Caption = "Ld" And Command1(2).Caption <> "En" And Command1(3).Caption <> "En" And Command1(4).Caption <> "En" Then
    Command1(1).Caption = "En"
    '否則變成IH
  ElseIf Command1(1).Caption = "Ld" And (Command1(2).Caption = "En" Or Command1(3).Caption = "En" Or Command1(4).Caption = "En") Then
    Command1(1).Caption = "Ih"
    '當為EN變成IH
  ElseIf Command1(1).Caption = "En" Then
    Command1(1).Caption = "Ih"
  End If
Case 2: '以下皆比照第一個
  If Command1(2).Caption = "Ih" Then
    Command1(2).Caption = "Ld"
  ElseIf Command1(2).Caption = "Ld" And Command1(1).Caption <> "En" And Command1(3).Caption <> "En" And Command1(4).Caption <> "En" Then
    Command1(2).Caption = "En"
  ElseIf Command1(2).Caption = "Ld" And (Command1(1).Caption = "En" Or Command1(3).Caption = "En" Or Command1(4).Caption = "En") Then
    Command1(2).Caption = "Ih"
  ElseIf Command1(2).Caption = "En" Then
    Command1(2).Caption = "Ih"
  End If
Case 3:
  If Command1(3).Caption = "Ih" Then
    Command1(3).Caption = "Ld"
  ElseIf Command1(3).Caption = "Ld" And Command1(1).Caption <> "En" And Command1(2).Caption <> "En" And Command1(4).Caption <> "En" Then
    Command1(3).Caption = "En"
  ElseIf Command1(3).Caption = "Ld" And (Command1(2).Caption = "En" Or Command1(1).Caption = "En" Or Command1(4).Caption = "En") Then
    Command1(3).Caption = "Ih"
  ElseIf Command1(3).Caption = "En" Then
    Command1(3).Caption = "Ih"
  End If
Case 4:
  If Command1(4).Caption = "Ih" Then
    Command1(4).Caption = "Ld"
  ElseIf Command1(4).Caption = "Ld" And Command1(1).Caption <> "En" And Command1(3).Caption <> "En" And Command1(2).Caption <> "En" Then
    Command1(4).Caption = "En"
  ElseIf Command1(4).Caption = "Ld" And (Command1(2).Caption = "En" Or Command1(3).Caption = "En" Or Command1(1).Caption = "En") Then
    Command1(4).Caption = "Ih"
  ElseIf Command1(4).Caption = "En" Then
    Command1(4).Caption = "Ih"
  End If
End Select

End Sub

Private Sub Command5_Click()
Command1(1).Caption = "Ih"
Command1(2).Caption = "Ih"
Command1(3).Caption = "Ih"
Command1(4).Caption = "Ih"
Randomize
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
For i = 1 To 8
a = Int(Rnd * 2)
Label6.Caption = a & Label6.Caption
b = Int(Rnd * 2)
Label7.Caption = b & Label7.Caption
C = Int(Rnd * 2)
Label8.Caption = C & Label8.Caption
d = Int(Rnd * 2)
Label9.Caption = d & Label9.Caption
Next i
End Sub

Private Sub Command6_Click()

If Command1(1).Caption = "En" Then
If Command1(2).Caption = "Ld" Then Label7.Caption = Label6.Caption
If Command1(3).Caption = "Ld" Then Label8.Caption = Label6.Caption
If Command1(4).Caption = "Ld" Then Label9.Caption = Label6.Caption
ElseIf Command1(2).Caption = "En" Then
If Command1(1).Caption = "Ld" Then Label6.Caption = Label7.Caption
If Command1(3).Caption = "Ld" Then Label8.Caption = Label7.Caption
If Command1(4).Caption = "Ld" Then Label9.Caption = Label7.Caption
ElseIf Command1(3).Caption = "En" Then
If Command1(1).Caption = "Ld" Then Label6.Caption = Label8.Caption
If Command1(2).Caption = "Ld" Then Label7.Caption = Label8.Caption
If Command1(4).Caption = "Ld" Then Label9.Caption = Label8.Caption
ElseIf Command1(4).Caption = "En" Then
If Command1(1).Caption = "Ld" Then Label6.Caption = Label9.Caption
If Command1(3).Caption = "Ld" Then Label8.Caption = Label9.Caption
If Command1(2).Caption = "Ld" Then Label7.Caption = Label9.Caption
End If


End Sub

Private Sub Command7_Click()
End
End Sub
