VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "解一元二次方程式"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   6060
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   1560
      TabIndex        =   19
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   1560
      TabIndex        =   18
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1560
      TabIndex        =   17
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4320
      TabIndex        =   16
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   4320
      TabIndex        =   15
      Top             =   4545
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "求解"
      Height          =   495
      Left            =   4440
      TabIndex        =   13
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "X="
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "Function Here"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Label Label10 
      Caption         =   "Input C="
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
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Input B="
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
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Input A="
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
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Line Line9 
      BorderWidth     =   3
      X1              =   120
      X2              =   6000
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   840
      X2              =   3600
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label9 
      Caption         =   "2A"
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
      Left            =   1920
      TabIndex        =   8
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "-4AC"
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
      Left            =   2640
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Line Line8 
      X1              =   2160
      X2              =   3480
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label7 
      Caption         =   "2"
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
      Left            =   2400
      TabIndex        =   6
      Top             =   1200
      Width           =   255
   End
   Begin VB.Line Line7 
      X1              =   1920
      X2              =   2160
      Y1              =   1680
      Y2              =   1200
   End
   Begin VB.Line Line6 
      X1              =   1800
      X2              =   1920
      Y1              =   1440
      Y2              =   1680
   End
   Begin VB.Line Line5 
      X1              =   1680
      X2              =   1800
      Y1              =   1560
      Y2              =   1440
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   1440
      X2              =   1680
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   1560
      X2              =   1560
      Y1              =   1320
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   1440
      X2              =   1680
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label2 
      Caption         =   "x="
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
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "2"
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
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "+Bx+C=0"
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
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Ax"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "-B"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label6 
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
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1320
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = Text3.Text
b = Text4.Text
c = Text5.Text

If a = 0 And b = 0 And c <> 0 Then Text1.Text = "無解": d = 0
If a = 0 And b = 0 And c = 0 Then Text1.Text = "無限多解": d = 0

If a = 0 And b <> 0 Then x = -c / b: d = 1
If a <> 0 And b <> 0 And b ^ 2 - 4 * a * c = 0 Then x = -b / (2 * a): d = 2
If a <> 0 And b <> 0 And b ^ 2 - 4 * a * c > 0 Then x = (-b + (b ^ 2 - 4 * a * c) ^ 0.5) / (2 * a): X1 = (-b - (b ^ 2 - 4 * a * c) ^ 0.5) / (2 * a): d = 3
If a <> 0 And b <> 0 And b ^ 2 - 4 * a * c < 0 Then x = Round(-b / (2 * a), 2) & "+" & Round(Sqr(4 * a * c - b ^ 2) / (2 * a), 2) & "i": X1 = Round(-b / (2 * a), 2) & "-" & Round(Sqr(4 * a * c - b ^ 2) / (2 * a), 2) & "i": d = 4

If d = 0 Then Text2.Text = ""
If d = 1 Then Text1.Text = x: Text2.Text = "只有一解"
If d = 2 Then Text1.Text = x: Text2.Text = "同根"
If d = 3 Then Text1.Text = x: Text2.Text = X1
If d = 4 Then Text1.Text = x: Text2.Text = X1


End Sub

