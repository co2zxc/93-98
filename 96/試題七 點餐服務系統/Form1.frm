VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "西式點餐-點餐服務系統"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   10995
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "總金額"
      Height          =   495
      Left            =   3360
      TabIndex        =   66
      Top             =   9000
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "數量清除"
      Height          =   495
      Left            =   720
      TabIndex        =   65
      Top             =   9000
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "冷飲"
      Height          =   2655
      Index           =   6
      Left            =   5520
      TabIndex        =   56
      Top             =   6000
      Width           =   5175
      Begin VB.VScrollBar VScroll7 
         Height          =   255
         Index           =   4
         Left            =   4440
         Max             =   0
         Min             =   20
         TabIndex        =   178
         Top             =   2040
         Width           =   255
      End
      Begin VB.VScrollBar VScroll7 
         Height          =   255
         Index           =   3
         Left            =   4440
         Max             =   0
         Min             =   20
         TabIndex        =   177
         Top             =   1680
         Width           =   255
      End
      Begin VB.VScrollBar VScroll7 
         Height          =   255
         Index           =   2
         Left            =   4440
         Max             =   0
         Min             =   20
         TabIndex        =   176
         Top             =   1320
         Width           =   255
      End
      Begin VB.VScrollBar VScroll7 
         Height          =   255
         Index           =   1
         Left            =   4440
         Max             =   0
         Min             =   20
         TabIndex        =   175
         Top             =   960
         Width           =   255
      End
      Begin VB.VScrollBar VScroll7 
         Height          =   255
         Index           =   0
         Left            =   4440
         Max             =   0
         Min             =   20
         TabIndex        =   174
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox Text14 
         Height          =   270
         Index           =   4
         Left            =   3480
         TabIndex        =   173
         Text            =   "0"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text14 
         Height          =   270
         Index           =   3
         Left            =   3480
         TabIndex        =   172
         Text            =   "0"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text14 
         Height          =   270
         Index           =   2
         Left            =   3480
         TabIndex        =   171
         Text            =   "0"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text14 
         Height          =   270
         Index           =   1
         Left            =   3480
         TabIndex        =   170
         Text            =   "0"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text14 
         Height          =   270
         Index           =   0
         Left            =   3480
         TabIndex        =   169
         Text            =   "0"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text13 
         Height          =   270
         Index           =   4
         Left            =   2040
         TabIndex        =   168
         Text            =   "70"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text13 
         Height          =   270
         Index           =   3
         Left            =   2040
         TabIndex        =   167
         Text            =   "90"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text13 
         Height          =   270
         Index           =   2
         Left            =   2040
         TabIndex        =   166
         Text            =   "100"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text13 
         Height          =   270
         Index           =   1
         Left            =   2040
         TabIndex        =   165
         Text            =   "70"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text13 
         Height          =   270
         Index           =   0
         Left            =   2040
         TabIndex        =   164
         Text            =   "50"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "冰金桔檸檬"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   57
         Left            =   120
         TabIndex        =   64
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "數量"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   56
         Left            =   3480
         TabIndex        =   63
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "單價"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   55
         Left            =   2160
         TabIndex        =   62
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "芒果汁"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   54
         Left            =   120
         TabIndex        =   61
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "冰拿鐵"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   53
         Left            =   120
         TabIndex        =   60
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "冰咖啡"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   52
         Left            =   120
         TabIndex        =   59
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "品名"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   51
         Left            =   360
         TabIndex        =   58
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "可口可樂"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   50
         Left            =   120
         TabIndex        =   57
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "熱飲"
      Height          =   2655
      Index           =   4
      Left            =   240
      TabIndex        =   47
      Top             =   6000
      Width           =   5175
      Begin VB.VScrollBar VScroll6 
         Height          =   255
         Index           =   4
         Left            =   4320
         Max             =   0
         Min             =   20
         TabIndex        =   163
         Top             =   2040
         Width           =   255
      End
      Begin VB.VScrollBar VScroll6 
         Height          =   255
         Index           =   3
         Left            =   4320
         Max             =   0
         Min             =   20
         TabIndex        =   162
         Top             =   1680
         Width           =   255
      End
      Begin VB.VScrollBar VScroll6 
         Height          =   255
         Index           =   2
         Left            =   4320
         Max             =   0
         Min             =   20
         TabIndex        =   161
         Top             =   1320
         Width           =   255
      End
      Begin VB.VScrollBar VScroll6 
         Height          =   255
         Index           =   1
         Left            =   4320
         Max             =   0
         Min             =   20
         TabIndex        =   160
         Top             =   960
         Width           =   255
      End
      Begin VB.VScrollBar VScroll6 
         Height          =   255
         Index           =   0
         Left            =   4320
         Max             =   0
         Min             =   20
         TabIndex        =   159
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox Text12 
         Height          =   270
         Index           =   4
         Left            =   3480
         TabIndex        =   158
         Text            =   "0"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text12 
         Height          =   270
         Index           =   3
         Left            =   3480
         TabIndex        =   157
         Text            =   "0"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text12 
         Height          =   270
         Index           =   2
         Left            =   3480
         TabIndex        =   156
         Text            =   "0"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text12 
         Height          =   270
         Index           =   1
         Left            =   3480
         TabIndex        =   155
         Text            =   "0"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text12 
         Height          =   270
         Index           =   0
         Left            =   3480
         TabIndex        =   154
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text11 
         Height          =   270
         Index           =   4
         Left            =   2040
         TabIndex        =   153
         Text            =   "100"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Text11 
         Height          =   270
         Index           =   3
         Left            =   2040
         TabIndex        =   152
         Text            =   "70"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text11 
         Height          =   270
         Index           =   2
         Left            =   2040
         TabIndex        =   151
         Text            =   "90"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text11 
         Height          =   270
         Index           =   1
         Left            =   2040
         TabIndex        =   150
         Text            =   "70"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text11 
         Height          =   270
         Index           =   0
         Left            =   2040
         TabIndex        =   149
         Text            =   "70"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "奶泡熱奶茶"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   41
         Left            =   120
         TabIndex        =   55
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "品名"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   40
         Left            =   360
         TabIndex        =   54
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "熱咖啡"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   39
         Left            =   120
         TabIndex        =   53
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "熱金桔檸檬梅子"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   38
         Left            =   120
         TabIndex        =   52
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "拿鐵熱咖啡"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   37
         Left            =   120
         TabIndex        =   51
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "單價"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   36
         Left            =   2160
         TabIndex        =   50
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "數量"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   35
         Left            =   3600
         TabIndex        =   49
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "蜂蜜柚子茶"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   34
         Left            =   120
         TabIndex        =   48
         Top             =   1320
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "甜點"
      Height          =   2535
      Index           =   5
      Left            =   7320
      TabIndex        =   38
      Top             =   3360
      Width           =   3495
      Begin VB.VScrollBar VScroll5 
         Height          =   255
         Index           =   4
         Left            =   3000
         Max             =   0
         Min             =   20
         TabIndex        =   148
         Top             =   2040
         Width           =   255
      End
      Begin VB.VScrollBar VScroll5 
         Height          =   255
         Index           =   3
         Left            =   3000
         Max             =   0
         Min             =   20
         TabIndex        =   147
         Top             =   1680
         Width           =   255
      End
      Begin VB.VScrollBar VScroll5 
         Height          =   255
         Index           =   2
         Left            =   3000
         Max             =   0
         Min             =   20
         TabIndex        =   146
         Top             =   1320
         Width           =   255
      End
      Begin VB.VScrollBar VScroll5 
         Height          =   255
         Index           =   1
         Left            =   3000
         Max             =   0
         Min             =   20
         TabIndex        =   145
         Top             =   960
         Width           =   255
      End
      Begin VB.VScrollBar VScroll5 
         Height          =   255
         Index           =   0
         Left            =   3000
         Max             =   0
         Min             =   20
         TabIndex        =   144
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   4
         Left            =   2400
         TabIndex        =   143
         Text            =   "0"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   3
         Left            =   2400
         TabIndex        =   142
         Text            =   "0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   2
         Left            =   2400
         TabIndex        =   141
         Text            =   "0"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   1
         Left            =   2400
         TabIndex        =   140
         Text            =   "0"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   0
         Left            =   2400
         TabIndex        =   139
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Height          =   270
         Index           =   4
         Left            =   1560
         TabIndex        =   138
         Text            =   "50"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Height          =   270
         Index           =   3
         Left            =   1560
         TabIndex        =   137
         Text            =   "50"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Height          =   270
         Index           =   2
         Left            =   1560
         TabIndex        =   136
         Text            =   "40"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Height          =   270
         Index           =   1
         Left            =   1560
         TabIndex        =   135
         Text            =   "50"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Height          =   270
         Index           =   0
         Left            =   1560
         TabIndex        =   134
         Text            =   "30"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "義式布丁"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   49
         Left            =   120
         TabIndex        =   46
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "數量"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   48
         Left            =   2520
         TabIndex        =   45
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "單價"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   47
         Left            =   1560
         TabIndex        =   44
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "柳橙水果凍"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   46
         Left            =   120
         TabIndex        =   43
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "提拉米蘇"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   45
         Left            =   120
         TabIndex        =   42
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "焦糖蛋糕"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   44
         Left            =   120
         TabIndex        =   41
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "品名"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   43
         Left            =   360
         TabIndex        =   40
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "雞蛋布丁"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   42
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "湯品"
      Height          =   2535
      Index           =   3
      Left            =   3720
      TabIndex        =   29
      Top             =   3360
      Width           =   3495
      Begin VB.VScrollBar VScroll4 
         Height          =   255
         Index           =   4
         Left            =   3120
         Max             =   0
         Min             =   20
         TabIndex        =   133
         Top             =   2040
         Width           =   255
      End
      Begin VB.VScrollBar VScroll4 
         Height          =   255
         Index           =   3
         Left            =   3120
         Max             =   0
         Min             =   20
         TabIndex        =   132
         Top             =   1680
         Width           =   255
      End
      Begin VB.VScrollBar VScroll4 
         Height          =   255
         Index           =   2
         Left            =   3120
         Max             =   0
         Min             =   20
         TabIndex        =   131
         Top             =   1320
         Width           =   255
      End
      Begin VB.VScrollBar VScroll4 
         Height          =   255
         Index           =   1
         Left            =   3120
         Max             =   0
         Min             =   20
         TabIndex        =   130
         Top             =   960
         Width           =   255
      End
      Begin VB.VScrollBar VScroll4 
         Height          =   255
         Index           =   0
         Left            =   3120
         Max             =   0
         Min             =   20
         TabIndex        =   129
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Index           =   4
         Left            =   2520
         TabIndex        =   128
         Text            =   "0"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Index           =   3
         Left            =   2520
         TabIndex        =   127
         Text            =   "0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Index           =   2
         Left            =   2520
         TabIndex        =   126
         Text            =   "0"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Index           =   1
         Left            =   2520
         TabIndex        =   125
         Text            =   "0"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Index           =   0
         Left            =   2520
         TabIndex        =   124
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Index           =   4
         Left            =   1560
         TabIndex        =   123
         Text            =   "100"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Index           =   3
         Left            =   1560
         TabIndex        =   122
         Text            =   "100"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Index           =   2
         Left            =   1560
         TabIndex        =   121
         Text            =   "100"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Index           =   1
         Left            =   1560
         TabIndex        =   120
         Text            =   "100"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Index           =   0
         Left            =   1560
         TabIndex        =   119
         Text            =   "100"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "烤洋蔥湯"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   33
         Left            =   120
         TabIndex        =   37
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "數量"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   32
         Left            =   2520
         TabIndex        =   36
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "單價"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   31
         Left            =   1560
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "脆皮濃湯"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   30
         Left            =   120
         TabIndex        =   34
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "南瓜湯"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   29
         Left            =   120
         TabIndex        =   33
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "海鮮燉魚湯"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "品名"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   27
         Left            =   360
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "雞蓉巧達"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "前菜"
      Height          =   2535
      Index           =   2
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   3495
      Begin VB.VScrollBar VScroll3 
         Height          =   255
         Index           =   4
         Left            =   3120
         Max             =   0
         Min             =   20
         TabIndex        =   118
         Top             =   2040
         Width           =   255
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   255
         Index           =   3
         Left            =   3120
         Max             =   0
         Min             =   20
         TabIndex        =   117
         Top             =   1680
         Width           =   255
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   255
         Index           =   2
         Left            =   3120
         Max             =   0
         Min             =   20
         TabIndex        =   116
         Top             =   1320
         Width           =   255
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   255
         Index           =   1
         Left            =   3120
         Max             =   0
         Min             =   20
         TabIndex        =   115
         Top             =   960
         Width           =   255
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   255
         Index           =   0
         Left            =   3120
         Max             =   0
         Min             =   20
         TabIndex        =   114
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Index           =   4
         Left            =   2520
         TabIndex        =   113
         Text            =   "0"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Index           =   3
         Left            =   2520
         TabIndex        =   112
         Text            =   "0"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Index           =   2
         Left            =   2520
         TabIndex        =   111
         Text            =   "0"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Index           =   1
         Left            =   2520
         TabIndex        =   110
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Index           =   0
         Left            =   2520
         TabIndex        =   109
         Text            =   "0"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Index           =   4
         Left            =   1560
         TabIndex        =   108
         Text            =   "80"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Index           =   3
         Left            =   1560
         TabIndex        =   107
         Text            =   "80"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Index           =   2
         Left            =   1560
         TabIndex        =   106
         Text            =   "80"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Index           =   1
         Left            =   1560
         TabIndex        =   105
         Text            =   "80"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Index           =   0
         Left            =   1560
         TabIndex        =   104
         Text            =   "80"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "洋蔥鱈魚肝"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "品名"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   360
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "泰式嫩菲力"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "香蒜烤田螺"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "黑菌鵝肝醬"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "單價"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   1680
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "數量"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   2520
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "煙燻鮭魚"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "沙拉"
      Height          =   3015
      Index           =   1
      Left            =   5400
      TabIndex        =   9
      Top             =   240
      Width           =   5175
      Begin VB.VScrollBar VScroll2 
         Height          =   255
         Index           =   5
         Left            =   4560
         Max             =   0
         Min             =   20
         TabIndex        =   103
         Top             =   2400
         Width           =   375
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   255
         Index           =   4
         Left            =   4560
         Max             =   0
         Min             =   20
         TabIndex        =   102
         Top             =   2040
         Width           =   375
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   255
         Index           =   3
         Left            =   4560
         Max             =   0
         Min             =   20
         TabIndex        =   101
         Top             =   1680
         Width           =   375
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   255
         Index           =   2
         Left            =   4560
         Max             =   0
         Min             =   20
         TabIndex        =   100
         Top             =   1320
         Width           =   375
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   255
         Index           =   1
         Left            =   4560
         Max             =   0
         Min             =   20
         TabIndex        =   99
         Top             =   960
         Width           =   375
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   255
         Index           =   0
         Left            =   4560
         Max             =   0
         Min             =   20
         TabIndex        =   98
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Index           =   5
         Left            =   3600
         TabIndex        =   97
         Text            =   "0"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Index           =   4
         Left            =   3600
         TabIndex        =   96
         Text            =   "0"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Index           =   3
         Left            =   3600
         TabIndex        =   95
         Text            =   "0"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Index           =   2
         Left            =   3600
         TabIndex        =   94
         Text            =   "0"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Index           =   1
         Left            =   3600
         TabIndex        =   93
         Text            =   "0"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Index           =   0
         Left            =   3600
         TabIndex        =   92
         Text            =   "0"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Index           =   5
         Left            =   2400
         TabIndex        =   91
         Text            =   "60"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Index           =   4
         Left            =   2400
         TabIndex        =   90
         Text            =   "60"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Index           =   3
         Left            =   2400
         TabIndex        =   89
         Text            =   "60"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Index           =   2
         Left            =   2400
         TabIndex        =   88
         Text            =   "60"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Index           =   1
         Left            =   2400
         TabIndex        =   87
         Text            =   "60"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Index           =   0
         Left            =   2400
         TabIndex        =   86
         Text            =   "60"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "義大利醬"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   360
         TabIndex        =   18
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "生菜沙拉"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   360
         TabIndex        =   17
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "品名"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "凱薩醬"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   360
         TabIndex        =   15
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "和風醬"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   360
         TabIndex        =   14
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "優格水果沙"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   360
         TabIndex        =   13
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "千島醬"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   360
         TabIndex        =   12
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "單價"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "數量"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "主餐"
      Height          =   3015
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   5
         Left            =   4680
         Max             =   0
         Min             =   20
         TabIndex        =   85
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   5
         Left            =   3720
         TabIndex        =   84
         Text            =   "0"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   5
         Left            =   2400
         TabIndex        =   83
         Text            =   "570"
         Top             =   2400
         Width           =   975
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   4
         Left            =   4680
         Max             =   0
         Min             =   20
         TabIndex        =   82
         Top             =   2040
         Width           =   255
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   3
         Left            =   4680
         Max             =   0
         Min             =   20
         TabIndex        =   81
         Top             =   1680
         Width           =   255
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   2
         Left            =   4680
         Max             =   0
         Min             =   20
         TabIndex        =   80
         Top             =   1320
         Width           =   255
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   1
         Left            =   4680
         Max             =   0
         Min             =   20
         TabIndex        =   79
         Top             =   960
         Width           =   255
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Index           =   0
         Left            =   4680
         Max             =   0
         Min             =   20
         TabIndex        =   78
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   4
         Left            =   3720
         TabIndex        =   77
         Text            =   "0"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   3
         Left            =   3720
         TabIndex        =   76
         Text            =   "0"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   2
         Left            =   3720
         TabIndex        =   75
         Text            =   "0"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   1
         Left            =   3720
         TabIndex        =   74
         Text            =   "0"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   0
         Left            =   3720
         TabIndex        =   73
         Text            =   "0"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   4
         Left            =   2400
         TabIndex        =   72
         Text            =   "300"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   3
         Left            =   2400
         TabIndex        =   71
         Text            =   "450"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   2
         Left            =   2400
         TabIndex        =   70
         Text            =   "430"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   1
         Left            =   2400
         TabIndex        =   69
         Text            =   "380"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   0
         Left            =   2400
         TabIndex        =   68
         Text            =   "250"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "特選菲力牛排"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   360
         TabIndex        =   19
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "數量"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   3480
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "單價"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "海陸大餐"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   6
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "法式藍帶豬排"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "什錦海鮮"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "特選沙朗牛排"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "品名"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "香酥脆皮雞排"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   67
      Top             =   9000
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() '複製貼上複製貼上  沒什麼技術可言@@
For i = 0 To 5
Text2(i).Text = 0
Text4(i).Text = 0
Next

For i = 0 To 4
Text6(i).Text = 0
Text8(i).Text = 0
Text10(i).Text = 0
Text12(i).Text = 0
Text14(i).Text = 0
Next

For i = 0 To 5
VScroll1(i).Value = 0
VScroll2(i).Value = 0
Next

For i = 0 To 4
VScroll3(i).Value = 0
VScroll4(i).Value = 0
VScroll5(i).Value = 0
VScroll6(i).Value = 0
VScroll7(i).Value = 0
Next


Label2.Caption = "等待客人點餐中"
End Sub

Private Sub Command2_Click()
Dim s(7), sum As Long


For i = 0 To 5
If Text2(i).Text > 20 Then Text2(i).Text = 20
If Text4(i).Text > 20 Then Text4(i).Text = 20
s(1) = s(1) + Text1(i).Text * Text2(i).Text
s(2) = s(2) + Text3(i).Text * Text4(i).Text
Next

For i1 = 0 To 4
If Text6(i1).Text > 20 Then Text6(i1).Text = 20
If Text8(i1).Text > 20 Then Text8(i1).Text = 20
If Text10(i1).Text > 20 Then Text10(i1).Text = 20
If Text12(i1).Text > 20 Then Text12(i1).Text = 20
If Text14(i1).Text > 20 Then Text14(i1).Text = 20
s(3) = s(3) + Text5(i1).Text * Text6(i1).Text
s(4) = s(4) + Text7(i1).Text * Text8(i1).Text
s(5) = s(5) + Text9(i1).Text * Text10(i1).Text
s(6) = s(6) + Text11(i1).Text * Text12(i1).Text
s(7) = s(7) + Text13(i1).Text * Text14(i1).Text
Next

For i = 1 To 7
sum = sum + s(i)
Next

sum = sum * 1.05

Label2.Caption = "總共: " & sum

End Sub

Private Sub Form_Load()
Label2.Caption = "等待客人點餐中"
End Sub

Private Sub VScroll1_Change(Index As Integer)
Select Case Index
Case 0:
Text2(0).Text = VScroll1(0).Value
Case 1:
Text2(1).Text = VScroll1(1).Value
Case 2:
Text2(2).Text = VScroll1(2).Value
Case 3:
Text2(3).Text = VScroll1(3).Value
Case 4:
Text2(4).Text = VScroll1(4).Value
Case 5:
Text2(5).Text = VScroll1(5).Value
End Select
End Sub

Private Sub VScroll2_Change(Index As Integer)
Select Case Index
Case 0:
Text4(0).Text = VScroll2(0).Value
Case 1:
Text4(1).Text = VScroll2(1).Value
Case 2:
Text4(2).Text = VScroll2(2).Value
Case 3:
Text4(3).Text = VScroll2(3).Value
Case 4:
Text4(4).Text = VScroll2(4).Value
Case 5:
Text4(5).Text = VScroll2(5).Value
End Select
End Sub

Private Sub VScroll3_Change(Index As Integer)
Select Case Index
Case 0:
Text6(0).Text = VScroll3(0).Value
Case 1:
Text6(1).Text = VScroll3(1).Value
Case 2:
Text6(2).Text = VScroll3(2).Value
Case 3:
Text6(3).Text = VScroll3(3).Value
Case 4:
Text6(4).Text = VScroll3(4).Value
Case 5:
Text6(5).Text = VScroll3(5).Value
End Select
End Sub

Private Sub VScroll4_Change(Index As Integer)
Select Case Index
Case 0:
Text8(0).Text = VScroll4(0).Value
Case 1:
Text8(1).Text = VScroll4(1).Value
Case 2:
Text8(2).Text = VScroll4(2).Value
Case 3:
Text8(3).Text = VScroll4(3).Value
Case 4:
Text8(4).Text = VScroll4(4).Value
Case 5:
Text8(5).Text = VScroll4(5).Value
End Select
End Sub

Private Sub VScroll5_Change(Index As Integer)
Select Case Index
Case 0:
Text10(0).Text = VScroll5(0).Value
Case 1:
Text10(1).Text = VScroll5(1).Value
Case 2:
Text10(2).Text = VScroll5(2).Value
Case 3:
Text10(3).Text = VScroll5(3).Value
Case 4:
Text10(4).Text = VScroll5(4).Value
Case 5:
Text10(5).Text = VScroll5(5).Value
End Select
End Sub

Private Sub VScroll6_Change(Index As Integer)
Select Case Index
Case 0:
Text12(0).Text = VScroll6(0).Value
Case 1:
Text12(1).Text = VScroll6(1).Value
Case 2:
Text12(2).Text = VScroll6(2).Value
Case 3:
Text12(3).Text = VScroll6(3).Value
Case 4:
Text12(4).Text = VScroll6(4).Value
Case 5:
Text12(5).Text = VScroll6(5).Value
End Select
End Sub

Private Sub VScroll7_Change(Index As Integer)
Select Case Index
Case 0:
Text14(0).Text = VScroll7(0).Value
Case 1:
Text14(1).Text = VScroll7(1).Value
Case 2:
Text14(2).Text = VScroll7(2).Value
Case 3:
Text14(3).Text = VScroll7(3).Value
Case 4:
Text14(4).Text = VScroll7(4).Value
Case 5:
Text14(5).Text = VScroll7(5).Value
End Select
End Sub
