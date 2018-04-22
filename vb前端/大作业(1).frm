VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "转出"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text12 
      Height          =   735
      Left            =   7800
      TabIndex        =   34
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text11 
      Height          =   735
      Left            =   4440
      TabIndex        =   32
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      Height          =   735
      Left            =   1560
      TabIndex        =   30
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      Height          =   735
      Left            =   9840
      TabIndex        =   28
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Height          =   735
      Left            =   7200
      TabIndex        =   26
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Height          =   735
      Left            =   4440
      TabIndex        =   24
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   735
      Left            =   1200
      TabIndex        =   22
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "功能选择"
      Height          =   975
      Left            =   5520
      TabIndex        =   17
      Top             =   5160
      Width           =   9015
      Begin VB.CommandButton Command3 
         Caption         =   "退出"
         Height          =   375
         Left            =   5520
         TabIndex        =   20
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "保存"
         Height          =   495
         Left            =   3120
         TabIndex        =   19
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "清屏"
         Height          =   495
         Left            =   480
         TabIndex        =   18
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "提示信息"
      Height          =   855
      Left            =   240
      TabIndex        =   16
      Top             =   5160
      Width           =   4695
   End
   Begin VB.TextBox Text15 
      Height          =   855
      Left            =   11280
      TabIndex        =   15
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox Text14 
      Height          =   735
      Left            =   6720
      TabIndex        =   13
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text13 
      Height          =   735
      Left            =   1680
      TabIndex        =   11
      Top             =   3600
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   12000
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   9720
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   6960
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   4320
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "入院日期"
      Height          =   735
      Left            =   6360
      TabIndex        =   33
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "入住科室"
      Height          =   735
      Left            =   2880
      TabIndex        =   31
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "医生"
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "预交金"
      Height          =   735
      Left            =   8880
      TabIndex        =   27
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "费用合计"
      Height          =   735
      Left            =   5880
      TabIndex        =   25
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "护理等级"
      Height          =   735
      Left            =   2760
      TabIndex        =   23
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "病情"
      Height          =   735
      Left            =   120
      TabIndex        =   21
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Line Line8 
      X1              =   14520
      X2              =   14520
      Y1              =   3480
      Y2              =   4560
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   120
      Y1              =   3480
      Y2              =   4560
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   14640
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   14640
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   14520
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   120
      Y2              =   3360
   End
   Begin VB.Line Line2 
      X1              =   14520
      X2              =   14520
      Y1              =   120
      Y2              =   3240
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   14520
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label15 
      Caption         =   "操作员"
      Height          =   615
      Left            =   9360
      TabIndex        =   14
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "转科时间"
      Height          =   615
      Left            =   5160
      TabIndex        =   12
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "专至"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "年龄"
      Height          =   735
      Left            =   10800
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "性别"
      Height          =   735
      Left            =   8640
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "姓名"
      Height          =   735
      Left            =   5760
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "住院号"
      Height          =   735
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "病号"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
