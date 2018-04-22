VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "病人信息"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "功能选择"
      Height          =   855
      Left            =   6120
      TabIndex        =   39
      Top             =   5400
      Width           =   8295
      Begin VB.CommandButton Command3 
         Caption         =   "退出"
         Height          =   495
         Left            =   4920
         TabIndex        =   42
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "保存"
         Height          =   495
         Left            =   2520
         TabIndex        =   41
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "清屏"
         Height          =   495
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "信息提示"
      Height          =   735
      Left            =   120
      TabIndex        =   38
      Top             =   5520
      Width           =   5535
   End
   Begin VB.TextBox Text19 
      Height          =   855
      Left            =   2160
      TabIndex        =   37
      Top             =   4320
      Width           =   7695
   End
   Begin VB.TextBox Text18 
      Height          =   615
      Left            =   8640
      TabIndex        =   35
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox Text17 
      Height          =   495
      Left            =   5400
      TabIndex        =   33
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text16 
      Height          =   495
      Left            =   2280
      TabIndex        =   31
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text15 
      Height          =   615
      Left            =   8640
      TabIndex        =   29
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox Text14 
      Height          =   615
      Left            =   5520
      TabIndex        =   27
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   2160
      TabIndex        =   25
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   8520
      TabIndex        =   23
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   5400
      TabIndex        =   21
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      Height          =   615
      Left            =   2040
      TabIndex        =   19
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      Height          =   855
      Left            =   10680
      TabIndex        =   17
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Height          =   855
      Left            =   7440
      TabIndex        =   15
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   975
      Left            =   4800
      TabIndex        =   13
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   975
      Left            =   1440
      TabIndex        =   11
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   855
      Left            =   13440
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   975
      Left            =   10560
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   975
      Left            =   7560
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   120
      Y1              =   2160
      Y2              =   5280
   End
   Begin VB.Line Line6 
      X1              =   14640
      X2              =   14640
      Y1              =   2160
      Y2              =   5280
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   14640
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line4 
      X1              =   14640
      X2              =   14640
      Y1              =   120
      Y2              =   2160
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   14760
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   0
      Y2              =   2160
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   14760
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Label Label19 
      Caption         =   "主要诊断"
      Height          =   495
      Left            =   120
      TabIndex        =   36
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label18 
      Caption         =   "责任护士"
      Height          =   375
      Left            =   7320
      TabIndex        =   34
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label17 
      Caption         =   "所在组别"
      Height          =   495
      Left            =   4080
      TabIndex        =   32
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "经治医生"
      Height          =   495
      Left            =   120
      TabIndex        =   30
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label15 
      Caption         =   "病情状态"
      Height          =   615
      Left            =   7200
      TabIndex        =   28
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "护理等级"
      Height          =   615
      Left            =   3960
      TabIndex        =   26
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "手术日期"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label12 
      Caption         =   "入科日期"
      Height          =   495
      Left            =   7200
      TabIndex        =   22
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "入院日期"
      Height          =   495
      Left            =   3840
      TabIndex        =   20
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "入住科室"
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "余额"
      Height          =   855
      Left            =   9000
      TabIndex        =   16
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "费用合计"
      Height          =   735
      Left            =   6000
      TabIndex        =   14
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "预交金"
      Height          =   975
      Left            =   3360
      TabIndex        =   12
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "费别"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "年龄"
      Height          =   735
      Left            =   12000
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "性别"
      Height          =   855
      Left            =   9000
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "姓名"
      Height          =   855
      Left            =   6120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "住院号"
      Height          =   735
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "床号"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text13_Change()

End Sub
