VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14730
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   14730
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   1335
      Left            =   840
      TabIndex        =   39
      Top             =   3720
      Width           =   9135
      Begin VB.ListBox List7 
         Height          =   1140
         Left            =   7080
         TabIndex        =   45
         Top             =   0
         Width           =   1335
      End
      Begin VB.ListBox List6 
         Height          =   1140
         Left            =   5880
         TabIndex        =   44
         Top             =   0
         Width           =   1095
      End
      Begin VB.ListBox List5 
         Height          =   1140
         Left            =   4560
         TabIndex        =   43
         Top             =   0
         Width           =   1215
      End
      Begin VB.ListBox List4 
         Height          =   1140
         Left            =   3240
         TabIndex        =   42
         Top             =   0
         Width           =   1335
      End
      Begin VB.ListBox List2 
         Height          =   1140
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   1575
      End
      Begin VB.ListBox List3 
         Height          =   1140
         Left            =   1680
         TabIndex        =   40
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "功能选择"
      Height          =   1095
      Left            =   4800
      TabIndex        =   35
      Top             =   5280
      Width           =   3975
      Begin VB.CommandButton Command3 
         Caption         =   "退出"
         Height          =   615
         Left            =   2400
         TabIndex        =   38
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "保存"
         Height          =   615
         Left            =   1320
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "清屏"
         Height          =   615
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "信息提示"
      Height          =   1215
      Left            =   720
      TabIndex        =   34
      Top             =   5160
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1215
      Left            =   600
      TabIndex        =   17
      Top             =   2160
      Width           =   9255
      Begin VB.TextBox Text13 
         Height          =   270
         Left            =   480
         TabIndex        =   33
         Top             =   840
         Width           =   7695
      End
      Begin VB.TextBox Text12 
         Height          =   270
         Left            =   7080
         TabIndex        =   31
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text11 
         Height          =   270
         Left            =   3480
         TabIndex        =   29
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   480
         TabIndex        =   27
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   7080
         TabIndex        =   25
         Top             =   120
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   4920
         TabIndex        =   23
         Top             =   120
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Height          =   240
         ItemData        =   "Form1.frx":0000
         Left            =   2520
         List            =   "Form1.frx":0002
         TabIndex        =   21
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   270
         Left            =   480
         TabIndex        =   19
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "诊断"
         Height          =   375
         Left            =   0
         TabIndex        =   32
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "责任护士"
         Height          =   255
         Left            =   6360
         TabIndex        =   30
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "所在组别"
         Height          =   255
         Left            =   2640
         TabIndex        =   28
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "医生"
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "癌情状态"
         Height          =   255
         Left            =   6240
         TabIndex        =   24
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "护理等级"
         Height          =   255
         Left            =   4080
         TabIndex        =   22
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "手术日期"
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "床位"
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   9255
      Begin VB.TextBox Text8 
         Height          =   390
         Left            =   6600
         TabIndex        =   16
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   3000
         TabIndex        =   14
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   8160
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   390
         Left            =   6600
         TabIndex        =   8
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   5040
         TabIndex        =   6
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   390
         Left            =   3000
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "预交金额"
         Height          =   375
         Left            =   5880
         TabIndex        =   15
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "科室"
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "入院日期"
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "费别"
         Height          =   375
         Left            =   7680
         TabIndex        =   9
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "年龄"
         Height          =   375
         Left            =   6120
         TabIndex        =   7
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "性别"
         Height          =   375
         Left            =   4560
         TabIndex        =   5
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "姓名"
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "住院号"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub Label13_Click()

End Sub

Private Sub Label16_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text6_Change()

End Sub
