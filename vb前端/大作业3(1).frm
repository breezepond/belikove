VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
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
      Height          =   615
      Left            =   6360
      TabIndex        =   11
      Top             =   5640
      Width           =   8055
      Begin VB.CommandButton Command3 
         Caption         =   "退出"
         Height          =   255
         Left            =   4920
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "保存"
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "清屏"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "信息提示"
      Height          =   735
      Left            =   360
      TabIndex        =   10
      Top             =   5640
      Width           =   5535
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   12360
      TabIndex        =   9
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   8880
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   6120
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   120
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   14760
      X2              =   14760
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   14880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   14760
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label5 
      Caption         =   "现分配床位号"
      Height          =   615
      Left            =   10320
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "性别"
      Height          =   615
      Left            =   7680
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "姓名"
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "床位"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "住院号"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
