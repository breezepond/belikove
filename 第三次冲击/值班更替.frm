VERSION 5.00
Begin VB.Form 值班更替 
   Caption         =   "值班更替"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "返回"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "换班"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   2520
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "值班更替.frx":0000
      Left            =   2040
      List            =   "值班更替.frx":000D
      TabIndex        =   4
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "换班护士名"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label label2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "护士名"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label time 
      Caption         =   "Label1"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "值班更替"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Label2.Caption = Combo1.Text
登录.N.Text = Combo1.Text

End Sub

Private Sub Command2_Click()
值班更替.Hide
引导.Show

End Sub

Private Sub Form_Load()
Label2.Caption = 登录.N.Text
 time.Caption = Format(Now, "yyyy-mm-dd hh:mm:ss")
End Sub


