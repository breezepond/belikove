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
   Begin VB.Label Label2 
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
      Left            =   3360
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "值班更替"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Label2.Caption = 登录.N.Text
End Sub

