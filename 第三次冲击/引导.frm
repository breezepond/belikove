VERSION 5.00
Begin VB.Form 引导 
   Caption         =   "引导"
   ClientHeight    =   6090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   10260
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Height          =   6135
      Left            =   0
      ScaleHeight     =   6075
      ScaleWidth      =   10155
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.CommandButton Command3 
         Caption         =   "充值"
         Height          =   1095
         Left            =   4080
         TabIndex        =   7
         Top             =   3480
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "用药登记"
         Height          =   1095
         Left            =   4080
         TabIndex        =   6
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "出院办理"
         Height          =   1095
         Left            =   6720
         TabIndex        =   5
         Top             =   1560
         Width           =   2535
      End
      Begin VB.CommandButton Command6 
         Caption         =   "费用结算"
         Height          =   1095
         Left            =   6720
         TabIndex        =   3
         Top             =   3480
         Width           =   2535
      End
      Begin VB.CommandButton Command5 
         Caption         =   "值班更替"
         Height          =   1095
         Left            =   960
         TabIndex        =   2
         Top             =   3480
         Width           =   2415
      End
      Begin VB.CommandButton Patientinform 
         Caption         =   "入院登记"
         Height          =   1095
         Left            =   960
         TabIndex        =   1
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label 标语 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "请选择"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   4
         Top             =   600
         Width           =   5535
      End
   End
End
Attribute VB_Name = "引导"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "请先进行相关费用的结算", vbInformation, "提示"
结算查询.Show
引导.Hide
End Sub

Private Sub Command2_Click()
引导.Hide
用药登记.Show
End Sub

Private Sub Command3_Click()
充值.Show
引导.Hide

End Sub

Private Sub Command5_Click()
值班更替.Show
引导.Hide

End Sub

Private Sub Command6_Click()
结算查询.Show
引导.Hide

End Sub

Private Sub Form_Load()
Picture1.Picture = LoadPicture(App.Path & "\forepic.jpg")
End Sub
Private Sub 引导_unload()
Dim drm As Form
For Each frm In froms
 Unload Form
 Next
End Sub

Private Sub getout_Click()
转出.Show
引导.Hide
End Sub

Private Sub getoutcancel_Click()
取消转出.Show
引导.Hide
End Sub

Private Sub Patientinform_Click()
入院登记.Show
引导.Hide
End Sub



