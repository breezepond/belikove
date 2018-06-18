VERSION 5.00
Begin VB.Form 出院办理 
   Caption         =   "出院办理"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6510
   LinkTopic       =   "Form2"
   ScaleHeight     =   3480
   ScaleWidth      =   6510
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "取消(&E)"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton Command3 
         Caption         =   "查询"
         Height          =   255
         Left            =   4680
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "住院号"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label9 
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "总费用"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label6 
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label5 
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "护理等级"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "病号"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "姓名"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "出院办理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub 取消出院_unload()
Dim drm As Form
For Each frm In froms
 Unload Form
 Next
End Sub

Private Sub Command2_Click()
引导.Show
取消出院.Hide
End Sub



Private Sub Text2_Change()

End Sub

Private Sub Text3_Change()

End Sub
