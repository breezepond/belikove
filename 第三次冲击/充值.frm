VERSION 5.00
Begin VB.Form ��ֵ 
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   6150
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   615
      Left            =   3960
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ֵ"
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "��ֵ���"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "����"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "��ֵ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
��Ժ����.Preprice.Caption = Val(Text2.Text) + Val(��Ժ����.Preprice.Caption)
        MsgBox "��ֵ�ɹ�", vbInformation, "��ʾ"
End Sub

Private Sub Command2_Click()
����.Show
��ֵ.Hide

End Sub

