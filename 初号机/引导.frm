VERSION 5.00
Begin VB.Form ���� 
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   10260
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox Picture1 
      Height          =   6135
      Left            =   0
      ScaleHeight     =   6075
      ScaleWidth      =   10155
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   1095
         Left            =   7320
         TabIndex        =   6
         Top             =   3480
         Width           =   2535
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   1095
         Left            =   4080
         TabIndex        =   5
         Top             =   3480
         Width           =   2415
      End
      Begin VB.CommandButton getout 
         Caption         =   "ת��"
         Height          =   1095
         Left            =   600
         TabIndex        =   4
         Top             =   3480
         Width           =   2535
      End
      Begin VB.CommandButton getoutcancel 
         Caption         =   "ȡ��ת��"
         Height          =   1095
         Left            =   7320
         TabIndex        =   3
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton outcancel 
         Caption         =   "ȡ����Ժ"
         Height          =   1095
         Left            =   4080
         TabIndex        =   2
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CommandButton Patientinform 
         Caption         =   "������Ϣ"
         Height          =   1095
         Left            =   600
         TabIndex        =   1
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label ���� 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   7
         Top             =   720
         Width           =   5535
      End
   End
End
Attribute VB_Name = "����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Picture1.Picture = LoadPicture(App.Path & "\forepic.jpg")
End Sub
Private Sub ����_unload()
Dim drm As Form
For Each frm In froms
 Unload Form
 Next
End Sub

Private Sub getout_Click()
ת��.Show
����.Hide
End Sub

Private Sub getoutcancel_Click()
ȡ��ת��.Show
����.Hide
End Sub

Private Sub Patientinform_Click()
������Ϣ.Show
����.Hide
End Sub

Private Sub outcancel_Click()
ȡ����Ժ.Show
����.Hide
End Sub

