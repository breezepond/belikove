VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ȡ��ת�� 
   Caption         =   "ȡ��ת��"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8610
   LinkTopic       =   "Form3"
   ScaleHeight     =   3030
   ScaleWidth      =   8610
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame4 
      Caption         =   "����ѡ��"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   4680
      TabIndex        =   23
      Top             =   2040
      Width           =   3855
      Begin VB.CommandButton Command1 
         Caption         =   "����(&R)"
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H000000FF&
         Caption         =   "����(&S)"
         Height          =   375
         Left            =   1440
         MaskColor       =   &H000000FF&
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "�˳�(&E)"
         Height          =   375
         Left            =   2520
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "��ʾ��Ϣ"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   2040
      Width           =   3975
      Begin VB.Label Label5 
         Caption         =   "��������ֱ�ӻس�����(ALT + S)"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.ComboBox Gender 
         Height          =   300
         ItemData        =   "ȡ��ת��.frx":0000
         Left            =   6360
         List            =   "ȡ��ת��.frx":000A
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
      Begin MSComCtl2.DTPicker Indate 
         Height          =   255
         Left            =   6720
         TabIndex        =   29
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   178847745
         CurrentDate     =   43272
      End
      Begin VB.ComboBox CareLevel 
         Height          =   300
         ItemData        =   "ȡ��ת��.frx":0016
         Left            =   960
         List            =   "ȡ��ת��.frx":0023
         TabIndex        =   28
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   960
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text2 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   2880
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text3 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   4440
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text5 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   7560
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text7 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   2880
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text8 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   4800
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text9 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   6600
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text10 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   960
         TabIndex        =   2
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text11 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   2880
         TabIndex        =   1
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "���úϼ�"
         Height          =   255
         Left            =   4080
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "סԺ��"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "����ȼ�"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "����"
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "����"
         Height          =   255
         Left            =   7080
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "�Ա�"
         Height          =   255
         Left            =   5880
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "����״̬"
         Height          =   255
         Left            =   2040
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "����"
         Height          =   255
         Left            =   3960
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Ԥ����"
         Height          =   255
         Left            =   5880
         TabIndex        =   13
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "��Ժ����"
         Height          =   255
         Left            =   5880
         TabIndex        =   12
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "��������"
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "ҽ��"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   495
      End
   End
End
Attribute VB_Name = "ȡ��ת��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ȡ��ת��_unload()
Dim drm As Form
For Each frm In froms
 Unload Form
 Next
End Sub

Private Sub Command3_Click()
ȡ��ת��.Hide
����.Show
End Sub

