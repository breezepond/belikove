VERSION 5.00
Begin VB.Form ��¼ 
   BackColor       =   &H00FFFFFF&
   Caption         =   "��ʿ����վ����ϵͳ�û���¼"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   5805
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox P 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   4320
      Width           =   3375
   End
   Begin VB.TextBox N 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   3840
      Width           =   3375
   End
   Begin VB.TextBox U 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   3360
      Width           =   3375
   End
   Begin VB.PictureBox P1 
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3315
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   -120
      Width           =   5775
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5760
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�û�����(&P):"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�û�����(&N):"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�û�����(&U):"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "��¼"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DbConnection As New ADODB.Connection
Private Sub Command1_Click()
Dim t1 As Integer
If P.Text = "123456" And U.Text <> "" And N.Text <> "" Then
t1 = MsgBox("��ӭ�㣡" & U.Text & "-" & N.Text, , "ϵͳ��ʾ")
��¼.Hide
����.Show
Else
t1 = MsgBox("������롢�����������Ƿ�ƥ��", , "ϵͳ��ʾ")
P.Text = ""
End If
a = N.Text
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If

End Sub

Private Sub Form_Load()
P1.Picture = LoadPicture(App.Path & "\P1.jpeg")
 DbConnection.ConnectionString = "Provider=SQLOLEDB.1;Server=10.251.30.78;Database=HOSBASE2018;Trusted_Connection=no;uid=L0G1n;Password=1qaz!QAZ"

End Sub
Private Sub ��½_unload()
Dim drm As Form
For Each frm In froms
 Unload Form
 Next
End Sub


