VERSION 5.00
Begin VB.Form ��Ժ���� 
   Caption         =   "��Ժ����"
   ClientHeight    =   4080
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   6840
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   6135
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   15
         Left            =   0
         TabIndex        =   13
         Top             =   720
         Width           =   5895
      End
      Begin VB.CommandButton btn_search 
         Caption         =   "��ѯ"
         Height          =   375
         Left            =   4680
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox PatientNo 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label bingh 
         Alignment       =   2  'Center
         Caption         =   "����"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2175
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   6135
      Begin VB.Label Preprice 
         Height          =   375
         Left            =   3960
         TabIndex        =   20
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Ԥ����"
         Height          =   375
         Left            =   3000
         TabIndex        =   19
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Age 
         Height          =   375
         Left            =   1080
         TabIndex        =   18
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "����"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label HosNo 
         Height          =   255
         Left            =   4080
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label TotalPrice 
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3960
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "�ܷ���"
         Height          =   375
         Left            =   3120
         TabIndex        =   8
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label PatientName 
         Height          =   375
         Left            =   960
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label CareLevel 
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "����ȼ�"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "סԺ��"
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "����"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton btn_clear 
      Caption         =   "ȷ��(&O)"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton btn_cls 
      Caption         =   "ȡ��(&E)"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "��Ժ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DbConnection As New ADODB.Connection

Private Sub btn_clear_Click()
X = MsgBox("ȷ��Ҫ���иò�����", vbYesNo, "��ʾ")
If X = vbNo Then
btn_cls_Click
End If
If TotalPrice.Caption > Preprice.Caption Then
MsgBox "���ò������Ƚ��г�ֵ", vbInformation, "����"
Else
Dim sqlCommand As String
Dim sqlCommand1 As String
    Dim recordSet As ADODB.recordSet
    sqlCommand = "delete tb_patient  WHERE HosNo='" + HosNo + "';"
  DbConnection.Open
    DbConnection.Execute sqlCommand, rowAffected
    If rowAffected = 1 Then
        MsgBox "�������Ժ", vbInformation, "��ʾ"
    Else
        MsgBox "����ʧ��", vbCritical, "����"
    End If
    DbConnection.Close
    End If
End Sub

Private Sub btn_search_Click()
 DbConnection.ConnectionString = "Provider=SQLOLEDB.1;Server=PC-20180428BGOQ;Database=HOSBASE2019;Trusted_Connection=no;Uid=L0G1n;Password=1qaz!QAZ"
Dim sqlCommand As String
    Dim recordSet As ADODB.recordSet
    sqlCommand = "select Name,CareLevel,TotalPrice,HosNo,Age,Preprice from tb_patient WHERE tb_patient.PatientNo='" + PatientNo + "';"
    DbConnection.Open
    Set recordSet = DbConnection.Execute(sqlCommand)
            If Not recordSet.EOF Then
        PatientName.Caption = recordSet.Fields("Name")
        CareLevel.Caption = recordSet.Fields("Carelevel")
        HosNo.Caption = recordSet.Fields("HosNo")
        Age.Caption = recordSet.Fields("Age")
       Preprice.Caption = recordSet.Fields("Preprice")
        End If
    DbConnection.Close
End Sub

Private Sub btn_cls_Click()
Dim ctrl As Control
For Each ctrl In Me.Controls
    If TypeOf ctrl Is TextBox Then   '�Ƿ�Ϊ�ı���TextBox
        ctrl.Text = ""
    End If
Next
PatientName.Caption = ""
HosNo.Caption = ""
CareLevel.Caption = ""
TotalPrice.Caption = ""
End Sub

Private Sub Command1_Click()
��Ժ����.Hide
����.Show
End Sub

Private Sub Form_Load()
 DbConnection.ConnectionString = "Provider=SQLOLEDB.1;Server=PC-20180428BGOQ;Database=HOSBASE2019;Trusted_Connection=no;Uid=L0G1n;Password=1qaz!QAZ"
End Sub


Private Sub ��Ժ����_unload()
Dim drm As Form
For Each frm In froms
 Unload Form
 Next
End Sub

