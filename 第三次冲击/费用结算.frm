VERSION 5.00
Begin VB.Form ���ý��� 
   Caption         =   "���ý���"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton btn_search 
      Caption         =   "��Ϣ��ѯ"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox PatientNo 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label PatientName 
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "����"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Width           =   615
   End
   Begin VB.Label TotalPrice 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "�ܽ��"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Ԥ����"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Preprice 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "����"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "���ý���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DbConnection As New ADODB.Connection

Private Sub btn_cls_Click()
Dim ctrl As Control
For Each ctrl In Me.Controls
    If TypeOf ctrl Is TextBox Then   '�Ƿ�Ϊ�ı���TextBox
        ctrl.Text = ""
    End If
Next
End Sub



Private Sub Btn_exit_Click()
ת��.Hide
����.Show
Dim ctrl As Control
For Each ctrl In Me.Controls
    If TypeOf ctrl Is TextBox Then   '�Ƿ�Ϊ�ı���TextBox
        ctrl.Text = ""
    End If
Next
End Sub

Private Sub btn_search_Click()
 Dim sqlCommand As String
    Dim recordSet As ADODB.recordSet
    sqlCommand = "SELECT P.Name,P.CareLevel ,P.TotalPrice ,P.PrePrice  FROM tb_Patient AS p WHERE p.patientNo='" + PatientNo + "';"
    DbConnection.Open
    Set recordSet = DbConnection.Execute(sqlCommand)
            If Not recordSet.EOF Then
        PatientName.Caption = recordSet.Fields("Name")
        TotalPrice.Caption = recordSet.Fields("TotalPrice")
        Preprice.Caption = recordSet.Fields("Preprice")
        End If
    DbConnection.Close
        If TotalPrice.Caption > Preprice.Caption Then
         MsgBox "���ֵ������", vbCritical, "����"
         ����.Show
         ���ý���.Hide
 End If
End Sub


Private Sub Command2_Click()

   Dim sqlCommand As String
    Dim rowAffected As Integer
      
       sqlCommand = " Update tb_Patient SET Totalprice-=Preprice where PatientNo='" + PatientNo + "';"
     DbConnection.Open
    DbConnection.Execute sqlCommand, rowAffected
    If rowAffected = 1 Then
     MsgBox "��ת����Ժ����", vbInformation, "��ʾ"
     ��Ժ����.Show
     ���ý���.Hide
    Else
        MsgBox "����ʧ��", vbCritical, "����"
    End If
 DbConnection.Close



End Sub

Private Sub Form_Load()
 DbConnection.ConnectionString = "Provider=SQLOLEDB.1;Server=PC-20180428BGOQ;Database=HOSBASE2019;Trusted_Connection=no;Uid=L0G1n;Password=1qaz!QAZ"
End Sub

Private Sub ���ý���_unload()
Dim drm As Form
For Each frm In froms
 Unload Form
 Next
End Sub

