VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ת�� 
   Caption         =   "ת��"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14745
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   14745
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ComboBox Gender 
      Height          =   300
      ItemData        =   "ת��.frx":0000
      Left            =   9960
      List            =   "ת��.frx":000A
      TabIndex        =   36
      Top             =   360
      Width           =   975
   End
   Begin MSComCtl2.DTPicker Indate 
      Height          =   375
      Left            =   7080
      TabIndex        =   35
      Top             =   1800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   178847745
      CurrentDate     =   43273
   End
   Begin MSComCtl2.DTPicker Todate 
      Height          =   255
      Left            =   6120
      TabIndex        =   34
      Top             =   3000
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Format          =   178847745
      CurrentDate     =   43273
   End
   Begin VB.ComboBox CareLevel 
      Height          =   300
      ItemData        =   "ת��.frx":0016
      Left            =   3840
      List            =   "ת��.frx":0023
      TabIndex        =   32
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Project 
      Height          =   375
      Left            =   3840
      TabIndex        =   29
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Doctor 
      Height          =   375
      Left            =   840
      TabIndex        =   27
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Preprice 
      Height          =   375
      Left            =   9840
      TabIndex        =   25
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Totalprice 
      Height          =   375
      Left            =   6960
      TabIndex        =   23
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Symptom 
      Height          =   375
      Left            =   840
      TabIndex        =   20
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "����ѡ��"
      Height          =   855
      Left            =   3720
      TabIndex        =   15
      Top             =   3840
      Width           =   10695
      Begin VB.CommandButton btn_Search 
         Caption         =   "��ѯ"
         Height          =   495
         Left            =   5640
         TabIndex        =   31
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Btn_exit 
         Caption         =   "�˳�"
         Height          =   495
         Left            =   8040
         TabIndex        =   18
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Btn_save 
         Caption         =   "����"
         Height          =   495
         Left            =   3120
         TabIndex        =   17
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Btn_cls 
         Caption         =   "����"
         Height          =   495
         Left            =   480
         TabIndex        =   16
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��ʾ��Ϣ"
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   3255
      Begin VB.Label Label16 
         Caption         =   "�������벡���ٽ�����ز���"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.TextBox Operator 
      Height          =   375
      Left            =   11520
      TabIndex        =   13
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox Getto 
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   2880
      Width           =   4335
   End
   Begin VB.TextBox Age 
      Height          =   375
      Left            =   12600
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox PatientName 
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox HosNo 
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox PatientNo 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "��Ժ����"
      Height          =   375
      Left            =   5880
      TabIndex        =   30
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "��ס����"
      Height          =   375
      Left            =   2760
      TabIndex        =   28
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "ҽ��"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Ԥ����"
      Height          =   375
      Left            =   9120
      TabIndex        =   24
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "���úϼ�"
      Height          =   495
      Left            =   6000
      TabIndex        =   22
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "����ȼ�"
      Height          =   255
      Left            =   2760
      TabIndex        =   21
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "����"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1080
      Width           =   495
   End
   Begin VB.Line Line8 
      X1              =   14520
      X2              =   14520
      Y1              =   3480
      Y2              =   3600
   End
   Begin VB.Line Line7 
      X1              =   0
      X2              =   0
      Y1              =   3360
      Y2              =   3600
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   14520
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   14520
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   14520
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   120
      Y2              =   3360
   End
   Begin VB.Line Line2 
      X1              =   14520
      X2              =   14520
      Y1              =   120
      Y2              =   3600
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   14520
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label15 
      Caption         =   "����Ա"
      Height          =   255
      Left            =   10800
      TabIndex        =   12
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label14 
      Caption         =   "ת��ʱ��"
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label13 
      Caption         =   "ת��"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "����"
      Height          =   375
      Left            =   11880
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "�Ա�"
      Height          =   375
      Left            =   9120
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "����"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "סԺ��"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "����"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "ת��"
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

Private Sub Btn_save_Click()
 Dim sqlCommand As String
    Dim rowAffected As Integer
    sqlCommand = "INSERT INTO tb_Patient ( PatientNo,HosNo,Name,Gender,Age,Symptom,CareLevel,TotalPrice,Preprice,Doctor,Project,Indate,GetTo,todate,Operator) VALUES ('" + PatientNo + "','" + HosNo + "','" + PatientName + "','" + Gender + "','" + Age + "','" + Symptom + "','" + CareLevel + "','" + TotalPrice + "','" + Preprice + "','" + Doctor + "','" + Project + "','" + Format(Indate.Value, "yyyy-mm-dd") + "','" + Format(Todate.Value, "yyyy-mm-dd") + "','" + Getto + "','" + Operator + "');"
    DbConnection.Open
    DbConnection.Execute sqlCommand, rowAffected
    If rowAffected = 1 Then
        MsgBox "���³ɹ�", vbInformation, "��ʾ"
    Else
        MsgBox "����ʧ��", vbCritical, "����"
    End If
    DbConnection.Close
End Sub

Private Sub Btn_save_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If PatientNo.Text = "" Then
     btn_save.Enabled = False
     End If
End Sub

Private Sub btn_search_Click()
 Dim sqlCommand As String
    Dim recordSet As ADODB.recordSet
    sqlCommand = "SELECT P.Name,P.HosNo,P.gender,P.symptom,P.Age,P.CareLevel ,P.TotalPrice ,P.PrePrice ,P.Doctor ,P.Project ,P.Indate  FROM tb_Patient AS p WHERE p.patientNo='" + PatientNo.Text + "';"
    DbConnection.Open
    Set recordSet = DbConnection.Execute(sqlCommand)
            If Not recordSet.EOF Then
        PatientName.Text = recordSet.Fields("Name")
        HosNo.Text = recordSet.Fields("HosNo")
        Gender.Text = recordSet.Fields("Gender")
        Symptom.Text = recordSet.Fields("Symptom")
        Age.Text = recordSet.Fields("Age")
        CareLevel.Text = recordSet.Fields("Carelevel")
        TotalPrice.Text = recordSet.Fields("TotalPrice")
        Symptom.Text = recordSet.Fields("Symptom")
        Doctor.Text = recordSet.Fields("Doctor")
        Project.Text = recordSet.Fields("Project")
        Indate.Text = recordSet.Fields("Indate")
        End If
    DbConnection.Close
End Sub

Private Sub Form_Load()
 DbConnection.ConnectionString = "Provider=SQLOLEDB.1;Server=PC-20180428BGOQ;Database=HOSBASE2019;Trusted_Connection=no;Uid=L0G1n;Password=1qaz!QAZ"
End Sub

Private Sub PatientNo_Change()
      If PatientNo.Text <> "" Then
     btn_save.Enabled = True
     End If
End Sub
Private Sub ת��_unload()
Dim drm As Form
For Each frm In froms
 Unload Form
 Next
End Sub

