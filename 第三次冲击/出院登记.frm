VERSION 5.00
Begin VB.Form 出院办理 
   Caption         =   "出院办理"
   ClientHeight    =   4080
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   6840
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "返回"
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
         Caption         =   "查询"
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
         Caption         =   "病号"
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
         Caption         =   "预交金"
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
         Caption         =   "年龄"
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
         Caption         =   "总费用"
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
         Caption         =   "护理等级"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "住院号"
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "姓名"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton btn_clear 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton btn_cls 
      Caption         =   "取消(&E)"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "出院办理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DbConnection As New ADODB.Connection

Private Sub btn_clear_Click()
X = MsgBox("确定要进行该操作吗？", vbYesNo, "提示")
If X = vbNo Then
btn_cls_Click
End If
If TotalPrice.Caption > Preprice.Caption Then
MsgBox "费用不足请先进行充值", vbInformation, "错误"
Else
Dim sqlCommand As String
Dim sqlCommand1 As String
    Dim recordSet As ADODB.recordSet
    sqlCommand = "delete tb_patient  WHERE HosNo='" + HosNo + "';"
  DbConnection.Open
    DbConnection.Execute sqlCommand, rowAffected
    If rowAffected = 1 Then
        MsgBox "已受理出院", vbInformation, "提示"
    Else
        MsgBox "办理失败", vbCritical, "错误"
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
    If TypeOf ctrl Is TextBox Then   '是否为文本框TextBox
        ctrl.Text = ""
    End If
Next
PatientName.Caption = ""
HosNo.Caption = ""
CareLevel.Caption = ""
TotalPrice.Caption = ""
End Sub

Private Sub Command1_Click()
出院办理.Hide
引导.Show
End Sub

Private Sub Form_Load()
 DbConnection.ConnectionString = "Provider=SQLOLEDB.1;Server=PC-20180428BGOQ;Database=HOSBASE2019;Trusted_Connection=no;Uid=L0G1n;Password=1qaz!QAZ"
End Sub


Private Sub 出院办理_unload()
Dim drm As Form
For Each frm In froms
 Unload Form
 Next
End Sub

