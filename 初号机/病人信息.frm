VERSION 5.00
Begin VB.Form 病人信息 
   Caption         =   "病人信息"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form4"
   ScaleHeight     =   4935
   ScaleWidth      =   8985
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame4 
      Caption         =   "功能选择"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   4440
      TabIndex        =   23
      Top             =   3360
      Width           =   3855
      Begin VB.CommandButton btn_back 
         Caption         =   "退出(&E)"
         Height          =   375
         Left            =   2520
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton btn_save 
         BackColor       =   &H000000FF&
         Caption         =   "保存(&S)"
         Height          =   375
         Left            =   1440
         MaskColor       =   &H000000FF&
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton btn_cls 
         Caption         =   "清屏(&R)"
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "提示信息"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   21
      Top             =   3360
      Width           =   3975
      Begin VB.Label Label5 
         Caption         =   "请输入病人病床号！"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   8415
      Begin VB.TextBox Gender 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   6480
         TabIndex        =   36
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Age 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   7680
         TabIndex        =   34
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox PatientName 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   5040
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox HosNo 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   2760
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox PatientNo 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   720
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Pricesort 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   720
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Preprice 
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   2760
         TabIndex        =   16
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TotalPrice 
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   5040
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "性别"
         Height          =   255
         Left            =   6000
         TabIndex        =   35
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "年龄"
         Height          =   255
         Left            =   7200
         TabIndex        =   33
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "姓名"
         Height          =   255
         Left            =   4200
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "住院号"
         Height          =   255
         Left            =   2040
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "病号"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "费别"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "预交金"
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "费用合计"
         Height          =   255
         Left            =   4200
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8415
      Begin VB.ComboBox CareLevel 
         Height          =   300
         Left            =   3120
         TabIndex        =   43
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Judgement 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   1080
         TabIndex        =   40
         Top             =   1800
         Width           =   3855
      End
      Begin VB.TextBox Prodate 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   6360
         TabIndex        =   38
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Project 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   960
         TabIndex        =   37
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox OperationDate 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   960
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox ResponseNurse 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   6360
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Symptom 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   6360
         TabIndex        =   4
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Doctor 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   960
         TabIndex        =   3
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Groups 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   3120
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Indate 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   3120
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label19 
         Caption         =   "主要诊断"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "所在组别"
         Height          =   255
         Left            =   2400
         TabIndex        =   41
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "入科日期"
         Height          =   255
         Left            =   5400
         TabIndex        =   39
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "责任护士"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5400
         TabIndex        =   13
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "护理等级"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "手术日期"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "病情状态"
         Height          =   255
         Left            =   5400
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "入院日期"
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "入住科室"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "经治医生"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
   End
End
Attribute VB_Name = "病人信息"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DbConnection As New ADODB.Connection

Private Sub btn_back_Click()
引导.Show
病人信息.Hide
End Sub

Private Sub Btn_cls_Click()
Dim ctrl As Control
For Each ctrl In Me.Controls
    If TypeOf ctrl Is TextBox Then   '是否为文本框TextBox
        ctrl.Text = ""
    End If
Next
End Sub



Private Sub Btn_exit_Click()
Form1.Hide
End Sub


Private Sub Btn_save_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If PatientName.Text = "" Then
     btn_save.Enabled = False
     End If
End Sub



Private Sub Form_Load()
 DbConnection.ConnectionString = "Provider=SQLOLEDB.1;Server=PC-20180428BGOQ;Database=HOSBASE2019;Trusted_Connection=no;Uid=L0G1n;Password=1qaz!QAZ"
End Sub

Private Sub PatientNo_Change()
      If PatientNo.Text <> "" Then
     btn_save.Enabled = True
     End If
End Sub

Private Sub 病人信息_unload()
Dim drm As Form
For Each frm In froms
 Unload Form
 Next
End Sub


Private Sub Btn_save_Click()
 Dim sqlCommand As String
    Dim rowAffected As Integer
    sqlCommand = "INSERT INTO tb_Patient ( PatientNo,HosNo,Name,Gender,Age, PriceSort, PrePrice, TotalPrice,Project,Indate,Prodate,Operationdate,CareLevel,Symptom,Doctor,Groups,ResponseNurse,Mainjudgement)VALUES( '" + PatientNo + "','" + HosNo + "','" + PatientName + "','" + Gender + "','" + Age + "', '" + Pricesort + "', '" + Preprice + "','" + TotalPrice + "','" + Project + "','" + Indate + "','" + Prodate + "','" + OperationDate + "','" + CareLevel + "','" + Symptom + "','" + Doctor + "','" + Groups + "','" + ResponseNurse + "','" + Mainjudgement + "');"
    DbConnection.Open
       DbConnection.Execute sqlCommand, rowAffected
    If rowAffected = 1 Then
        MsgBox "更新成功", vbInformation, "提示"
    Else
        MsgBox "更新失败", vbCritical, "错误"
    End If
    DbConnection.Close
End Sub
