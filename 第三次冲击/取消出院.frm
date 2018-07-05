VERSION 5.00
Begin VB.Form 出院办理 
   Caption         =   "出院办理"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6510
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   3480
   ScaleWidth      =   6510
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "出院办理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DbConnection As New ADODB.Connection
Private Sub 出院办理_unload()
Dim drm As Form
For Each frm In froms
 Unload Form
 Next
End Sub

Private Sub Command2_Click()
引导.Show
取消出院.Hide
End Sub


Private Sub btn_cls_Click()
Dim ctrl As Control
For Each ctrl In Me.Controls
    If TypeOf ctrl Is TextBox Then   '是否为文本框TextBox
        ctrl.Text = ""
    End If
Next
End Sub

Private Sub btn_search_Click()

 Dim sqlCommand As String
    Dim recordSet As ADODB.recordSet
    sqlCommand = "SELECT P.Name,P.HosNo,P.CareLevel ,P.TotalPrice FROM tb_Patient AS p WHERE p.HosNo='" + HosNo + "';"
    DbConnection.Open
    Set recordSet = DbConnection.Execute(sqlCommand)
            If Not recordSet.EOF Then
        PatientName.Caption = recordSet.Fields("Name")
        TotalPrice.Caption = recordSet.Fields("TotalPrice")
        CareLevel.Caption = recordSet.Fields("Carelevel")
        PatientNo.Caption = recordSet.Fields("PatientNo")
        End If
    DbConnection.Close
End Sub

Private Sub Form_Load()
 DbConnection.ConnectionString = "Provider=SQLOLEDB.1;Server=PC-20180428BGOQ;Database=HOSBASE2019;Trusted_Connection=no;Uid=L0G1n;Password=1qaz!QAZ"
End Sub


