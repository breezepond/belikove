VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form �����ѯ 
   Caption         =   "�����ѯ"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10215
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   10215
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "�ص���Ժ"
      Height          =   735
      Left            =   4080
      TabIndex        =   8
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton calculation 
      Caption         =   "����"
      Height          =   735
      Left            =   6240
      TabIndex        =   4
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton btn_search 
      Caption         =   "��ѯ��ϸ"
      Height          =   735
      Left            =   8280
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox PatientNo 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   3120
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid list 
      Height          =   2655
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label totalprice 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "�ܷ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "�����ѯ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DbConnection As New ADODB.Connection

Private Sub calculation_Click()
    DbConnection.ConnectionString = "Provider=SQLOLEDB.1;Server=(local);Database=HOSBASE2019;Integrated Security=sspi"
    List.Enabled = False
 Dim sqlCommand As String
    Dim recordSet As ADODB.recordSet
       sqlCommand = "select SUM(Price*counts) as total from [" + PatientNo + "]"
    DbConnection.Open
    Set recordSet = DbConnection.Execute(sqlCommand)
            If Not recordSet.EOF Then
   TotalPrice.Caption = recordSet.Fields("total")
   ��Ժ����.TotalPrice.Caption = TotalPrice.Caption
        End If
    DbConnection.Close
    
End Sub


Private Sub Command1_Click()
����.Show
�����ѯ.Hide

End Sub

Private Sub Command2_Click()
��Ժ����.Show
�����ѯ.Hide
End Sub

'˽�з������������룻
Private Sub Form_Load()
    DbConnection.ConnectionString = "Provider=SQLOLEDB.1;Server=(local);Database=HOSBASE2019;Integrated Security=sspi"
    List.Enabled = False
End Sub

'˽�з�����������밴ť��
Private Sub btn_search_Click()
    Dim sqlCommand As String
    Dim recordSet As New ADODB.recordSet
    sqlCommand = "select * from [" + PatientNo + "];"
    DbConnection.Open
    DbConnection.CursorLocation = adUseClient
    recordSet.Open sqlCommand, DbConnection
    If Not recordSet.EOF Then
        Set List.DataSource = recordSet
        List.Columns("MedicineName").Caption = "ҩ��"
        List.Columns("Price").Caption = "�۸�"
        List.Columns("Factory").Caption = "����"
        List.Columns("Counts").Caption = "����"
        List.Enabled = True
    End If
    Set recordSet = Nothing
    Set DbConnection = Nothing
End Sub

'˽�з����������������
Private Sub dgd_Course_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    List.Col = 1
    txb_CurrentCourseName.Text = dgd_Course.Text
End Sub

