VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ��ҩ�Ǽ� 
   Caption         =   "��ҩ�Ǽ�"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   10215
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox price 
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   615
      Left            =   360
      TabIndex        =   12
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�Ǽ�"
      Height          =   615
      Left            =   7320
      TabIndex        =   11
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox counts 
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   4320
      Width           =   1815
   End
   Begin VB.ComboBox factory 
      Height          =   300
      Left            =   7440
      TabIndex        =   8
      Text            =   "fjtcm"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.ComboBox medicinename 
      Height          =   300
      ItemData        =   "��ҩ�Ǽ�.frx":0000
      Left            =   1560
      List            =   "��ҩ�Ǽ�.frx":0007
      TabIndex        =   5
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ѯ"
      Height          =   495
      Left            =   8280
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid list 
      Height          =   2055
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3625
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
   Begin VB.TextBox PatientNo 
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "����"
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "������"
      Height          =   255
      Left            =   6480
      TabIndex        =   7
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "�۸�"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "ҩ��"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "����"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "��ҩ�Ǽ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DbConnection As New ADODB.Connection
Private Sub Command1_Click()
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
End Sub

Private Sub Command2_Click()
DbConnection.ConnectionString = "Provider=SQLOLEDB.1;Server=PC-20180428BGOQ;Database=HOSBASE2019;Trusted_Connection=no;Uid=L0G1n;Password=1qaz!QAZ"
    Dim sqlCommand As String
    Dim recordSet As ADODB.recordSet
    sqlCommand = "INSERT INTO [" + PatientNo + "] (MedicineName,Price,Factory,Counts) VALUES ('" + medicinename + "'," + price + ",'" + factory + "'," + counts + ");"
    DbConnection.Open
       DbConnection.Execute sqlCommand, rowAffected
    If rowAffected = 1 Then
        MsgBox "���³ɹ�", vbInformation, "��ʾ"
    Else
        MsgBox "����ʧ��", vbCritical, "����"
    End If
    DbConnection.Close
End Sub

Private Sub Command3_Click()
����.Show
��ҩ�Ǽ�.Hide
End Sub

Private Sub Form_Load()
DbConnection.ConnectionString = "Provider=SQLOLEDB.1;Server=PC-20180428BGOQ;Database=HOSBASE2019;Trusted_Connection=no;Uid=L0G1n;Password=1qaz!QAZ"
End Sub
