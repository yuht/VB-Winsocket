VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Socket Client"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox textget 
      Height          =   870
      Left            =   4590
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1125
      Width           =   1770
   End
   Begin VB.TextBox Textsend 
      Height          =   870
      Left            =   4500
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   90
      Width           =   1770
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1485
      TabIndex        =   3
      Text            =   "192.168.1.131"
      Top             =   225
      Width           =   1680
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   360
      Left            =   1755
      TabIndex        =   1
      Top             =   2070
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�˳�"
      Height          =   360
      Left            =   405
      TabIndex        =   0
      Top             =   2025
      Width           =   990
   End
   Begin MSWinsockLib.Winsock Winsockclient 
      Left            =   990
      Top             =   990
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   195
      Left            =   585
      TabIndex        =   2
      Top             =   225
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�ͻ�������ʹ�õĿؼ�����:
'��1��Command1���˳���ť��
'��2��Command2�����Ӱ�ť��
'��3��Winsockclient���ͻ�Winsock��
'��4��Text1���������ı���
'��5��Textsend�����������ı���
'��6��Textget�����������ı���
'�ͻ��������Դ��������:
Private Sub Command1_Click()
    End
End Sub
Private Sub Command2_Click()
    Winsockclient.Connect
End Sub
Private Sub Form_Load()
    Textsend.Visible = False
    textget.Visible = False
    Winsockclient.RemotePort = 1001
    Winsockclient.RemoteHost = "192.168.1.131"
End Sub
Private Sub Text1_Change()
    Winsockclient.RemoteHost = Text1.Text
End Sub
Private Sub textsend_Change()
    Winsockclient.SendData Textsend.Text
End Sub
Private Sub Winsockclient_Close()
    Winsockclient.Close
    End
End Sub
Private Sub winsockclient_Connect()
    Textsend.Visible = True
    textget.Visible = True
    Command2.Visible = False
End Sub
Private Sub winsockclient_DataArrival(ByVal bytesTotal As Long)
    Dim tmpstr As String
    Winsockclient.GetData tmpstr
    textget.Text = tmpstr
End Sub
