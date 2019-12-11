VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Socket Server"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
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
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox textget 
      Height          =   1320
      Left            =   1395
      TabIndex        =   2
      Text            =   "textget"
      Top             =   1485
      Width           =   2220
   End
   Begin VB.CommandButton Command1 
      Caption         =   "退出"
      Height          =   360
      Left            =   90
      TabIndex        =   1
      Top             =   1260
      Width           =   990
   End
   Begin VB.TextBox textsend 
      Height          =   1050
      Left            =   1395
      TabIndex        =   0
      Text            =   "textsend"
      Top             =   180
      Width           =   2220
   End
   Begin MSWinsockLib.Winsock WinsockServer 
      Left            =   630
      Top             =   315
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'服务器程序使用的控件如下:
'（1）Command1：退出按钮；
'（2）textsend：发送数据文本框；
'（3）Winsockserver： 服务器Winsock；
'（4）textget ：接收数据文本框。
'服务器程序的界面如图所示?
'服务器程序的源代码如下:
Private Sub Command1_Click()
    End
End Sub
Private Sub Form_Load()
    textsend.Visible = False
    textget.Visible = False
    WinsockServer.LocalPort = 1001
    WinsockServer.Listen
End Sub

Private Sub textsend_Change()
    WinsockServer.SendData textsend.Text
End Sub
Private Sub Winsockserver_Close()
    WinsockServer.Close
    End
End Sub

Private Sub Winsockserver_ConnectionRequest(ByVal requestID As Long)
    textsend.Visible = True
    textget.Visible = True
    If WinsockServer.State <> sckClosed Then WinsockServer.Close
    WinsockServer.Accept requestID
End Sub

Private Sub Winsockserver_DataArrival(ByVal bytesTotal As Long)
    Dim tmpstr As String
    
    WinsockServer.GetData tmpstr
    textget.Text = tmpstr
End Sub
