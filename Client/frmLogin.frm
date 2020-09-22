VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "VegaCOMM Login"
   ClientHeight    =   1650
   ClientLeft      =   4635
   ClientTop       =   1815
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   1650
   ScaleWidth      =   4950
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   315
      Left            =   60
      TabIndex        =   6
      Top             =   1320
      Width           =   1755
   End
   Begin VB.TextBox txtLogin 
      Height          =   375
      Index           =   2
      Left            =   1860
      TabIndex        =   5
      Top             =   840
      Width           =   2835
   End
   Begin VB.TextBox txtLogin 
      Height          =   375
      Index           =   1
      Left            =   1860
      TabIndex        =   4
      Top             =   480
      Width           =   2835
   End
   Begin VB.TextBox txtLogin 
      Height          =   375
      Index           =   0
      Left            =   1860
      TabIndex        =   3
      Top             =   120
      Width           =   2835
   End
   Begin VB.Label lblLogin 
      Caption         =   "Password"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   900
      Width           =   1695
   End
   Begin VB.Label lblLogin 
      Caption         =   "Handle"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   1695
   End
   Begin VB.Label lblLogin 
      Caption         =   "Server Address"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   1695
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConnect_Click()

frmVegaClient.sckClient.RemoteHost = txtLogin(0).Text
frmVegaClient.sckClient.RemotePort = 1001
frmVegaClient.sckClient.Connect

'Unload frmLogin

End Sub

Private Sub Form_Load()

txtLogin(0).Text = "127.0.0.1"

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmLogin = Nothing

End Sub
