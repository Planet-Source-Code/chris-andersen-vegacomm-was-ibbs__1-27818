VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   Caption         =   "VegaCOMM Chat"
   ClientHeight    =   5040
   ClientLeft      =   3840
   ClientTop       =   2295
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   8085
   Begin VB.CommandButton cmdChangeChannel 
      Caption         =   "ChangeChannel"
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.ComboBox cboChannels 
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Text            =   "Lobby"
      Top             =   360
      Width           =   2055
   End
   Begin VB.ListBox lstChatUsers 
      Height          =   3765
      Left            =   6000
      TabIndex        =   3
      Top             =   780
      Width           =   2055
   End
   Begin VB.CommandButton cmdSendChat 
      Caption         =   "Send Chat Message"
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox txtChatToSend 
      Height          =   435
      Left            =   60
      TabIndex        =   1
      Top             =   4140
      Width           =   5895
   End
   Begin RichTextLib.RichTextBox rtfChatText 
      Height          =   3315
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5847
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmChat.frx":0000
   End
   Begin VB.Label lblUsers 
      Caption         =   "Users In Channel"
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblChannelList 
      Caption         =   "Select A Channel..."
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   120
      Width           =   1635
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChangeChannel_Click()

' Send channel change message to server
strPreviousChatChannel = strCurrentChatChannel
strCurrentChatChannel = cboChannels.Text

VSock.SendData "VChatChannel1||" & strPreviousChatChannel & "||" & strCurrentChatChannel

End Sub

Private Sub cmdSendChat_Click()

' Send chat message to server
VSock.SendData "VChat1||" & strHandle & "||" & strCurrentChatChannel & "||" & txtChatToSend

End Sub

Private Sub Form_Load()

cboChannels.AddItem "Lobby"
strCurrentChatChannel = "Lobby"
' Send chat on message to server
VSock.SendData "VChatStatus1||On"

End Sub

Private Sub Form_Unload(Cancel As Integer)

' Send chat off message to server
VSock.SendData "VChatStatus1||Off||" & strCurrentChatChannel
Set frmChat = Nothing

End Sub
