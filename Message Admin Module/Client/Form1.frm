VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Message To Admin Test"
   ClientHeight    =   1290
   ClientLeft      =   3435
   ClientTop       =   2010
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   1290
   ScaleWidth      =   6585
   Begin VB.CommandButton Command1 
      Caption         =   "Send Message"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   780
      Width           =   2715
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Message:"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   1515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

VSock.SendData "MsgAdmin1||" & Text1.Text

End Sub

Private Sub Form_Load()
Set VSock = GetObject(, "VCClient.clsClientSocket")

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set Form1 = Nothing

End Sub
