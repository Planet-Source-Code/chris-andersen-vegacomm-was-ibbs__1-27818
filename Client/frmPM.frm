VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPM 
   Caption         =   "VC Private Message"
   ClientHeight    =   2250
   ClientLeft      =   2955
   ClientTop       =   1620
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   5400
   Begin VB.CommandButton cmdSendPM 
      Caption         =   "Send PM"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1980
      Width           =   2355
   End
   Begin VB.TextBox txtPMToSend 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   1620
      Width           =   5415
   End
   Begin RichTextLib.RichTextBox rtfPMText 
      Height          =   1515
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2672
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmPM.frx":0000
   End
End
Attribute VB_Name = "frmPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This is the form used for Private Messaging a
' single user

Private Sub cmdSendPM_Click()

' Write text to PM text box
With rtfPMText
    .SelColor = &H8080&
    .SelBold = 2
    .SelText = "[" & strHandle & "]"
    .SelColor = vbBlack
    .SelBold = 0
    .SelText = txtPMToSend.Text & vbCrLf
End With

' send PM
VSock.SendData "VPM1||" & strHandle & "||" & Me.Tag & "||" & txtPMToSend

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set frmPM = Nothing

End Sub
