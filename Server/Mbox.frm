VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00A0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   146
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   307
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox But1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   2
      Left            =   2565
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   7
      Top             =   1755
      Width           =   1140
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "No way !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1140
      End
   End
   Begin VB.PictureBox But1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   1
      Left            =   1350
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   5
      Top             =   1755
      Width           =   1140
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "No way !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1140
      End
   End
   Begin VB.PictureBox But1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   135
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   3
      Top             =   1755
      Width           =   1140
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "No way !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1140
      End
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   60
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   0
      Top             =   0
      Width           =   4560
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0F0C0&
         Height          =   195
         Left            =   60
         TabIndex        =   2
         Top             =   0
         Width           =   4470
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   45
      Top             =   1035
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mbox.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mbox.frx":0E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mbox.frx":1CA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mbox.frx":2AFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mbox.frx":3950
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mbox.frx":565C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mbox.frx":64B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mbox.frx":7084
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mbox.frx":7960
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3960
      Top             =   810
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   405
      Width           =   3735
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "MBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bord1&, Bord2&

Private Sub But1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
But1(Index).Line (0, 0)-(But1(Index).Width - 1, But1(Index).Height - 1), Bord2, B
But1(Index).Line (0, But1(Index).Height - 1)-(But1(Index).Width, But1(Index).Height - 1), Bord1
But1(Index).Line (But1(Index).Width - 1, 0)-(But1(Index).Width - 1, But1(Index).Height - 1), Bord1
But1(Index).Move (But1(Index).Left + 1), 118
End Sub

Private Sub But1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
But1(Index).Line (0, 0)-(But1(Index).Width - 1, But1(Index).Height - 1), Bord1, B
But1(Index).Line (0, But1(Index).Height - 1)-(But1(Index).Width, But1(Index).Height - 1), Bord2
But1(Index).Line (But1(Index).Width - 1, 0)-(But1(Index).Width - 1, But1(Index).Height - 1), Bord2
But1(Index).Move (But1(Index).Left - 1), 117
mbReturn = Index
MBox.Hide
End Sub

Private Sub Form_Activate()
Bord1 = But1(0).Point(0, 0)
Bord2 = But1(0).Point(But1(0).Width - 1, But1(0).Height - 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)

Set MBox = Nothing

End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
But1(Index).Line (0, 0)-(But1(Index).Width - 1, But1(Index).Height - 1), Bord2, B
But1(Index).Line (0, But1(Index).Height - 1)-(But1(Index).Width, But1(Index).Height - 1), Bord1
But1(Index).Line (But1(Index).Width - 1, 0)-(But1(Index).Width - 1, But1(Index).Height - 1), Bord1
But1(Index).Move (But1(Index).Left + 1), 118

End Sub

Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
But1(Index).Line (0, 0)-(But1(Index).Width - 1, But1(Index).Height - 1), Bord1, B
But1(Index).Line (0, But1(Index).Height - 1)-(But1(Index).Width, But1(Index).Height - 1), Bord2
But1(Index).Line (But1(Index).Width - 1, 0)-(But1(Index).Width - 1, But1(Index).Height - 1), Bord2
But1(Index).Move (But1(Index).Left - 1), 117
mbReturn = Index
MBox.Hide
End Sub
