VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsers 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Users Online"
   ClientHeight    =   5385
   ClientLeft      =   1755
   ClientTop       =   3810
   ClientWidth     =   2265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   2265
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView tvUserList 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   9340
      _Version        =   393217
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

' Add Main Nodes
Set tvNode = tvUserList.Nodes.Add(, , "ALL", "All Users (0)")
Set tvNode = Nothing


End Sub

Private Sub Form_Unload(Cancel As Integer)

Cancel = -1
frmUsers.Hide

'Set frmUsers = Nothing

End Sub

Private Sub tvUserList_DblClick()

' Open Private MEssage for selected user

' If item double clicked is not All Users parent node
' then PM selected user
If tvUserList.SelectedItem.Key <> "ALL" Then
    Set frmPMessage = New frmPM
    Load frmPMessage
    frmPMessage.Show
    frmPMessage.Caption = "VC Private Message With " & tvUserList.SelectedItem.Text
    ' set Tag to whom user is PMing for this
    ' window so we can search by Tag to put
    ' incoming pm text to correct window when
    ' more than one PM session is open on client
    frmPMessage.Tag = tvUserList.SelectedItem.Text
End If

End Sub
