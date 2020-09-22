VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmVegaClient 
   Caption         =   "VegaCOMM Client Test"
   ClientHeight    =   2625
   ClientLeft      =   4665
   ClientTop       =   4170
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   6585
   Begin VB.CommandButton Command2 
      Caption         =   "Show Users"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   795
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   5355
   End
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   660
      Top             =   1860
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Incoming Data"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   1695
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "System"
      Begin VB.Menu mnuChat 
         Caption         =   "VegaCOMM Chat"
      End
   End
   Begin VB.Menu mnuModules 
      Caption         =   "Modules"
      Begin VB.Menu mnuInstall 
         Caption         =   "Module Administration"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmVegaClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SendCodes As clsSendCodes

Private Sub Command1_Click()

VSock.SendData Text1.Text

End Sub

Private Sub Command2_Click()

Load frmUsers
frmUsers.Show

End Sub

Private Sub Command3_Click()



End Sub

Private Sub mnuChat_Click()

' Load built in chat room
Load frmChat
frmChat.Show

End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
' Retrieve incoming data
Dim strData As String

sckClient.GetData strData
Text2.Text = strData

' Start nw isntance of sendcode class and run
' PrcessSendCode routine to process incoming data
Set SendCodes = New clsSendCodes
SendCodes.ProcessSendCode strData
Set SendCodes = Nothing

End Sub

Private Sub Form_Load()

' Load installed modules
InitializeModules

' Show login form
Load frmLogin
frmLogin.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)

' Empty form and objects from memory
Set colModules = Nothing
Set VSock = Nothing
Set frmVegaClient = Nothing
Set frmUsers = Nothing
End

End Sub

Private Sub mnuInstall_Click(Index As Integer)
' Runs StartModule function of selected module

On Error GoTo ErrHandler
If Index = 0 Then
    Load frmModules
    frmModules.Show
Else
    ' Search collection
    For Each ptr In colModules
        If ptr.ModuleName = mnuInstall(Index).Caption Then
            ' Run StartModule
            ptr.objModule.StartModule
            Exit Sub
        End If
    Next ptr
End If
Exit Sub

ErrHandler:

If Err.Number = 438 Then
    ' StartModule function not found in module
    Msbox "Module does not have a startup form.", mbBlue, mbOkOnly, mbExclamation, "VegaCOMM Module Error"
    Resume Next
End If

End Sub
