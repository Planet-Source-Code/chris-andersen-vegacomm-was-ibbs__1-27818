VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmVega 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "VegaCOMM Test"
   ClientHeight    =   3375
   ClientLeft      =   3120
   ClientTop       =   5400
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   1395
      Left            =   1680
      Picture         =   "frmVega.frx":0000
      ScaleHeight     =   1335
      ScaleWidth      =   5955
      TabIndex        =   1
      Top             =   1920
      Width           =   6015
   End
   Begin RichTextLib.RichTextBox VOutput 
      Height          =   1815
      Left            =   1680
      TabIndex        =   0
      Top             =   60
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3201
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmVega.frx":52E5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   1200
      Top             =   2940
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuModules 
      Caption         =   "Modules"
      Begin VB.Menu mnuInstall 
         Caption         =   "Module Administration"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmVega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SendCodes As clsSendCodes
Public lngSockCount As Long

Private Sub Form_Load()

' Set up winsock to listen for connections.
sckServer(0).LocalPort = 1001
sckServer(0).Listen

' Show startup information for output box
With VOutput
    .SelColor = vbBlack
    .SelBold = 2
    .SelText = "VegaCOMM Server Version: "
    .SelColor = &H8080&
    .SelText = App.Major & "." & App.Minor & " Revision " & App.Revision & vbCrLf
    
    ' Load installed modules
    InitializeModules
    
    .SelColor = vbBlue
    .SelBold = 2
    .SelText = TimeStamp & "Server Status: Test Mode" & vbCrLf
    .SelColor = vbBlack
    .SelBold = 0
    .SelText = "This server is the VegCOMM server Alpha test. Testing all functionality before building beta."
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
' Program ending. Clean up objects from memory

Set colModules = Nothing
Set Users = Nothing
Set VSock = Nothing
Set frmVega = Nothing
Set Users = Nothing
Set propUsers = Nothing
End

End Sub

Private Sub mnuInstall_Click(Index As Integer)
' Runs StartModule function of selected module

On Error GoTo errhandler
If Index = 0 Then
    ' Module Administration was selected
    Load frmModules
    frmModules.Show
Else
 ' A module was selected
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

errhandler:

If Err.Number = 438 Then
    ' StartModule function not found in module
    Msbox "Module does not have a startup form.", mbBlue, mbOkOnly, mbExclamation, "VegaCOMM Module Error"
    Resume Next
End If

End Sub

Private Sub sckServer_Close(Index As Integer)
' Socket has been closed

On Error GoTo errhandler

' Remove disconnected socket/user from collection
Users.RemoveUser CStr(Index)

' Close the connection to free up for next new
' connection
sckServer(Index).Close
errhandler:

Resume Next

End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
' Listen for new connections to the server
Dim i As Long

'Check to see if any open socks are being used
'and use unused ones for new connections to save memory
For i = 1 To sckServer.UBound
    If sckServer(i).State <> 7 Then
        sckServer(i).Close
        sckServer(i).Accept requestID
        ' Send Message to client to request Logon
        ' information
        ' **Structure: VGetLogon1
        VSock.SendData "VGetLogon1", i
        Exit Sub
    End If
Next

'If all open socks are being used, create a new one.
lngSockCount = lngSockCount + 1
Load sckServer(lngSockCount)
sckServer(lngSockCount).Accept requestID
' Send message to client requesting logon information
' **Structure: VGetLogon1
VSock.SendData "VGetLogon1", lngSockCount

End Sub


Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
' Get incoming data and pass it to a new
' instance or the Sendcode processing class
Dim strData As String
Dim lngSocket As Long

lngSocket = Index

sckServer(Index).GetData strData

frmVega.VOutput.SelText = frmVega.VOutput.SelText & strData

Set SendCodes = New clsSendCodes
SendCodes.ProcessSendCode strData, lngSocket
Set SendCodes = Nothing

End Sub
