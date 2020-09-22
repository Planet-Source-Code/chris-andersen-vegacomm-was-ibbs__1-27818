VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUserMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type GUIDs
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'Declares needed to register object in the ROT
Private Const ACTIVEOBJECT_STRONG = 0
Private Const ACTIVEOBJECT_WEAK = 1
Private Declare Function CLSIDFromProgID Lib "ole32.dll" (ByVal ProgID As Long, rclsid As GUIDs) As Long
Private Declare Function CoDisconnectObject Lib "ole32.dll" (ByVal pUnk As IUnknown, pvReserved As Long) As Long
Private Declare Function RegisterActiveObject Lib "oleaut32.dll" (ByVal pUnk As IUnknown, rclsid As GUIDs, ByVal dwFlags As Long, pdwRegister As Long) As Long
Private Declare Function RevokeActiveObject Lib "oleaut32.dll" (ByVal dwRegister As Long, ByVal pvReserved As Long) As Long
Dim OLEInstance As Long


Private Sub Class_Initialize()
'This code is responsible for creating the entry in the rot

Dim mGUID As GUIDs
Dim lp As Long
OLEInstance = 0
lp = CLSIDFromProgID(StrPtr("UserModule.clsUserMod"), mGUID)
If lp = 0 Then
    lp = RegisterActiveObject(Me, mGUID, ACTIVEOBJECT_WEAK, OLEInstance)
End If

End Sub

Private Sub Class_Terminate()
'Once we are done with the main program, lets clean up the rot
'by removing the entry for our ActiveX Server

If OLEInstance <> 0 Then
    RevokeActiveObject OLEInstance, 0
End If

CoDisconnectObject Me, 0

End Sub

Public Function UserCount() As Long

' Get count of uses in users collection
UserCount = colUsers.Count

End Function

Public Sub AddUser(Socket As Long, Handle As String, Privalages As Long, IP As String)
' Add user to users collection

' Make the SockIndex the key which will always be unique.
With User
    .Handle = Handle
    .SockIndex = Socket
    .IPAddress = IP
    .UserPrivileges = Privalages
    .ChatStatus = "Off"
    .ChatChannel = "Lobby"
End With

colUsers.Add User, CStr(Socket)
Set User = Nothing

End Sub

Public Function GetHandleFromSocket(Socket As Long) As String
' Returns handle of user connected to passed socket
Dim strHandle As String

For Each User In colUsers
    DoEvents
    If User.SockIndex = Socket Then
        strHandle = User.Handle
        Exit For
    End If
Next User

Set User = Nothing
GetHandleFromSocket = strHandle

End Function

Public Sub RemoveUser(Socket As Long)
' Remove selected user from collection

colUsers.Remove CStr(Socket)

End Sub

Public Function UserCollection() As Collection

Set UserCollection = colUsers

End Function

Public Function UserList() As String

' Return a full list of logged on users delimited by ||
For Each User In colUsers
    DoEvents
    UserList = UserList & User.Handle & "||"
Next User

Set User = Nothing

End Function

Public Sub ChangeChannel(Socket As Long, Channel As String)
' Change chat channel for user on selected socket

'For Each User In colUsers
'    DoEvents
'    If User.SockIndex = Socket Then
'        User.ChatChannel = Channel
'        End Sub
'    End If
'Next User

End Sub

Public Sub ChatStatus(Socket As Long, Status As String)
' Change chat status on or off

'For Each User In colUsers
'    DoEvents
'    If User.SockIndex = Socket Then
'        User.ChatStatus = Status
'        End Sub
'    End If
'Next User

End Sub
