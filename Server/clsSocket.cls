VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSocket"
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

' Declares needed to register object in the ROT
Private Const ACTIVEOBJECT_STRONG = 0
Private Const ACTIVEOBJECT_WEAK = 1
Private Declare Function CLSIDFromProgID Lib "ole32.dll" (ByVal ProgID As Long, rclsid As GUIDs) As Long
Private Declare Function CoDisconnectObject Lib "ole32.dll" (ByVal pUnk As IUnknown, pvReserved As Long) As Long
Private Declare Function RegisterActiveObject Lib "oleaut32.dll" (ByVal pUnk As IUnknown, rclsid As GUIDs, ByVal dwFlags As Long, pdwRegister As Long) As Long
Private Declare Function RevokeActiveObject Lib "oleaut32.dll" (ByVal dwRegister As Long, ByVal pvReserved As Long) As Long

Private OLEInstance As Long

Private Sub Class_Initialize()
Dim typGUID As GUIDs
Dim lp As Long

OLEInstance = 0
' This code is responsible for creating the entry in the rot
lp = CLSIDFromProgID(StrPtr("VegaCOMM.clsSocket"), typGUID)
If lp = 0 Then
    lp = RegisterActiveObject(Me, typGUID, ACTIVEOBJECT_WEAK, OLEInstance)
End If

End Sub


Private Sub Class_Terminate()
' Once we are done with the main program, lets clean up the rot
' by removing the entry for our ActiveX Server

If OLEInstance <> 0 Then
    RevokeActiveObject OLEInstance, 0
End If

CoDisconnectObject Me, 0

End Sub

Public Sub SendData(strData As String, Socket As Long)
' Send data to selected socket

frmVega.sckServer(Socket).SendData strData

End Sub

Public Function GetIP(Socket As Long) As String
' Returns IP address of selected socket
GetIP = frmVega.sckServer(Socket).RemoteHostIP

End Function
