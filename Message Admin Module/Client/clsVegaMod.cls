VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVegaMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' Declaration for object that will hold instance
' to the UsersModule. This is optional
Dim objUsers As Object
Dim UserProps As Object


Public Sub ProcessData(strData As String, blnMatch As Boolean, strLogEvent As String, lngSocket As Long)
' This is the procedure called by VegaCOMM
' when the socket receives data and then passes
' into this procedure for processing
' Parameters:
' strData - The incoming data for processing from
' the socket
' blnMatch - Set this parameter to true if the
' incoming data has a sendcode this module will
' will use
' strLogEvent - Text VegaCOMM will write to the servers
' Event Log
' lngSocket - The socket number the incoming data



End Sub

Public Sub StartModule()
' Used by VegaCOMM to start the module
' from the Module menu on the client and server
Load Form1
Form1.Show

End Sub


Private Sub Class_Initialize()
' Optional. Get current instance of Users Module
'Set objUsers = GetObject(, "UserModule.clsUserMod")
' Optional. Get colelction Properties list class
'Set UserProps = CreateObject("UserModules.clsUserProperties")
' Optional. Get Sockte Class isntance for Sending data


End Sub

Private Sub Class_Terminate()
' Optional. close out pointers to these objects
' from your module.
'Set objUsers = Nothing
'Set UserProps = Nothing
Set VSock = Nothing

End Sub
