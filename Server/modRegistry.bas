Attribute VB_Name = "modRegistry"
Option Explicit
' This code is used to register/unregister a
' module.
Declare Function LoadLibraryRegister _
        Lib "kernel32" Alias "LoadLibraryA" _
        (ByVal lpLibFileName As String) As Long
  
Declare Function CreateThreadForRegister _
        Lib "kernel32" Alias "CreateThread" _
        (lpThreadAttributes As Any, ByVal dwStackSize _
        As Long, ByVal lpStartAddress As Long, _
        ByVal lParameter As Long, ByVal dwCreationFlags As Long, _
        lpThreadID As Long) As Long
   
Declare Function WaitForSingleObject _
        Lib "kernel32" (ByVal hHandle As Long, _
        ByVal dwMilliseconds As Long) As Long
   
Declare Function GetProcAddressRegister _
        Lib "kernel32" Alias "GetProcAddress" _
        (ByVal hModule As Long, ByVal lpProcName As String) _
        As Long

Declare Function FreeLibraryRegister Lib _
        "kernel32" Alias "FreeLibrary" (ByVal hLibModule As Long) _
        As Long

Declare Function CloseHandle Lib "kernel32" _
        (ByVal hObject As Long) As Long

Declare Function GetExitCodeThread Lib "kernel32" _
        (ByVal hThread As Long, lpExitCode As Long) As Long

Declare Sub ExitThread Lib "kernel32" _
        (ByVal dwExitCode As Long)

Public Function RegSvr32(ByVal FileName As String, bUnReg As _
   Boolean) As Boolean
' Pass in True for bUnReg to unregister, false to register

Dim lLib As Long
Dim lProcAddress As Long
Dim lThreadID As Long
Dim lSuccess As Long
Dim lExitCode As Long
Dim lThread As Long
Dim bAns As Boolean
Dim sPurpose As String

sPurpose = IIf(bUnReg, "DllUnregisterServer", _
  "DllRegisterServer")
' File not found
If Dir(FileName) = "" Then Exit Function

lLib = LoadLibraryRegister(FileName)
' could not load file
If lLib = 0 Then Exit Function

lProcAddress = GetProcAddressRegister(lLib, sPurpose)

If lProcAddress = 0 Then
  ' Not an ActiveX Component
   FreeLibraryRegister lLib
   Exit Function
Else
   lThread = CreateThreadForRegister(ByVal 0&, 0&, ByVal lProcAddress, ByVal 0&, 0&, lThread)
   If lThread Then
        lSuccess = (WaitForSingleObject(lThread, 10000) = 0)
        If Not lSuccess Then
           Call GetExitCodeThread(lThread, lExitCode)
           Call ExitThread(lExitCode)
           bAns = False
           Exit Function
        Else
           bAns = True
        End If
        CloseHandle lThread
        FreeLibraryRegister lLib
   End If
End If
    RegSvr32 = bAns
End Function


