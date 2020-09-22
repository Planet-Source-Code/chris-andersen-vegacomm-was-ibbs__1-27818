Attribute VB_Name = "modModuleInfo"
Option Explicit
Public itmModules As ListItem

Public Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" _
    (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, _
    lpData As Any) As Long
Public Declare Function GetFileVersionInfoSize Lib "version.dll" Alias _
    "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock _
    As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 _
    As Any, ByVal lpString2 As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, _
    Source As Any, ByVal Length As Long)
Public Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersion As Long
    
    dwFileVersionMS As Long
    dwFileVersionLS As Long
    dwProductVersionMS As Long
    dwProductVersionLS As Long
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type
Public Const VFT_APP = &H1
Public Const VFT_DLL = &H2
Public Const VFT_DRV = &H3
Public Const VFT_VXD = &H5

' HIWORD and LOWORD are API macros defined below.
Public Function HIWORD(ByVal dwValue As Long) As Long

Dim hexstr As String
hexstr = Right("00000000" & Hex(dwValue), 8)
HIWORD = CLng("&H" & Left(hexstr, 4))

End Function
Public Function LOWORD(ByVal dwValue As Long) As Long

Dim hexstr As String
hexstr = Right("00000000" & Hex(dwValue), 8)
LOWORD = CLng("&H" & Right(hexstr, 4))

End Function

' This nifty subroutine swaps two byte values without needing a buffer variable.
' This technique, which uses Xor, works as long as the two values to be swapped are
' numeric and of the same data type (here, both Byte).
Public Sub SwapByte(byte1 As Byte, byte2 As Byte)

byte1 = byte1 Xor byte2
byte2 = byte1 Xor byte2
byte1 = byte1 Xor byte2

End Sub

' This function creates a hexadecimal string to represent a number, but it
' outputs a string of a fixed number of digits.  Extra zeros are added to make
' the string the proper length.  The "&H" prefix is not put into the string.
Public Function FixedHex(ByVal hexval As Long, ByVal nDigits As Long) As String

FixedHex = Right("00000000" & Hex(hexval), nDigits)

End Function

Public Sub GetModuleInfo(strFile As String)
' use this Subroutine to get extended information
' for a module including comments added by the
' author.
Dim vffi As VS_FIXEDFILEINFO  ' version info structure
Dim buffer() As Byte          ' buffer for version info resource
Dim pData As Long             ' pointer to version info data
Dim nDataLen As Long          ' length of info pointed at by pData
Dim cpl(0 To 3) As Byte       ' buffer for code page & language
Dim cplstr As String          ' 8-digit hex string of cpl
Dim dispstr As String         ' string used to display version information
Dim retval As Long            ' generic return value
Dim strInternalName As String
' First, get the size of the version info resource.  If this function fails, then
' identify that the file isn't a 32-bit executable/DLL/etc.
nDataLen = GetFileVersionInfoSize(strFile, pData)
If nDataLen = 0 Then
    Debug.Print "Not a 32-bit executable!"
    Exit Sub
End If

' Make the buffer large enough to hold the version info resource.
ReDim buffer(0 To nDataLen - 1) As Byte
' Get the version information resource.
retval = GetFileVersionInfo(strFile, 0, nDataLen, buffer(0))
' Get a pointer to a structure that holds a bunch of data.
retval = VerQueryValue(buffer(0), "\", pData, nDataLen)
' Copy that structure into the one we can access.
CopyMemory vffi, ByVal pData, nDataLen
' Display the full version number of the file.
dispstr = Trim(Str(HIWORD(vffi.dwFileVersionMS))) & "." & _
    Trim(Str(LOWORD(vffi.dwFileVersionMS))) & "." & _
    Trim(Str(HIWORD(vffi.dwFileVersionLS))) & "." & _
    Trim(Str(LOWORD(vffi.dwFileVersionLS)))

' Check the type of file it is
' To see if it is an allowable module
Select Case vffi.dwFileType
Case VFT_DLL
    dispstr = "Dynamic Link Library (DLL)"
Case Else
    dispstr = "Not a valid VegaCOMM Module!"
End Select

' Before reading any strings out of the resource, we must first determine the code page
' and language.  The code to get this information follows.
retval = VerQueryValue(buffer(0), "\VarFileInfo\Translation", pData, nDataLen)
' Copy that informtion into the byte array.
CopyMemory cpl(0), ByVal pData, 4
' It is necessary to swap the first two bytes, as well as the last two bytes.
SwapByte cpl(0), cpl(1)
SwapByte cpl(2), cpl(3)
' Convert those four bytes into a 8-digit hexadecimal string.
cplstr = FixedHex(cpl(0), 2) & FixedHex(cpl(1), 2) & FixedHex(cpl(2), 2) & _
    FixedHex(cpl(3), 2)
' cplstr now represents the code page and language to read strings as.
' Copy that data into a string for display.

' DLL information is allocated in memory, let
' us now add the appropriate info to the
' module listview
' Get Internal name
retval = VerQueryValue(buffer(0), "\StringFileInfo\" & cplstr & "\InternalName", _
    pData, nDataLen)
strInternalName = Space(nDataLen)
retval = lstrcpy(strInternalName, pData)

' If module installed then print installed for
' status
If ModuleInstalled(strInternalName) Then
    Set itmModules = frmModules.lvwModules.ListItems.Add(, , "Installed")
Else
    ' Module not in isntalled list. Print status
    ' of Not Installed
    Set itmModules = frmModules.lvwModules.ListItems.Add(, , "Not Installed")
End If
' Add internal name to Name column
itmModules.SubItems(1) = strInternalName

' Get Module Comments
retval = VerQueryValue(buffer(0), "\StringFileInfo\" & cplstr & "\Comments", _
    pData, nDataLen)
dispstr = Space(nDataLen)
retval = lstrcpy(dispstr, pData)
' IF comments found, add them to the Description column
If dispstr <> "" Then
    itmModules.SubItems(2) = dispstr
Else
    ' No comments found
    itmModules.SubItems(2) = "<No Comments Found>"
End If

' Add file location to Module list
itmModules.SubItems(3) = strFile

Set itmModules = Nothing

End Sub

Public Function ModuleInstalled(strModuleName As String) As Boolean
' Check if this module name passed in is installed
Dim strName As String
Dim strData As String
Dim xmlDoc As MSXML.DOMDocument
Dim xmlRoot As MSXML.IXMLDOMNode
Dim xmlNode As MSXML.IXMLDOMNode

' We have some Nulls at the end of the name
' so we trim it out
strName = Trim(strModuleName)
Do Until Asc(Right(strName, 1)) <> 0
    strName = Left(strName, Len(strName) - 1)
    DoEvents
Loop

' Open module list XML document
Set xmlDoc = New MSXML.DOMDocument
xmlDoc.async = False
xmlDoc.Load App.Path & "\modules\vegamodules.xml"
Set xmlRoot = xmlDoc.selectSingleNode("VegaCOMM")

' Loop through all Module nodes and see if module is installed
For Each xmlNode In xmlRoot.childNodes
    strData = xmlNode.selectSingleNode("@class").Text
    If strData = strName & ".clsVegaMod" Then
        ModuleInstalled = True
        Exit For
    End If
Next xmlNode

Set xmlRoot = Nothing
Set xmlNode = Nothing
Set xmlDoc = Nothing

End Function
