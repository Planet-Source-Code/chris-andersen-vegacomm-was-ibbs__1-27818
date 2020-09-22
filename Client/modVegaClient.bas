Attribute VB_Name = "modVegaClient"
Option Explicit

Public VSock As clsClientSocket
Public ptr As New clsModuleCol
Public colModules As New Collection
Public lngMnuCount As Long
Public strHandle As String
Public tvNode As Node
Public frmPMessage As frmPM
Public strCurrentChatChannel As String
Public strPreviousChatChannel As String

Sub Main()
' Load the main form

Load frmVegaClient
frmVegaClient.Show
Set VSock = CreateObject("VCClient.clsClientSocket")
Do Until Not (VSock Is Nothing)
    DoEvents
Loop

End Sub

Public Sub InitializeModules()
' This sub will load all included modules and
' installed third party modules
Dim strModuleName As String
Dim strMsg As String
Dim xmlDoc As MSXML.DOMDocument
Dim xmlRoot As MSXML.IXMLDOMNode
Dim xmlNode As MSXML.IXMLDOMNode
Dim vntName As Variant

'On Error GoTo errhandler

' Search modules list(vegamodules.xml) and initialize
' isntalled modules here
' Open installed module list xml document
Set xmlDoc = New MSXML.DOMDocument
xmlDoc.async = False
xmlDoc.Load App.Path & "\Modules\vegamodules.xml"
Set xmlRoot = xmlDoc.selectSingleNode("VegaCOMM")

' Loop through all Module nodes and get installed modules
For Each xmlNode In xmlRoot.childNodes
    strModuleName = xmlNode.selectSingleNode("@class").Text
    If strModuleName <> "BASE" Then
        ' This node is not the base node that is always in the
        ' document.
        Set ptr.objModule = CreateObject(strModuleName)
        
        ' If this module is invalid(not registered), do not add
        ' to Module collection
        If ptr.objModule Is Nothing Then
            strMsg = MsgBox("VegaCOMM could not initialize Module: " & strModuleName, vbCritical)
        Else
            ' get module name
            vntName = Split(strModuleName, ".")
            ptr.ModuleName = vntName(0)
            colModules.Add ptr, strModuleName
            ' Add module to Modules menu
            lngMnuCount = lngMnuCount + 1
            Load frmVegaClient.mnuInstall(lngMnuCount)
            frmVegaClient.mnuInstall(lngMnuCount).Caption = vntName(0)
            frmVegaClient.mnuInstall(lngMnuCount).Enabled = True
        End If
    
        Set ptr = Nothing
    End If
Next xmlNode

Set xmlRoot = Nothing
Set xmlNode = Nothing
Set xmlDoc = Nothing

Exit Sub

errhandler:

Resume Next

End Sub
