Attribute VB_Name = "modVega"
Option Explicit

Public VSock As clsSocket
Public Users As Object
Public propUsers As Object
Public Modules() As Object
Public colModules As New Collection
Public lngMnuCount As Long
Public ptr As New clsModuleCollection

Sub Main()
' Load the main form and the Socket class

Load frmVega
frmVega.Show
Set VSock = CreateObject("VegaCOMM.clsSocket")


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

On Error GoTo errhandler

' Load user dll to get it into the ROT
Set Users = CreateObject("UserModule.clsUserMod")
Set propUsers = CreateObject("UserModule.clsUserProperties")

' See if users module loaded successfully
If Users Is Nothing Or propUsers Is Nothing Then
    ' Not loaded. Inform user to reinstall VegaCOMM and close
    ' VegaCOMM since the users module is necessary for
    ' operation.
    strMsg = MsgBox("There was a problem loading the Users Module. Please re-install VegaCOMM!", vbCritical)
    'End
End If

' Search modules list(vegamodules.xml) and initialize
' isntalled modules here
' Open installed module list xml document
Set xmlDoc = New MSXML.DOMDocument
xmlDoc.async = False
xmlDoc.Load App.Path & "\modules\vegamodules.xml"
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
            ' Add to output
            With frmVega.VOutput
                .SelColor = vbRed
                .SelBold = 0
                .SelText = strModuleName & " Module could not be loaded...." & vbCrLf
            End With
        Else
            ' get module name
            vntName = Split(strModuleName, ".")
            ptr.ModuleName = vntName(0)
            colModules.Add ptr, strModuleName
            
            ' Add to output
            With frmVega.VOutput
                .SelColor = vbBlue
                .SelBold = 0
                .SelText = strModuleName & " Module loaded....." & vbCrLf
            End With
            ' Add module to Modules menu
            lngMnuCount = lngMnuCount + 1
            Load frmVega.mnuInstall(lngMnuCount)
            frmVega.mnuInstall(lngMnuCount).Caption = vntName(0)
            frmVega.mnuInstall(lngMnuCount).Enabled = True
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

Public Function TimeStamp() As String

' Format Now for adding to server output
TimeStamp = "(" & Now() & ")  "

End Function

Public Function GetChannelList() As String

' Returns a list of chat channels
Dim xmlDoc As MSXML.DOMDocument
Dim xmlRoot As MSXML.IXMLDOMNode
Dim xmlNode As MSXML.IXMLDOMNode
Dim strChannelList As String

' Open chat channel list xml document
Set xmlDoc = New MSXML.DOMDocument
xmlDoc.async = False
xmlDoc.Load App.Path & "\system\chatchannels.xml"
Set xmlRoot = xmlDoc.selectSingleNode("ChatChannels")

' Loop through all Channel nodes and get chat channels
For Each xmlNode In xmlRoot.childNodes
    strChannelList = strChannelList & xmlNode.selectSingleNode("@name").Text & "||"
Next xmlNode

Set xmlRoot = Nothing
Set xmlNode = Nothing
Set xmlDoc = Nothing

GetChannelList = strChannelList

End Function
