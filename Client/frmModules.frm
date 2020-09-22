VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmModules 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Module Administrator"
   ClientHeight    =   3255
   ClientLeft      =   210
   ClientTop       =   825
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwModules 
      Height          =   2955
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   5212
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Status"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Module"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Location"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Right Click On A Highlighted Module To Install Or Un-Install"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   5055
   End
   Begin VB.Menu mnuModule 
      Caption         =   "Module"
      Begin VB.Menu mnuInstall 
         Caption         =   "Install Module"
      End
      Begin VB.Menu mnuUninstall 
         Caption         =   "Un-Install Module"
      End
   End
End
Attribute VB_Name = "frmModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
' Load list of modules into Listview
Dim myPath As String
Dim myName As String

' Set path to look into the 'modules' dir for all
' dll's
myPath = App.Path & "\modules\*.dll"
' Get first file
myName = Dir(myPath)

' Loop through each file in directory
Do While myName <> ""
    DoEvents
    If myName <> "." And myName <> ".." Then
        ' Get the modules comments
        GetModuleInfo App.Path & "\modules\" & myName
    End If
    ' Get next file
    myName = Dir
Loop

End Sub

Private Sub lvwModules_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

' If user right clicks an item, show Module
' administration menu.
If Button = 2 Then
    PopupMenu mnuModule
End If

End Sub

Private Sub mnuInstall_Click()
' Install the selected module
Dim strModuleToInstall As String
Dim blnRegister As Boolean
Dim msgError As String
Dim xmlDoc As MSXML.DOMDocument
Dim xmlRoot As MSXML.IXMLDOMNode
Dim xmlNode As MSXML.IXMLDOMNode

' If module already isntalled skip isntallation process
If lvwModules.SelectedItem = "Installed" Then
    Exit Sub
End If

strModuleToInstall = lvwModules.SelectedItem.SubItems(1) & ".clsVegaMod"

' Register it if not registered
blnRegister = RegSvr32(lvwModules.SelectedItem.SubItems(3), False)
If blnRegister Then
    ' Append new module node to vegamodules.xml
    Set xmlDoc = New MSXML.DOMDocument
    xmlDoc.async = False
    xmlDoc.Load App.Path & "\modules\vegamodules.xml"
    Set xmlRoot = xmlDoc.selectSingleNode("VegaCOMM/Module")
    Set xmlNode = xmlRoot.cloneNode(True)
    xmlDoc.firstChild.appendChild xmlNode
    xmlDoc.selectSingleNode("VegaCOMM/Module/@class").Text = strModuleToInstall
    xmlDoc.save App.Path & "\modules\vegamodules.xml"
    
    ' Update status in list
    lvwModules.SelectedItem.Text = "Installed"
    
    ' Add module to module collection
    Set ptr.objModule = CreateObject(strModuleToInstall)
    ptr.ModuleName = lvwModules.SelectedItem.SubItems(1)
    colModules.Add ptr, strModuleToInstall
    Set ptr = Nothing
    
    ' Add module to Modules menu
    lngMnuCount = lngMnuCount + 1
    Load frmVegaClient.mnuInstall(lngMnuCount)
    frmVegaClient.mnuInstall(lngMnuCount).Caption = lvwModules.SelectedItem.SubItems(1)
    frmVegaClient.mnuInstall(lngMnuCount).Enabled = True
Else
    ' There was a problem registering
    msgError = MsgBox("Unable To Install Module!", vbCritical)
End If

Set xmlRoot = Nothing
Set xmlNode = Nothing
Set xmlDoc = Nothing

End Sub

Private Sub mnuUninstall_Click()
' Un-Install selected module
Dim xmlDoc As MSXML.DOMDocument
Dim xmlRoot As MSXML.IXMLDOMNode
Dim xmlNode As MSXML.IXMLDOMNode
Dim strModule As String

On Error GoTo errhandler

strModule = lvwModules.SelectedItem.SubItems(1) & ".clsVegaMod"

' First lets unregister the selected module
RegSvr32 lvwModules.SelectedItem.SubItems(3), True

' Remove the module node from vegamodules.xml
Set xmlDoc = New MSXML.DOMDocument
xmlDoc.async = False
xmlDoc.Load App.Path & "\modules\vegamodules.xml"
Set xmlRoot = xmlDoc.selectSingleNode("VegaCOMM")
For Each xmlNode In xmlRoot.childNodes
    If xmlNode.selectSingleNode("@class").Text = strModule Then
        xmlDoc.firstChild.removeChild xmlNode
        xmlDoc.save App.Path & "\modules\vegamodules.xml"
        Exit For
    End If
Next xmlNode
    
' Change the Modules status on the Module list form
lvwModules.SelectedItem.Text = "Not Installed"

' Remove module from module collection
colModules.Remove strModule

Set xmlRoot = Nothing
Set xmlNode = Nothing
Set xmlDoc = Nothing
Exit Sub

errhandler:

If Err.Number = 5 Then
    ' Module is already un-installed
    Msbox "Module already un-installed.", mbBlue, mbOkOnly, mbExclamation, "VegaCOMM Module Error"
    Resume Next
End If

End Sub
