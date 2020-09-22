VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSendCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ProcessSendCode(strData As String)
' This sub accepts the incoming winsock data
' passes it into each installed module till it
' finds a module that uses the sendcode. This is the heart and soul
' of how VegaCOMM works.
Dim vntData As Variant
Dim strSendCode As String
Dim strFrom As String
Dim strMessage As String
Dim i As Long
Dim blnPMSession As Boolean
Dim blnMatch As Boolean

On Error GoTo ErrHandler

' Split incoming data and get the sendcode
vntData = Split(strData, "||")
strSendCode = vntData(0)

' Find out which process to perform based on
' SendCode.
Select Case strSendCode
' ********************Login SendCodes*******************8
    Case "VGetLogon1"
        ' Request from server to get logon info
        ' **Structure: VGetLogon1
        ' Send handle and password
        ' **Structure: VLogin1||Handle||Password
        VSock.SendData "VLogin1||" & frmLogin.txtLogin(1) & "||" & frmLogin.txtLogin(2)
    
    Case "VLogin2"
        ' Response from server if login is
        ' correct or incorrect
        ' **Structure: VLogin2||Correct/Incorrect
        strMessage = vntData(1)
        If strMessage = "Correct" Then
            ' Login is correct
            ' Save handle to global variable
            strHandle = frmLogin.txtLogin(1)
            ' Send message bask to server to
            ' let it know it is ready for user list
            ' **Structure: VUserLogged1||Handle
            VSock.SendData "VUserLogged1||" & strHandle
        Else
            ' Login is incorrect
        End If
    Case "VUserList1"
        ' Incoming list of all users
        ' Add to user list
        
        For i = 1 To UBound(vntData) - 1
            ' Add each user to user list
            DoEvents
            strFrom = vntData(i)
            Set tvNode = frmUsers.tvUserList.Nodes.Add("ALL", tvwChild, , strFrom)
        Next i
        
        Set tvNode = Nothing
        
        ' Change usercount
        frmUsers.tvUserList.Nodes.Item("ALL").Text = "All Users (" & frmUsers.tvUserList.Nodes.Count - 1 & ")"
    Case "VAddUser1"
        ' Adding a new single user
        ' **Structure: VAddUser1||Handle
        strFrom = vntData(1)
        Set tvNode = frmUsers.tvUserList.Nodes.Add("ALL", tvwChild, , strFrom)
        Set tvNode = Nothing
        
        ' Change usercount
        frmUsers.tvUserList.Nodes.Item("ALL").Text = "All Users (" & frmUsers.tvUserList.Nodes.Count - 1 & ")"
    
' *********************Chat SendCodes***************
    Case "VChannelList1"
        ' Incoming list of channels. Add them to
        ' chat windows channel list
        ' **Structure: VChannelList1||Channel||Channel||...
        For i = 1 To UBound(vntData) - 1
            DoEvents
            frmChat.cboChannels.AddItem vntData(i)
        Next i
        
        ' Send request to server to get users in
        ' lobby
        ' **Structure: VLobbyUsers1
        VSock.SendData "VLobbyUsers1"
    
    Case "VLobbyUsers2"
        ' Incoming list of users in lobby. Add to
        ' Chat user list
        ' **Structure: VLobbyUsers2||Handle||Handle||..
        frmChat.lstChatUsers.Clear
        
        For i = 1 To UBound(vntData) - 1
            DoEvents
            frmChat.lstChatUsers.AddItem vntData(i)
        Next i
        
    Case "VAddLobbyUser1"
        ' Add new user to lobby caht users list
        ' **Structure: VAddLobbyUser1||Handle
        frmChat.lstChatUsers.AddItem vntData(1)
    
    Case "VChatUserList1"
        ' Get incoming list of users for current channel
        ' **Structure: VChatUserList1||Handle||Handle||Handle||etc..
        ' Add users to user list
        frmChat.lstChatUsers.Clear
        
        For i = 1 To UBound(vntData) - 1
            DoEvents
            frmChat.lstChatUsers.AddItem vntData(i)
        Next i
        
        ' Send message to server to add this user
        ' to everyones user list in channel
        ' **Structure: VChatAddUser1||Handle||Channel
        VSock.SendData "VChatAddUser1||" & strHandle & "||" & strCurrentChatChannel
    
    Case "VChatAddUser2"
        ' Add this user to this users in chat list
        ' **Structure: VChatAddUser2||Handle
        
        frmChat.lstChatUsers.AddItem vntData(1)
        
    Case "VChatRemoveUser1"
        ' Remove this user from chat user list
        ' **Structure: VChatRemoveUser1||Handle
        strFrom = vntData(1)
        
        For i = 0 To frmChat.lstChatUsers.ListCount + 1
            DoEvents
            If frmChat.lstChatUsers.List(i) = strFrom Then
                frmChat.lstChatUsers.RemoveItem (i)
            End If
        Next i
        
    Case "VChat2"
        ' Incoming chat text from server
        ' Add message to chat window
        ' **Structure: VChat2||From||Message
        strFrom = vntData(1)
        strMessage = vntData(2)
        
        With frmChat.rtfChatText
            If strFrom = strHandle Then
                ' this message is from this user
                ' so show Handle on chat in
                ' gold
                .SelColor = &H8080&
            Else
                ' It is from someone else
                ' Show it blue
                .SelColor = vbBlue
            End If
            .SelBold = 2
            .SelText = "[" & strFrom & "] "
            .SelColor = vbBlack
            .SelBold = 0
            .SelText = strMessage & vbCrLf
        End With
    
    Case "VPM2"
        ' incoming Private Message
        ' **Structure: VPM2||From||Message
        strFrom = vntData(1)
        strMessage = vntData(2)
        
        ' Check if a window for the from user
        ' is already open
        For i = 0 To Forms.Count - 1
            DoEvents
            If Forms(i).Tag = strFrom Then
                ' A PM session is found. Send
                ' message to that one
                With Forms(i).rtfPMText
                    .SelColor = vbBlue
                    .SelBold = 2
                    .SelText = "[" & strFrom & "]"
                    .SelColor = vbBlack
                    .SelBold = 0
                    .SelText = strMessage & vbCrLf
                End With
                
                blnPMSession = True
                Exit For
            End If
        Next i
        
        If Not blnPMSession Then
            ' Session not found. Create new
            ' PM window.
            Set frmPMessage = New frmPM
            Load frmPMessage
            frmPMessage.Show
            frmPMessage.Caption = "VC Private Message With " & strFrom
            frmPMessage.Tag = strFrom
            ' Send message to new window
            With frmPMessage.rtfPMText
                .SelColor = vbBlue
                .SelBold = 2
                .SelText = "[" & strFrom & "]"
                .SelColor = vbBlack
                .SelBold = 0
                .SelText = strMessage & vbCrLf
            End With
        End If
        
    Case Else
        ' Sendcode not used internally
        ' Pass data into main function of each installed
        ' module (ProcessData) until a match is found and the
        ' Match parameter = TRUE
        For Each ptr In colModules
            DoEvents
            ptr.objModule.ProcessData strData, blnMatch
            'If blnMatch returns true, stop module loop
            If blnMatch Then Exit For
        Next ptr
        
End Select

Exit Sub

ErrHandler:
Select Case Err.Number
    Case 438
        Resume Next
    Case Else
        Msbox "An error has occured: " & Err.Number & " " & Err.Description, mbBlue, mbOkOnly, mbExclamation, "VegaCOMM Server Error"
        Resume Next
End Select
End Sub

