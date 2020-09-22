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

Public Sub ProcessSendCode(strData As String, Socket As Long)
' This sub accepts the incoming winsock data
' passes it into each installed module till it
' finds a module that uses the sendcode. This is the heart and soul
' of how VegaCOMM works.
Dim blnMatch As Boolean
Dim strLogEvent As String
Dim vntSplitData
Dim strSendCode As String
Dim vcConn As ADODB.Connection
Dim vcRS As ADODB.Recordset
Dim UserConnString As String
Dim strHandle As String
Dim strPassword As String
Dim lngPrivalages As Long
Dim errTrack As String
Dim strChatStatus As String
Dim strChannel As String
Dim strFrom As String
Dim strMessage As String
Dim strNewMessage As String
Dim strTo As String
Dim strOldChannel As String
Dim strIP As String

On Error GoTo errhandler

' Split up incoming data
vntSplitData = Split(strData, "||")
strSendCode = vntSplitData(0)
'MsgBox Socket

' Check if first instance of strSendCode
' returns a SendCode we use internally
Select Case strSendCode
'**************************Login SendCodes******************************
    Case "VLogin1"
        ' SendCode for user login
        ' Check if user's password is valid
        ' If valid add user to user list in User module
        ' **Structure: VLogin1||Handle||Password
        
        ' Put split data into variables
        strHandle = vntSplitData(1)
        strPassword = vntSplitData(2)
                
        ' Open system database and compare login's to find if login is
        ' correct
        UserConnString = "Provider=Microsoft.JET.OLEDB.3.51;Data Source=" & App.Path & "\System\Users.mdb"
        Set vcConn = New ADODB.Connection
        vcConn.CursorLocation = adUseClient
        vcConn.ConnectionString = UserConnString
        vcConn.Open
        
        ' Run query searching for matching Handle and password
        Set vcRS = New ADODB.Recordset
        vcRS.Open "Select * FROM Users WHERE Handle='" & strHandle & "' AND Password='" & strPassword & "'", vcConn, adOpenKeyset, adLockReadOnly
    
        ' If login correct send login ok message otherwise
        ' send login failed
        If vcRS.RecordCount > 0 Then
            ' login correct
            ' Send message back to client
            ' **Structure: VLogin2||Correct
            VSock.SendData "VLogin2||Correct", Socket
        Else
            ' login incorrect
            ' Send message back to client
            ' **Structure: VLogin2||Incorrect
            VSock.SendData "VLogin2||Incorrect", Socket
            Exit Sub
        End If
        ' Get Privalages from database
        lngPrivalages = vcRS!Privalages
        
        ' Close out database stuff
        vcRS.Close
        vcConn.Close
        Set vcRS = Nothing
        Set vcConn = Nothing
                
        ' Add User information to collection
        strIP = VSock.GetIP(Socket)
        Users.AddUser Socket, strHandle, lngPrivalages, strIP
    
    Case "VUserLogged1"
        ' Received from client after successfull login
        ' Update User list for all users and send full user list
        ' to new client
        ' **Structure: VUserLogged1||Handle
        
        ' Loop through each user in User Module
        For Each propUsers In Users.UserCollection
            DoEvents
            If propUsers.SockIndex <> Socket Then
                ' Not this user. Send new users handle to add to user
                ' list for this socket's client
                ' **Structure: VAddUser1||Handle
                strHandle = vntSplitData(1)
                ' Send message
                VSock.SendData "VAddUser1||" & strHandle, propUsers.SockIndex
            Else
                Dim strUsers As String
                ' Send full user logged on list to new client
                ' **Structure: VUserList1||Handle||Handle||Handle||etc
                strUsers = Users.UserList
                
                ' Send message
                VSock.SendData "VUserList1||" & strUsers, Socket
            End If
        Next propUsers

' **************************Chat SendCodes******************************
    Case "VChatStatus1"
        ' Received from client telling if VegaCOMM's built in chat window
        ' has been opened and closed to change the user's chat status
        ' to on or off. Only send chat data to users with status of on
        ' **Structure: VChatStatus1||On/Off||If Off then this will have Channel
        strChatStatus = vntSplitData(1)
        
        For Each propUsers In Users.UserCollection
            ' When User Module finds user for this socket
            ' change status for them
            DoEvents
            If propUsers.SockIndex = Socket Then
                propUsers.ChatStatus = strChatStatus
                ' Set chat channel to default Lobby if status is off
                If strChatStatus = "Off" Then
                    propUsers.ChatChannel = "Lobby"
                End If
                
                Exit For
            End If
        Next propUsers
        
        If strChatStatus = "Off" Then
            ' Send RemoveUser message to eveyone in this users channel
            ' **Structure: VChatRemoveUser1||Handle
            ' Get handle of this socket
            strHandle = Users.GetHandleFromSocket(Socket)
            
            For Each propUsers In Users.UserCollection
                DoEvents
                If propUsers.ChatStatus = "On" And propUsers.ChatChannel = vntSplitData(2) Then
                    VSock.SendData "VChatRemoveUser1||" & strHandle, propUsers.SockIndex
                End If
            Next propUsers
        End If
        
        ' Check if chat is on
        If strChatStatus = "On" Then
            ' If chat is on, send Chat channel list
            ' **Structure: VChannelList1||Channel||Channel||etc...
            strMessage = GetChannelList
            VSock.SendData "VChannelList1||" & strMessage, Socket
        End If
    
    Case "VLobbyUsers1"
        ' after client receives channel list, this message comes to server
        ' to get list of all users in Lobby Chat
        ' **Structure: VLobbyUsers1
        strHandle = Users.GetHandleFromSocket(Socket)
        
        For Each propUsers In Users.UserCollection
            ' Build list of user in lobby
            DoEvents
            If propUsers.ChatChannel = "Lobby" And propUsers.ChatStatus = "On" Then
                strMessage = strMessage & propUsers.Handle & "||"
            End If
            
            If propUsers.ChatChannel = "Lobby" And propUsers.ChatStatus = "On" And propUsers.SockIndex <> Socket Then
                ' Send message to users in lobby to add this users to
                ' lobby list
                ' **Structure: VAddLobbyUser1||Handle
                VSock.SendData "VAddLobbyUser1||" & strHandle, propUsers.SockIndex
            End If
        Next propUsers
        
        ' Send list to this user
        ' **Structure: VLobbyUsers2
        VSock.SendData "VLobbyUsers2||" & strMessage, Socket
        
    Case "VChatChannel1"
        ' Changes selected channel for user so he will only receive chat
        ' messages from that channel
        ' **Structure: VChatChannel1||OldChannel||NewChannel
        ' TODO: Add to structure to also get previous channel
        ' for determining where to send message to remove user from
        ' channel user list.
        strChannel = vntSplitData(2)
        strOldChannel = vntSplitData(1)
                
        For Each propUsers In Users.UserCollection
            ' When User Module finds user for this socket
            ' change channel for them
            DoEvents
            If propUsers.SockIndex = Socket Then
                propUsers.ChatChannel = strChannel
                strHandle = propUsers.Handle
                Exit For
            End If
        Next propUsers
        
        For Each propUsers In Users.UserCollection
            DoEvents
            If propUsers.SockIndex <> Socket And propUsers.ChatChannel = strOldChannel Then
                ' Send message to remove this user from everyone elses channel
                ' userlist in old channel
                ' **Structure: VChatRemoveUser1||Handle
                VSock.SendData "VChatRemoveUser1||" & strHandle, propUsers.SockIndex
            End If
            
            'Build user list
            If propUsers.ChatChannel = strChannel Then
                strMessage = strMessage & propUsers.Handle & "||"
            End If
        Next propUsers
        
        ' then send userlist of all users in new channel
        ' **Structure: VChatUserList1||Handle||Handle||Handle||etc..
        VSock.SendData "VChatUserList1||" & strMessage, Socket
        
    Case "VChatAddUser1"
        ' Once client has gotten full list of users in channel, send
        ' message to add this user to everyones list in that channel
        ' **Structure: VChatAddUser1||Handle||Channel
        strHandle = vntSplitData(1)
        strChannel = vntSplitData(2)
        
        For Each propUsers In Users.UserCollection
            DoEvents
            If propUsers.ChatChannel = strChannel And propUsers.Handle <> strHandle Then
                ' **Structure: VChatAddUser2||Handle
                VSock.SendData "VChatAddUser2||" & strHandle, propUsers.SockIndex
            End If
        Next propUsers
        
    Case "VChat1"
        ' SendCode for built in chat functionality
        ' Take incoming data and send to all logged on
        ' users that have a chat status of ON and are in the same
        ' chat channel
        ' **Structure: VChat1||From||Channel||Message
        
        ' Put split data into variables
        strFrom = vntSplitData(1)
        strChannel = vntSplitData(2)
        strMessage = vntSplitData(3)
        
        ' Build the new message to send back to clients with new Sendcode
        ' **Structure: VChat2||From||Message
        strNewMessage = "VChat2||" & strFrom & "||" & strMessage
        
        ' Send message to all clients with chat status on and in same
        ' channel
        For Each propUsers In Users.UserCollection
            DoEvents
            If propUsers.ChatStatus = "On" And propUsers.ChatChannel = strChannel Then
                ' Send the new string of data to each user matching
                ' criteria.
                VSock.SendData strNewMessage, propUsers.SockIndex
            End If
        Next propUsers

' **********************Private Message SendCodes***********************
    Case "VPM1"
        ' SendCode for Private Messaging
        ' **Structure: VPM1||From||To||Message
        strFrom = vntSplitData(1)
        strTo = vntSplitData(2)
        strMessage = vntSplitData(3)
        
        ' Build new SendCode message to send to client
        ' **Structure: VPM2||From||Message
        strNewMessage = "VPM2||" & strFrom & "||" & strMessage
        
        ' Get Socket of To User and Forward message to that user
        For Each propUsers In Users.UserCollection
            DoEvents
            If propUsers.Handle = strTo Then
                VSock.SendData strNewMessage, propUsers.SockIndex
                Exit For
            End If
        Next propUsers

' *****************************Mail SendCodes************************
    Case "VSendMail1"
        ' SendCode for a new V-Mail
        ' **Structure: VSendMail1||From||To||Message
    
    Case "VGetMail1"
        ' Sencode for getting users V-mail
        ' **Structure: VGetMail1||Handle

' *******************External/Invalid SendCodes***********************
    Case Else
        ' Sendcode not used internally
        ' Pass data into main function of each installed
        ' module (ProcessData) until a match is found and the
        ' Match parameter = TRUE
        For Each ptr In colModules
            DoEvents
            ptr.objModule.ProcessData strData, blnMatch, strLogEvent, Socket
            'If blnMatch returns true, stop module loop
            If blnMatch Then Exit For
        Next ptr
        
        If Not blnMatch Then
            ' Send a message back to client that server does not support
            ' the module
            ' **Structure: VNoModule1
            VSock.SendData "VNoModule1", Socket
        End If
        
        ' Check if Modules returned anything to print
        ' to the Server Event Log
        If Len(Trim(strLogEvent)) > 0 Then
            ' Data returned. Write to Server Event Log Screen.
            frmVega.VOutput.SelText = strLogEvent & vbCrLf
        End If
End Select

Exit Sub

errhandler:
Select Case Err.Number
    Case 438
        Resume Next
    Case Else
        'Msbox "An error has occured: " & Err.Number & " " & Err.Description, mbBlue, mbOkOnly, mbExclamation, "VegaCOMM Server Error"
        Resume Next
End Select

End Sub
