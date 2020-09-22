VERSION 5.00
Begin VB.Form frmWebserver 
   Caption         =   "Winsock API - Webserver"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3645
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frLog 
      Caption         =   "Log:"
      Height          =   3630
      Left            =   15
      TabIndex        =   2
      Top             =   -15
      Width           =   7890
      Begin VB.TextBox txtLog 
         Height          =   3360
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   210
         Width           =   7770
      End
   End
   Begin VB.CommandButton cmdMSG 
      Caption         =   "Receive"
      Enabled         =   0   'False
      Height          =   405
      Index           =   1
      Left            =   1575
      TabIndex        =   1
      Top             =   3435
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton cmdMSG 
      Caption         =   "Accept"
      Enabled         =   0   'False
      Height          =   405
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   3435
      Visible         =   0   'False
      Width           =   1410
   End
End
Attribute VB_Name = "frmWebserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//For detailed HTTP RFC, check: http://www.ics.uci.edu/pub/ietf/http/rfc1945.html


'//Webserver code
Private Sub cmdMSG_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim lngReturn As Long, lngBytes As Long
    Dim I As Long
    Dim aSock As Sock, aUser As User, strData As String * 1024
    Dim lngPosition1 As Long, lngPosition2 As Long
    Dim strFileName As String, strHeader As String
    Dim strIP As String, lngExists As Long
    Dim strAuthData As String, strAuthName As String, strAuthPass As String '//Auth scheme
    Dim lngPos As Long, strPost As String, strPostData As String '//Post data
    
    Select Case Index
        Case 0 '//Accept
            Debug.Print "con"
            Set aSock = New Sock
            lngReadSocket = accept(lngSocketNumber, Remote_Buffer, Len(Remote_Buffer))

            lngReturn = WSAAsyncSelect(lngReadSocket, cmdMSG(1).hwnd, ByVal &H202, ByVal FD_READ Or FD_CLOSE)
            
            aSock.Socket = lngReadSocket
            aSock.SSocket = lngReadSocket
            Socks.Add aSock, aSock.SSocket
            
            strIP = mdlTools.GetIP(lngReadSocket)
            lngExists = mdlTools.ClientExists(strIP)
            
            '//Add to log
            Call AddToLog("Incoming:" & strIP)
            
            If lngExists = -1 Then '//does not exist...
                Set aUser = New User
                With aUser
                    .Authenticated = False
                    .IPAddress = strIP
                    .SocketID = lngReadSocket
                End With
                Users.Add aUser
            End If
            
            Set aUser = Nothing
            Set aSock = Nothing
        
        Case 1 '//Receive data
            For Each aSock In Socks
            'For I = 1 To Socks.Count
                'Set aSock = New Sock
                
                '//Set socket to current I
                'Set aSock = Socks.Item(I)
                
                '//Get the number of bytes in the buffer
                lngBytes = recv(aSock.Socket, strData, 1024, 0)
            
                '//If there is something
                If lngBytes > 0 Then
                    strData = Left$(strData, lngBytes)
                    
                    '//Log data
                    AddToLog (strData)

                    '//Get first 3 chars
                    strHeader = Left(strData, 3)
                    
                        
                    Select Case UCase(strHeader)
                        '/////////////////////////////////
                        Case "POS"
                            lngPosition1 = InStr(1, strData, " ")   '//Get the first space...in the received data
                            lngPosition2 = InStr(lngPosition1 + 1, strData, " ") '//Get the second space
                            strFileName = Mid$(strData, lngPosition1 + 1, (lngPosition2 - lngPosition1) - 1)
                          

                            lngPos = InStr(strData, vbCrLf & vbCrLf)
                            lngPos = lngPos + Len(vbCrLf & vbCrLf)
                            lngPos = lngPos - 1
                    
                           If lngPos < Len(strData) Then
                              strPost = Trim$(Right$(strData, Len(strData) - lngPos))
                           End If
                        
                           If strPost <> "" Then
                              If strPostData = "" Then
                                 strPostData = strPostData & strPost
                              Else
                                 strPostData = strPostData & "&" & strPost
                              End If
                           End If
                            '//Make the postdata correctly...no weird chars
                            Call CorrectFormat(strPostData)
                            Debug.Print "Char is:" & Asc(Left$(Right$(strPostData, 2), 1))
                            Dim strAction As String
                            lngPosition1 = InStr(1, strPostData, "cmdSubmitbuttonaction", vbTextCompare)
                            '//22 = submitbuttonaction=
                            lngPosition1 = lngPosition1 + 22
                            lngPosition2 = InStr(2, strPostData, Chr$(13), vbTextCompare)
                            Debug.Print "position2: " & lngPosition2
                            
                            If lngPosition2 = 0 Then '//Check if the data ends with an "&", use that as terminator
                                lngPosition2 = InStr(2, strPostData, "&", vbTextCompare)
                                strAction = LCase$(Mid$(strPostData, lngPosition1, lngPosition2 - lngPosition1))
                            Else '//end of sentence. End with Chr(13)
                                strAction = LCase$(Mid$(strPostData, lngPosition1, lngPosition2 - lngPosition1))
                            End If
                            
                            Debug.Print strPostData
                            strAction = Trim$(strAction)

                            'Call mdlQWServer.ProcessPostData(strPostData, strAction)
                            '//strPostData constains the form information.
                            '//Can be used now.
                            
                            If InStr(strData, "Authorization: Basic ") Then '//String has been found
                                lngPosition1 = InStr(1, strData, "Authorization: Basic ", vbTextCompare) + Len("Authorization: Basic ")
                                lngPosition2 = InStr(lngPosition1, strData, vbCrLf, vbTextCompare)
                                strAuthData = Base64Decode(Mid$(strData, lngPosition1, lngPosition2 - lngPosition1))
                                lngPosition1 = InStr(strAuthData, ":")
                                strAuthName = Left$(strAuthData, lngPosition1 - 1)
                                strAuthPass = Right$(strAuthData, Len(strAuthData) - lngPosition1)
                                
                                '//Password correct... no special stuff here. Just to show it works.
                                If strAuthName = "admin" And strAuthPass = "123" Then
                                    '//Get the IP address, and retrieve the index of the user with it
                                    strIP = mdlTools.GetIP(aSock.Socket)
                                    lngExists = mdlTools.ClientExists(strIP)
                                    
                                    '//Get the item from the collection
                                    Set aUser = Users(lngExists)
                                    
                                    '//Check if the user has already been authenticated, if not, send "ok"
                                    If aUser.Authenticated = False Then
                                        '//Send 200 (ok)
                                        'Call mdlWebserver.SendHTTPStatus(200, aSock.Socket)
                                        aUser.Authenticated = True

                                        '//update the collection...I can't simply update it??
                                        Users.Remove lngExists
                                        Users.Add aUser
                                    Else
                                        '//Send page
                                        Call mdlWebserver.SendPage(aSock.Socket, strFileName)
                                    End If
                                    
                                Else
                                    '//Incorrect password has been entered!
                                    '//Send 401, forbidden
                                    Call mdlWebserver.SendHTTPStatus(401, aSock.Socket)
                                End If
                            Else
                                    '//Send 401, forbidden
                                    Call mdlWebserver.SendHTTPStatus(401, aSock.Socket)
                            End If

                            '//Close after something has been send. Always do this!
                            '//Or the browser will wait...and wait
                            closesocket aSock.Socket
                            Socks.Remove aSock.SSocket
                        
                        
                        '/////////////////////////////////
                        Case "GET"
                            lngPosition1 = InStr(1, strData, " ")   '//Get the first space...in the received data
                            lngPosition2 = InStr(lngPosition1 + 1, strData, " ") '//Get the second space
                            strFileName = Mid$(strData, lngPosition1 + 1, (lngPosition2 - lngPosition1) - 1)

                            If InStr(strData, "Authorization: Basic ") Then '//String has been found
                                lngPosition1 = InStr(1, strData, "Authorization: Basic ", vbTextCompare) + Len("Authorization: Basic ")
                                lngPosition2 = InStr(lngPosition1, strData, vbCrLf, vbTextCompare)
                                strAuthData = Base64Decode(Mid$(strData, lngPosition1, lngPosition2 - lngPosition1))
                                lngPosition1 = InStr(strAuthData, ":")
                                strAuthName = Left$(strAuthData, lngPosition1 - 1)
                                strAuthPass = Right$(strAuthData, Len(strAuthData) - lngPosition1)
                                
                                '//Password correct... no special stuff here. Just to show it works.
                                If strAuthName = "admin" And strAuthPass = "123" Then
                                    '//Get the IP address, and retrieve the index of the user with it
                                    strIP = mdlTools.GetIP(aSock.Socket)
                                    lngExists = mdlTools.ClientExists(strIP)
                                    
                                    '//Get the item from the collection
                                    Set aUser = Users(lngExists)
                                    
                                    '//Check if the user has already been authenticated, if not, send "ok"
                                    If aUser.Authenticated = False Then
                                        '//Send 200 (ok)
                                        'Call mdlWebserver.SendHTTPStatus(200, aSock.Socket)
                                        aUser.Authenticated = True

                                        '//update the collection...I can't simply update it??
                                        Users.Remove lngExists
                                        Users.Add aUser
                                    Else
                                        '//Send page
                                        Call mdlWebserver.SendPage(aSock.Socket, strFileName)
                                    End If
                                    
                                Else
                                    '//Incorrect password has been entered!
                                    '//Send 401, forbidden
                                    Call mdlWebserver.SendHTTPStatus(401, aSock.Socket)
                                End If
                            Else
                                    '//Send 401, forbidden
                                    Call mdlWebserver.SendHTTPStatus(401, aSock.Socket)
                            End If
                            
                            '//Close after something has been send. Always do this!
                            '//Or the browser will wait...and wait
                            closesocket aSock.Socket
                            Socks.Remove aSock.SSocket
                    End Select
                End If
            Next
            'Next I
    
    End Select
End Sub

Private Sub Form_Load()
    strHomeDir = App.Path & "\html\"
    strDefaultFile = "index.html"
    mdlWebserver.StartWebserver
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdlWebserver.StopWebserver
    Unload Me
    End
End Sub

Public Sub AddToLog(strText As String)
    txtLog.Text = txtLog.Text & strText & vbCrLf
End Sub
