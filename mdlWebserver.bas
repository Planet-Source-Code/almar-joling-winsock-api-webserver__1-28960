Attribute VB_Name = "mdlWebserver"
Option Explicit
Public Users As Collection
Public Socks As Collection

Public lngSocketNumber As Long
Public lngReadSocket As Long
Public Remote_Buffer As sockaddr
Public Read_Buffer As String * 1024

Public sockBuffer As sockaddr
Public udtWSAData As WSADataType

Public strHomeDir As String
Public strDefaultFile As String

Public Sub StartWebserver()
    Dim lngReturn As Long   '//Contains returned value by functions

    '//Create collections
    Set Users = New Collection
    Set Socks = New Collection

    '//Startup winsock
    lngReturn = WSACleanup()
    lngReturn = WSAStartup(&H101, udtWSAData)
        
    '//An error occured
    If (lngReturn = SOCKET_ERROR) Then
        MsgBox Error
        Exit Sub
    End If
    
    '//Try to create our TCP socket
    lngSocketNumber = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    
    '//An error occured while trying...
    If (lngSocketNumber = SOCKET_ERROR) Then
        Exit Sub
    End If
        
    '//Configure our socket
    With sockBuffer
        .sin_family = AF_INET
        .sin_port = htons(80)      '// Port 80 is HTTP
        .sin_addr = 0
        .sin_zero = String$(8, 0)
    End With
                        
    '//Bind socket...
    lngReturn = bind(lngSocketNumber, sockBuffer, sockaddr_size)
    
    '//An error occured while trying to bind
    If Not lngReturn = 0 Then
        lngReturn = WSACleanup()
        Exit Sub
    End If
    
    '//Start listening for connections
    lngReturn = listen(lngSocketNumber, 1)
    
    '//Perform Async functions, with our form's hwnd
    lngReturn = WSAAsyncSelect(lngSocketNumber, frmWebserver.cmdMSG(0).hwnd, &H202, FD_CONNECT Or FD_ACCEPT)
    
    frmWebserver.txtLog.Text = "Listening on port: " & htons(sockBuffer.sin_port) & vbCrLf
End Sub

Public Sub StopWebserver()
    Dim lngReturn As Long
    
    '//Stop server
    lngReturn = WSACleanup()
End Sub

Public Sub SendHTTPStatus(intStatus As Integer, lngSocket As Long)
    Dim strHeader As String
    Dim strContent As String
    strContent = vbCrLf & "Server: Winsock API server" & vbCrLf & "Server-Version: 1.0" & vbCrLf & "Made by: Almar Joling"

    Select Case intStatus
        Case 200
            strHeader = "HTTP/1.0 200 OK" & strContent & vbCrLf & vbCrLf & vbCrLf
        Case 400
            strHeader = "HTTP/1.0 400 Bad Request" & strContent & vbCrLf & vbCrLf & vbCrLf
        Case 401
            strHeader = "HTTP/1.0 401 Unauthorized" & vbCrLf & "WWW-Authenticate: Basic realm=" & Chr$(34) & "Quadrant Wars Remote-Admin" & Chr$(34) & vbCrLf & vbCrLf & vbCrLf
        Case 404
            strHeader = "HTTP/1.0 404 File Not Found" & "" & vbCrLf & vbCrLf & vbCrLf
    End Select
    
    '//Send the data
    Call SendData(lngSocket, strHeader)
End Sub

Public Function SendData(lngSocket As Long, ByVal strMessage As String) As Long
    Dim TheMsg() As Byte, sTemp As String
    TheMsg = ""
    
    sTemp = StrConv(strMessage, vbFromUnicode)
    TheMsg = sTemp
    SendData = send(lngSocket, TheMsg(0), UBound(TheMsg) + 1, 0)
End Function

Public Sub SendPage(lngSocket As Long, strFileName As String)
    Dim strValidFile As String
    Dim lngFreeFile As Long
    Dim strFileData As String
    
    '//Check file validity, etc
    If Left$(strFileName, 1) = "/" Then
        If Len(strFileName) = 1 Then   '//Default document
            strValidFile = strHomeDir & strDefaultFile
        Else
            strValidFile = strHomeDir & Mid$(strFileName, 2)
        End If
    Else
        Call SendHTTPStatus(404, lngSocket)
        Exit Sub
    End If
    
    If FileExists(strValidFile) = False Then
        '//Send error
        Call SendHTTPStatus(404, lngSocket)
        Exit Sub
    End If

    '//Check if the ".." security issue isn't used.
    If InStr(1, strValidFile, "..") Then
        '//Send error
        Call SendHTTPStatus(400, lngSocket)
        Exit Sub
    End If

    '//Get a freefile
    lngFreeFile = FreeFile
    Open strValidFile For Binary As #lngFreeFile
    
    '//Redim array to fit filesize
    strFileData = Space$(LOF(lngFreeFile))
    'ReDim bFileData(LOF(lngFreeFile))
    
    '//Get filedata
    Get #lngFreeFile, , strFileData '//bFileData
    
    '//Close File
    Close #lngFreeFile
    
    '//send the page
    Call SendData(lngSocket, strFileData)
End Sub
