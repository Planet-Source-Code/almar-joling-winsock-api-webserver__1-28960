Attribute VB_Name = "mdlTools"
Option Explicit

Public Function FileExists(strFileName As String) As Boolean
    '//check if file really exists
    If Len(Dir$(strFileName)) > 1 Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

Public Function Base64Decode(base64String As String)
  Const Base64CodeBase = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength As Long, Out As String, groupBegin As Long
  
  dataLength = Len(base64String)
  Out = ""

  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string."
    Exit Function
  End If

  ' Now decode each group:
  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, groupData
    ' Each data group encodes up To 3 actual bytes.
    numDataBytes = 3
    groupData = 0

    For CharCounter = 0 To 3
      ' Convert each character into 6 bits of data, And add it To
      ' an integer For temporary storage.  If a character is a '=', there
      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
      ' the whole string.)

      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(Base64CodeBase, thisChar) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      groupData = 64 * groupData + thisData
    Next

    ' Convert 3-byte integer into up To 3 characters
    Dim OneChar
    For CharCounter = 1 To numDataBytes
      Select Case CharCounter
        Case 1: OneChar = groupData \ 65536
        Case 2: OneChar = (groupData And 65535) \ 256
        Case 3: OneChar = (groupData And 255)
      End Select
      Out = Out & Chr(OneChar)
    Next
  Next

  Base64Decode = Out
End Function

Public Function ClientExists(strIP As String) As Long
    Dim I As Long
    Dim aUser As User
    ClientExists = -1
    
    For Each aUser In Users
        I = I + 1   '//Increment with one
        If aUser.IPAddress = strIP Then '//If they match
            ClientExists = I    '//Return I
            Exit Function   '//Exit
        End If
    Next
End Function

Public Function GetIP(lngSocket As Long) As String
    Dim lngPosition1 As Long
    Dim strIPAddress As String
    
    strIPAddress = GetPeerAddress(lngSocket)

    lngPosition1 = InStr(1, strIPAddress, ":")
    
    '//Return everything before the ':'
    GetIP = Left$(strIPAddress, lngPosition1 - 1)
End Function

Public Function CorrectFormat(strData As String) As String
      strData = Replace$(strData, "%22", Chr$(34))
      strData = Replace$(strData, "%3C", "<")
      strData = Replace$(strData, "%3E", ">")
      strData = Replace$(strData, "+", " ")
      strData = Replace$(strData, "%0D%0A", "<br>")
      strData = Replace$(strData, "%21", "!")
      strData = Replace$(strData, "%22", "&quot;")
      strData = Replace$(strData, "%20", " ")
      strData = Replace$(strData, "%A7", "§")
      strData = Replace$(strData, "%24", "$")
      strData = Replace$(strData, "%25", "%")
      strData = Replace$(strData, "%26", "&")
      strData = Replace$(strData, "%2F", "/")
      strData = Replace$(strData, "%28", "(")
      strData = Replace$(strData, "%29", ")")
      strData = Replace$(strData, "%3D", "=")
      strData = Replace$(strData, "%3F", "?")
      strData = Replace$(strData, "%B2", "²")
      strData = Replace$(strData, "%B3", "³")
      strData = Replace$(strData, "%7B", "{")
      strData = Replace$(strData, "%5B", "[")
      strData = Replace$(strData, "%5D", "]")
      strData = Replace$(strData, "%7D", "}")
      strData = Replace$(strData, "%5C", "\")
      strData = Replace$(strData, "%DF", "ß")
      strData = Replace$(strData, "%23", "#")
      strData = Replace$(strData, "%27", "'")
      strData = Replace$(strData, "%3A", ":")
      strData = Replace$(strData, "%2C", ",")
      strData = Replace$(strData, "%3B", ";")
      strData = Replace$(strData, "%60", "`")
      strData = Replace$(strData, "%7E", "~")
      strData = Replace$(strData, "%2B", "+")
      strData = Replace$(strData, "%B4", "´")
      CorrectFormat = strData
End Function
