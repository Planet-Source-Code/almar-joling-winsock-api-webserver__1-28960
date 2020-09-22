VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarSocket As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarIPAddress As String 'local copy
Private mvarHostName As String 'local copy
Private mvarNick As String 'local copy
Private mvarAuthenticated As Boolean 'local copy

Public Property Let Nick(ByVal vData As String)
    '//Used when assigning a value to the property, on the left side of an assignment.
    '//Syntax: X.Nick = 5
    
    mvarNick = vData
End Property


Public Property Get Nick() As String
    '//Used when retrieving value of a property, on the right side of an assignment.
    '//Syntax: Debug.Print X.Nick
    
    Nick = mvarNick
End Property



Public Property Let HostName(ByVal vData As String)
    '//Used when assigning a value to the property, on the left side of an assignment.
    '//Syntax: X.HostName = 5
    
    mvarHostName = vData
End Property


Public Property Get HostName() As String
    '//Used when retrieving value of a property, on the right side of an assignment.
    '//Syntax: Debug.Print X.HostName
    
    HostName = mvarHostName
End Property



Public Property Let IPAddress(ByVal vData As String)
    '//Used when assigning a value to the property, on the left side of an assignment.
    '//Syntax: X.IPAddress = 5
    
    mvarIPAddress = vData
End Property


Public Property Get IPAddress() As String
    '//Used when retrieving value of a property, on the right side of an assignment.
    '//Syntax: Debug.Print X.IPAddress
    
    IPAddress = mvarIPAddress
End Property



Public Property Let SocketID(ByVal vData As Long)
    '//Used when assigning a value to the property, on the left side of an assignment.
    '//Syntax: X.Socket = 5
    
    mvarSocket = vData
End Property


Public Property Get SocketID() As Long
    '//Used when retrieving value of a property, on the right side of an assignment.
    '//Syntax: Debug.Print X.Socket
    
    SocketID = mvarSocket
End Property


Public Property Let Authenticated(ByVal vData As Boolean)
    '//Used when assigning a value to the property, on the left side of an assignment.
    '//Syntax: X.Authenticated = True
    
    mvarAuthenticated = vData
End Property


Public Property Get Authenticated() As Boolean
    '//Used when retrieving value of a property, on the right side of an assignment.
    '//Syntax: Debug.Print X.Socket
    
    Authenticated = mvarAuthenticated
End Property

