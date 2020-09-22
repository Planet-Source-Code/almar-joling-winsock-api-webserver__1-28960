VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarSocket As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarSSocket As String 'local copy
Public Property Let SSocket(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SSocket = 5
    mvarSSocket = vData
End Property


Public Property Get SSocket() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SSocket
    SSocket = mvarSSocket
End Property



Public Property Let Socket(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Socket = 5
    mvarSocket = vData
End Property


Public Property Get Socket() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Socket
    Socket = mvarSocket
End Property
