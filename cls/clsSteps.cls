VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSteps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarCaption As String 'local copy
Private mvarTamanho As Long 'local copy
Private mvarCorNormal As Long 'local copy
Private mvarCorCompleto As Long 'local copy
Private mvarKey As Variant
Public Property Let Key(ByVal vdata As Variant)
    mvarKey = vdata
End Property
Public Property Get Key() As Variant
    Key = mvarKey
End Property
Public Property Let CorCompleto(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CorCompleto = 5
    mvarCorCompleto = vdata
End Property


Public Property Get CorCompleto() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CorCompleto
    CorCompleto = mvarCorCompleto
End Property



Public Property Let CorNormal(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CorNormal = 5
    mvarCorNormal = vdata
End Property


Public Property Get CorNormal() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CorNormal
    CorNormal = mvarCorNormal
End Property



Public Property Let Tamanho(ByVal vdata As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tamanho = 5
    mvarTamanho = vdata
End Property


Public Property Get Tamanho() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tamanho
    Tamanho = mvarTamanho
End Property








Public Property Let Caption(ByVal vdata As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Caption = 5
    mvarCaption = vdata
    
End Property


Public Property Get Caption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
End Property



Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

End Sub
