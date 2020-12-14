VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TLimpiezaItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Defino la posición por separado porque vb no me deja usar WorldPos
Private mMap As Integer
Private mX As Integer
Private mY As Integer
Private mTime As Long 'Item 'Time of item to clear

Public Property Let Map(ByVal Value As Integer)
100     mMap = Value
End Property

Public Property Let X(ByVal Value As Integer)
100     mX = Value
End Property

Public Property Let Y(ByVal Value As Integer)
100     mY = Value
End Property

Public Property Get Map() As Integer
100     Map = mMap
End Property

Public Property Get X() As Integer
100     X = mX
End Property

Public Property Get Y() As Integer
100     Y = mY
End Property

Public Property Let Time(Value As Long)
100     mTime = Value
End Property

Public Property Get Time() As Long
100     Time = mTime
End Property

Public Property Get Indice() As String
100     Indice = GetIndiceByPos(mMap, mX, mY)
End Property
