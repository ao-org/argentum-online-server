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
        
        On Error GoTo Map_Err
    
        
100     mMap = Value
        
        Exit Property

Map_Err:
102     Call TraceError(Err.Number, Err.Description, "TLimpiezaItem.Map", Erl)

        
End Property

Public Property Let X(ByVal Value As Integer)
        
        On Error GoTo X_Err
    
        
100     mX = Value
        
        Exit Property

X_Err:
102     Call TraceError(Err.Number, Err.Description, "TLimpiezaItem.X", Erl)

        
End Property

Public Property Let Y(ByVal Value As Integer)
        
        On Error GoTo Y_Err
    
        
100     mY = Value
        
        Exit Property

Y_Err:
102     Call TraceError(Err.Number, Err.Description, "TLimpiezaItem.Y", Erl)

        
End Property

Public Property Get Map() As Integer
        
        On Error GoTo Map_Err
    
        
100     Map = mMap
        
        Exit Property

Map_Err:
102     Call TraceError(Err.Number, Err.Description, "TLimpiezaItem.Map", Erl)

        
End Property

Public Property Get X() As Integer
        
        On Error GoTo X_Err
    
        
100     X = mX
        
        Exit Property

X_Err:
102     Call TraceError(Err.Number, Err.Description, "TLimpiezaItem.X", Erl)

        
End Property

Public Property Get Y() As Integer
        
        On Error GoTo Y_Err
    
        
100     Y = mY
        
        Exit Property

Y_Err:
102     Call TraceError(Err.Number, Err.Description, "TLimpiezaItem.Y", Erl)

        
End Property

Public Property Let Time(Value As Long)
        
        On Error GoTo Time_Err
    
        
100     mTime = Value
        
        Exit Property

Time_Err:
102     Call TraceError(Err.Number, Err.Description, "TLimpiezaItem.Time", Erl)

        
End Property

Public Property Get Time() As Long
        
        On Error GoTo Time_Err
    
        
100     Time = mTime
        
        Exit Property

Time_Err:
102     Call TraceError(Err.Number, Err.Description, "TLimpiezaItem.Time", Erl)

        
End Property

Public Property Get Indice() As String
        
        On Error GoTo Indice_Err
    
        
100     Indice = GetIndiceByPos(mMap, mX, mY)
        
        Exit Property

Indice_Err:
102     Call TraceError(Err.Number, Err.Description, "TLimpiezaItem.Indice", Erl)

        
End Property
