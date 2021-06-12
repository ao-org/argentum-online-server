VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrafteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Items() As Integer

Public Tipo As Byte

Public Precio As Long

Public Probabilidad As Byte

Public Resultado As Integer

Public Sub SetItems(Data() As Integer)
    m_Items = Data
End Sub

Public Property Get Tamaño() As Integer
    Tamaño = UBound(m_Items)
End Property

Public Property Get Items(ByVal Index As Integer) As Integer
    Items = m_Items(Index)
End Property
