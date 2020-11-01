VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_InterGTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsInterGTC.cls
'
' Designed and implemented by Loopzer
'**************************************************************
 
'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at [url=http://www.affero.org/oagpl.html]http://www.affero.org/oagpl.html[/url]
'**************************************************************************
 
Option Explicit
 
Public Intervalo As Long
Private Tiempo As Long
Private Puede As Boolean
 
Public Sub Init(Inter As Long)
    Intervalo = Inter
    Tiempo = GetTickCount + Intervalo
End Sub
 
Public Function Puedo() As Boolean
 
    If GetTickCount < Intervalo - 1 Then '  Me fijo que el tiempo no vuelva a cerro.
        Tiempo = GetTickCount + Intervalo
    End If
    If GetTickCount >= Tiempo Then
        Puede = True
        Tiempo = GetTickCount + Intervalo
    End If
    
Puedo = Puede
Puede = False
 
End Function
