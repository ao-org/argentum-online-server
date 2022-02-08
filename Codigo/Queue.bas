Attribute VB_Name = "Queue"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Option Explicit

Public Type t_Vertice

    X As Integer
    Y As Integer

End Type

Private Const MAXELEM As Integer = 1000

Private m_array()     As t_Vertice

Private m_lastelem    As Integer

Private m_firstelem   As Integer

Private m_size        As Integer

Public Function IsEmpty() As Boolean
        
        On Error GoTo IsEmpty_Err
        
100     IsEmpty = m_size = 0

        
        Exit Function

IsEmpty_Err:
102     Call TraceError(Err.Number, Err.Description, "Queue.IsEmpty", Erl)
104
        
End Function

Public Function IsFull() As Boolean
        
        On Error GoTo IsFull_Err
        
100     IsFull = m_lastelem = MAXELEM

        
        Exit Function

IsFull_Err:
102     Call TraceError(Err.Number, Err.Description, "Queue.IsFull", Erl)
104
        
End Function

Public Function Push(ByRef Vertice As t_Vertice) As Boolean
        
        On Error GoTo Push_Err
        

100     If Not IsFull Then
    
102         If IsEmpty Then m_firstelem = 1
    
104         m_lastelem = m_lastelem + 1
106         m_size = m_size + 1
108         m_array(m_lastelem) = Vertice
    
110         Push = True
        Else
112         Push = False

        End If

        
        Exit Function

Push_Err:
114     Call TraceError(Err.Number, Err.Description, "Queue.Push", Erl)
116
        
End Function

Public Function Pop() As t_Vertice
        
        On Error GoTo Pop_Err
        

100     If Not IsEmpty Then
    
102         Pop = m_array(m_firstelem)
104         m_firstelem = m_firstelem + 1
106         m_size = m_size - 1
    
108         If m_firstelem > m_lastelem And m_size = 0 Then
110             m_lastelem = 0
112             m_firstelem = 0
114             m_size = 0

            End If
   
        End If

        
        Exit Function

Pop_Err:
116     Call TraceError(Err.Number, Err.Description, "Queue.Pop", Erl)
118
        
End Function

Public Sub InitQueue()
        
        On Error GoTo InitQueue_Err
        
100     ReDim m_array(MAXELEM) As t_Vertice
102     m_lastelem = 0
104     m_firstelem = 0
106     m_size = 0

        
        Exit Sub

InitQueue_Err:
108     Call TraceError(Err.Number, Err.Description, "Queue.InitQueue", Erl)
110
        
End Sub

