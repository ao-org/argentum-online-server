Attribute VB_Name = "Queue"
Option Explicit

Public Type tVertice

    x As Integer
    Y As Integer

End Type

Private Const MAXELEM As Integer = 1000

Private m_array()     As tVertice

Private m_lastelem    As Integer

Private m_firstelem   As Integer

Private m_size        As Integer

Public Function IsEmpty() As Boolean
        
        On Error GoTo IsEmpty_Err
        
100     IsEmpty = m_size = 0

        
        Exit Function

IsEmpty_Err:
        Call RegistrarError(Err.Number, Err.description, "Queue.IsEmpty", Erl)
        Resume Next
        
End Function

Public Function IsFull() As Boolean
        
        On Error GoTo IsFull_Err
        
100     IsFull = m_lastelem = MAXELEM

        
        Exit Function

IsFull_Err:
        Call RegistrarError(Err.Number, Err.description, "Queue.IsFull", Erl)
        Resume Next
        
End Function

Public Function Push(ByRef Vertice As tVertice) As Boolean
        
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
        Call RegistrarError(Err.Number, Err.description, "Queue.Push", Erl)
        Resume Next
        
End Function

Public Function Pop() As tVertice
        
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
        Call RegistrarError(Err.Number, Err.description, "Queue.Pop", Erl)
        Resume Next
        
End Function

Public Sub InitQueue()
        
        On Error GoTo InitQueue_Err
        
100     ReDim m_array(MAXELEM) As tVertice
102     m_lastelem = 0
104     m_firstelem = 0
106     m_size = 0

        
        Exit Sub

InitQueue_Err:
        Call RegistrarError(Err.Number, Err.description, "Queue.InitQueue", Erl)
        Resume Next
        
End Sub

