Attribute VB_Name = "modRetoPendiente"
Option Explicit

Public Type t_RetoPendienteJugador
    nombre As String
    Aceptado As Boolean
    CurIndex As t_UserReference
End Type

Public Type t_RetoPendiente
    Oferente As t_UserReference
    Jugadores() As Integer
    Apuesta As Long
    PocionesMaximas As Integer
    CaenItems As Boolean
End Type

Public RetosPendientes() As t_RetoPendiente
Public TotalRetosPendientes As Integer

Public Function NuevoRetoPendiente(ByVal Oferente As Integer, _
                                   ByVal Apuesta As Long, _
                                   ByVal PocionesMaximas As Integer, _
                                   ByVal CaenItems As Boolean) As Integer

    If TotalRetosPendientes = 0 Then
        ReDim RetosPendientes(1 To 1)
        TotalRetosPendientes = 1
    Else
        TotalRetosPendientes = TotalRetosPendientes + 1
        ReDim Preserve RetosPendientes(1 To TotalRetosPendientes)
    End If

    With RetosPendientes(TotalRetosPendientes)
        Call SetUserRef(.Oferente, Oferente)
        .Apuesta = Apuesta
        .PocionesMaximas = PocionesMaximas
        .CaenItems = CaenItems
    End With

    NuevoRetoPendiente = TotalRetosPendientes
End Function

Public Sub AgregarJugadorAReto(ByVal RetoIndex As Integer, ByVal UserIndex As Integer)
    Dim n As Integer
    With RetosPendientes(RetoIndex)
        If (Not Not .Jugadores) = 0 Then
            ReDim .Jugadores(1 To 1)
            n = 1
        Else
            n = UBound(.Jugadores) + 1
            ReDim Preserve .Jugadores(1 To n)
        End If

        .Jugadores(n) = UserIndex
    End With
End Sub

Public Function IndiceJugadorEnRetoPendiente(ByVal UserIndex As Integer, ByVal RetoIndex As Integer) As Integer
    Dim i As Integer
    IndiceJugadorEnRetoPendiente = -1
    
    If RetoIndex <= 0 Then Exit Function
    
    With RetosPendientes(RetoIndex)
        For i = LBound(.Jugadores) To UBound(.Jugadores)
            If .Jugadores(i) = UserIndex Then
                IndiceJugadorEnRetoPendiente = i
                Exit Function
            End If
        Next i
    End With
End Function

Public Sub MensajeATodosRetoPendiente(ByVal RetoIndex As Integer, ByVal Mensaje As String, ByVal Font As e_FontTypeNames)
    Dim i As Integer
    Dim u As Integer

    With RetosPendientes(RetoIndex)
        ' Enviar a todos los jugadores invitados
        For i = LBound(.Jugadores) To UBound(.Jugadores)
            u = .Jugadores(i)
            
            If u > 0 And UserList(u).flags.UserLogged Then
                Call WriteConsoleMsg(u, Mensaje, Font)
            End If
        Next i
        If IsValidUserRef(.Oferente) Then
            Call WriteConsoleMsg(.Oferente.ArrayIndex, Mensaje, Font)
        End If
    End With
End Sub

Public Sub BuscarSalaDesdeRetoPendiente(ByVal RetoIndex As Integer)
    Dim i As Integer
    Dim UserIdx As Integer
    With RetosPendientes(RetoIndex)
        For i = LBound(.Jugadores) To UBound(.Jugadores)
            UserIdx = .Jugadores(i)
            If UserIdx > 0 And UserList(UserIdx).flags.UserLogged Then
                UserList(UserIdx).flags.EnReto = True
            End If
        Next i
        
        If IsValidUserRef(.Oferente) Then
            UserList(.Oferente.ArrayIndex).flags.EnReto = True
        End If
        Call BuscarSala(.Oferente.ArrayIndex)
    End With
End Sub
