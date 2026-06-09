Attribute VB_Name = "ModRetos"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit

Private Const MAX_BET = 10000000
Public Retos As t_Retos
Private WaitingList As New Dictionary

Public Sub LoadChallengeInfo()
    Dim File As clsIniManager
    Set File = New clsIniManager
    Call File.Initialize(DatPath & "Retos.dat")
    With Retos
        .TamañoMaximoEquipo = val(File.GetValue("Retos", "MaximoEquipo"))
        .ApuestaMinima = val(File.GetValue("Retos", "ApuestaMinima"))
        .ImpuestoApuesta = val(File.GetValue("Retos", "ImpuestoApuesta"))
        .DuracionMaxima = val(File.GetValue("Retos", "DuracionMaxima"))
        #If DEBUGGING Then
            .TiempoConteo = 3
        #Else
            .TiempoConteo = val(File.GetValue("Retos", "TiempoConteo"))
        #End If
        .TotalSalas = val(File.GetValue("Salas", "Cantidad"))
        If .TotalSalas <= 0 Then Exit Sub
        ReDim .Salas(1 To .TotalSalas)
        .SalasLibres = .TotalSalas
        .AnchoSala = val(File.GetValue("Salas", "Ancho"))
        .AltoSala = val(File.GetValue("Salas", "Alto"))
        Dim Sala As Integer, SalaStr As String
        For Sala = 1 To .TotalSalas
            SalaStr = "Sala" & Sala
            With .Salas(Sala)
                .PosIzquierda.Map = val(File.GetValue(SalaStr, "Mapa"))
                .PosIzquierda.x = val(File.GetValue(SalaStr, "X"))
                .PosIzquierda.y = val(File.GetValue(SalaStr, "Y"))
                .PosDerecha.Map = .PosIzquierda.Map
                .PosDerecha.x = .PosIzquierda.x + Retos.AnchoSala - 1
                .PosDerecha.y = .PosIzquierda.y + Retos.AltoSala - 1
            End With
        Next
    End With
    Set File = Nothing
End Sub

Public Sub CreateChallenge(ByVal UserIndex As Integer, PlayersStr As String, ByVal Bet As Long, ByVal MaxPotions As Integer, Optional ByVal ItemsDrop As Boolean = False)
    On Error GoTo ErrHandler
        
    With UserList(UserIndex)
        If .flags.SolicitudReto.Estado <> e_SolicitudRetoEstado.Libre Then
            Call CancelChallengeRequest(UserIndex, .name & " ha cancelado la solicitud.")
        ElseIf IsValidUserRef(.flags.AceptoReto) Then
            Call CancelChallengeRequest(.flags.AceptoReto.ArrayIndex, .name & " ha cancelado su admisión.")
        End If
        Dim TamanoReal As Byte: TamanoReal = Retos.TamañoMaximoEquipo * 2 - 1
        If LenB(PlayersStr) <= 0 Then Exit Sub
        Dim Jugadores() As String: Jugadores = Split(PlayersStr, ";", TamanoReal)
        If UBound(Jugadores) > TamanoReal - 1 Or UBound(Jugadores) Mod 2 = 1 Then Exit Sub
        Dim MaxIndexEquipo As Integer: MaxIndexEquipo = UBound(Jugadores) \ 2
        If Bet < Retos.ApuestaMinima Or Bet > MAX_BET Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(MSG_APUESTA_MINIMA_MONEDAS_ORO, PonerPuntos(Retos.ApuestaMinima), e_FontTypeNames.FONTTYPE_INFO)) ' Msg1958=La apuesta mínima es de ¬1 monedas de oro.
            Exit Sub
        End If
        If Not CanChallengeWithMessage(UserIndex) Then Exit Sub
        If .Stats.GLD < Bet Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_TIENES_ORO_SUFICIENTE, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If MaxPotions >= 0 Then
            If TieneObjetos(38, MaxPotions + 1, UserIndex) Then
                Call WriteLocaleMsg(UserIndex, MSG_TIENES_DEMASIADAS_POCIONES_ROJAS_CANTIDAD_MAXIMA_1443, e_FontTypeNames.FONTTYPE_INFO, MaxPotions)
                Exit Sub
            End If
        End If
        With .flags.SolicitudReto
            .Apuesta = Bet
            .PocionesMaximas = MaxPotions
            .CaenItems = ItemsDrop
            ReDim .Jugadores(0 To UBound(Jugadores))
            Dim i       As Integer, tIndex As t_UserReference
            Dim Equipo1 As String, Equipo2 As String
            Equipo1 = UserList(UserIndex).name
            For i = 0 To UBound(.Jugadores)
                With .Jugadores(i)
                    If EsGmChar(Jugadores(i)) Then
                        Call WriteLocaleMsg(UserIndex, MSG_PUEDES_JUGAR_RETOS_ADMINISTRADORES, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    tIndex = NameIndex(Jugadores(i))
                    If Not IsValidUserRef(tIndex) Then
                        Call WriteLocaleMsg(UserIndex, MSG_NO_USUARIO_PUEDE_JUGAR_RETO_MOMENTO, e_FontTypeNames.FONTTYPE_INFO, Jugadores(i))
                        Exit Sub
                    End If
                    If Not CanChallenge(tIndex.ArrayIndex) Then
                        Call WriteLocaleMsg(UserIndex, MSG_NO_USUARIO_PUEDE_JUGAR_RETO_MOMENTO, e_FontTypeNames.FONTTYPE_INFO, UserList(tIndex.ArrayIndex).name)
                        Exit Sub
                    End If
                    .CurIndex = tIndex
                    .nombre = UserList(.CurIndex.ArrayIndex).name
                    .Aceptado = False
                    If i Mod 2 Then
                        Equipo1 = Equipo1 & IIf((i + 1) \ 2 < MaxIndexEquipo, ", ", " y ") & .nombre
                    Else
                        If LenB(Equipo2) > 0 Then
                            Equipo2 = Equipo2 & IIf(i \ 2 < MaxIndexEquipo, ", ", " y ") & .nombre
                        Else
                            Equipo2 = .nombre
                        End If
                    End If
                End With
            Next
            Dim Texto1 As String, Texto2 As String, Texto3 As String
            Texto1 = UserList(UserIndex).name & "(" & UserList(UserIndex).Stats.ELV & ") te invita a jugar el siguiente reto:"
            Texto2 = Equipo1 & " vs " & Equipo2 & ". Apuesta: " & PonerPuntos(Bet) & " monedas de oro" & IIf(ItemsDrop, " y los items.", ".")
            Texto3 = "Escribe /ACEPTAR " & UCase$(UserList(UserIndex).name) & " para participar en el reto."
            If MaxPotions >= 0 Then
                Texto2 = Texto2 & " Máximo " & MaxPotions & " pociones rojas."
            End If
            For i = 0 To UBound(.Jugadores)
                With .Jugadores(i)
                    Call WriteConsoleMsg(.CurIndex.ArrayIndex, Texto1, e_FontTypeNames.FONTTYPE_INFO)
                    Call WriteConsoleMsg(.CurIndex.ArrayIndex, Texto2, e_FontTypeNames.FONTTYPE_New_Naranja)
                    Call WriteConsoleMsg(.CurIndex.ArrayIndex, Texto3, e_FontTypeNames.FONTTYPE_INFO)
                End With
            Next
            .Estado = e_SolicitudRetoEstado.Enviada
        End With
        Call WriteLocaleMsg(UserIndex, MSG_ENVIADO_SOLICITUD_SIGUIENTE_RETO, e_FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, Texto2, e_FontTypeNames.FONTTYPE_New_Naranja)
        Call WriteLocaleMsg(UserIndex, MSG_ESCRIBE_CANCELAR_ANULAR_SOLICITUD, e_FontTypeNames.FONTTYPE_New_Gris)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.CreateChallenge", Erl)
End Sub

Public Sub AcceptChallenge(ByVal UserIndex As Integer, ChallengerName As String)
    On Error GoTo ErrHandler
    If Not CanChallengeWithMessage(UserIndex) Then Exit Sub
    If EsGmChar(ChallengerName) Then Exit Sub
    Dim ChallengerRef As t_UserReference
    ChallengerRef = NameIndex(ChallengerName)
    If Not IsValidUserRef(ChallengerRef) Then Exit Sub
    Dim Challenger As Integer
    Challenger = ChallengerRef.ArrayIndex
    
    With UserList(Challenger).flags.SolicitudReto
        If .Estado <> e_SolicitudRetoEstado.Enviada Then
            Call WriteConsoleMsg(UserIndex, "Ese reto ya no existe o fue cancelado.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim i As Integer
        Dim Found As Boolean
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i).CurIndex.ArrayIndex = UserIndex Then
                .Jugadores(i).Aceptado = True
                Found = True
                Exit For
            End If
        Next
        If Not Found Then Exit Sub
        
        ' Check if everyone accepted
        For i = 0 To UBound(.Jugadores)
            If Not .Jugadores(i).Aceptado Then
                Call WriteConsoleMsg(UserIndex, "Aceptaste el reto. Esperando a los demás...", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Next
        Dim Sala As Integer
        Sala = FindFreeRoom()
        If Sala <= 0 Then
            Call WriteConsoleMsg(UserIndex, "No hay salas de reto disponibles en este momento.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call StartChallenge(Challenger, Sala)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.AcceptChallenge", Erl)
End Sub

Public Sub CancelChallengeRequest(ByVal Challenger As Integer, Message As String)
    On Error GoTo ErrHandler
    With UserList(Challenger).flags.SolicitudReto
        If .Estado = e_SolicitudRetoEstado.EnCola Then
            Call WaitingList.Remove(Challenger)
        End If
        .Estado = e_SolicitudRetoEstado.Libre
        Dim i As Integer, tUser As t_UserReference
        
        ' Send to invited players
        For i = 0 To UBound(.Jugadores)
            tUser = NameIndex(.Jugadores(i).nombre)
            If IsValidUserRef(tUser) Then
                Call WriteConsoleMsg(tUser.ArrayIndex, Message, e_FontTypeNames.FONTTYPE_WARNING)
                Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(MSG_RETO_SIDO_CANCELADO, vbNullString, e_FontTypeNames.FONTTYPE_WARNING)) ' Msg1965=El reto ha sido cancelado.
                If .Jugadores(i).Aceptado Then
                    Call SetUserRef(UserList(tUser.ArrayIndex).flags.AceptoReto, 0)
                End If
            End If
        Next
        
        ' And to the challenger separately
        Call WriteConsoleMsg(Challenger, Message, e_FontTypeNames.FONTTYPE_WARNING)
        Call WriteConsoleMsg(Challenger, PrepareMessageLocaleMsg(MSG_RETO_SIDO_CANCELADO, vbNullString, e_FontTypeNames.FONTTYPE_WARNING)) ' Msg1965=El reto ha sido cancelado.
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.CancelChallengeRequest", Erl)
End Sub

Public Sub FindRoom(ByVal Oferente As Integer)
    On Error GoTo ErrHandler
    With UserList(Oferente).flags.SolicitudReto
        If Retos.SalasLibres <= 0 Then
            Call WaitingList.Add(Oferente, 0)
            Call SendMessageToAllInRequest(Oferente, "No hay salas disponibles. El reto comenzará cuando se desocupe una sala.", e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        Dim Sala As Integer, SalaAleatoria As Integer
        SalaAleatoria = RandomNumber(1, Retos.SalasLibres)
        For Sala = 1 To Retos.TotalSalas
            If Not Retos.Salas(Sala).EnUso Then
                SalaAleatoria = SalaAleatoria - 1
                If SalaAleatoria = 0 Then Exit For
            End If
        Next
        Call StartChallenge(Oferente, Sala)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.FindRoom", Erl)
End Sub

Private Sub StartChallenge(ByVal Challenger As Integer, ByVal Room As Integer)
    On Error GoTo ErrHandler
    With UserList(Challenger).flags.SolicitudReto
        ' Last check if everyone can enter/pay
        If Not AllCanChallenge(Challenger) Then Exit Sub
        Dim Bet As Long, BetStr As String
        Bet = .Apuesta
        BetStr = PonerPuntos(Bet)
        ' Calculate team size
        Retos.Salas(Room).TamañoEquipoIzq = UBound(.Jugadores) \ 2 + 1
        Retos.Salas(Room).TamañoEquipoDer = Retos.Salas(Room).TamañoEquipoIzq
        ' Reserve space for players (including challenger)
        ReDim Retos.Salas(Room).Jugadores(0 To UBound(.Jugadores) + 1)
        ' Coin flip (50-50) to decide if challenger goes at start or end of list
        Dim Coin As Byte
        Coin = RandomNumber(0, 1)
        Dim CurIndex As Integer
        If Coin = 0 Then
            ' Add challenger at start (their team plays on the left)
            Call SetUserRef(Retos.Salas(Room).Jugadores(CurIndex), Challenger)
            CurIndex = CurIndex + 1
        End If
        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            Retos.Salas(Room).Jugadores(CurIndex) = .Jugadores(i).CurIndex
            Call SetUserRef(UserList(.Jugadores(i).CurIndex.ArrayIndex).flags.AceptoReto, 0)
            CurIndex = CurIndex + 1
        Next i
        If Coin = 1 Then
            ' Add challenger at end (their team plays on the right)
            Call SetUserRef(Retos.Salas(Room).Jugadores(CurIndex), Challenger)
        End If
    End With
    With Retos.Salas(Room)
        .EnUso = True
        .Puntaje = 0
        .Ronda = 1
        .Apuesta = Bet
        .TiempoRestante = Retos.DuracionMaxima
        .CaenItems = UserList(Challenger).flags.SolicitudReto.CaenItems
        Dim tUser As t_UserReference
        For i = 0 To UBound(.Jugadores)
            tUser = .Jugadores(i)
            ' Charge the bet
            UserList(tUser.ArrayIndex).Stats.GLD = UserList(tUser.ArrayIndex).Stats.GLD - Bet
            Call WriteUpdateGold(tUser.ArrayIndex)
            Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(MSG_OTORGAS_MONEDAS_ORO_POZO_RETO, BetStr, e_FontTypeNames.FONTTYPE_New_Rojo_Salmon)) ' Msg1966=Otorgas ¬1 monedas de oro al pozo del reto.
            ' Dismount
            If UserList(tUser.ArrayIndex).flags.Montado <> 0 Then
                Call DoMontar(tUser.ArrayIndex, ObjData(UserList(tUser.ArrayIndex).invent.EquippedSaddleObjIndex), UserList(tUser.ArrayIndex).invent.EquippedSaddleSlot)
            End If
            ' Stop navigating
            If UserList(tUser.ArrayIndex).flags.Nadando <> 0 Or UserList(tUser.ArrayIndex).flags.Navegando <> 0 Then
                Call DoNavega(tUser.ArrayIndex, ObjData(UserList(tUser.ArrayIndex).invent.EquippedShipObjIndex), UserList(tUser.ArrayIndex).invent.EquippedShipSlot)
            End If
            ' Assign flags
            With UserList(tUser.ArrayIndex).flags
                .EnReto = True
                .EquipoReto = IIf(i Mod 2, e_EquipoReto.Derecha, e_EquipoReto.Izquierda)
                .SalaReto = Room
                ' Save position
                .LastPos = UserList(tUser.ArrayIndex).pos
            End With
            Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(MSG_COMENZADO_RETO, vbNullString, e_FontTypeNames.FONTTYPE_New_Rojo_Salmon)) ' Msg1967=¡Ha comenzado el reto!
            Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(MSG_ADMITIR_DERROTA_ESCRIBE_ABANDONAR, vbNullString, e_FontTypeNames.FONTTYPE_New_Gris)) ' Msg1968=Para admitir la derrota escribe /ABANDONAR.
        Next
    End With
    Retos.SalasLibres = Retos.SalasLibres - 1
    Call StartRound(Room)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.StartChallenge", Erl)
End Sub

Private Sub StartRound(ByVal Room As Integer)
    With Retos.Salas(Room)
        Dim i As Integer, tUser As t_UserReference
        For i = 0 To UBound(.Jugadores)
            tUser = .Jugadores(i)
            If tUser.ArrayIndex <> 0 Then
                Call ReviveAndClean(tUser.ArrayIndex)
                ' Use round number and index to decide which side
                If (.Ronda + i) Mod 2 = 1 Then
                    Call WarpToLegalPos(tUser.ArrayIndex, .PosIzquierda.Map, .PosIzquierda.x, .PosIzquierda.y, True)
                Else
                    Call WarpToLegalPos(tUser.ArrayIndex, .PosDerecha.Map, .PosDerecha.x, .PosDerecha.y, True)
                End If
                ' If countdown is enabled
                If Retos.TiempoConteo > 0 Then
                    ' Set the countdown
                    UserList(tUser.ArrayIndex).Counters.CuentaRegresiva = Retos.TiempoConteo
                    ' Stop the player
                    Call WriteStopped(tUser.ArrayIndex, True)
                End If
                Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(MSG_COMIENZA_RONDA_N, .Ronda, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1969=Comienza la ronda Nº¬1
            End If
        Next
    End With
End Sub

Public Sub DiesInChallenge(ByVal UserIndex As Integer)
    On Error GoTo ErrorHandler
    Dim Room As Integer, Team As e_EquipoReto
    With UserList(UserIndex)
        Room = .flags.SalaReto
        Team = .flags.EquipoReto
    End With
    With Retos.Salas(Room)
        Dim Idx As Integer
        ' Right team is at odd indices
        If Team = e_EquipoReto.Derecha Then Idx = 1
        For Idx = Idx To UBound(.Jugadores) Step 2
            If .Jugadores(Idx).ArrayIndex <> 0 Then
                ' If there's still someone alive on the team
                If UserList(.Jugadores(Idx).ArrayIndex).flags.Muerto = 0 Then
                    Exit Sub
                End If
            End If
        Next
        ' All dead, opposing team wins
        Call ProcessRoundWon(Room, OpposingTeam(Team))
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.DiesInChallenge", Erl)
End Sub

Private Sub ProcessRoundWon(ByVal Room As Integer, ByVal Team As e_EquipoReto)
    With Retos.Salas(Room)
        ' Add or subtract score depending on the team
        If Team = e_EquipoReto.Derecha Then
            .Puntaje = .Puntaje + 1
        Else
            .Puntaje = .Puntaje - 1
        End If
        ' If third round ended or a team got 2 consecutive wins
        If .Ronda >= 3 Or Abs(.Puntaje) >= 2 Then
            Call FinalizeChallenge(Room)
            Exit Sub
        End If
        ' Increase round number
        .Ronda = .Ronda + 1
        ' Get current team size (in case someone left)
        Dim TeamSize As Integer, TeamSize2 As Integer
        TeamSize = GetTeamSize(Room, Team)
        ' Less calculations in the loop
        TeamSize2 = TeamSize * 2
        ' Get winning team names
        Dim i As Integer, names As String
        For i = IIf(Team = e_EquipoReto.Izquierda, 0, 1) To TeamSize2 - 1 Step 2
            If .Jugadores(i).ArrayIndex <> 0 Then
                names = names & UserList(.Jugadores(i).ArrayIndex).name
                If i < TeamSize2 - 2 Then
                    names = names & IIf(i > TeamSize2 - 5, " y ", ", ")
                End If
            End If
        Next
        ' Inform round winner
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i).ArrayIndex <> 0 Then
                Call WriteConsoleMsg(.Jugadores(i).ArrayIndex, PrepareMessageLocaleMsg(MSG_RONDA, names, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1970=Esta ronda es para ¬1.
                Call WriteConsoleMsg(.Jugadores(i).ArrayIndex, "", 0) ' Vertical space
            End If
        Next
        ' Start next round
        Call StartRound(Room)
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.ProcessRoundWon", Erl)
End Sub

Public Sub FinalizeChallenge(ByVal Room As Integer, Optional ByVal TimeExpired As Boolean)
    On Error GoTo ErrorHandler
    With Retos.Salas(Room)
        ' Calculate total prize gold
        Dim TotalGold As Long, gold As Long, GoldStr As String
        TotalGold = .Apuesta * (UBound(.Jugadores) + 1)
        ' Discount tax
        TotalGold = TotalGold * (1 - Retos.ImpuestoApuesta)
        ' Decide result based on score
        Dim i                 As Integer, tUser As t_UserReference, Team1 As String, Team2 As String
        Dim eloTotalIzquierda As Long, eloTotalDerecha As Long, winsIzquierda As Long, winsDerecha As Long
        Dim allEligible       As Boolean
        allEligible = True
        For i = 0 To UBound(.Jugadores)
            tUser = .Jugadores(i)
            If tUser.ArrayIndex <> 0 Then
                allEligible = allEligible And (UserList(tUser.ArrayIndex).Stats.ELV >= 33)
                If i Mod 2 = 0 Then
                    eloTotalIzquierda = eloTotalIzquierda + UserList(tUser.ArrayIndex).Stats.ELO
                Else
                    eloTotalDerecha = eloTotalDerecha + UserList(tUser.ArrayIndex).Stats.ELO
                End If
            End If
        Next i
        ' Draw
        If .Puntaje = 0 Then
            ' Pay everyone who didn't abandon
            gold = TotalGold \ (UBound(.Jugadores) + 1)
            GoldStr = PonerPuntos(gold)
            ' No winners, ELO bonus doesn't apply
            winsIzquierda = 0
            winsDerecha = 0
            For i = 0 To UBound(.Jugadores)
                tUser = .Jugadores(i)
                If IsValidUserRef(tUser) Then
                    UserList(tUser.ArrayIndex).Stats.GLD = UserList(tUser.ArrayIndex).Stats.GLD + gold
                    Call WriteUpdateGold(tUser.ArrayIndex)
                    Call WriteLocaleMsg(tUser.ArrayIndex, "29", e_FontTypeNames.FONTTYPE_MP, GoldStr) ' Has ganado X monedas de oro
                    Call ReviveAndClean(tUser.ArrayIndex)
                    Call DevolverPosAnterior(tUser.ArrayIndex)
                    ' Reset flags
                    UserList(tUser.ArrayIndex).Counters.CuentaRegresiva = -1
                    With UserList(tUser.ArrayIndex).flags
                        .EnReto = False
                        .SalaReto = 0
                        .EquipoReto = 0
                        .SolicitudReto.Estado = e_SolicitudRetoEstado.Libre
                    End With
                    ' Names
                    If i Mod 2 Then
                        If LenB(Team2) > 0 Then
                            Team2 = Team2 & IIf((i + 1) \ 2 < .TamañoEquipoDer - 2, ", ", " y ") & UserList(tUser.ArrayIndex).name
                        Else
                            Team2 = UserList(tUser.ArrayIndex).name
                        End If
                    Else
                        If LenB(Team1) > 0 Then
                            Team1 = Team1 & IIf(i \ 2 < .TamañoEquipoIzq - 2, ", ", " y ") & UserList(tUser.ArrayIndex).name
                        Else
                            Team1 = UserList(tUser.ArrayIndex).name
                        End If
                    End If
                End If
            Next
            ' Global announcement
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_RETOS_VS_NINGUNO_PUDO_VENCER_RIVAL, Team1 & "¬" & Team2, e_FontTypeNames.FONTTYPE_INFO))
            Call RoomFreed(Room)
        ' There was a winner
        Else
            Dim Winner As e_EquipoReto
            If .Puntaje < 0 Then
                Winner = e_EquipoReto.Izquierda
                winsIzquierda = .TamañoEquipoDer
                winsDerecha = -.TamañoEquipoIzq
            Else
                Winner = e_EquipoReto.Derecha
                winsIzquierda = -.TamañoEquipoDer
                winsDerecha = .TamañoEquipoIzq
            End If
            ' Pay winners who didn't abandon
            gold = TotalGold \ GetTeamSize(Room, Winner)
            GoldStr = PonerPuntos(gold)
            For i = 0 To UBound(.Jugadores)
                tUser = .Jugadores(i)
                If IsValidUserRef(tUser) Then
                    Call ReviveAndClean(tUser.ArrayIndex)
                    If UserList(tUser.ArrayIndex).flags.EquipoReto = Winner Then
                        UserList(tUser.ArrayIndex).Stats.GLD = UserList(tUser.ArrayIndex).Stats.GLD + gold
                        Call WriteUpdateGold(tUser.ArrayIndex)
                        Call WriteLocaleMsg(tUser.ArrayIndex, "29", e_FontTypeNames.FONTTYPE_MP, GoldStr) ' Has ganado X monedas de oro
                        If .CaenItems Then
                            Call WarpToLegalPos(tUser.ArrayIndex, .PosIzquierda.Map, .PosIzquierda.x, .PosIzquierda.y, True)
                        Else
                            With UserList(tUser.ArrayIndex).flags
                                .EnReto = False
                                .SalaReto = 0
                                .EquipoReto = 0
                                .SolicitudReto.Estado = e_SolicitudRetoEstado.Libre
                            End With
                            Call DevolverPosAnterior(tUser.ArrayIndex)
                        End If
                    Else
                        If .CaenItems Then
                            Call DropItemsAtPos(tUser.ArrayIndex, ((.PosDerecha.x - .PosIzquierda.x) \ 2) + .PosIzquierda.x, ((.PosDerecha.y - .PosIzquierda.y) \ 2) + _
                                    .PosIzquierda.y)
                        End If
                        With UserList(tUser.ArrayIndex).flags
                            .EnReto = False
                            .SalaReto = 0
                            .EquipoReto = 0
                            .SolicitudReto.Estado = e_SolicitudRetoEstado.Libre
                        End With
                        Call DevolverPosAnterior(tUser.ArrayIndex)
                    End If
                    ' Reset flags
                    UserList(tUser.ArrayIndex).Counters.CuentaRegresiva = -1
                    If TimeExpired Then
                        Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(MSG_AGOTADO_TIEMPO_RETO, vbNullString, e_FontTypeNames.FONTTYPE_New_Gris))
                    End If
                    ' Names
                    If i Mod 2 Then
                        If LenB(Team2) > 0 Then
                            Team2 = Team2 & IIf((i + 1) \ 2 < .TamañoEquipoDer - 2, ", ", " y ") & UserList(tUser.ArrayIndex).name
                        Else
                            Team2 = UserList(tUser.ArrayIndex).name
                        End If
                    Else
                        If LenB(Team1) > 0 Then
                            Team1 = Team1 & IIf(i \ 2 < .TamañoEquipoIzq - 2, ", ", " y ") & UserList(tUser.ArrayIndex).name
                        Else
                            Team1 = UserList(tUser.ArrayIndex).name
                        End If
                    End If
                End If
            Next
            Dim WinnerTeam As String, LoserTeam As String
            WinnerTeam = IIf(Winner = e_EquipoReto.Izquierda, Team1, Team2)
            LoserTeam = IIf(Winner = e_EquipoReto.Izquierda, Team2, Team1)
            ' Global announcement
            If UBound(.Jugadores) > 1 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_RETOS_EQUIPO_VENCIO_EQUIPO_QUEDO_BOTIN_MONEDAS, WinnerTeam & "¬" & LoserTeam & "¬" & PonerPuntos(.Apuesta), e_FontTypeNames.FONTTYPE_INFO))
            Else ' 1 vs 1
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_RETOS_VENCIO_QUEDO_BOTIN_MONEDAS_ORO, WinnerTeam & "¬" & LoserTeam & "¬" & PonerPuntos(.Apuesta), e_FontTypeNames.FONTTYPE_INFO))
            End If
            If .CaenItems Then
                Call StartItemDeposit(Room)
            Else
                Call RoomFreed(Room)
            End If
        End If
        ' Update ELO for each player, inspired by `400 Algorithm`
        ' https://en.wikipedia.org/wiki/Elo_rating_system
        Dim eloDiff As Long
        For i = 0 To UBound(.Jugadores)
            tUser = .Jugadores(i)
            If IsValidUserRef(tUser) Then
                If allEligible Then
                    If i Mod 2 = 0 Then ' Left team players
                        eloDiff = winsIzquierda * (eloTotalDerecha * 0.1)
                    Else
                        eloDiff = winsDerecha * (eloTotalIzquierda * 0.1)
                    End If
                    If eloDiff > 0 Then
                        Call SendData(SendTarget.ToIndex, tUser.ArrayIndex, PrepareMessageLocaleMsg(MSG_HAS_GANADO_PUNTOS_ELO, Abs(eloDiff), e_FontTypeNames.FONTTYPE_ROSA))
                    Else
                        If UserList(tUser.ArrayIndex).Stats.ELO < Abs(eloDiff) Then
                            eloDiff = -UserList(tUser.ArrayIndex).Stats.ELO
                        End If
                        Call SendData(SendTarget.ToIndex, tUser.ArrayIndex, PrepareMessageLocaleMsg(MSG_HAS_PERDIDO_PUNTOS_ELO, Abs(eloDiff), e_FontTypeNames.FONTTYPE_ROSA))
                    End If
                    UserList(tUser.ArrayIndex).Stats.ELO = UserList(tUser.ArrayIndex).Stats.ELO + eloDiff
                Else ' Someone is below level 33
                    Call SendData(SendTarget.ToIndex, tUser.ArrayIndex, PrepareMessageLocaleMsg(MSG_PARTICIPANTE_RETO_TIENE_NIVEL_MENOR_ELO_PERMANECE, vbNullString, e_FontTypeNames.FONTTYPE_INFOIAO))
                End If
            End If
        Next i
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.FinalizeChallenge", Erl)
End Sub

Public Sub DropItemsAtPos(ByVal UserIndex As Integer, ByVal x As Byte, ByVal y As Byte)
    On Error GoTo DropItemsAtPos_Err
    Dim i        As Byte
    Dim NewPos   As t_WorldPos
    Dim MyObj    As t_Obj
    Dim ItemIdx  As Integer
    Dim ItemsPos As t_WorldPos
    With UserList(UserIndex)
        ItemsPos.Map = .pos.Map
        ItemsPos.x = x
        ItemsPos.y = y
        For i = 1 To .CurrentInventorySlots
            ItemIdx = .invent.Object(i).ObjIndex
            If ItemIdx > 0 Then
                If ItemSeCae(ItemIdx) And PirataCaeItem(UserIndex, i) And (Not EsNewbie(UserIndex) Or Not ItemNewbie(ItemIdx)) Then
                    NewPos.x = 0
                    NewPos.y = 0
                    MyObj.amount = .invent.Object(i).amount
                    MyObj.ObjIndex = ItemIdx
                    MyObj.ElementalTags = .invent.Object(i).ElementalTags
                    Call Tilelibre(ItemsPos, NewPos, MyObj, True, True, False)
                    If NewPos.x <> 0 And NewPos.y <> 0 Then
                        Call DropObj(UserIndex, i, MyObj.amount, NewPos.Map, NewPos.x, NewPos.y)
                        ' If no space, burn the item from inventory (no free backpacks)
                    Else
                        Call QuitarUserInvItem(UserIndex, i, MyObj.amount)
                        Call UpdateUserInv(False, UserIndex, i)
                    End If
                End If
            End If
        Next i
    End With
    Exit Sub
DropItemsAtPos_Err:
    Call TraceError(Err.Number, Err.Description, "ModRetos.DropItemsAtPos", Erl)
End Sub

Public Sub StartItemDeposit(ByVal Room As Integer)
    Dim i      As Byte
    Dim Winner As e_EquipoReto
    With Retos.Salas(Room)
        If .Puntaje < 0 Then
            Winner = e_EquipoReto.Izquierda
        Else
            Winner = e_EquipoReto.Derecha
        End If
        For i = 0 To UBound(.Jugadores)
            If UserList(.Jugadores(i).ArrayIndex).flags.EquipoReto = Winner Then
                Call WriteConsoleMsg(.Jugadores(i).ArrayIndex, PrepareMessageLocaleMsg(MSG_TIENES_MINUTO_LEVANTAR_ITEMS_PISO, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1972=Tienes 1 minuto para levantar los items del piso.
            End If
        Next i
        Dim pos As t_WorldPos
        pos.Map = .PosIzquierda.Map
        pos.x = ((.PosDerecha.x - .PosIzquierda.x) \ 2) + .PosIzquierda.x
        pos.y = ((.PosDerecha.y - .PosIzquierda.y) \ 2) + .PosIzquierda.y
        ' Spawn a banker
        .IndexBanquero = SpawnNpc(3, pos, True, False)
        #If DEBUGGING Then
            .TiempoItems = 20
        #Else
            .TiempoItems = 60
        #End If
    End With
End Sub

Public Sub EndItemPickupTime(ByVal Room As Integer)
    Dim Winner As e_EquipoReto
    With Retos.Salas(Room)
        ' Kill the banker
        Call QuitarNPC(.IndexBanquero, eChallenge)
        If .Puntaje < 0 Then
            Winner = e_EquipoReto.Izquierda
        Else
            Winner = e_EquipoReto.Derecha
        End If
        Dim i As Byte
        For i = 0 To UBound(.Jugadores)
            If IsValidUserRef(.Jugadores(i)) Then
                If UserList(.Jugadores(i).ArrayIndex).flags.EquipoReto = Winner Then
                    UserList(.Jugadores(i).ArrayIndex).flags.EnReto = False
                    Call DevolverPosAnterior(.Jugadores(i).ArrayIndex)
                End If
            End If
        Next i
        .TiempoItems = 0
        Dim x As Byte
        Dim y As Byte
        For x = .PosIzquierda.x To .PosDerecha.x
            For y = .PosIzquierda.y To .PosDerecha.y
                Call EraseObj(GetMaxInvOBJ(), .PosIzquierda.Map, x, y)
            Next y
        Next x
    End With
    Call RoomFreed(Room)
End Sub

Public Sub AbandonChallenge(ByVal UserIndex As Integer, Optional ByVal Disconnected As Boolean)
    Dim Room As Integer, Team As e_EquipoReto
    With UserList(UserIndex)
        Room = .flags.SalaReto
        Team = .flags.EquipoReto
        .Counters.CuentaRegresiva = -1
        .flags.EnReto = False
    End With
    With Retos.Salas(Room)
        If .CaenItems And Abs(.Puntaje) >= 2 Then
            If .Puntaje < 0 Then
                .TamañoEquipoIzq = .TamañoEquipoIzq - 1
                If .TamañoEquipoIzq <= 0 Then
                    EndItemPickupTime (Room)
                End If
            Else
                .TamañoEquipoDer = .TamañoEquipoDer - 1
                If .TamañoEquipoDer <= 0 Then
                    EndItemPickupTime (Room)
                End If
            End If
            Exit Sub
        End If
        If Not Disconnected Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(MSG_HAS_ABANDONADO_RETO, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1973=Has abandonado el reto.
        End If
        ' Subtract a team member and if it reaches zero process the defeat
        If Team = e_EquipoReto.Izquierda Then
            If .TamañoEquipoIzq > 1 Then
                .TamañoEquipoIzq = .TamañoEquipoIzq - 1
            Else
                .Puntaje = 123 ' Force positive score
                Call FinalizeChallenge(Room)
                Exit Sub
            End If
        Else
            If .TamañoEquipoDer > 1 Then
                .TamañoEquipoDer = .TamañoEquipoDer - 1
            Else
                .Puntaje = -123 ' Force negative score
                Call FinalizeChallenge(Room)
                Exit Sub
            End If
        End If
        Call ReviveAndClean(UserIndex)
        Call DevolverPosAnterior(UserIndex)
        Dim Message As String
        If Disconnected Then
            Message = UserList(UserIndex).name & " es descalificado por desconectarse."
        Else
            Message = UserList(UserIndex).name & " ha abandonado el reto."
        End If
        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i).ArrayIndex = UserIndex Then
                Call SetUserRef(.Jugadores(i), 0)
            Else
                Call WriteConsoleMsg(.Jugadores(i).ArrayIndex, Message, e_FontTypeNames.FONTTYPE_New_Gris)
            End If
        Next
    End With
End Sub

Private Sub RoomFreed(ByVal Room As Integer)
    On Error GoTo ErrHandler
    Retos.Salas(Room).EnUso = False
    Retos.SalasLibres = Retos.SalasLibres + 1
    If WaitingList.count > 0 Then
        Dim Challenger As Integer
        Challenger = WaitingList.Keys(0)
        Call WaitingList.Remove(Challenger)
        Call StartChallenge(Challenger, Room)
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.RoomFreed", Erl)
End Sub

Public Function CanChallenge(ByVal UserIndex As Integer) As Boolean
    With UserList(UserIndex)
        If .flags.EnReto Then Exit Function
        If .flags.EnConsulta Then Exit Function
        If .pos.Map = 0 Or .pos.x = 0 Or .pos.y = 0 Then Exit Function
        If MapInfo(.pos.Map).Seguro = 0 Then Exit Function
        If .flags.EnTorneo Then Exit Function
        If MapData(.pos.Map, .pos.x, .pos.y).trigger = CARCEL Then Exit Function
    End With
    CanChallenge = True
End Function

Public Function CanChallengeWithMessage(ByVal UserIndex As Integer) As Boolean
    With UserList(UserIndex)
        If .flags.EnReto Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(MSG_ENCUENTRAS_RETO, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1974=Ya te encuentras en un reto.
            Exit Function
        End If
        If .flags.EnConsulta Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(MSG_NO_PUEDES_ACCEDER_RETO_CONSULTA, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1975=No puedes acceder a un reto si estás en consulta.
            Exit Function
        End If
        If .flags.jugando_captura = 1 Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(MSG_NO_PUEDES_JUGAR_RETO_ESTANDO_EVENTO, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1976=No puedes jugar un reto estando en un evento.
            Exit Function
        End If
        If Not esCiudad(.pos.Map) Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(MSG_NO_PUEDES_PARTICIPAR_RETO_MAPA_INSEGURO, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1977=No puedes participar de un reto en un mapa inseguro.
            Exit Function
        End If
        If .flags.EnTorneo Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(MSG_NO_PUEDES_IR_RETO_PARTICIPAS_TORNEO, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1978=No puedes ir a un reto si participas de un torneo.
            Exit Function
        End If
        If MapData(.pos.Map, .pos.x, .pos.y).trigger = CARCEL Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(MSG_ENCARCELADO, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1979=¡Estás encarcelado!
            Exit Function
        End If
    End With
    CanChallengeWithMessage = True
End Function

Private Function GetPlayerIndexInRequest(ByVal UserIndex As Integer, ByVal Challenger As Integer) As Integer
    With UserList(Challenger).flags.SolicitudReto
        GetPlayerIndexInRequest = -1
        If .Estado <> e_SolicitudRetoEstado.Enviada Then Exit Function
        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i).nombre = UserList(UserIndex).name Then
                GetPlayerIndexInRequest = i
                Exit Function
            End If
        Next
    End With
End Function

Private Sub SendMessageToAllInRequest(ByVal Challenger As Integer, Message As String, ByVal Font As e_FontTypeNames)
    With UserList(Challenger).flags.SolicitudReto
        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i).Aceptado Then
                Call WriteConsoleMsg(.Jugadores(i).CurIndex.ArrayIndex, Message, Font)
            End If
        Next
        Call WriteConsoleMsg(Challenger, Message, Font)
    End With
End Sub

Private Function AllCanChallenge(ByVal Challenger As Integer) As Boolean
    On Error GoTo ErrHandler
    With UserList(Challenger).flags.SolicitudReto
        If Not CanChallenge(Challenger) Then
            Call CancelChallengeRequest(Challenger, UserList(Challenger).name & " no puede entrar al reto en este momento.")
            Exit Function
        ElseIf UserList(Challenger).Stats.GLD < .Apuesta Then
            Call CancelChallengeRequest(Challenger, UserList(Challenger).name & " no tiene las monedas de oro suficientes.")
            Exit Function
        ElseIf .PocionesMaximas > 0 Then
            If TieneObjetos(38, .PocionesMaximas + 1, Challenger) Then
                Call CancelChallengeRequest(Challenger, UserList(Challenger).name & " tiene demasiadas pociones rojas (Cantidad máxima: " & .PocionesMaximas & ").")
                Exit Function
            End If
        End If
        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            If Not CanChallenge(.Jugadores(i).CurIndex.ArrayIndex) Then
                Call CancelChallengeRequest(Challenger, UserList(.Jugadores(i).CurIndex.ArrayIndex).name & " no puede entrar al reto en este momento.")
                Exit Function
            ElseIf UserList(.Jugadores(i).CurIndex.ArrayIndex).Stats.GLD < .Apuesta Then
                Call CancelChallengeRequest(Challenger, UserList(.Jugadores(i).CurIndex.ArrayIndex).name & " no tiene las monedas de oro suficientes.")
                Exit Function
            ElseIf .PocionesMaximas > 0 Then
                If TieneObjetos(38, .PocionesMaximas + 1, .Jugadores(i).CurIndex.ArrayIndex) Then
                    Call CancelChallengeRequest(Challenger, UserList(.Jugadores(i).CurIndex.ArrayIndex).name & " tiene demasiadas pociones rojas (Cantidad máxima: " & .PocionesMaximas & ").")
                    Exit Function
                End If
            End If
        Next
        AllCanChallenge = True
    End With
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.AllCanChallenge", Erl)
End Function

Private Function OpposingTeam(ByVal Team As e_EquipoReto) As e_EquipoReto
    If Team = e_EquipoReto.Izquierda Then
        OpposingTeam = e_EquipoReto.Derecha
    Else
        OpposingTeam = e_EquipoReto.Izquierda
    End If
End Function

Private Function GetTeamSize(ByVal Room As Integer, ByVal Team As e_EquipoReto) As Integer
    If Team = e_EquipoReto.Izquierda Then
        GetTeamSize = Retos.Salas(Room).TamañoEquipoIzq
    Else
        GetTeamSize = Retos.Salas(Room).TamañoEquipoDer
    End If
End Function

Private Sub ReviveAndClean(ByVal UserIndex As Integer)
    Call WriteStopped(UserIndex, False)
    ' If alive
    If UserList(UserIndex).flags.Muerto = 0 Then
        Call LimpiarEstadosAlterados(UserIndex)
    End If
    ' If dead revive, otherwise heal
    Call RevivirUsuario(UserIndex)
End Sub

Public Function GetPendingChallengeIndexForUser(ByVal UserIndex As Integer) As Integer
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).flags.SolicitudReto.Estado <> e_SolicitudRetoEstado.Libre Then
            ' Is challenger
            If i = UserIndex Then
                GetPendingChallengeIndexForUser = i
                Exit Function
            End If
            ' Is in that challenger's list
            Dim j As Integer
            With UserList(i).flags.SolicitudReto
                For j = 0 To UBound(.Jugadores)
                    If .Jugadores(j).CurIndex.ArrayIndex = UserIndex Then
                        GetPendingChallengeIndexForUser = i
                        Exit Function
                    End If
                Next j
            End With
        End If
    Next i
End Function

Private Function FindFreeRoom() As Integer
    Dim i As Integer
    For i = LBound(Retos.Salas) To UBound(Retos.Salas)
        If Not Retos.Salas(i).EnUso Then
            FindFreeRoom = i
            Exit Function
        End If
    Next i
    FindFreeRoom = 0
End Function
