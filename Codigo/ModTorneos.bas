Attribute VB_Name = "ModTorneos"

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
'

Private Const TORNEO_MAPA As Integer = 272
Private Const TORNEO_SLOT1_X As Byte = 16
Private Const TORNEO_SLOT1_Y As Byte = 45
Private Const TORNEO_SLOT2_X As Byte = 32
Private Const TORNEO_SLOT2_Y As Byte = 56

Public Type t_Torneo
    ' Estado
    HayTorneoActivo     As Boolean
    Started             As Boolean
    nombre              As String
    reglas              As String
    
    ' Configuracion
    NivelMinimo         As Byte
    NivelMaximo         As Byte
    cupos               As Byte
    costo               As Long
    Mapa                As Integer
    x                   As Byte
    y                   As Byte
    
    ' Clases permitidas (indexado por e_Class, 0 a 11)
    ClasesPermitidas(11) As Boolean
    ClasesTexto          As String
    
    ' Participantes
    participantes       As Byte
    IndexParticipantes() As Integer
    LastPosMap()        As Integer
    LastPosX()          As Byte
    LastPosY()          As Byte
End Type

Public Torneo        As t_Torneo
Public MensajeTorneo As String

Public Sub IniciarTorneo()
    On Error GoTo IniciarTorneo_Err
    If Torneo.Started Then
        LogInfoServidor "Invalid call IniciarTorneo for a Torneo that already started"
        Debug.Assert False
        Exit Sub
    End If
    
    Dim i          As Long
    Dim inscriptos As Byte
    Dim clase      As e_Class
    
    inscriptos = 0
    Torneo.ClasesTexto = ""
    For clase = 0 To 11
        If Torneo.ClasesPermitidas(clase) Then
            Torneo.ClasesTexto = Torneo.ClasesTexto & ClaseToString(clase) & ","
        End If
    Next clase
    If Len(Torneo.ClasesTexto) > 0 Then
        Torneo.ClasesTexto = Left$(Torneo.ClasesTexto, Len(Torneo.ClasesTexto) - 1)
    End If
    
    If Not Torneo.HayTorneoActivo Then
        ReDim Torneo.IndexParticipantes(1 To Torneo.cupos)
        ReDim Torneo.LastPosMap(1 To Torneo.cupos)
        ReDim Torneo.LastPosX(1 To Torneo.cupos)
        ReDim Torneo.LastPosY(1 To Torneo.cupos)
        Torneo.HayTorneoActivo = True
    Else
        For i = 1 To Torneo.cupos
            If Torneo.IndexParticipantes(i) > 0 Then
                inscriptos = inscriptos + 1
            End If
        Next i
    End If
    Torneo.Started = True
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_EVENTO_ABIERTAS_INSCRIPCIONES_CARACTERISTICAS_NIVEL_ENTRE_INSCRIPTOS, Torneo.nombre & "¬" & Torneo.NivelMinimo & "¬" & Torneo.NivelMaximo & "¬" & inscriptos & "¬" & Torneo.cupos _
            & "¬" & PonerPuntos(Torneo.costo) & "¬" & Torneo.reglas, e_FontTypeNames.FONTTYPE_CITIZEN))
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_EVENTO_CLASES_PARTICIPANTES_ESCRIBI_PARTICIPAR_INGRESAR_EVENTO, Torneo.ClasesTexto, e_FontTypeNames.FONTTYPE_CITIZEN))
    Exit Sub
IniciarTorneo_Err:
    Call TraceError(Err.Number, Err.Description, "ModTorneos.IniciarTorneo", Erl)
End Sub

Public Sub ParticiparTorneo(ByVal UserIndex As Integer)
    On Error GoTo ParticiparTorneo_Err
    
    ' Verificar que hay un torneo activo
    If Not Torneo.HayTorneoActivo Then
        Call WriteLocaleMsg(UserIndex, MSG_NO_HAY_TORNEO_ACTIVO, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    ' Verificar que el jugador no esté inscripto
    If UserList(UserIndex).flags.EnTorneo Then
        Call WriteLocaleMsg(UserIndex, MSG_REGISTERED_IN_TOURNAMENT, e_FontTypeNames.FONTTYPE_INFOIAO)
        Exit Sub
    End If
    
    ' Verificar nivel
    If UserList(UserIndex).Stats.ELV < Torneo.NivelMinimo Or UserList(UserIndex).Stats.ELV > Torneo.NivelMaximo Then
        Call WriteLocaleMsg(UserIndex, MSG_NIVEL_NO_PERMITIDO_TORNEO, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    ' Verificar clase
    If Not ClasePermitidaEnTorneo(UserList(UserIndex).clase) Then
        Call WriteLocaleMsg(UserIndex, MSG_CLASE_NO_PERMITIDA_TORNEO, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    ' Verificar cupo disponible
    Dim IndexVacio As Byte
    IndexVacio = BuscarIndexFreeTorneo()
    If IndexVacio = 0 Then
        Call WriteLocaleMsg(UserIndex, MSG_TORNEO_SIN_CUPO, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    ' Verificar que tiene oro suficiente
    If UserList(UserIndex).Stats.GLD < Torneo.costo Then
        Call WriteLocaleMsg(UserIndex, MSG_UTILIZAR_COMANDO_NECESITAS_MONEDAS_ORO, e_FontTypeNames.FONTTYPE_INFO, Torneo.costo)
        Exit Sub
    End If
    
    ' Cobrar inscripcion
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Torneo.costo
    Call WriteUpdateGold(UserIndex)
    
    ' Registrar participante
    Torneo.IndexParticipantes(IndexVacio) = UserIndex
    Torneo.participantes = Torneo.participantes + 1
    UserList(UserIndex).flags.EnTorneo = True
    
    ' Guardar posicion actual para poder devolver al jugador si se cancela
    Torneo.LastPosMap(IndexVacio) = UserList(UserIndex).pos.Map
    Torneo.LastPosX(IndexVacio) = UserList(UserIndex).pos.x
    Torneo.LastPosY(IndexVacio) = UserList(UserIndex).pos.y
    
    Call WriteLocaleMsg(UserIndex, MSG_REGISTERED_IN_TOURNAMENT, e_FontTypeNames.FONTTYPE_INFOIAO)
    
    ' Si se llenó el cupo, arrancar el torneo
    If Torneo.participantes >= Torneo.cupos Then
        Call ComenzarTorneoOk
    End If
    
    Exit Sub
ParticiparTorneo_Err:
    Call TraceError(Err.Number, Err.Description, "ModTorneos.ParticiparTorneo", Erl)
End Sub

Private Function ClasePermitidaEnTorneo(ByVal clase As e_Class) As Boolean
    ClasePermitidaEnTorneo = Torneo.ClasesPermitidas(clase)
End Function

Private Function ClaseToString(ByVal clase As e_Class) As String
    Select Case clase
        Case e_Class.Mage:       ClaseToString = "Mago"
        Case e_Class.Cleric:     ClaseToString = "Clerigo"
        Case e_Class.Warrior:    ClaseToString = "Guerrero"
        Case e_Class.Assasin:    ClaseToString = "Asesino"
        Case e_Class.Bard:       ClaseToString = "Bardo"
        Case e_Class.Druid:      ClaseToString = "Druida"
        Case e_Class.Paladin:    ClaseToString = "Paladin"
        Case e_Class.Hunter:     ClaseToString = "Cazador"
        Case e_Class.Trabajador: ClaseToString = "Trabajador"
        Case e_Class.Pirat:      ClaseToString = "Pirata"
        Case e_Class.Thief:      ClaseToString = "Ladron"
        Case e_Class.Bandit:     ClaseToString = "Bandido"
        Case Else:               ClaseToString = ""
    End Select
End Function


Public Function BuscarIndexFreeTorneo() As Byte
    On Error GoTo BuscarIndexFreeTorneo_Err
    Dim i As Byte
    For i = 1 To Torneo.cupos
        If Torneo.IndexParticipantes(i) = 0 Then
            BuscarIndexFreeTorneo = i
            Exit Function
        End If
    Next i
    BuscarIndexFreeTorneo = 0
    Exit Function
BuscarIndexFreeTorneo_Err:
    Call TraceError(Err.Number, Err.Description, "ModTorneos.BuscarIndexFreeTorneo", Erl)
End Function

Public Sub BorrarIndexInTorneo(ByVal Index As Integer)
    On Error GoTo BorrarIndexInTorneo_Err
    Dim i As Byte
    For i = 1 To Torneo.cupos
        If Torneo.IndexParticipantes(i) = Index Then
            Torneo.IndexParticipantes(i) = 0
            Exit For
        End If
    Next i
    Torneo.participantes = Torneo.participantes - 1
    Exit Sub
BorrarIndexInTorneo_Err:
    Call TraceError(Err.Number, Err.Description, "ModTorneos.BorrarIndexInTorneo", Erl)
End Sub

Public Sub ComenzarTorneoOk()
    On Error GoTo ComenzarTorneoOk_Err
    Dim i       As Long
    Dim nombres As String
    Dim x       As Byte
    Dim y       As Byte
        
    For i = 1 To Torneo.participantes
        nombres = nombres & UserList(Torneo.IndexParticipantes(i)).name & ", "
        Torneo.LastPosMap(i) = UserList(Torneo.IndexParticipantes(i)).pos.Map
        Torneo.LastPosX(i) = UserList(Torneo.IndexParticipantes(i)).pos.x
        Torneo.LastPosY(i) = UserList(Torneo.IndexParticipantes(i)).pos.y
        
        ' TODO: En PRs posteriores generalizar posiciones para 4/8/16/32 cupos.
        ' Por ahora solo soporta 1v1 (2 cupos): jugador 1 en 16,45 - jugador 2 en 32,56
        If i = 1 Then
            x = TORNEO_SLOT1_X
            y = TORNEO_SLOT1_Y
        Else
            x = TORNEO_SLOT2_X
            y = TORNEO_SLOT2_Y
        End If
        Call WarpUserChar(Torneo.IndexParticipantes(i), TORNEO_MAPA, x, y, True)
    Next i
    
    If Len(nombres) > 0 Then
        nombres = Left$(nombres, Len(nombres) - 2)
    End If
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_EVENTO_ELEGIDOS_PARTICIPAR_DAMOS_INICIO_EVENTO, nombres, e_FontTypeNames.FONTTYPE_CITIZEN)) 'Msg1676=Evento> Los elegidos para participar son: ¬1 damos inicio al evento.
    Exit Sub
ComenzarTorneoOk_Err:
    Call TraceError(Err.Number, Err.Description, "ModTorneos.ComenzarTorneoOk", Erl)
End Sub

Public Sub ResetearTorneo(ByVal UserIndex As Integer, Optional ByVal Reembolsar As Boolean = True)
    On Error GoTo ResetearTorneo_Err
    
    Debug.Assert Torneo.Started
    If Not Torneo.Started Then
        LogInfoServidor "Invalid call ResetearTorneo for a Torneo that has not started yet"
        Exit Sub
    End If
    
    Dim i As Byte
    
    ' Devolver jugadores a su posicion original
    For i = 1 To Torneo.participantes
        If Torneo.IndexParticipantes(i) > 0 Then
            ' Devolver oro solo si se cancela el torneo
            If Reembolsar Then
                UserList(Torneo.IndexParticipantes(i)).Stats.GLD = UserList(Torneo.IndexParticipantes(i)).Stats.GLD + Torneo.costo
                Call WriteUpdateGold(Torneo.IndexParticipantes(i))
            End If
            ' Devolver posicion y limpiar flag
            UserList(Torneo.IndexParticipantes(i)).flags.EnTorneo = False
            Call WarpUserChar(Torneo.IndexParticipantes(i), Torneo.LastPosMap(i), Torneo.LastPosX(i), Torneo.LastPosY(i), True)
        End If
    Next i
    
    ' Limpiar estado
    Torneo.HayTorneoActivo = False
    Torneo.Started = False
    Torneo.nombre = ""
    Torneo.reglas = ""
    
    ' Limpiar configuracion
    Torneo.NivelMinimo = 0
    Torneo.NivelMaximo = 0
    Torneo.cupos = 0
    Torneo.costo = 0
    Torneo.Mapa = 0
    Torneo.x = 0
    Torneo.y = 0
    
    ' Limpiar clases
    Dim clase As e_Class
    For clase = 0 To 11
        Torneo.ClasesPermitidas(clase) = False
    Next clase
    Torneo.ClasesTexto = ""
    
    ' Limpiar participantes
    Torneo.participantes = 0
    ReDim Torneo.IndexParticipantes(1 To 1)
    ReDim Torneo.LastPosMap(1 To 1)
    ReDim Torneo.LastPosX(1 To 1)
    ReDim Torneo.LastPosY(1 To 1)
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_EVENTO_FINALIZADO_1677, vbNullString, e_FontTypeNames.FONTTYPE_CITIZEN))
    
    If UserIndex > 0 Then
        Call WriteLocaleMsg(UserIndex, MSG_TORNEO_RESETEADO_CORRECTAMENTE, e_FontTypeNames.FONTTYPE_INFOIAO)
    End If
    Exit Sub
    
ResetearTorneo_Err:
    Call TraceError(Err.Number, Err.Description, "ModTorneos.ResetearTorneo", Erl)
End Sub
