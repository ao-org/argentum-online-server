Attribute VB_Name = "ModTorneos"

' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
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
Public Type t_Torneo
    HayTorneoaActivo As Boolean
    NivelMinimo As Byte
    NivelMaximo As Byte
    cupos As Byte
    costo As Long
    mago As Byte
    clerico As Byte
    guerrero As Byte
    asesino As Byte
    bardo As Byte
    druido As Byte
    Paladin As Byte
    cazador As Byte
    Trabajador As Byte
    Pirata As Byte
    Ladron As Byte
    Bandido As Byte
    ClasesTexto As String
    participantes As Byte
    IndexParticipantes() As Integer
    Mapa As Integer
    x As Byte
    y As Byte
    nombre As String
    reglas As String
End Type

Public Torneo        As t_Torneo
Public MensajeTorneo As String

Public Sub IniciarTorneo()
    On Error GoTo IniciarTorneo_Err
    Dim i          As Long
    Dim inscriptos As Byte
    inscriptos = 0
    If Torneo.mago > 0 Then Torneo.ClasesTexto = "Mago,"
    If Torneo.clerico > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Clerigo,"
    If Torneo.guerrero > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Guerrero,"
    If Torneo.asesino > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Asesino,"
    If Torneo.bardo > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Bardo,"
    If Torneo.druido > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Druida,"
    If Torneo.Paladin > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Paladin,"
    If Torneo.cazador > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Cazador,"
    If Torneo.Trabajador > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Trabajador,"
    If Torneo.Pirata > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Pirata,"
    If Torneo.Ladron > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Ladron,"
    If Torneo.Bandido > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Bandido"
    If Not Torneo.HayTorneoaActivo Then
        ReDim Torneo.IndexParticipantes(1 To Torneo.cupos)
        Torneo.HayTorneoaActivo = True
    Else
        For i = 1 To Torneo.cupos
            If Torneo.IndexParticipantes(i) > 0 Then
                inscriptos = inscriptos + 1
            End If
        Next i
    End If
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1674, Torneo.nombre & "¬" & Torneo.NivelMinimo & "¬" & Torneo.NivelMaximo & "¬" & inscriptos & "¬" & Torneo.cupos _
            & "¬" & PonerPuntos(Torneo.costo) & "¬" & Torneo.reglas, e_FontTypeNames.FONTTYPE_CITIZEN)) 'Msg1674=Evento> Están abiertas las inscripciones para: ¬1: características: Nivel entre: ¬2/¬3. Inscriptos: ¬4/¬5. Precio de inscripción: ¬6 monedas de oro. Reglas: ¬7.
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1675, Torneo.ClasesTexto, e_FontTypeNames.FONTTYPE_CITIZEN)) 'Msg1675=Evento> Clases participantes: ¬1. Escribí /PARTICIPAR para ingresar al evento.
    Exit Sub
IniciarTorneo_Err:
    Call TraceError(Err.Number, Err.Description, "ModTorneos.IniciarTorneo", Erl)
End Sub

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
    Dim nombres As String
    Dim x       As Byte
    Dim y       As Byte
    For i = 1 To Torneo.participantes
        nombres = nombres & UserList(Torneo.IndexParticipantes(i)).name & ", "
        x = Torneo.x
        y = Torneo.y
        Call FindLegalPos(Torneo.IndexParticipantes(i), Torneo.Mapa, x, y)
        Call WarpUserChar(Torneo.IndexParticipantes(i), Torneo.Mapa, x, y, True)
    Next i
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1676, nombres, e_FontTypeNames.FONTTYPE_CITIZEN)) 'Msg1676=Evento> Los elegidos para participar son: ¬1 damos inicio al evento.
    Exit Sub
ComenzarTorneoOk_Err:
    Call TraceError(Err.Number, Err.Description, "ModTorneos.ComenzarTorneoOk", Erl)
End Sub

Public Sub ResetearTorneo()
    On Error GoTo ResetearTorneo_Err
    Dim i As Byte
    Torneo.HayTorneoaActivo = False
    Torneo.NivelMinimo = 0
    Torneo.NivelMaximo = 0
    Torneo.cupos = 0
    Torneo.costo = 0
    Torneo.mago = 0
    Torneo.clerico = 0
    Torneo.guerrero = 0
    Torneo.asesino = 0
    Torneo.bardo = 0
    Torneo.druido = 0
    Torneo.Paladin = 0
    Torneo.cazador = 0
    Torneo.Trabajador = 0
    Torneo.Pirata = 0
    Torneo.Ladron = 0
    Torneo.Bandido = 0
    Torneo.ClasesTexto = ""
    Torneo.Mapa = 0
    Torneo.x = 0
    Torneo.y = 0
    Torneo.nombre = ""
    Torneo.reglas = 0
    For i = 1 To Torneo.participantes
        UserList(Torneo.IndexParticipantes(i)).flags.EnTorneo = False
    Next i
    Torneo.participantes = 0
    ReDim Torneo.IndexParticipantes(1 To 1)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1677, vbNullString, e_FontTypeNames.FONTTYPE_CITIZEN)) 'Msg1677=Eventos> Evento Finalizado.
    Exit Sub
ResetearTorneo_Err:
    Call TraceError(Err.Number, Err.Description, "ModTorneos.ResetearTorneo", Erl)
End Sub
