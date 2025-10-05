Attribute VB_Name = "ModEventos"
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
Option Explicit
Public HoraEvento           As Byte
Public TiempoRestanteEvento As Byte
Public EventoActivo         As Boolean
Public EventoAcutal         As t_EventoPropiedades
Public Evento(0 To 23)      As t_EventoPropiedades

Public Type t_EventoPropiedades
    Tipo As Byte
    Duracion As Byte
    multiplicacion As Byte
End Type

Public ExpMultOld         As Integer
Public OroMultOld         As Integer
Public DropMultOld        As Integer
Public RecoleccionMultOld As Double
Public PublicidadEvento   As String
Enum TipoEvento
    Invasion
End Enum

Public Sub CheckEvento(ByVal Hora As Byte)
    On Error GoTo CheckEvento_Err
    If EventoActivo = True Then Exit Sub
    Dim aviso As String
    aviso = "Eventos> Nuevo evento iniciado: "
    PublicidadEvento = "Evento en curso>"
    Select Case Evento(Hora).Tipo
        Case 1
        Case 2
        Case 3
        Case 4
        Case 5
        Case 6
        Case 7
        Case Else
            EventoActivo = False
            Exit Sub
    End Select
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(553, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno
    Call AgregarAConsola(aviso)
    EventoAcutal.Duracion = Evento(Hora).Duracion
    EventoAcutal.multiplicacion = Evento(Hora).multiplicacion
    EventoAcutal.Tipo = Evento(Hora).Tipo
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(aviso, e_FontTypeNames.FONTTYPE_New_Eventos))
    TiempoRestanteEvento = Evento(Hora).Duracion
    frmMain.Evento.Enabled = True
    EventoActivo = True
    Exit Sub
CheckEvento_Err:
    Call TraceError(Err.Number, Err.Description, "ModEventos.CheckEvento", Erl)
End Sub

Public Sub FinalizarEvento()
    On Error GoTo FinalizarEvento_Err
    frmMain.Evento.Enabled = False
    EventoActivo = False
    Select Case EventoAcutal.Tipo
        Case 1
        Case 2
        Case 3
        Case 4
        Case 5
        Case 6
        Case 7
        Case Else
            Exit Sub
    End Select
    Call AgregarAConsola("Eventos > Evento finalizado.")
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1563, vbNullString, e_FontTypeNames.FONTTYPE_New_Eventos)) 'Msg1563=Eventos > Evento finalizado.
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(551, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno
    Exit Sub
FinalizarEvento_Err:
    Call TraceError(Err.Number, Err.Description, "ModEventos.FinalizarEvento", Erl)
End Sub

Public Function DescribirEvento(ByVal Hora As Byte) As String
    On Error GoTo DescribirEvento_Err
    Dim aviso As String
    aviso = "("
    Select Case Evento(Hora).Tipo
        Case 1
            aviso = aviso & "Oro multiplicado por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).Duracion & " minutos"
        Case 2
            aviso = aviso & "Experiencia multiplicada por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).Duracion & " minutos"
        Case 3
            aviso = aviso & "Recoleccion multiplicada por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).Duracion & " minutos"
        Case 4
            aviso = aviso & "Dropeo multiplicado por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).Duracion & " minutos"
        Case 5
            aviso = aviso & "Oro y experiencia multiplicados por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).Duracion & " minutos"
        Case 6
            aviso = aviso & "Oro, experiencia y recoleccion multiplicados por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).Duracion & " minutos"
        Case 7
            aviso = aviso & "Oro, experiencia, recoleccion y dropeo multiplicados por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).Duracion & " minutos"
        Case Else
            aviso = aviso & "sin información"
    End Select
    aviso = aviso & ")"
    DescribirEvento = aviso
    Exit Function
DescribirEvento_Err:
    Call TraceError(Err.Number, Err.Description, "ModEventos.DescribirEvento", Erl)
End Function

Public Sub CargarEventos()
    On Error GoTo CargarEventos_Err
    Dim i          As Byte
    Dim EventoStrg As String
    For i = 0 To 23
        EventoStrg = GetVar(IniPath & "Configuracion.ini", "EVENTOS", i)
        Evento(i).Tipo = val(ReadField(1, EventoStrg, Asc("-")))
        Evento(i).Duracion = val(ReadField(2, EventoStrg, Asc("-")))
        Evento(i).multiplicacion = val(ReadField(3, EventoStrg, Asc("-")))
    Next i
    Exit Sub
CargarEventos_Err:
    Call TraceError(Err.Number, Err.Description, "ModEventos.CargarEventos", Erl)
End Sub

Public Sub ForzarEvento(ByVal Tipo As Byte, ByVal Duracion As Byte, ByVal multi As Byte, ByVal Quien As String)
    On Error GoTo ForzarEvento_Err
    Dim tUser As t_UserReference
    tUser = NameIndex(Quien)
    If Not IsValidUserRef(tUser) Then
        Call LogError("Failed to force event, unknown user: " & Quien)
        Exit Sub
    End If
    If Tipo > 3 Or Tipo < 1 Then
        Call WriteLocaleMsg(tUser.ArrayIndex, 2071, e_FontTypeNames.FONTTYPE_New_Eventos) ' Msg2071="Tipo de evento invalido."
        Exit Sub
    End If
    If Duracion > 59 Then
        Call WriteLocaleMsg(tUser.ArrayIndex, 2072, e_FontTypeNames.FONTTYPE_New_Eventos) ' Msg2072="Duracion invalida. maxima 59 minutos."
        Exit Sub
    End If
    If (Tipo = 1 And multi > 2) Then
        Call WriteLocaleMsg(tUser.ArrayIndex, 2073, e_FontTypeNames.FONTTYPE_New_Eventos) ' Msg2073="Multiplicacion invalida. maxima x2."
        Exit Sub
    End If
    If (Tipo = 2 And multi > 2) Then
        Call WriteLocaleMsg(tUser.ArrayIndex, 2074, e_FontTypeNames.FONTTYPE_New_Eventos) ' Msg2074="Multiplicacion invalida. maxima x2."
        Exit Sub
    End If
    If (Tipo = 3 And multi > 5) Then
        Call WriteLocaleMsg(tUser.ArrayIndex, 2075, e_FontTypeNames.FONTTYPE_New_Eventos) ' Msg2075="Multiplicacion invalida. maxima x5."
        Exit Sub
    End If
    Dim aviso As String
    aviso = "Eventos> " & Quien & " inicio un nuevo evento: "
    PublicidadEvento = "Evento en curso>"
    Select Case Tipo
        Case 1
        Case 2
        Case 3
        Case 4
        Case 5
        Case 6
        Case 7
        Case Else
            EventoActivo = False
            Exit Sub
    End Select
    Call AgregarAConsola(aviso)
    EventoAcutal.Duracion = Duracion
    EventoAcutal.multiplicacion = multi
    EventoAcutal.Tipo = Tipo
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(aviso, e_FontTypeNames.FONTTYPE_New_Eventos))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(553, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno
    TiempoRestanteEvento = Duracion
    frmMain.Evento.Enabled = True
    EventoActivo = True
    Exit Sub
ForzarEvento_Err:
    Call TraceError(Err.Number, Err.Description, "ModEventos.ForzarEvento", Erl)
End Sub

Public Sub IniciarEvento(ByVal Tipo As TipoEvento, ByVal data As Variant)
    Select Case Tipo
        Case TipoEvento.Invasion
            Call IniciarInvasion(data)
    End Select
End Sub
