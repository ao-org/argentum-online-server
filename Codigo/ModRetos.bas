Attribute VB_Name = "ModRetos"
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
Option Explicit
Private Const APUESTA_MAXIMA = 100000000
Public Retos          As t_Retos
Private ListaDeEspera As New Dictionary

Public Sub CargarInfoRetos()
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

Public Sub CrearReto(ByVal UserIndex As Integer, JugadoresStr As String, ByVal Apuesta As Long, ByVal PocionesMaximas As Integer, Optional ByVal CaenItems As Boolean = False)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        If .flags.SolicitudReto.Estado <> e_SolicitudRetoEstado.Libre Then
            Call CancelarSolicitudReto(UserIndex, .name & " ha cancelado la solicitud.")
        ElseIf IsValidUserRef(.flags.AceptoReto) Then
            Call CancelarSolicitudReto(.flags.AceptoReto.ArrayIndex, .name & " ha cancelado su admisión.")
        End If
        Dim TamanoReal As Byte: TamanoReal = Retos.TamañoMaximoEquipo * 2 - 1
        If LenB(JugadoresStr) <= 0 Then Exit Sub
        Dim Jugadores() As String: Jugadores = Split(JugadoresStr, ";", TamanoReal)
        If UBound(Jugadores) > TamanoReal - 1 Or UBound(Jugadores) Mod 2 = 1 Then Exit Sub
        Dim MaxIndexEquipo As Integer: MaxIndexEquipo = UBound(Jugadores) \ 2
        If Apuesta < Retos.ApuestaMinima Or Apuesta > APUESTA_MAXIMA Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1958, PonerPuntos(Retos.ApuestaMinima), e_FontTypeNames.FONTTYPE_INFO)) ' Msg1958=La apuesta mínima es de ¬1 monedas de oro.
            Exit Sub
        End If
        If Not PuedeRetoConMensaje(UserIndex) Then Exit Sub
        If .Stats.GLD < Apuesta Then
            ' Msg588=No tienes el oro suficiente.
            Call WriteLocaleMsg(UserIndex, 588, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If PocionesMaximas >= 0 Then
            If TieneObjetos(38, PocionesMaximas + 1, UserIndex) Then
                Call WriteLocaleMsg(UserIndex, 1443, e_FontTypeNames.FONTTYPE_INFO, PocionesMaximas) ' Msg1443=Tienes demasiadas pociones rojas (Cantidad máxima: ¬1).
                Exit Sub
            End If
        End If
        With .flags.SolicitudReto
            .Apuesta = Apuesta
            .PocionesMaximas = PocionesMaximas
            .CaenItems = CaenItems
            ReDim .Jugadores(0 To UBound(Jugadores))
            Dim i       As Integer, tIndex As t_UserReference
            Dim Equipo1 As String, Equipo2 As String
            Equipo1 = UserList(UserIndex).name
            For i = 0 To UBound(.Jugadores)
                With .Jugadores(i)
                    If EsGmChar(Jugadores(i)) Then
                        Call WriteLocaleMsg(UserIndex, 1444, e_FontTypeNames.FONTTYPE_INFO) ' Msg1444=¡No puedes jugar retos con administradores!
                        Exit Sub
                    End If
                    tIndex = NameIndex(Jugadores(i))
                    If Not IsValidUserRef(tIndex) Then
                        Call WriteLocaleMsg(UserIndex, 1445, e_FontTypeNames.FONTTYPE_INFO, Jugadores(i)) ' Msg1445=El usuario ¬1 no puede jugar un reto en este momento.
                        Exit Sub
                    End If
                    If Not PuedeReto(tIndex.ArrayIndex) Then
                        Call WriteLocaleMsg(UserIndex, 1445, e_FontTypeNames.FONTTYPE_INFO, UserList(tIndex.ArrayIndex).name) ' Msg1445=El usuario ¬1 no puede jugar un reto en este momento.
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
            Texto2 = Equipo1 & " vs " & Equipo2 & ". Apuesta: " & PonerPuntos(Apuesta) & " monedas de oro" & IIf(CaenItems, " y los items.", ".")
            Texto3 = "Escribe /ACEPTAR " & UCase$(UserList(UserIndex).name) & " para participar en el reto."
            If PocionesMaximas >= 0 Then
                Texto2 = Texto2 & " Máximo " & PocionesMaximas & " pociones rojas."
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
        Call WriteLocaleMsg(UserIndex, 1446, e_FontTypeNames.FONTTYPE_INFO) ' Msg1446=Has enviado una solicitud para el siguiente reto:
        Call WriteConsoleMsg(UserIndex, Texto2, e_FontTypeNames.FONTTYPE_New_Naranja)
        Call WriteLocaleMsg(UserIndex, 1447, e_FontTypeNames.FONTTYPE_New_Gris) ' Msg1447=Escribe /CANCELAR para anular la solicitud.
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.CrearReto", Erl)
End Sub

Public Sub AceptarReto(ByVal UserIndex As Integer, OferenteName As String)
    On Error GoTo ErrHandler
    If Not PuedeRetoConMensaje(UserIndex) Then Exit Sub
    With UserList(UserIndex)
        If .flags.SolicitudReto.Estado <> e_SolicitudRetoEstado.Libre Then
            Call CancelarSolicitudReto(UserIndex, .name & " ha cancelado la solicitud.")
        ElseIf IsValidUserRef(.flags.AceptoReto) Then
            Call CancelarSolicitudReto(.flags.AceptoReto.ArrayIndex, .name & " ha cancelado su admisión.")
        End If
    End With
    If EsGmChar(OferenteName) Then
        Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1959, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1959=¡No puedes jugar retos con administradores!
        Exit Sub
    End If
    Dim Oferente As t_UserReference
    Oferente = NameIndex(OferenteName)
    If Not IsValidUserRef(Oferente) Then
        Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1960, OferenteName, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1960=¬1 no está conectado.
        Exit Sub
    End If
    With UserList(Oferente.ArrayIndex).flags.SolicitudReto
        Dim JugadorIndex As Integer
        JugadorIndex = IndiceJugadorEnSolicitud(UserIndex, Oferente.ArrayIndex)
        If JugadorIndex < 0 Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1961, UserList(Oferente.ArrayIndex).name, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1961=¬1 no te ha invitado a ningún reto o ha sido cancelado.
            Exit Sub
        End If
        If UserList(UserIndex).Stats.GLD < .Apuesta Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1962, PonerPuntos(.Apuesta), e_FontTypeNames.FONTTYPE_INFO)) ' Msg1962=Necesitas al menos ¬1 monedas de oro para aceptar este reto.
            Exit Sub
        End If
        If .PocionesMaximas >= 0 Then
            If TieneObjetos(38, .PocionesMaximas + 1, UserIndex) Then
                Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1963, .PocionesMaximas, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1963=Tienes demasiadas pociones rojas (Cantidad máxima: ¬1).
                Exit Sub
            End If
        End If
        Call MensajeATodosSolicitud(Oferente.ArrayIndex, UserList(UserIndex).name & " ha aceptado el reto.", e_FontTypeNames.FONTTYPE_INFO)
        .Jugadores(JugadorIndex).Aceptado = True
        Call SetUserRef(.Jugadores(JugadorIndex).CurIndex, UserIndex)
        UserList(UserIndex).flags.AceptoReto = Oferente
        Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1964, UserList(Oferente.ArrayIndex).name, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1964=Has aceptado el reto de ¬1.
        Dim FaltanAceptar As String
        Dim i             As Integer
        For i = 0 To UBound(.Jugadores)
            If Not .Jugadores(i).Aceptado Then
                FaltanAceptar = FaltanAceptar & .Jugadores(i).nombre & " - "
            End If
        Next
        If LenB(FaltanAceptar) > 0 Then
            FaltanAceptar = Left$(FaltanAceptar, Len(FaltanAceptar) - 3)
            Call MensajeATodosSolicitud(Oferente.ArrayIndex, "Faltan aceptar: " & FaltanAceptar, e_FontTypeNames.FONTTYPE_New_Gris)
            Exit Sub
        End If
        Call MensajeATodosSolicitud(Oferente.ArrayIndex, "Todos los jugadores han aceptado el reto. Buscando sala...", e_FontTypeNames.FONTTYPE_New_Gris)
        Call BuscarSala(Oferente.ArrayIndex)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.AceptarReto", Erl)
End Sub

Public Sub CancelarSolicitudReto(ByVal Oferente As Integer, mensaje As String)
    On Error GoTo ErrHandler
    With UserList(Oferente).flags.SolicitudReto
        If .Estado = e_SolicitudRetoEstado.EnCola Then
            Call ListaDeEspera.Remove(Oferente)
        End If
        .Estado = e_SolicitudRetoEstado.Libre
        Dim i As Integer, tUser As t_UserReference
        ' Enviamos a los invitados
        For i = 0 To UBound(.Jugadores)
            tUser = NameIndex(.Jugadores(i).nombre)
            If IsValidUserRef(tUser) Then
                Call WriteConsoleMsg(tUser.ArrayIndex, mensaje, e_FontTypeNames.FONTTYPE_WARNING)
                Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(1965, vbNullString, e_FontTypeNames.FONTTYPE_WARNING)) ' Msg1965=El reto ha sido cancelado.
                If .Jugadores(i).Aceptado Then
                    Call SetUserRef(UserList(tUser.ArrayIndex).flags.AceptoReto, 0)
                End If
            End If
        Next
        ' Y al oferente por separado
        Call WriteConsoleMsg(Oferente, mensaje, e_FontTypeNames.FONTTYPE_WARNING)
        Call WriteConsoleMsg(Oferente, PrepareMessageLocaleMsg(1965, vbNullString, e_FontTypeNames.FONTTYPE_WARNING)) ' Msg1965=El reto ha sido cancelado.
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.CancelarSolicitudReto", Erl)
End Sub

Private Sub BuscarSala(ByVal Oferente As Integer)
    On Error GoTo ErrHandler
    With UserList(Oferente).flags.SolicitudReto
        If Retos.SalasLibres <= 0 Then
            Call ListaDeEspera.Add(Oferente, 0)
            Call MensajeATodosSolicitud(Oferente, "No hay salas disponibles. El reto comenzará cuando se desocupe una sala.", e_FontTypeNames.FONTTYPE_FIGHT)
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
        Call IniciarReto(Oferente, Sala)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.BuscarSala", Erl)
End Sub

Private Sub IniciarReto(ByVal Oferente As Integer, ByVal Sala As Integer)
    On Error GoTo ErrHandler
    With UserList(Oferente).flags.SolicitudReto
        ' Última comprobación de si todos pueden entrar/pagar
        If Not TodosPuedenReto(Oferente) Then Exit Sub
        Dim Apuesta As Long, ApuestaStr As String
        Apuesta = .Apuesta
        ApuestaStr = PonerPuntos(Apuesta)
        ' Calculamos el tamaño del equipo
        Retos.Salas(Sala).TamañoEquipoIzq = UBound(.Jugadores) \ 2 + 1
        Retos.Salas(Sala).TamañoEquipoDer = Retos.Salas(Sala).TamañoEquipoIzq
        ' Reservamos espacio para los jugadores (incluyendo al oferente)
        ReDim Retos.Salas(Sala).Jugadores(0 To UBound(.Jugadores) + 1)
        ' Tiramos una moneda (50-50) y decidimos si agregar al oferente al inicio o al final de la lista
        Dim Moneda As Byte
        Moneda = RandomNumber(0, 1)
        Dim CurIndex As Integer
        If Moneda = 0 Then
            ' Agregamos al oferente al inicio (su equipo juega a la izquierda)
            Call SetUserRef(Retos.Salas(Sala).Jugadores(CurIndex), Oferente)
            CurIndex = CurIndex + 1
        End If
        Dim i As Integer
        ' Agregamos los jugadores alternando 1 y 1 (en los índices pares está el equipo izquierdo y en los impares el derecho - el array empieza en cero)
        For i = 0 To UBound(.Jugadores)
            Retos.Salas(Sala).Jugadores(CurIndex) = .Jugadores(i).CurIndex
            CurIndex = CurIndex + 1
            ' Reset flag
            Call SetUserRef(UserList(.Jugadores(i).CurIndex.ArrayIndex).flags.AceptoReto, 0)
        Next
        If Moneda = 1 Then
            ' Agregamos al oferente al final (su equipo juega a la derecha)
            Call SetUserRef(Retos.Salas(Sala).Jugadores(CurIndex), Oferente)
        End If
        ' Reset estado de la solicitud, ya que no la necesitamos más
        .Estado = e_SolicitudRetoEstado.Libre
    End With
    With Retos.Salas(Sala)
        .EnUso = True
        .Puntaje = 0
        .Ronda = 1
        .Apuesta = Apuesta
        .TiempoRestante = Retos.DuracionMaxima
        .CaenItems = UserList(Oferente).flags.SolicitudReto.CaenItems
        Dim tUser As t_UserReference
        For i = 0 To UBound(.Jugadores)
            tUser = .Jugadores(i)
            ' Le cobramos
            UserList(tUser.ArrayIndex).Stats.GLD = UserList(tUser.ArrayIndex).Stats.GLD - Apuesta
            Call WriteUpdateGold(tUser.ArrayIndex)
            Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(1966, ApuestaStr, e_FontTypeNames.FONTTYPE_New_Rojo_Salmon)) ' Msg1966=Otorgas ¬1 monedas de oro al pozo del reto.
            ' Desmontamos
            If UserList(tUser.ArrayIndex).flags.Montado <> 0 Then
                Call DoMontar(tUser.ArrayIndex, ObjData(UserList(tUser.ArrayIndex).invent.EquippedSaddleObjIndex), UserList(tUser.ArrayIndex).invent.EquippedSaddleSlot)
            End If
            ' Dejamos de navegar
            If UserList(tUser.ArrayIndex).flags.Nadando <> 0 Or UserList(tUser.ArrayIndex).flags.Navegando <> 0 Then
                Call DoNavega(tUser.ArrayIndex, ObjData(UserList(tUser.ArrayIndex).invent.EquippedShipObjIndex), UserList(tUser.ArrayIndex).invent.EquippedShipSlot)
            End If
            ' Asignamos flags
            With UserList(tUser.ArrayIndex).flags
                .EnReto = True
                .EquipoReto = IIf(i Mod 2, e_EquipoReto.Derecha, e_EquipoReto.Izquierda)
                .SalaReto = Sala
                ' Guardar posición
                .LastPos = UserList(tUser.ArrayIndex).pos
            End With
            Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(1967, vbNullString, e_FontTypeNames.FONTTYPE_New_Rojo_Salmon)) ' Msg1967=¡Ha comenzado el reto!
            Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(1968, vbNullString, e_FontTypeNames.FONTTYPE_New_Gris)) ' Msg1968=Para admitir la derrota escribe /ABANDONAR.
        Next
    End With
    Retos.SalasLibres = Retos.SalasLibres - 1
    Call iniciarRonda(Sala)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.IniciarReto", Erl)
End Sub

Private Sub iniciarRonda(ByVal Sala As Integer)
    With Retos.Salas(Sala)
        Dim i As Integer, tUser As t_UserReference
        For i = 0 To UBound(.Jugadores)
            tUser = .Jugadores(i)
            If tUser.ArrayIndex <> 0 Then
                Call RevivirYLimpiar(tUser.ArrayIndex)
                ' Usando el número de ronda y el índice, decidimos el lado al que corresponde
                If (.Ronda + i) Mod 2 = 1 Then
                    ' Lado izquierdo
                    Call WarpToLegalPos(tUser.ArrayIndex, .PosIzquierda.Map, .PosIzquierda.x, .PosIzquierda.y, True)
                Else
                    ' Lado derecho
                    Call WarpToLegalPos(tUser.ArrayIndex, .PosDerecha.Map, .PosDerecha.x, .PosDerecha.y, True)
                End If
                ' Si usamos el conteo
                If Retos.TiempoConteo > 0 Then
                    ' Le ponemos el conteo
                    UserList(tUser.ArrayIndex).Counters.CuentaRegresiva = Retos.TiempoConteo
                    ' Lo stoppeamos
                    Call WriteStopped(tUser.ArrayIndex, True)
                End If
                Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(1969, .Ronda, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1969=Comienza la ronda Nº¬1
            End If
        Next
    End With
End Sub

Public Sub MuereEnReto(ByVal UserIndex As Integer)
    On Error GoTo ErrorHandler
    Dim Sala As Integer, Equipo As e_EquipoReto
    With UserList(UserIndex)
        Sala = .flags.SalaReto
        Equipo = .flags.EquipoReto
    End With
    With Retos.Salas(Sala)
        Dim CurIndex As Integer
        ' El equipo derecho está en índices pares
        If Equipo = e_EquipoReto.Derecha Then CurIndex = 1
        For CurIndex = CurIndex To UBound(.Jugadores) Step 2
            If .Jugadores(CurIndex).ArrayIndex <> 0 Then
                ' Si todavía hay alguno vivo del equipo
                If UserList(.Jugadores(CurIndex).ArrayIndex).flags.Muerto = 0 Then
                    Exit Sub
                End If
            End If
        Next
        ' Están todos muertos, ganó el equipo contrario
        Call ProcesarRondaGanada(Sala, EquipoContrario(Equipo))
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.MuereEnReto", Erl)
End Sub

Private Sub ProcesarRondaGanada(ByVal Sala As Integer, ByVal Equipo As e_EquipoReto)
    With Retos.Salas(Sala)
        ' Sumamos puntaje o restamos según el equipo
        If Equipo = e_EquipoReto.Derecha Then
            .Puntaje = .Puntaje + 1
        Else
            .Puntaje = .Puntaje - 1
        End If
        ' Si terminó la tercer ronda o bien algún equipo obtuvo 2 victorias seguidas
        If .Ronda >= 3 Or Abs(.Puntaje) >= 2 Then
            Call FinalizarReto(Sala)
            Exit Sub
        End If
        ' Aumentamos el número de ronda
        .Ronda = .Ronda + 1
        ' Obtenemos el tamaño actual del equipo (por si alguno abandonó)
        Dim TamañoEquipo As Integer, TamañoEquipo2 As Integer
        TamañoEquipo = ObtenerTamañoEquipo(Sala, Equipo)
        ' Menos cálculos en el bucle
        TamañoEquipo2 = TamañoEquipo * 2
        ' Obtenemos los nombres del equipo ganador
        Dim i As Integer, nombres As String
        For i = IIf(Equipo = e_EquipoReto.Izquierda, 0, 1) To TamañoEquipo2 - 1 Step 2
            If .Jugadores(i).ArrayIndex <> 0 Then
                nombres = nombres & UserList(.Jugadores(i).ArrayIndex).name
                If i < TamañoEquipo2 - 2 Then
                    nombres = nombres & IIf(i > TamañoEquipo2 - 5, " y ", ", ")
                End If
            End If
        Next
        ' Informamos el ganador de esta ronda
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i).ArrayIndex <> 0 Then
                Call WriteConsoleMsg(.Jugadores(i).ArrayIndex, PrepareMessageLocaleMsg(1970, nombres, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1970=Esta ronda es para ¬1.
                Call WriteConsoleMsg(.Jugadores(i).ArrayIndex, "", 0) ' Dejamos un espacio vertical
            End If
        Next
        ' Iniciamos la próxima ronda
        Call iniciarRonda(Sala)
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.ProcesarRondaGanada", Erl)
End Sub

Public Sub FinalizarReto(ByVal Sala As Integer, Optional ByVal TiempoAgotado As Boolean)
    On Error GoTo ErrorHandler
    With Retos.Salas(Sala)
        ' Calculamos el oro total del premio
        Dim OroTotal As Long, Oro As Long, OroStr As String
        OroTotal = .Apuesta * (UBound(.Jugadores) + 1)
        ' Descontamos el impuesto
        OroTotal = OroTotal * (1 - Retos.ImpuestoApuesta)
        ' Decidimos el resultado del reto según el puntaje:
        Dim i                 As Integer, tUser As t_UserReference, Equipo1 As String, Equipo2 As String
        Dim eloTotalIzquierda As Long, eloTotalDerecha As Long, winsIzquierda As Long, winsDerecha As Long
        Dim todosMayorA35     As Boolean
        todosMayorA35 = True
        For i = 0 To UBound(.Jugadores)
            tUser = .Jugadores(i)
            If tUser.ArrayIndex <> 0 Then
                todosMayorA35 = todosMayorA35 And (UserList(tUser.ArrayIndex).Stats.ELV >= 35)
                If i Mod 2 = 0 Then
                    eloTotalIzquierda = eloTotalIzquierda + UserList(tUser.ArrayIndex).Stats.ELO
                Else
                    eloTotalDerecha = eloTotalDerecha + UserList(tUser.ArrayIndex).Stats.ELO
                End If
            End If
        Next i
        ' Empate
        If .Puntaje = 0 Then
            ' Pagamos a todos los que no abandonaron
            Oro = OroTotal \ (UBound(.Jugadores) + 1)
            OroStr = PonerPuntos(Oro)
            ' No hubo ganadores, entonces el ELO no les da el bonus.
            winsIzquierda = 0
            winsDerecha = 0
            For i = 0 To UBound(.Jugadores)
                tUser = .Jugadores(i)
                If IsValidUserRef(tUser) Then
                    UserList(tUser.ArrayIndex).Stats.GLD = UserList(tUser.ArrayIndex).Stats.GLD + Oro
                    Call WriteUpdateGold(tUser.ArrayIndex)
                    Call WriteLocaleMsg(tUser.ArrayIndex, "29", e_FontTypeNames.FONTTYPE_MP, OroStr) ' Has ganado X monedas de oro
                    Call RevivirYLimpiar(tUser.ArrayIndex)
                    Call DevolverPosAnterior(tUser.ArrayIndex)
                    ' Reset flags
                    UserList(tUser.ArrayIndex).Counters.CuentaRegresiva = -1
                    UserList(tUser.ArrayIndex).flags.EnReto = False
                    ' Nombres
                    If i Mod 2 Then
                        If LenB(Equipo2) > 0 Then
                            Equipo2 = Equipo2 & IIf((i + 1) \ 2 < .TamañoEquipoDer - 2, ", ", " y ") & UserList(tUser.ArrayIndex).name
                        Else
                            Equipo2 = UserList(tUser.ArrayIndex).name
                        End If
                    Else
                        If LenB(Equipo1) > 0 Then
                            Equipo1 = Equipo2 & IIf(i \ 2 < .TamañoEquipoIzq - 2, ", ", " y ") & UserList(tUser.ArrayIndex).name
                        Else
                            Equipo1 = UserList(tUser.ArrayIndex).name
                        End If
                    End If
                End If
            Next
            ' Anuncio global
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1670, Equipo1 & "¬" & Equipo2, e_FontTypeNames.FONTTYPE_INFO)) 'Msg1670=Retos » ¬1 vs ¬2. Ninguno pudo vencer a su rival.
            Call SalaLiberada(Sala)
            ' Hubo un ganador
        Else
            Dim Ganador As e_EquipoReto
            If .Puntaje < 0 Then
                Ganador = e_EquipoReto.Izquierda
                winsIzquierda = .TamañoEquipoDer
                winsDerecha = -.TamañoEquipoIzq
            Else
                Ganador = e_EquipoReto.Derecha
                winsIzquierda = -.TamañoEquipoDer
                winsDerecha = .TamañoEquipoIzq
            End If
            ' Pagamos a los ganadores que no abandonaron
            Oro = OroTotal \ ObtenerTamañoEquipo(Sala, Ganador)
            OroStr = PonerPuntos(Oro)
            For i = 0 To UBound(.Jugadores)
                tUser = .Jugadores(i)
                If IsValidUserRef(tUser) Then
                    Call RevivirYLimpiar(tUser.ArrayIndex)
                    If UserList(tUser.ArrayIndex).flags.EquipoReto = Ganador Then
                        UserList(tUser.ArrayIndex).Stats.GLD = UserList(tUser.ArrayIndex).Stats.GLD + Oro
                        Call WriteUpdateGold(tUser.ArrayIndex)
                        Call WriteLocaleMsg(tUser.ArrayIndex, "29", e_FontTypeNames.FONTTYPE_MP, OroStr) ' Has ganado X monedas de oro
                        If .CaenItems Then
                            Call WarpToLegalPos(tUser.ArrayIndex, .PosIzquierda.Map, .PosIzquierda.x, .PosIzquierda.y, True)
                        Else
                            UserList(tUser.ArrayIndex).flags.EnReto = False
                            Call DevolverPosAnterior(tUser.ArrayIndex)
                        End If
                    Else
                        If .CaenItems Then
                            Call TirarItemsEnPos(tUser.ArrayIndex, ((.PosDerecha.x - .PosIzquierda.x) \ 2) + .PosIzquierda.x, ((.PosDerecha.y - .PosIzquierda.y) \ 2) + _
                                    .PosIzquierda.y)
                        End If
                        UserList(tUser.ArrayIndex).flags.EnReto = False
                        Call DevolverPosAnterior(tUser.ArrayIndex)
                    End If
                    ' Reset flags
                    UserList(tUser.ArrayIndex).Counters.CuentaRegresiva = -1
                    If TiempoAgotado Then
                        Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(1971, vbNullString, e_FontTypeNames.FONTTYPE_New_Gris)) ' Msg1971=Se ha agotado el tiempo del reto.
                    End If
                    ' Nombres
                    If i Mod 2 Then
                        If LenB(Equipo2) > 0 Then
                            Equipo2 = Equipo2 & IIf((i + 1) \ 2 < .TamañoEquipoDer - 2, ", ", " y ") & UserList(tUser.ArrayIndex).name
                        Else
                            Equipo2 = UserList(tUser.ArrayIndex).name
                        End If
                    Else
                        If LenB(Equipo1) > 0 Then
                            Equipo1 = Equipo1 & IIf(i \ 2 < .TamañoEquipoIzq - 2, ", ", " y ") & UserList(tUser.ArrayIndex).name
                        Else
                            Equipo1 = UserList(tUser.ArrayIndex).name
                        End If
                    End If
                End If
            Next
            Dim equipoGanador As String, equipoPerdedor As String
            equipoGanador = IIf(Ganador = e_EquipoReto.Izquierda, Equipo1, Equipo2)
            equipoPerdedor = IIf(Ganador = e_EquipoReto.Izquierda, Equipo2, Equipo1)
            ' Anuncio global
            If UBound(.Jugadores) > 1 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1671, equipoGanador & "¬" & equipoPerdedor & "¬" & PonerPuntos(.Apuesta), _
                        e_FontTypeNames.FONTTYPE_INFO)) 'Msg1671=Retos » El equipo ¬1 venció al equipo ¬2 y se quedó con el botín de: ¬3 monedas de oro.
            Else ' 1 vs 1
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1672, equipoGanador & "¬" & equipoPerdedor & "¬" & PonerPuntos(.Apuesta), _
                        e_FontTypeNames.FONTTYPE_INFO)) 'Msg1672=Retos » ¬1 venció a ¬2 y se quedó con el botín de: ¬3 monedas de oro.
            End If
            If .CaenItems Then
                Call IniciarDepositoItems(Sala)
            Else
                Call SalaLiberada(Sala)
            End If
        End If
        ' Actualizamos el ELO de cada jugador, inspirados en `Algoritmo de 400`
        ' https://en.wikipedia.org/wiki/Elo_rating_system
        Dim eloDiff As Long
        For i = 0 To UBound(.Jugadores)
            tUser = .Jugadores(i)
            If IsValidUserRef(tUser) Then
                If todosMayorA35 Then
                    If i Mod 2 = 0 Then ' Jugadores en el equipo Izquierdo
                        eloDiff = winsIzquierda * (eloTotalDerecha * 0.1)
                    Else
                        eloDiff = winsDerecha * (eloTotalIzquierda * 0.1)
                    End If
                    If eloDiff > 0 Then
                        Call SendData(SendTarget.ToIndex, tUser.ArrayIndex, PrepareMessageLocaleMsg(1695, Abs(eloDiff), e_FontTypeNames.FONTTYPE_ROSA)) 'Msg1695=Has ganado ¬1 puntos de ELO!
                    Else
                        If UserList(tUser.ArrayIndex).Stats.ELO < Abs(eloDiff) Then
                            eloDiff = -UserList(tUser.ArrayIndex).Stats.ELO
                        End If
                        Call SendData(SendTarget.ToIndex, tUser.ArrayIndex, PrepareMessageLocaleMsg(1696, Abs(eloDiff), e_FontTypeNames.FONTTYPE_ROSA)) 'Msg1696=Has perdido ¬1 puntos de ELO!
                    End If
                    UserList(tUser.ArrayIndex).Stats.ELO = UserList(tUser.ArrayIndex).Stats.ELO + eloDiff
                Else ' Alguno es menor a level 35
                    Call SendData(SendTarget.ToIndex, tUser.ArrayIndex, PrepareMessageLocaleMsg(1697, vbNullString, e_FontTypeNames.FONTTYPE_INFOIAO)) 'Msg1697=Al menos un participante del reto tiene nivel menor a 35, tu ELO permanece igual.
                End If
            End If
        Next i
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.FinalizarReto", Erl)
End Sub

Public Sub TirarItemsEnPos(ByVal UserIndex As Integer, ByVal x As Byte, ByVal y As Byte)
    On Error GoTo TirarItemsEnPos_Err
    Dim i         As Byte
    Dim NuevaPos  As t_WorldPos
    Dim MiObj     As t_Obj
    Dim ItemIndex As Integer
    Dim posItems  As t_WorldPos
    With UserList(UserIndex)
        posItems.Map = .pos.Map
        posItems.x = x
        posItems.y = y
        For i = 1 To .CurrentInventorySlots
            ItemIndex = .invent.Object(i).ObjIndex
            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) And PirataCaeItem(UserIndex, i) And (Not EsNewbie(UserIndex) Or Not ItemNewbie(ItemIndex)) Then
                    NuevaPos.x = 0
                    NuevaPos.y = 0
                    MiObj.amount = .invent.Object(i).amount
                    MiObj.ObjIndex = ItemIndex
                    MiObj.ElementalTags = .invent.Object(i).ElementalTags
                    Call Tilelibre(posItems, NuevaPos, MiObj, True, True, False)
                    If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then
                        Call DropObj(UserIndex, i, MiObj.amount, NuevaPos.Map, NuevaPos.x, NuevaPos.y)
                        '  Si no hay lugar, quemamos el item del inventario (nada de mochilas gratis)
                    Else
                        Call QuitarUserInvItem(UserIndex, i, MiObj.amount)
                        Call UpdateUserInv(False, UserIndex, i)
                    End If
                End If
            End If
        Next i
    End With
    Exit Sub
TirarItemsEnPos_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.TirarItemsEnPos", Erl)
End Sub

Public Sub IniciarDepositoItems(ByVal Sala As Integer)
    Dim i       As Byte
    Dim Ganador As e_EquipoReto
    With Retos.Salas(Sala)
        If .Puntaje < 0 Then
            Ganador = e_EquipoReto.Izquierda
        Else
            Ganador = e_EquipoReto.Derecha
        End If
        For i = 0 To UBound(.Jugadores)
            If UserList(.Jugadores(i).ArrayIndex).flags.EquipoReto = Ganador Then
                Call WriteConsoleMsg(.Jugadores(i).ArrayIndex, PrepareMessageLocaleMsg(1972, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1972=Tienes 1 minuto para levantar los items del piso.
            End If
        Next i
        Dim pos As t_WorldPos
        pos.Map = .PosIzquierda.Map
        pos.x = ((.PosDerecha.x - .PosIzquierda.x) \ 2) + .PosIzquierda.x
        pos.y = ((.PosDerecha.y - .PosIzquierda.y) \ 2) + .PosIzquierda.y
        'Spawneo un banquero.
        .IndexBanquero = SpawnNpc(3, pos, True, False)
        #If DEBUGGING Then
            .TiempoItems = 20
        #Else
            .TiempoItems = 60
        #End If
    End With
End Sub

Public Sub TerminarTiempoAgarrarItems(ByVal Sala As Integer)
    Dim Ganador As e_EquipoReto
    With Retos.Salas(Sala)
        'Mato al banquero
        Call QuitarNPC(.IndexBanquero, eChallenge)
        If .Puntaje < 0 Then
            Ganador = e_EquipoReto.Izquierda
        Else
            Ganador = e_EquipoReto.Derecha
        End If
        Dim i As Byte
        For i = 0 To UBound(.Jugadores)
            If IsValidUserRef(.Jugadores(i)) Then
                If UserList(.Jugadores(i).ArrayIndex).flags.EquipoReto = Ganador Then
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
    Call SalaLiberada(Sala)
End Sub

Public Sub AbandonarReto(ByVal UserIndex As Integer, Optional ByVal Desconexion As Boolean)
    Dim Sala As Integer, Equipo As e_EquipoReto
    With UserList(UserIndex)
        Sala = .flags.SalaReto
        Equipo = .flags.EquipoReto
        .Counters.CuentaRegresiva = -1
        .flags.EnReto = False
    End With
    With Retos.Salas(Sala)
        If .CaenItems And Abs(.Puntaje) >= 2 Then
            If .Puntaje < 0 Then
                .TamañoEquipoIzq = .TamañoEquipoIzq - 1
                If .TamañoEquipoIzq <= 0 Then
                    TerminarTiempoAgarrarItems (Sala)
                End If
            Else
                .TamañoEquipoDer = .TamañoEquipoDer - 1
                If .TamañoEquipoDer <= 0 Then
                    TerminarTiempoAgarrarItems (Sala)
                End If
            End If
            Exit Sub
        End If
        If Not Desconexion Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1973, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1973=Has abandonado el reto.
        End If
        ' Restamos un miembro al equipo y si llega a cero entonces procesamos la derrota
        If Equipo = e_EquipoReto.Izquierda Then
            If .TamañoEquipoIzq > 1 Then
                .TamañoEquipoIzq = .TamañoEquipoIzq - 1
            Else
                .Puntaje = 123 ' Forzamos puntaje positivo
                Call FinalizarReto(Sala)
                Exit Sub
            End If
        Else
            If .TamañoEquipoDer > 1 Then
                .TamañoEquipoDer = .TamañoEquipoDer - 1
            Else
                .Puntaje = -123 ' Forzamos puntaje negativo
                Call FinalizarReto(Sala)
                Exit Sub
            End If
        End If
        Call RevivirYLimpiar(UserIndex)
        Call DevolverPosAnterior(UserIndex)
        Dim texto As String
        If Desconexion Then
            texto = UserList(UserIndex).name & " es descalificado por desconectarse."
        Else
            texto = UserList(UserIndex).name & " ha abandonado el reto."
        End If
        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i).ArrayIndex = UserIndex Then
                Call SetUserRef(.Jugadores(i), 0)
            Else
                Call WriteConsoleMsg(.Jugadores(i).ArrayIndex, texto, e_FontTypeNames.FONTTYPE_New_Gris)
            End If
        Next
    End With
End Sub

Private Sub SalaLiberada(ByVal Sala As Integer)
    On Error GoTo ErrHandler
    Retos.Salas(Sala).EnUso = False
    Retos.SalasLibres = Retos.SalasLibres + 1
    If ListaDeEspera.count > 0 Then
        Dim Oferente As Integer
        Oferente = ListaDeEspera.Keys(0)
        Call ListaDeEspera.Remove(Oferente)
        Call IniciarReto(Oferente, Sala)
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.SalaLiberada", Erl)
End Sub

Public Function PuedeReto(ByVal UserIndex As Integer) As Boolean
    With UserList(UserIndex)
        If .flags.EnReto Then Exit Function
        If .flags.EnConsulta Then Exit Function
        If .pos.Map = 0 Or .pos.x = 0 Or .pos.y = 0 Then Exit Function
        If MapInfo(.pos.Map).Seguro = 0 Then Exit Function
        If .flags.EnTorneo Then Exit Function
        If MapData(.pos.Map, .pos.x, .pos.y).trigger = CARCEL Then Exit Function
    End With
    PuedeReto = True
End Function

Public Function PuedeRetoConMensaje(ByVal UserIndex As Integer) As Boolean
    With UserList(UserIndex)
        If .flags.EnReto Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1974, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1974=Ya te encuentras en un reto.
            Exit Function
        End If
        If .flags.EnConsulta Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1975, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1975=No puedes acceder a un reto si estás en consulta.
            Exit Function
        End If
        If .flags.jugando_captura = 1 Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1976, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1976=No puedes jugar un reto estando en un evento.
            Exit Function
        End If
        If Not esCiudad(.pos.Map) Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1977, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1977=No puedes participar de un reto en un mapa inseguro.
            Exit Function
        End If
        If .flags.EnTorneo Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1978, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1978=No puedes ir a un reto si participas de un torneo.
            Exit Function
        End If
        If MapData(.pos.Map, .pos.x, .pos.y).trigger = CARCEL Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1979, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1979=¡Estás encarcelado!
            Exit Function
        End If
    End With
    PuedeRetoConMensaje = True
End Function

Private Function IndiceJugadorEnSolicitud(ByVal UserIndex As Integer, ByVal Oferente As Integer) As Integer
    With UserList(Oferente).flags.SolicitudReto
        IndiceJugadorEnSolicitud = -1
        If .Estado <> e_SolicitudRetoEstado.Enviada Then Exit Function
        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i).nombre = UserList(UserIndex).name Then
                IndiceJugadorEnSolicitud = i
                Exit Function
            End If
        Next
    End With
End Function

Private Sub MensajeATodosSolicitud(ByVal Oferente As Integer, mensaje As String, ByVal Fuente As e_FontTypeNames)
    With UserList(Oferente).flags.SolicitudReto
        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i).Aceptado Then
                Call WriteConsoleMsg(.Jugadores(i).CurIndex.ArrayIndex, mensaje, Fuente)
            End If
        Next
        Call WriteConsoleMsg(Oferente, mensaje, Fuente)
    End With
End Sub

Private Function TodosPuedenReto(ByVal Oferente As Integer) As Boolean
    On Error GoTo ErrHandler
    With UserList(Oferente).flags.SolicitudReto
        If Not PuedeReto(Oferente) Then
            Call CancelarSolicitudReto(Oferente, UserList(Oferente).name & " no puede entrar al reto en este momento.")
            Exit Function
        ElseIf UserList(Oferente).Stats.GLD < .Apuesta Then
            Call CancelarSolicitudReto(Oferente, UserList(Oferente).name & " no tiene las monedas de oro suficientes.")
            Exit Function
        ElseIf .PocionesMaximas >= 0 Then
            If TieneObjetos(38, .PocionesMaximas + 1, Oferente) Then
                Call CancelarSolicitudReto(Oferente, UserList(Oferente).name & " tiene demasiadas pociones rojas (Cantidad máxima: " & .PocionesMaximas & ").")
                Exit Function
            End If
        End If
        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            If Not PuedeReto(.Jugadores(i).CurIndex.ArrayIndex) Then
                Call CancelarSolicitudReto(Oferente, UserList(.Jugadores(i).CurIndex.ArrayIndex).name & " no puede entrar al reto en este momento.")
                Exit Function
            ElseIf UserList(.Jugadores(i).CurIndex.ArrayIndex).Stats.GLD < .Apuesta Then
                Call CancelarSolicitudReto(Oferente, UserList(.Jugadores(i).CurIndex.ArrayIndex).name & " no tiene las monedas de oro suficientes.")
                Exit Function
            ElseIf .PocionesMaximas >= 0 Then
                If TieneObjetos(38, .PocionesMaximas + 1, Oferente) Then
                    Call CancelarSolicitudReto(Oferente, UserList(.Jugadores(i).CurIndex.ArrayIndex).name & " tiene demasiadas pociones rojas (Cantidad máxima: " & _
                            .PocionesMaximas & ").")
                    Exit Function
                End If
            End If
        Next
        TodosPuedenReto = True
    End With
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.TodosPuedenReto", Erl)
End Function

Private Function EquipoContrario(ByVal Equipo As e_EquipoReto) As e_EquipoReto
    If Equipo = e_EquipoReto.Izquierda Then
        EquipoContrario = e_EquipoReto.Derecha
    Else
        EquipoContrario = e_EquipoReto.Izquierda
    End If
End Function

Private Function ObtenerTamañoEquipo(ByVal Sala As Integer, ByVal Equipo As e_EquipoReto) As Integer
    If Equipo = e_EquipoReto.Izquierda Then
        ObtenerTamañoEquipo = Retos.Salas(Sala).TamañoEquipoIzq
    Else
        ObtenerTamañoEquipo = Retos.Salas(Sala).TamañoEquipoDer
    End If
End Function

Private Sub RevivirYLimpiar(ByVal UserIndex As Integer)
    Call WriteStopped(UserIndex, False)
    ' Si está vivo
    If UserList(UserIndex).flags.Muerto = 0 Then
        Call LimpiarEstadosAlterados(UserIndex)
    End If
    ' Si está muerto lo revivimos, sino lo curamos
    Call RevivirUsuario(UserIndex)
End Sub
