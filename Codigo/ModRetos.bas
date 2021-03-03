Attribute VB_Name = "ModRetos"
Option Explicit

Private Const APUESTA_MAXIMA = 100000000

Public Retos As tRetos
Private ListaDeEspera As New Dictionary

Public Sub CargarInfoRetos()
    Dim File As clsIniReader
    Set File = New clsIniReader

    Call File.Initialize(DatPath & "Retos.dat")
    
    With Retos

        .TamañoMaximoEquipo = val(File.GetValue("Retos", "MaximoEquipo"))
        .ApuestaMinima = val(File.GetValue("Retos", "ApuestaMinima"))
        .ImpuestoApuesta = val(File.GetValue("Retos", "ImpuestoApuesta"))
        .DuracionMaxima = val(File.GetValue("Retos", "DuracionMaxima"))
        .TiempoConteo = val(File.GetValue("Retos", "TiempoConteo"))
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
                .PosIzquierda.X = val(File.GetValue(SalaStr, "X"))
                .PosIzquierda.Y = val(File.GetValue(SalaStr, "Y"))
                
                .PosDerecha.Map = .PosIzquierda.Map
                .PosDerecha.X = .PosIzquierda.X + Retos.AnchoSala - 1
                .PosDerecha.Y = .PosIzquierda.Y + Retos.AltoSala - 1
            End With
        Next
        
    End With
    
    Set File = Nothing
End Sub

Public Sub CrearReto(ByVal UserIndex As Integer, JugadoresStr As String, ByVal Apuesta As Long)
    
    On Error GoTo ErrHandler
    
    With UserList(UserIndex)

        If .flags.SolicitudReto.estado <> SolicitudRetoEstado.Libre Then
            Call CancelarSolicitudReto(UserIndex, .name & " ha cancelado la solicitud.")

        ElseIf .flags.AceptoReto > 0 Then
            Call CancelarSolicitudReto(.flags.AceptoReto, .name & " ha cancelado su admisión.")
        End If

        If LenB(JugadoresStr) <= 0 Then Exit Sub
    
        Dim Jugadores() As String
        Jugadores = Split(JugadoresStr, ";", 5)
        
        If UBound(Jugadores) > Retos.TamañoMaximoEquipo - 2 Or UBound(Jugadores) Mod 2 = 1 Then Exit Sub
        
        Dim MaxIndexEquipo As Integer
        MaxIndexEquipo = UBound(Jugadores) \ 2
    
        If Apuesta < Retos.ApuestaMinima Or Apuesta > APUESTA_MAXIMA Then
            Call WriteConsoleMsg(UserIndex, "La apuesta mínima es de " & PonerPuntos(Retos.ApuestaMinima) & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Not PuedeRetoConMensaje(UserIndex) Then Exit Sub

        If .Stats.GLD < Apuesta Then
            Call WriteConsoleMsg(UserIndex, "No tienes el oro suficiente.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        With .flags.SolicitudReto
            .Apuesta = Apuesta
            
            ReDim .Jugadores(0 To UBound(Jugadores))
            
            Dim i As Integer, tIndex As Integer
            Dim Equipo1 As String, Equipo2 As String
            
            Equipo1 = UserList(UserIndex).name

            For i = 0 To UBound(.Jugadores)
                With .Jugadores(i)
                    If EsGmChar(Jugadores(i)) Then
                        Call WriteConsoleMsg(UserIndex, "¡No puedes jugar retos con administradores!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    tIndex = NameIndex(Jugadores(i))
                    
                    If tIndex <= 0 Then
                        Call WriteConsoleMsg(UserIndex, "El usuario " & Jugadores(i) & " no está conectado.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    If Not PuedeReto(tIndex) Then
                        Call WriteConsoleMsg(UserIndex, "El usuario " & UserList(tIndex).name & " no puede jugar un reto en este momento.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    .CurIndex = tIndex
                    .nombre = UserList(.CurIndex).name
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
            Texto1 = UserList(UserIndex).name & " te invita a jugar el siguiente reto:"
            Texto2 = Equipo1 & " vs. " & Equipo2 & ". Apuesta: " & PonerPuntos(Apuesta) & " monedas de oro."
            Texto3 = "Escribe /ACEPTAR " & UCase$(UserList(UserIndex).name) & " para participar en el reto."

            For i = 0 To UBound(.Jugadores)
                With .Jugadores(i)
                    Call WriteConsoleMsg(.CurIndex, Texto1, FontTypeNames.FONTTYPE_INFOBOLD)
                    Call WriteConsoleMsg(.CurIndex, Texto2, FontTypeNames.FONTTYPE_New_Naranja)
                    Call WriteConsoleMsg(.CurIndex, Texto3, FontTypeNames.FONTTYPE_INFOBOLD)
                End With
            Next

            .estado = SolicitudRetoEstado.Enviada
        End With

        Call WriteConsoleMsg(UserIndex, "Has enviado una solicitud para el siguiente reto:", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, Texto2, FontTypeNames.FONTTYPE_New_Naranja)
        Call WriteConsoleMsg(UserIndex, "Escribe /CANCELAR para anular la solicitud.", FontTypeNames.FONTTYPE_New_Gris)
    
    End With
    
    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "ModRetos.CrearReto", Erl)

End Sub

Public Sub AceptarReto(ByVal UserIndex As Integer, OferenteName As String)

    On Error GoTo ErrHandler

    If Not PuedeRetoConMensaje(UserIndex) Then Exit Sub
    
    With UserList(UserIndex)
        If .flags.SolicitudReto.estado <> SolicitudRetoEstado.Libre Then
            Call CancelarSolicitudReto(UserIndex, .name & " ha cancelado la solicitud.")
            
        ElseIf .flags.AceptoReto > 0 Then
            Call CancelarSolicitudReto(.flags.AceptoReto, .name & " ha cancelado su admisión.")
        End If
    End With
    
    If EsGmChar(OferenteName) Then
        Call WriteConsoleMsg(UserIndex, "¡No puedes jugar retos con administradores!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If

    Dim Oferente As Integer
    Oferente = NameIndex(OferenteName)
    
    If Oferente <= 0 Then
        Call WriteConsoleMsg(UserIndex, OferenteName & " no está conectado.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If

    With UserList(Oferente).flags.SolicitudReto

        Dim JugadorIndex As Integer
        JugadorIndex = IndiceJugadorEnSolicitud(UserIndex, Oferente)
        
        If JugadorIndex < 0 Then
            Call WriteConsoleMsg(UserIndex, UserList(Oferente).name & " no te ha invitado a ningún reto o ha sido cancelado.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).Stats.GLD < .Apuesta Then
            Call WriteConsoleMsg(UserIndex, "Necesitas al menos " & PonerPuntos(.Apuesta) & " monedas de oro para aceptar este reto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call MensajeATodosSolicitud(Oferente, UserList(UserIndex).name & " ha aceptado el reto.", FontTypeNames.FONTTYPE_INFO)
        
        .Jugadores(JugadorIndex).Aceptado = True
        .Jugadores(JugadorIndex).CurIndex = UserIndex
        UserList(UserIndex).flags.AceptoReto = Oferente
        
        Call WriteConsoleMsg(UserIndex, "Has aceptado el reto de " & UserList(Oferente).name & ".", FontTypeNames.FONTTYPE_INFO)
        
        Dim FaltanAceptar As String

        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            If Not .Jugadores(i).Aceptado Then
                FaltanAceptar = FaltanAceptar & " - "
            End If
        Next
        
        If LenB(FaltanAceptar) > 0 Then
            FaltanAceptar = Left$(FaltanAceptar, Len(FaltanAceptar) - 3)
            Call MensajeATodosSolicitud(Oferente, "Faltan aceptar: " & FaltanAceptar, FontTypeNames.FONTTYPE_New_Gris)
            Exit Sub
        End If
        
        Call MensajeATodosSolicitud(Oferente, "Todos los jugadores han aceptado el reto. Buscando sala...", FontTypeNames.FONTTYPE_New_Gris)

        Call BuscarSala(Oferente)

    End With
    
    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "ModRetos.AceptarReto", Erl)
End Sub

Public Sub CancelarSolicitudReto(ByVal Oferente As Integer, Mensaje As String)
    
    On Error GoTo ErrHandler
    
    With UserList(Oferente).flags.SolicitudReto
    
        If .estado = SolicitudRetoEstado.EnCola Then
            Call ListaDeEspera.Remove(Oferente)
        End If

        .estado = SolicitudRetoEstado.Libre
        
        Dim i As Integer, tIndex As Integer

        For i = 0 To UBound(.Jugadores)

            tIndex = NameIndex(.Jugadores(i).nombre)
            
            If tIndex > 0 Then
                Call WriteConsoleMsg(tIndex, Mensaje, FontTypeNames.FONTTYPE_WARNING)
                Call WriteConsoleMsg(tIndex, "El reto ha sido cancelado.", FontTypeNames.FONTTYPE_WARNING)

                If .Jugadores(i).Aceptado Then
                    UserList(tIndex).flags.AceptoReto = 0
                End If
            End If

        Next
    
    End With
    
    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "ModRetos.CancelarSolicitudReto", Erl)
    
End Sub

Private Sub BuscarSala(ByVal Oferente As Integer)

    On Error GoTo ErrHandler
    
    With UserList(Oferente).flags.SolicitudReto

        If Retos.SalasLibres <= 0 Then
            Call ListaDeEspera.Add(Oferente, 0)
            Call MensajeATodosSolicitud(Oferente, "No hay salas disponibles. El reto comenzará cuando se desocupe una sala.", FontTypeNames.FONTTYPE_FIGHT)
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
    Call RegistrarError(Err.Number, Err.Description, "ModRetos.BuscarSala", Erl)
End Sub

Private Sub IniciarReto(ByVal Oferente As Integer, ByVal Sala As Integer)

    On Error GoTo ErrHandler
    
    With UserList(Oferente).flags.SolicitudReto
    
        ' Última comprobación de si todos pueden entrar/pagar
        If Not TodosPuedenReto(Oferente) Then Exit Sub
        
        Dim Apuesta As Integer
        Apuesta = .Apuesta

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
            Retos.Salas(Sala).Jugadores(CurIndex) = Oferente
            CurIndex = CurIndex + 1
        End If
        
        Dim i As Integer
        
        ' Agregamos los jugadores alternando 1 y 1 (en los índices pares está el equipo izquierdo y en los impares el derecho - el array empieza en cero)
        For i = 0 To UBound(.Jugadores)
            Retos.Salas(Sala).Jugadores(CurIndex) = .Jugadores(i).CurIndex
            CurIndex = CurIndex + 1
            ' Reset flag
            UserList(.Jugadores(i).CurIndex).flags.AceptoReto = 0
        Next
        
        If Moneda = 1 Then
            ' Agregamos al oferente al final (su equipo juega a la derecha)
            Retos.Salas(Sala).Jugadores(CurIndex) = Oferente
        End If
        
        ' Reset estado de la solicitud, ya que no la necesitamos más
        .estado = SolicitudRetoEstado.Libre
        
    End With

    With Retos.Salas(Sala)
        .EnUso = True
        .Puntaje = 0
        .Ronda = 1
        .Apuesta = Apuesta
        .TiempoRestante = Retos.DuracionMaxima

        Dim tIndex As Integer

        For i = 0 To UBound(.Jugadores)

            tIndex = .Jugadores(i)

            ' Le cobramos
            UserList(tIndex).Stats.GLD = UserList(tIndex).Stats.GLD - Apuesta
            Call WriteUpdateGold(tIndex)
            
            ' Desmontamos
            If UserList(tIndex).flags.Montado <> 0 Then
                Call DoMontar(tIndex, ObjData(UserList(tIndex).Invent.MonturaObjIndex), UserList(tIndex).Invent.MonturaSlot)
            End If
            
            ' Dejamos de navegar
            If UserList(tIndex).flags.Nadando <> 0 Or UserList(tIndex).flags.Navegando <> 0 Then
                Call DoNavega(tIndex, ObjData(UserList(tIndex).Invent.BarcoObjIndex), UserList(tIndex).Invent.BarcoSlot)
            End If
            
            ' Asignamos flags
            With UserList(tIndex).flags
                .EnReto = True
                .EquipoReto = IIf(i Mod 2, EquipoReto.Izquierda, EquipoReto.Derecha)
                .SalaReto = Sala
                ' Guardar posición
                .LastPos = UserList(tIndex).Pos
            End With
            
            Call WriteConsoleMsg(tIndex, "¡Ha comenzado el reto!", FontTypeNames.FONTTYPE_New_Rojo_Salmon)
            Call WriteConsoleMsg(tIndex, "Para admitir la derrota escribe /ABANDONAR.", FontTypeNames.FONTTYPE_New_Gris)

        Next

    End With

    Call IniciarRonda(Sala)

    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "ModRetos.IniciarReto", Erl)
    
End Sub

Private Sub IniciarRonda(ByVal Sala As Integer)

    With Retos.Salas(Sala)
    
        Dim i As Integer, tIndex As Integer
        
        For i = 0 To UBound(.Jugadores)

            tIndex = .Jugadores(i)

            If tIndex <> 0 Then

                Call RevivirYLimpiar(tIndex)

                ' Usando el número de ronda y el índice, decidimos el lado al que corresponde
                If (.Ronda + i) Mod 2 = 1 Then
                    ' Lado izquierdo
                    Call WarpToLegalPos(tIndex, .PosIzquierda.Map, .PosIzquierda.X, .PosIzquierda.Y, True)
                Else
                    ' Lado derecho
                    Call WarpToLegalPos(tIndex, .PosDerecha.Map, .PosDerecha.X, .PosDerecha.Y, True)
                End If

                ' Si usamos el conteo
                If Retos.TiempoConteo > 0 Then
                    ' Le ponemos el conteo
                    UserList(tIndex).Counters.CuentaRegresiva = Retos.TiempoConteo
                    ' Lo stoppeamos
                    Call WriteStopped(tIndex, True)
                End If
                
                Call WriteConsoleMsg(tIndex, "Comienza la ronda Nº" & .Ronda, FontTypeNames.FONTTYPE_GUILD)

            End If
        Next
    
    End With
    
End Sub

Public Sub MuereEnReto(ByVal UserIndex As Integer)

    Dim Sala As Integer, Equipo As EquipoReto

    With UserList(UserIndex)
        Sala = .flags.SalaReto
        Equipo = .flags.EquipoReto
    End With
    
    With Retos.Salas(Sala)
    
        Dim CurIndex As Integer
        
        ' El equipo derecho está en índices pares
        If Equipo = EquipoReto.Derecha Then CurIndex = 1
        
        For CurIndex = CurIndex To UBound(.Jugadores) Step 2
            If .Jugadores(CurIndex) <> 0 Then
                ' Si todavía hay alguno vivo del equipo
                If UserList(.Jugadores(CurIndex)).flags.Muerto = 0 Then
                    Exit Sub
                End If
            End If
        Next
        
        ' Están todos muertos, ganó el equipo contrario
        Call ProcesarRondaGanada(Sala, EquipoContrario(Equipo))
    
    End With

End Sub

Private Sub ProcesarRondaGanada(ByVal Sala As Integer, ByVal Equipo As EquipoReto)

    With Retos.Salas(Sala)

        ' Sumamos puntaje o restamos según el equipo
        If Equipo = EquipoReto.Derecha Then
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
        For i = IIf(Equipo = EquipoReto.Izquierda, 0, 1) To TamañoEquipo2 - 1 Step 2

            If .Jugadores(i) <> 0 Then
                nombres = nombres & UserList(.Jugadores(i)).name
                
                If i < TamañoEquipo2 - 2 Then
                    nombres = nombres & IIf(i > TamañoEquipo2 - 5, " y ", ", ")
                End If
            End If
        Next
        
        ' Informamos el ganador de esta ronda
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i) <> 0 Then
                Call WriteConsoleMsg(.Jugadores(i), "Esta ronda es para " & nombres & ".", FontTypeNames.FONTTYPE_GUILD)
                Call WriteConsoleMsg(.Jugadores(i), "", 0) ' Dejamos un espacio vertical
            End If
        Next
        
        ' Iniciamos la próxima ronda
        Call IniciarRonda(Sala)
    
    End With

End Sub

Public Sub FinalizarReto(ByVal Sala As Integer, Optional ByVal TiempoAgotado As Boolean)
    
    With Retos.Salas(Sala)
    
        ' Calculamos el oro total del premio
        Dim OroTotal As Long, Oro As Long, OroStr As String
        OroTotal = .Apuesta * (UBound(.Jugadores) + 1)
        
        ' Descontamos el impuesto
        OroTotal = Fix(OroTotal * (1 - Retos.ImpuestoApuesta))
    
        ' Decidimos el resultado del reto según el puntaje:
        Dim i As Integer, tIndex As Integer, Equipo1 As String, Equipo2 As String
        
        ' Empate
        If .Puntaje = 0 Then
            ' Pagamos a todos los que no abandonaron
            Oro = OroTotal \ (UBound(.Jugadores) + 1)
            OroStr = PonerPuntos(Oro)

            For i = 0 To UBound(.Jugadores)
                tIndex = .Jugadores(i)

                If tIndex <> 0 Then
                    UserList(tIndex).Stats.GLD = UserList(tIndex).Stats.GLD + Oro
                    Call WriteUpdateGold(tIndex)
                    Call WriteLocaleMsg(tIndex, "29", FontTypeNames.FONTTYPE_INFO, OroStr) ' Has ganado X monedas de oro
                    
                    Call RevivirYLimpiar(tIndex)

                    Call DevolverPosAnterior(tIndex)
                    
                    ' Reset flags
                    UserList(tIndex).flags.EnReto = False
                    
                    ' Nombres
                    If i Mod 2 Then
                        Equipo1 = Equipo1 & IIf((i + 1) \ 2 < .TamañoEquipoDer - 1, ", ", " y ") & UserList(tIndex).name
                    Else
                        If LenB(Equipo2) > 0 Then
                            Equipo2 = Equipo2 & IIf(i \ 2 < .TamañoEquipoIzq - 1, ", ", " y ") & UserList(tIndex).name
                        Else
                            Equipo2 = UserList(tIndex).name
                        End If
                    End If
                End If
            Next
            
            ' Anuncio global
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Equipo1 & " vs. " & Equipo2 & ". Apuesta: " & PonerPuntos(.Apuesta) & _
                                                " monedas de oro. Resultado: empate por tiempo.", FontTypeNames.FONTTYPE_INFO))

        ' Hubo un ganador
        Else
            Dim Ganador As EquipoReto
            
            If .Puntaje < 0 Then
                Ganador = EquipoReto.Izquierda
            Else
                Ganador = EquipoReto.Derecha
            End If

            ' Pagamos a los ganadores que no abandonaron
            Oro = OroTotal \ ObtenerTamañoEquipo(Sala, Ganador)
            OroStr = PonerPuntos(Oro)

            For i = 0 To UBound(.Jugadores)
                tIndex = .Jugadores(i)

                If tIndex <> 0 Then
                    If UserList(tIndex).flags.EquipoReto = Ganador Then
                        UserList(tIndex).Stats.GLD = UserList(tIndex).Stats.GLD + Oro
                        Call WriteUpdateGold(tIndex)
                        Call WriteLocaleMsg(tIndex, "29", FontTypeNames.FONTTYPE_INFO, OroStr) ' Has ganado X monedas de oro
                    End If
                    
                    Call RevivirYLimpiar(tIndex)
    
                    Call DevolverPosAnterior(tIndex)
                    
                    ' Reset flags
                    UserList(tIndex).flags.EnReto = False
                    
                    If TiempoAgotado Then
                        Call WriteConsoleMsg(tIndex, "Se ha agotado el tiempo del reto.", FontTypeNames.FONTTYPE_New_Gris)
                    End If

                    ' Nombres
                    If i Mod 2 Then
                        Equipo1 = Equipo1 & IIf((i + 1) \ 2 < .TamañoEquipoDer - 1, ", ", " y ") & UserList(tIndex).name
                    Else
                        If LenB(Equipo2) > 0 Then
                            Equipo2 = Equipo2 & IIf(i \ 2 < .TamañoEquipoIzq - 1, ", ", " y ") & UserList(tIndex).name
                        Else
                            Equipo2 = UserList(tIndex).name
                        End If
                    End If
                End If
            Next

            ' Anuncio global
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Equipo1 & " vs. " & Equipo2 & ". Apuesta: " & PonerPuntos(.Apuesta) & " monedas de oro. " & _
                                IIf(UBound(.Jugadores) > 1, "Ganadores: ", "Ganador: ") & IIf(Ganador = EquipoReto.Izquierda, Equipo1, Equipo2) & ".", FontTypeNames.FONTTYPE_INFO))
        End If
    
    End With
    
    Call SalaLiberada(Sala)
    
End Sub

Public Sub AbandonarReto(ByVal UserIndex As Integer, Optional ByVal Desconexion As Boolean)
    
    Dim Sala As Integer, Equipo As EquipoReto

    With UserList(UserIndex)
        Sala = .flags.SalaReto
        Equipo = .flags.EquipoReto

        .flags.EnReto = False
    End With
    
    With Retos.Salas(Sala)

        If Not Desconexion Then
            Call WriteConsoleMsg(UserIndex, "Has abandonado el reto.", FontTypeNames.FONTTYPE_INFO)
        End If
        
        ' Restamos un miembro al equipo y si llega a cero entonces procesamos la derrota
        If Equipo = EquipoReto.Izquierda Then
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
            If .Jugadores(i) = UserIndex Then
                .Jugadores(i) = 0
            Else
                Call WriteConsoleMsg(.Jugadores(i), texto, FontTypeNames.FONTTYPE_New_Gris)
            End If
        Next
    
    End With
    
End Sub

Private Sub SalaLiberada(ByVal Sala As Integer)

    On Error GoTo ErrHandler
    
    Retos.Salas(Sala).EnUso = False
    
    If ListaDeEspera.Count > 0 Then
    
        Dim Oferente As Integer
        Oferente = ListaDeEspera.Keys(0)
        Call ListaDeEspera.Remove(Oferente)
        
        Call IniciarReto(Oferente, Sala)

    End If
    
    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "ModRetos.SalaLiberada", Erl)
    
End Sub

Public Function PuedeReto(ByVal UserIndex As Integer) As Boolean
    
    With UserList(UserIndex)
        
        If .flags.EnReto Then Exit Function
        
        If .flags.EnConsulta Then Exit Function
        
        If MapInfo(.Pos.Map).Seguro = 0 Then Exit Function
        
        If .flags.EnTorneo Then Exit Function
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = CARCEL Then Exit Function
        
    End With
    
    PuedeReto = True
    
End Function

Public Function PuedeRetoConMensaje(ByVal UserIndex As Integer) As Boolean

    With UserList(UserIndex)
        
        If .flags.EnReto Then
            Call WriteConsoleMsg(UserIndex, "Ya te encuentras en un reto.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        If .flags.EnConsulta Then
            Call WriteConsoleMsg(UserIndex, "No puedes acceder a un reto si estás en consulta.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        If MapInfo(.Pos.Map).Seguro = 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes  estar en un mapa seguro.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        If .flags.EnTorneo Then
            Call WriteConsoleMsg(UserIndex, "No puedes ir a un reto si participas de un torneo.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = CARCEL Then
            Call WriteConsoleMsg(UserIndex, "¡Estás encarcelado!", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
    End With

    PuedeRetoConMensaje = True

End Function

Private Function IndiceJugadorEnSolicitud(ByVal UserIndex As Integer, ByVal Oferente As Integer) As Integer

    With UserList(Oferente).flags.SolicitudReto

        If Not .estado <> SolicitudRetoEstado.Enviada Then Exit Function

        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i).nombre = UserList(UserIndex).name Then
                IndiceJugadorEnSolicitud = i
                Exit Function
            End If
        Next
    
    End With

End Function

Private Sub MensajeATodosSolicitud(ByVal Oferente As Integer, Mensaje As String, ByVal Fuente As FontTypeNames)
    
    With UserList(Oferente).flags.SolicitudReto

        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i).Aceptado Then
                Call WriteConsoleMsg(.Jugadores(i).CurIndex, Mensaje, Fuente)
            End If
        Next
        
        Call WriteConsoleMsg(Oferente, Mensaje, Fuente)

    End With
    
End Sub

Private Function TodosPuedenReto(ByVal Oferente As Integer) As Boolean

    On Error GoTo ErrHandler
    
    With UserList(Oferente).flags.SolicitudReto

        Dim i As Integer

        For i = 0 To UBound(.Jugadores)
            If Not PuedeReto(.Jugadores(i).CurIndex) Then
                Call CancelarSolicitudReto(Oferente, UserList(.Jugadores(i).CurIndex).name & " no puede entrar al reto en este momento.")
                Exit Function

            ElseIf UserList(.Jugadores(i).CurIndex).Stats.GLD < .Apuesta Then
                Call CancelarSolicitudReto(Oferente, UserList(.Jugadores(i).CurIndex).name & " no tiene las monedas de oro suficientes.")
                Exit Function
            End If
        Next
        
        TodosPuedenReto = True
    
    End With
    
    Exit Function
    
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "ModRetos.TodosPuedenReto", Erl)
    
End Function

Private Function EquipoContrario(ByVal Equipo As EquipoReto) As EquipoReto
    If Equipo = EquipoReto.Izquierda Then
        EquipoContrario = EquipoReto.Derecha
    Else
        EquipoContrario = EquipoReto.Izquierda
    End If
End Function

Private Function ObtenerTamañoEquipo(ByVal Sala As Integer, ByVal Equipo As EquipoReto) As Integer
    If Equipo = EquipoReto.Izquierda Then
        ObtenerTamañoEquipo = Retos.Salas(Sala).TamañoEquipoIzq
    Else
        ObtenerTamañoEquipo = Retos.Salas(Sala).TamañoEquipoDer
    End If
End Function

Private Sub RevivirYLimpiar(ByVal UserIndex As Integer)

    ' Si está vivo
    If UserList(UserIndex).flags.Muerto = 0 Then
        Call LimpiarEstadosAlterados(UserIndex)
    End If

    ' Si está muerto lo revivimos, sino lo curamos
    Call RevivirUsuario(UserIndex)

End Sub
