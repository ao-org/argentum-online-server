Attribute VB_Name = "ModEventos"

Public HoraEvento           As Byte

Public TiempoRestanteEvento As Byte

Public EventoActivo         As Boolean

Public EventoAcutal         As EventoPropiedades

Public Evento(0 To 23)      As EventoPropiedades

Public Type EventoPropiedades

    Tipo As Byte
    duracion As Byte
    multiplicacion As Byte

End Type

Public ExpMultOld         As Integer

Public OroMultOld         As Integer

Public DropMultOld        As Integer

Public RecoleccionMultOld As Integer

Public PublicidadEvento   As String

Public Sub CheckEvento(ByVal Hora As Byte)

    If EventoActivo = True Then Exit Sub

    Dim aviso As String

    aviso = "Eventos> Nuevo evento iniciado: "
    PublicidadEvento = "Evento en curso>"

    Select Case Evento(Hora).Tipo

        Case 1
            OroMult = OroMult * Evento(Hora).multiplicacion
            aviso = aviso & " Oro multiplicado por " & Evento(Hora).multiplicacion & " - Duración del evento: " & Evento(Hora).duracion & " minutos."
            PublicidadEvento = PublicidadEvento & " Oro multiplicado por " & Evento(Hora).multiplicacion

        Case 2
            ExpMult = ExpMult * Evento(Hora).multiplicacion
            aviso = aviso & " Experiencia multiplicada por " & Evento(Hora).multiplicacion & " - Duración del evento: " & Evento(Hora).duracion & " minutos."
            PublicidadEvento = PublicidadEvento & " Experiencia multiplicada por " & Evento(Hora).multiplicacion

        Case 3
            RecoleccionMult = RecoleccionMult * Evento(Hora).multiplicacion
            aviso = aviso & " Recoleccion multiplicada por " & Evento(Hora).multiplicacion & " - Duración del evento: " & Evento(Hora).duracion & " minutos."
            PublicidadEvento = PublicidadEvento & " Recoleccion multiplicada por " & Evento(Hora).multiplicacion

        Case 4
            DropMult = DropMult / Evento(Hora).multiplicacion
            aviso = aviso & " Dropeo multiplicado por " & Evento(Hora).multiplicacion & " - Duración del evento: " & Evento(Hora).duracion & " minutos."
            PublicidadEvento = PublicidadEvento & " Dropeo multiplicado por " & Evento(Hora).multiplicacion

        Case 5
            ExpMult = ExpMult * Evento(Hora).multiplicacion
            OroMult = OroMult * Evento(Hora).multiplicacion
            aviso = aviso & " Oro y experiencia multiplicados por " & Evento(Hora).multiplicacion & " - Duración del evento: " & Evento(Hora).duracion & " minutos."
            PublicidadEvento = PublicidadEvento & " Oro y experiencia multiplicados por " & Evento(Hora).multiplicacion

        Case 6
            ExpMult = ExpMult * Evento(Hora).multiplicacion
            OroMult = OroMult * Evento(Hora).multiplicacion
            RecoleccionMult = RecoleccionMult * Evento(Hora).multiplicacion
            aviso = aviso & " Oro, experiencia y recoleccion multiplicados por " & Evento(Hora).multiplicacion & " - Duración del evento: " & Evento(Hora).duracion & " minutos."
            PublicidadEvento = PublicidadEvento & " Oro, experiencia y recoleccion multiplicados por " & Evento(Hora).multiplicacion

        Case 7
            ExpMult = ExpMult * Evento(Hora).multiplicacion
            OroMult = OroMult * Evento(Hora).multiplicacion
            DropMult = DropMult / Evento(Hora).multiplicacion
            RecoleccionMult = RecoleccionMult * Evento(Hora).multiplicacion
            aviso = aviso & " Oro, experiencia, recoleccion y dropeo multiplicados por " & Evento(Hora).multiplicacion & " - Duración del evento: " & Evento(Hora).duracion & " minutos."
            PublicidadEvento = PublicidadEvento & " Oro, experiencia, recoleccion y dropeo multiplicados por " & Evento(Hora).multiplicacion

        Case Else

            EventoActivo = False
            Exit Sub
        
    End Select

    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(553, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno

    Call AgregarAConsola(aviso)

    EventoAcutal.duracion = Evento(Hora).duracion
    EventoAcutal.multiplicacion = Evento(Hora).multiplicacion
    EventoAcutal.Tipo = Evento(Hora).Tipo

    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(aviso, FontTypeNames.FONTTYPE_New_Eventos))
    TiempoRestanteEvento = Evento(Hora).duracion
    frmMain.Evento.Enabled = True
    EventoActivo = True
        
End Sub

Public Sub FinalizarEvento()
    frmMain.Evento.Enabled = False
    EventoActivo = False

    Select Case EventoAcutal.Tipo

        Case 1
            OroMult = OroMultOld
       
        Case 2
            ExpMult = ExpMultOld
       
        Case 3
            RecoleccionMult = RecoleccionMultOld
  
        Case 4
            DropMult = DropMultOld
        
        Case 5
            ExpMult = ExpMultOld
            OroMult = OroMultOld

        Case 6
            ExpMult = ExpMultOld
            OroMult = OroMultOld
            RecoleccionMult = RecoleccionMultOld

        Case 7
            ExpMult = ExpMultOld
            OroMult = OroMultOld
            DropMult = DropMultOld
            RecoleccionMult = RecoleccionMultOld

        Case Else
            Exit Sub
        
    End Select

    Call AgregarAConsola("Eventos > Evento finalizado.")
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos > Evento finalizado.", FontTypeNames.FONTTYPE_New_Eventos))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(551, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno

End Sub

Public Function DescribirEvento(ByVal Hora As Byte) As String

    Dim aviso As String

    aviso = "("

    Select Case Evento(Hora).Tipo

        Case 1

            aviso = aviso & "Oro multiplicado por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).duracion & " minutos"

        Case 2
        
            aviso = aviso & "Experiencia multiplicada por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).duracion & " minutos"

        Case 3
            aviso = aviso & "Recoleccion multiplicada por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).duracion & " minutos"

        Case 4
            aviso = aviso & "Dropeo multiplicado por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).duracion & " minutos"
       
        Case 5
            aviso = aviso & "Oro y experiencia multiplicados por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).duracion & " minutos"

        Case 6

            aviso = aviso & "Oro, experiencia y recoleccion multiplicados por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).duracion & " minutos"

        Case 7
            aviso = aviso & "Oro, experiencia, recoleccion y dropeo multiplicados por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).duracion & " minutos"

        Case Else
            aviso = aviso & "sin información"
        
    End Select

    aviso = aviso & ")"

    DescribirEvento = aviso

End Function

Public Sub CargarEventos()

    Dim i          As Byte

    Dim EventoStrg As String

    For i = 0 To 23
        EventoStrg = GetVar(IniPath & "Configuracion.ini", "EVENTOS", i)
        Evento(i).Tipo = val(ReadField(1, EventoStrg, Asc("-")))
        Evento(i).duracion = val(ReadField(2, EventoStrg, Asc("-")))
        Evento(i).multiplicacion = val(ReadField(3, EventoStrg, Asc("-")))
    Next i

    ExpMultOld = ExpMult
    OroMultOld = OroMult
    DropMultOld = DropMult
    RecoleccionMultOld = RecoleccionMult

End Sub

Public Sub ForzarEvento(ByVal Tipo As Byte, ByVal duracion As Byte, ByVal multi As Byte, ByVal Quien As String)

    If Tipo > 7 Or Tipo < 1 Then
        Call WriteConsoleMsg(NameIndex(Quien), "Tipo de evento invalido.", FontTypeNames.FONTTYPE_New_Eventos)
        Exit Sub

    End If
 
    If duracion > 59 Then
        Call WriteConsoleMsg(NameIndex(Quien), "Duracion invalida, maxima 59 minutos.", FontTypeNames.FONTTYPE_New_Eventos)
        Exit Sub

    End If

    If multi > 10 Then
        Call WriteConsoleMsg(NameIndex(Quien), "Multiplicacion invalida, maxima x10.", FontTypeNames.FONTTYPE_New_Eventos)
        Exit Sub

    End If

    Dim aviso As String

    aviso = "Eventos> " & Quien & " inicio un nuevo evento: "
    PublicidadEvento = "Evento en curso>"

    Select Case Tipo

        Case 1
            OroMult = OroMult * multi
            aviso = aviso & " Oro multiplicado por " & multi & " - Duración del evento: " & duracion & " minutos."
            PublicidadEvento = PublicidadEvento & " Oro multiplicado por " & multi

        Case 2
            ExpMult = ExpMult * multi
            aviso = aviso & " Experiencia multiplicada por " & multi & " - Duración del evento: " & duracion & " minutos."
            PublicidadEvento = PublicidadEvento & " Experiencia multiplicada por " & multi

        Case 3
            RecoleccionMult = RecoleccionMult * multi
            aviso = aviso & " Recoleccion multiplicada por " & multi & " - Duración del evento: " & duracion & " minutos."
            PublicidadEvento = PublicidadEvento & " Recoleccion multiplicada por " & multi

        Case 4
            DropMult = DropMult / multi
            aviso = aviso & " Dropeo multiplicado por " & multi & " - Duración del evento: " & duracion & " minutos."
            PublicidadEvento = PublicidadEvento & " Dropeo multiplicado por " & multi

        Case 5
            ExpMult = ExpMult * multi
            OroMult = OroMult * multi
            aviso = aviso & " Oro y experiencia multiplicados por " & multi & " - Duración del evento: " & duracion & " minutos."
            PublicidadEvento = PublicidadEvento & " Oro y experiencia multiplicados por " & multi

        Case 6
            ExpMult = ExpMult * multi
            OroMult = OroMult * multi
            RecoleccionMult = RecoleccionMult * multi
            aviso = aviso & " Oro, experiencia y recoleccion multiplicados por " & multi & " - Duración del evento: " & duracion & " minutos."
            PublicidadEvento = PublicidadEvento & " Oro, experiencia y recoleccion multiplicados por " & multi

        Case 7
            ExpMult = ExpMult * multi
            OroMult = OroMult * multi
            DropMult = DropMult / multi
            RecoleccionMult = RecoleccionMult * multi
            aviso = aviso & " Oro, experiencia, recoleccion y dropeo multiplicados por " & multi & " - Duración del evento: " & duracion & " minutos."
            PublicidadEvento = PublicidadEvento & " Oro, experiencia, recoleccion y dropeo multiplicados por " & multi

        Case Else

            EventoActivo = False
            Exit Sub
        
    End Select

    Call AgregarAConsola(aviso)

    EventoAcutal.duracion = duracion
    EventoAcutal.multiplicacion = multi
    EventoAcutal.Tipo = Tipo

    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(aviso, FontTypeNames.FONTTYPE_New_Eventos))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(553, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno
    TiempoRestanteEvento = duracion
    frmMain.Evento.Enabled = True
    EventoActivo = True

End Sub

Public Sub EventoDeBoss()

End Sub

