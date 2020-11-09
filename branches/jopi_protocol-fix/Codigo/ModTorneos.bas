Attribute VB_Name = "ModTorneos"

Public Type tTorneo

    HayTorneoaActivo As Boolean
    NivelMinimo As Byte
    nivelmaximo As Byte
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
    ClasesTexto As String
    Participantes As Byte
    IndexParticipantes() As Integer
    Mapa As Integer
    x As Byte
    Y As Byte
    nombre As String
    reglas As String

End Type

Public Torneo        As tTorneo

Public MensajeTorneo As String

Public Sub IniciarTorneo()

    ReDim Torneo.IndexParticipantes(1 To Torneo.cupos)

    If Torneo.mago > 0 Then Torneo.ClasesTexto = "Mago,"
    If Torneo.clerico > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Clerigo,"
    If Torneo.guerrero > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Guerrero,"
    If Torneo.asesino > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Asesino,"
    If Torneo.bardo > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Bardo,"
    If Torneo.druido > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Druida,"
    If Torneo.Paladin > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Paladin,"
    If Torneo.cazador > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Cazador,"
    If Torneo.Trabajador > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Trabajador"

    Torneo.HayTorneoaActivo = True

    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Evento> Se abren las inscripciones para: " & Torneo.nombre & ": características: Nivel entre: " & Torneo.NivelMinimo & "/" & Torneo.nivelmaximo & ". Cupos disponibles: " & Torneo.cupos & " personajes. Precio de inscripción: " & Torneo.costo & " monedas de oro. Reglas: " & Torneo.reglas & ".", FontTypeNames.FONTTYPE_CITIZEN))
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Clases participantes: " & Torneo.ClasesTexto & ". Escribí /PARTICIPAR para ingresar al evento. ", FontTypeNames.FONTTYPE_CITIZEN))

End Sub

Public Sub ParticiparTorneo(ByVal UserIndex As Integer)

    Dim IndexVacio As Byte
    
    IndexVacio = BuscarIndexFreeTorneo
    Torneo.IndexParticipantes(IndexVacio) = UserIndex
    
    Torneo.Participantes = Torneo.Participantes + 1
    UserList(UserIndex).flags.EnTorneo = True
    
    Call WriteConsoleMsg(UserIndex, "¡Ya estas anotado! Solo debes aguardar hasta que seas enviado a la sala de espera.", FontTypeNames.FONTTYPE_INFOIAO)
    
End Sub

Public Function BuscarIndexFreeTorneo() As Byte

    Dim i As Byte

    For i = 1 To Torneo.cupos

        If Torneo.IndexParticipantes(i) = 0 Then
            BuscarIndexFreeTorneo = i
            Exit For

        End If

    Next i
    
End Function

Public Sub BorrarIndexInTorneo(ByVal Index As Integer)

    Dim i As Byte

    For i = 1 To Torneo.cupos

        If Torneo.IndexParticipantes(i) = Index Then
            Torneo.IndexParticipantes(i) = 0
            Exit For

        End If

    Next i

    Torneo.Participantes = Torneo.Participantes - 1
    
End Sub

Public Sub ComenzarTorneoOk()

    Dim nombres As String

    Dim x       As Byte

    Dim Y       As Byte

    For i = 1 To Torneo.Participantes
    
        nombres = nombres & UserList(Torneo.IndexParticipantes(i)).name & ", "
        x = Torneo.x
        Y = Torneo.Y
        Call FindLegalPos(Torneo.IndexParticipantes(i), Torneo.Mapa, x, Y)
        Call WarpUserChar(Torneo.IndexParticipantes(i), Torneo.Mapa, x, Y, True)
        ' Call WriteConsoleMsg(Torneo.IndexParticipantes(i), "¡Ya estas participado! Solo debes aguardar aquí hasta que seas convocado al torneo.", FontTypeNames.FONTTYPE_INFO)
    Next i

    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Evento> Los elegidos para participar son: " & nombres & " damos inicio al evento.", FontTypeNames.FONTTYPE_CITIZEN))

End Sub

Public Sub ResetearTorneo()

    Dim i As Byte

    Torneo.HayTorneoaActivo = False
    Torneo.NivelMinimo = 0
    Torneo.nivelmaximo = 0
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
    Torneo.ClasesTexto = ""
    Torneo.Mapa = 0
    Torneo.x = 0
    Torneo.Y = 0
    Torneo.Trabajador = 0
    Torneo.nombre = ""
    Torneo.reglas = 0
    
    For i = 1 To Torneo.Participantes
        UserList(Torneo.IndexParticipantes(i)).flags.EnTorneo = False
    Next i

    Torneo.Participantes = 0
    ReDim Torneo.IndexParticipantes(1 To 1)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Evento Finalizado. ", FontTypeNames.FONTTYPE_CITIZEN))

End Sub
