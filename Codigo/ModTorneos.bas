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
    Pirata As Byte
    Ladron As Byte
    Bandido As Byte
    ClasesTexto As String
    Participantes As Byte
    IndexParticipantes() As Integer
    Mapa As Integer
    X As Byte
    Y As Byte
    nombre As String
    reglas As String

End Type

Public Torneo        As tTorneo

Public MensajeTorneo As String

Public Sub IniciarTorneo()
        On Error GoTo IniciarTorneo_Err
        
        Dim i As Long
        Dim inscriptos As Byte
        inscriptos = 0
        
102     If Torneo.mago > 0 Then Torneo.ClasesTexto = "Mago,"
104     If Torneo.clerico > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Clerigo,"
106     If Torneo.guerrero > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Guerrero,"
108     If Torneo.asesino > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Asesino,"
110     If Torneo.bardo > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Bardo,"
112     If Torneo.druido > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Druida,"
114     If Torneo.Paladin > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Paladin,"
116     If Torneo.cazador > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Cazador,"
118     If Torneo.Trabajador > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Trabajador,"
        If Torneo.Pirata > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Pirata,"
        If Torneo.Ladron > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Ladron,"
        If Torneo.Bandido > 0 Then Torneo.ClasesTexto = Torneo.ClasesTexto & "Bandido"

        If Not Torneo.HayTorneoaActivo Then
119         ReDim Torneo.IndexParticipantes(1 To Torneo.cupos)
120         Torneo.HayTorneoaActivo = True
        Else
            For i = 1 To Torneo.cupos
                If Torneo.IndexParticipantes(i) > 0 Then
                    inscriptos = inscriptos + 1
                End If
            Next i
        End If

122     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Evento> Están abiretas las inscripciones para: " & Torneo.nombre & ": características: Nivel entre: " & Torneo.NivelMinimo & "/" & Torneo.nivelmaximo & ". Inscriptos: " & inscriptos & "/" & Torneo.cupos & ". Precio de inscripción: " & PonerPuntos(Torneo.costo) & " monedas de oro. Reglas: " & Torneo.reglas & ".", FontTypeNames.FONTTYPE_CITIZEN))
124     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Evento> Clases participantes: " & Torneo.ClasesTexto & ". Escribí /PARTICIPAR para ingresar al evento. ", FontTypeNames.FONTTYPE_CITIZEN))

        
        Exit Sub

IniciarTorneo_Err:
126     Call RegistrarError(Err.Number, Err.Description, "ModTorneos.IniciarTorneo", Erl)
128
        
End Sub

Public Sub ParticiparTorneo(ByVal UserIndex As Integer)
        
        On Error GoTo ParticiparTorneo_Err
        

        Dim IndexVacio As Byte
    
100     IndexVacio = BuscarIndexFreeTorneo
102     Torneo.IndexParticipantes(IndexVacio) = UserIndex
    
104     Torneo.Participantes = Torneo.Participantes + 1
106     UserList(UserIndex).flags.EnTorneo = True
    
108     Call WriteConsoleMsg(UserIndex, "¡Ya estas anotado! Solo debes aguardar hasta que seas enviado a la sala de espera.", FontTypeNames.FONTTYPE_INFOIAO)
    
        
        Exit Sub

ParticiparTorneo_Err:
110     Call RegistrarError(Err.Number, Err.Description, "ModTorneos.ParticiparTorneo", Erl)
112
        
End Sub

Public Function BuscarIndexFreeTorneo() As Byte
        
        On Error GoTo BuscarIndexFreeTorneo_Err
        

        Dim i As Byte

100     For i = 1 To Torneo.cupos

102         If Torneo.IndexParticipantes(i) = 0 Then
104             BuscarIndexFreeTorneo = i
                Exit For

            End If

106     Next i
    
        
        Exit Function

BuscarIndexFreeTorneo_Err:
108     Call RegistrarError(Err.Number, Err.Description, "ModTorneos.BuscarIndexFreeTorneo", Erl)
110
        
End Function

Public Sub BorrarIndexInTorneo(ByVal Index As Integer)
        
        On Error GoTo BorrarIndexInTorneo_Err
        

        Dim i As Byte

100     For i = 1 To Torneo.cupos

102         If Torneo.IndexParticipantes(i) = Index Then
104             Torneo.IndexParticipantes(i) = 0
                Exit For

            End If

106     Next i

108     Torneo.Participantes = Torneo.Participantes - 1
    
        
        Exit Sub

BorrarIndexInTorneo_Err:
110     Call RegistrarError(Err.Number, Err.Description, "ModTorneos.BorrarIndexInTorneo", Erl)
112
        
End Sub

Public Sub ComenzarTorneoOk()
        
        On Error GoTo ComenzarTorneoOk_Err
        

        Dim nombres As String

        Dim X       As Byte

        Dim Y       As Byte

100     For i = 1 To Torneo.Participantes
    
102         nombres = nombres & UserList(Torneo.IndexParticipantes(i)).Name & ", "
104         X = Torneo.X
106         Y = Torneo.Y
108         Call FindLegalPos(Torneo.IndexParticipantes(i), Torneo.Mapa, X, Y)
110         Call WarpUserChar(Torneo.IndexParticipantes(i), Torneo.Mapa, X, Y, True)
            ' Call WriteConsoleMsg(Torneo.IndexParticipantes(i), "¡Ya estas participado! Solo debes aguardar aquí hasta que seas convocado al torneo.", FontTypeNames.FONTTYPE_INFO)
112     Next i

114     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Evento> Los elegidos para participar son: " & nombres & " damos inicio al evento.", FontTypeNames.FONTTYPE_CITIZEN))

        
        Exit Sub

ComenzarTorneoOk_Err:
116     Call RegistrarError(Err.Number, Err.Description, "ModTorneos.ComenzarTorneoOk", Erl)
118
        
End Sub

Public Sub ResetearTorneo()
        
        On Error GoTo ResetearTorneo_Err
        

        Dim i As Byte

100     Torneo.HayTorneoaActivo = False
102     Torneo.NivelMinimo = 0
104     Torneo.nivelmaximo = 0
106     Torneo.cupos = 0
108     Torneo.costo = 0
110     Torneo.mago = 0
112     Torneo.clerico = 0
114     Torneo.guerrero = 0
116     Torneo.asesino = 0
118     Torneo.bardo = 0
120     Torneo.druido = 0
122     Torneo.Paladin = 0
124     Torneo.cazador = 0
        Torneo.Trabajador = 0
        Torneo.Pirata = 0
        Torneo.Ladron = 0
        Torneo.Bandido = 0
126     Torneo.ClasesTexto = ""
128     Torneo.Mapa = 0
130     Torneo.X = 0
132     Torneo.Y = 0
134
136     Torneo.nombre = ""
138     Torneo.reglas = 0
    
140     For i = 1 To Torneo.Participantes
142         UserList(Torneo.IndexParticipantes(i)).flags.EnTorneo = False
144     Next i

146     Torneo.Participantes = 0
148     ReDim Torneo.IndexParticipantes(1 To 1)
150     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Evento Finalizado. ", FontTypeNames.FONTTYPE_CITIZEN))

        
        Exit Sub

ResetearTorneo_Err:
152     Call RegistrarError(Err.Number, Err.Description, "ModTorneos.ResetearTorneo", Erl)
154
        
End Sub
