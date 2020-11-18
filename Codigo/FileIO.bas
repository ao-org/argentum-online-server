Attribute VB_Name = "ES"
'Argentum Online 0.11.6
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private Type Position

    x As Integer
    Y As Integer

End Type

'Item type
Private Type tItem

    ObjIndex As Integer
    Amount As Integer

End Type

Private Type tWorldPos

    Map As Integer
    x As Byte
    Y As Byte

End Type

Private Type Grh

    GrhIndex As Long
    FrameCounter As Single
    speed As Single
    Started As Byte
    alpha_blend As Boolean
    angle As Single

End Type

Private Type GrhData

    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Integer
    Frames() As Integer
    speed As Integer
    mini_map_color As Long

End Type

Private Type tMapHeader

    NumeroBloqueados As Long
    NumeroLayers(1 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long

End Type

Private Type tDatosBloqueados

    x As Integer
    Y As Integer

End Type

Private Type tDatosGrh

    x As Integer
    Y As Integer
    GrhIndex As Long

End Type

Private Type tDatosTrigger

    x As Integer
    Y As Integer
    trigger As Integer

End Type

Private Type tDatosLuces

    x As Integer
    Y As Integer
    Color As Long
    Rango As Byte

End Type

Private Type tDatosParticulas

    x As Integer
    Y As Integer
    Particula As Long

End Type

Private Type tDatosNPC

    x As Integer
    Y As Integer
    NpcIndex As Integer

End Type

Private Type tDatosObjs

    x As Integer
    Y As Integer
    ObjIndex As Integer
    ObjAmmount As Integer

End Type

Private Type tDatosTE

    x As Integer
    Y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer

End Type

Private Type tMapSize

    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer

End Type

Private Type tMapDat

    map_name As String
    backup_mode As Byte
    restrict_mode As String
    music_numberHi As Long
    music_numberLow As Long
    Seguro As Byte
    zone As String
    terrain As String
    ambient As String
    base_light As Long
    letter_grh As Long
    extra1 As Long
    extra2 As Long
    extra3 As String
    lluvia As Byte
    Nieve As Byte
    niebla As Byte

End Type

Private MapSize As tMapSize
Private MapDat  As tMapDat

Private Type UltimoError
    Componente As String
    Contador As Byte
    ErrorCode As Long
End Type: Private HistorialError As UltimoError

Public Sub CargarSpawnList()
        
        On Error GoTo CargarSpawnList_Err
        

        Dim n As Integer, LoopC As Integer

100     n = val(GetVar(DatPath & "npcs.dat", "INIT", "NumNPCs"))
102     ReDim SpawnList(n) As tCriaturasEntrenador

104     For LoopC = 1 To n

106         SpawnList(LoopC).NpcIndex = LoopC
108         SpawnList(LoopC).NpcName = GetVar(DatPath & "npcs.dat", "NPC" & LoopC, "Name")

110         If Len(SpawnList(LoopC).NpcName) = 0 Then
112             SpawnList(LoopC).NpcName = "Nada"
            End If
            
114     Next LoopC

        
        Exit Sub

CargarSpawnList_Err:
116     Call RegistrarError(Err.Number, Err.description, "ES.CargarSpawnList", Erl)
118     Resume Next
        
End Sub

Function EsAdmin(ByRef name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        
        On Error GoTo EsAdmin_Err
        
100     EsAdmin = (val(Administradores.GetValue("Admin", name)) = 1)

        
        Exit Function

EsAdmin_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.EsAdmin", Erl)
        Resume Next
        
End Function

Function EsDios(ByRef name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        
        On Error GoTo EsDios_Err
        
100     EsDios = (val(Administradores.GetValue("Dios", name)) = 1)

        
        Exit Function

EsDios_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.EsDios", Erl)
        Resume Next
        
End Function

Function EsSemiDios(ByRef name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        
        On Error GoTo EsSemiDios_Err
        
100     EsSemiDios = (val(Administradores.GetValue("SemiDios", name)) = 1)

        
        Exit Function

EsSemiDios_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.EsSemiDios", Erl)
        Resume Next
        
End Function

Function EsConsejero(ByRef name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        
        On Error GoTo EsConsejero_Err
        
100     EsConsejero = (val(Administradores.GetValue("Consejero", name)) = 1)

        
        Exit Function

EsConsejero_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.EsConsejero", Erl)
        Resume Next
        
End Function

Function EsRolesMaster(ByRef name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        
        On Error GoTo EsRolesMaster_Err
        
100     EsRolesMaster = (val(Administradores.GetValue("RM", name)) = 1)

        
        Exit Function

EsRolesMaster_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.EsRolesMaster", Erl)
        Resume Next
        
End Function

Public Function EsGmChar(ByRef name As String) As Boolean
        '***************************************************
        'Author: ZaMa
        'Last Modification: 27/03/2011
        'Returns true if char is administrative user.
        '***************************************************
        
        On Error GoTo EsGmChar_Err
        
    
        Dim EsGM As Boolean
    
        ' Admin?
100     EsGM = EsAdmin(name)

        ' Dios?
102     If Not EsGM Then EsGM = EsDios(name)

        ' Semidios?
104     If Not EsGM Then EsGM = EsSemiDios(name)

        ' Consejero?
106     If Not EsGM Then EsGM = EsConsejero(name)

108     EsGmChar = EsGM

        
        Exit Function

EsGmChar_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.EsGmChar", Erl)
        Resume Next
        
End Function

Public Sub loadAdministrativeUsers()
        'Admines     => Admin
        'Dioses      => Dios
        'SemiDioses  => SemiDios
        'Especiales  => Especial
        'Consejeros  => Consejero
        'RoleMasters => RM
        ' If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando Administradores/Dioses/Gms."
        
        On Error GoTo loadAdministrativeUsers_Err
        

        'Si esta mierda tuviese array asociativos el codigo seria tan lindo.
        Dim buf  As Integer

        Dim i    As Long

        Dim name As String
       
        ' Public container
100     Set Administradores = New clsIniReader
    
        ' Server ini info file
        Dim ServerIni As clsIniReader

102     Set ServerIni = New clsIniReader
    
104     Call ServerIni.Initialize(IniPath & "Server.ini")
       
        ' Admines
106     buf = val(ServerIni.GetValue("INIT", "Admines"))
    
108     For i = 1 To buf
110         name = UCase$(ServerIni.GetValue("Admines", "Admin" & i))
        
112         If Left$(name, 1) = "*" Or Left$(name, 1) = "+" Then name = Right$(name, Len(name) - 1)
        
            ' Add key
114         Call Administradores.ChangeValue("Admin", name, "1")

116     Next i
    
        ' Dioses
118     buf = val(ServerIni.GetValue("INIT", "Dioses"))
    
120     For i = 1 To buf
122         name = UCase$(ServerIni.GetValue("Dioses", "Dios" & i))
        
124         If Left$(name, 1) = "*" Or Left$(name, 1) = "+" Then name = Right$(name, Len(name) - 1)
        
            ' Add key
126         Call Administradores.ChangeValue("Dios", name, "1")
        
128     Next i
        
        ' SemiDioses
130     buf = val(ServerIni.GetValue("INIT", "SemiDioses"))
    
132     For i = 1 To buf
134         name = UCase$(ServerIni.GetValue("SemiDioses", "SemiDios" & i))
        
136         If Left$(name, 1) = "*" Or Left$(name, 1) = "+" Then name = Right$(name, Len(name) - 1)
        
            ' Add key
138         Call Administradores.ChangeValue("SemiDios", name, "1")
        
140     Next i
    
        ' Consejeros
142     buf = val(ServerIni.GetValue("INIT", "Consejeros"))
        
144     For i = 1 To buf
146         name = UCase$(ServerIni.GetValue("Consejeros", "Consejero" & i))
        
148         If Left$(name, 1) = "*" Or Left$(name, 1) = "+" Then name = Right$(name, Len(name) - 1)
        
            ' Add key
150         Call Administradores.ChangeValue("Consejero", name, "1")
        
152     Next i
    
        ' RolesMasters
154     buf = val(ServerIni.GetValue("INIT", "RolesMasters"))
        
156     For i = 1 To buf
158         name = UCase$(ServerIni.GetValue("RolesMasters", "RM" & i))
        
160         If Left$(name, 1) = "*" Or Left$(name, 1) = "+" Then name = Right$(name, Len(name) - 1)
        
            ' Add key
162         Call Administradores.ChangeValue("RM", name, "1")
164     Next i
    
166     Set ServerIni = Nothing

        'If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & Time & " - Los Administradores/Dioses/Gms se han cargado correctamente."

        
        Exit Sub

loadAdministrativeUsers_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.loadAdministrativeUsers", Erl)
        Resume Next
        
End Sub

Public Function GetCharPrivs(ByRef UserName As String) As PlayerType
        '****************************************************
        'Author: ZaMa
        'Last Modification: 18/11/2010
        'Reads the user's charfile and retrieves its privs.
        '***************************************************
        
        On Error GoTo GetCharPrivs_Err
        

        Dim privs As PlayerType

100     If EsAdmin(UserName) Then
102         privs = PlayerType.Admin
        
104     ElseIf EsDios(UserName) Then
106         privs = PlayerType.Dios

108     ElseIf EsSemiDios(UserName) Then
110         privs = PlayerType.SemiDios
        
112     ElseIf EsConsejero(UserName) Then
114         privs = PlayerType.Consejero
    
        Else
116         privs = PlayerType.user

        End If

118     GetCharPrivs = privs

        
        Exit Function

GetCharPrivs_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.GetCharPrivs", Erl)
        Resume Next
        
End Function

Public Function TxtDimension(ByVal name As String) As Long
        
        On Error GoTo TxtDimension_Err
        

        Dim n As Integer, cad As String, Tam As Long

100     n = FreeFile(1)
102     Open name For Input As #n
104     Tam = 0

106     Do While Not EOF(n)
108         Tam = Tam + 1
110         Line Input #n, cad
        Loop
112     Close n
114     TxtDimension = Tam

        
        Exit Function

TxtDimension_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.TxtDimension", Erl)
        Resume Next
        
End Function

Public Sub CargarForbidenWords()
        
        On Error GoTo CargarForbidenWords_Err
        

        Dim Size As Integer

100     Size = TxtDimension(DatPath & "NombresInvalidos.txt")
    
102     If Size = 0 Then
104         ReDim ForbidenNames(0)
            Exit Sub

        End If
    
106     ReDim ForbidenNames(1 To Size)

        Dim n As Integer, i As Integer

108     n = FreeFile(1)
110     Open DatPath & "NombresInvalidos.txt" For Input As #n
    
112     For i = 1 To UBound(ForbidenNames)
114         Line Input #n, ForbidenNames(i)
116     Next i
    
118     Close n

        
        Exit Sub

CargarForbidenWords_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.CargarForbidenWords", Erl)
        Resume Next
        
End Sub

Public Sub CargarHechizos()

    '###################################################
    '#               ATENCION PELIGRO                  #
    '###################################################
    '
    '  ¡¡¡¡ NO USAR GetVar PARA LEER Hechizos.dat !!!!
    '
    'El que ose desafiar esta LEY, se las tendrá que ver
    'con migo. Para leer Hechizos.dat se deberá usar
    'la nueva clase clsLeerInis.
    '
    'Alejo
    '
    '###################################################

    On Error GoTo Errhandler

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."

    Dim Hechizo As Integer

    Dim Leer    As New clsIniReader

    Call Leer.Initialize(DatPath & "Hechizos.dat")

    'obtiene el numero de hechizos
    NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))

    ReDim Hechizos(1 To NumeroHechizos) As tHechizo

    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumeroHechizos
    frmCargando.cargar.Value = 0

    'Llena la lista
    For Hechizo = 1 To NumeroHechizos

        Hechizos(Hechizo).Velocidad = val(Leer.GetValue("Hechizo" & Hechizo, "Velocidad"))
    
        'Materializacion
        Hechizos(Hechizo).MaterializaObj = val(Leer.GetValue("Hechizo" & Hechizo, "MaterializaObj"))
        Hechizos(Hechizo).MaterializaCant = val(Leer.GetValue("Hechizo" & Hechizo, "MaterializaCant"))
        'Materializacion
    
        'Screen Efecto
        Hechizos(Hechizo).ScreenColor = val(Leer.GetValue("Hechizo" & Hechizo, "ScreenColor"))
        Hechizos(Hechizo).TimeEfect = val(Leer.GetValue("Hechizo" & Hechizo, "TimeEfect"))
        'Screen Efecto

        Hechizos(Hechizo).TeleportX = val(Leer.GetValue("Hechizo" & Hechizo, "Teleport"))
        Hechizos(Hechizo).TeleportXMap = val(Leer.GetValue("Hechizo" & Hechizo, "TeleportMap"))
        Hechizos(Hechizo).TeleportXX = val(Leer.GetValue("Hechizo" & Hechizo, "TeleportX"))
        Hechizos(Hechizo).TeleportXY = val(Leer.GetValue("Hechizo" & Hechizo, "TeleportY"))

        Hechizos(Hechizo).nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
        Hechizos(Hechizo).Desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
        Hechizos(Hechizo).PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
    
        Hechizos(Hechizo).HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
        Hechizos(Hechizo).TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
        Hechizos(Hechizo).PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
    
        Hechizos(Hechizo).NecesitaObj = val(Leer.GetValue("Hechizo" & Hechizo, "NecesitaObj"))
        Hechizos(Hechizo).NecesitaObj2 = val(Leer.GetValue("Hechizo" & Hechizo, "NecesitaObj2"))
    
        Hechizos(Hechizo).Tipo = val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
        Hechizos(Hechizo).wav = val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
        Hechizos(Hechizo).FXgrh = val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
        Hechizos(Hechizo).Particle = val(Leer.GetValue("Hechizo" & Hechizo, "Particle"))
        Hechizos(Hechizo).ParticleViaje = val(Leer.GetValue("Hechizo" & Hechizo, "ParticleViaje"))
        Hechizos(Hechizo).TimeParticula = val(Leer.GetValue("Hechizo" & Hechizo, "TimeParticula"))
        Hechizos(Hechizo).desencantar = val(Leer.GetValue("Hechizo" & Hechizo, "desencantar"))
        Hechizos(Hechizo).Sanacion = val(Leer.GetValue("Hechizo" & Hechizo, "Sanacion"))
        Hechizos(Hechizo).AntiRm = val(Leer.GetValue("Hechizo" & Hechizo, "AntiRm"))
        'Hechizos de area
        Hechizos(Hechizo).AreaRadio = val(Leer.GetValue("Hechizo" & Hechizo, "AreaRadio"))
        Hechizos(Hechizo).AreaAfecta = val(Leer.GetValue("Hechizo" & Hechizo, "AreaAfecta"))
        'Hechizos de area
    
        Hechizos(Hechizo).incinera = val(Leer.GetValue("Hechizo" & Hechizo, "Incinera"))
    
        Hechizos(Hechizo).AutoLanzar = val(Leer.GetValue("Hechizo" & Hechizo, "AutoLanzar"))
    
        Hechizos(Hechizo).CoolDown = val(Leer.GetValue("Hechizo" & Hechizo, "CoolDown"))
    
        Hechizos(Hechizo).loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
    
        '    Hechizos(Hechizo).Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
    
        Hechizos(Hechizo).SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
        Hechizos(Hechizo).MinHp = val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
        Hechizos(Hechizo).MaxHp = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
    
        Hechizos(Hechizo).SubeMana = val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
        Hechizos(Hechizo).MiMana = val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
        Hechizos(Hechizo).MaMana = val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
    
        Hechizos(Hechizo).SubeSta = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
        Hechizos(Hechizo).MinSta = val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
        Hechizos(Hechizo).MaxSta = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
    
        Hechizos(Hechizo).SubeHam = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
        Hechizos(Hechizo).MinHam = val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
        Hechizos(Hechizo).MaxHam = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
    
        Hechizos(Hechizo).SubeSed = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
        Hechizos(Hechizo).MinSed = val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
        Hechizos(Hechizo).MaxSed = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
    
        Hechizos(Hechizo).SubeAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
        Hechizos(Hechizo).MinAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
        Hechizos(Hechizo).MaxAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
    
        Hechizos(Hechizo).SubeFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
        Hechizos(Hechizo).MinFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
        Hechizos(Hechizo).MaxFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
    
        Hechizos(Hechizo).SubeCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
        Hechizos(Hechizo).MinCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
        Hechizos(Hechizo).MaxCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
    
        Hechizos(Hechizo).Invisibilidad = val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
        Hechizos(Hechizo).Paraliza = val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
        Hechizos(Hechizo).Inmoviliza = val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
    
        Hechizos(Hechizo).RemoverParalisis = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
        Hechizos(Hechizo).RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
        Hechizos(Hechizo).RemueveInvisibilidadParcial = val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
    
        Hechizos(Hechizo).CuraVeneno = val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
        Hechizos(Hechizo).Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
        Hechizos(Hechizo).Maldicion = val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
        Hechizos(Hechizo).RemoverMaldicion = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
        Hechizos(Hechizo).Bendicion = val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
        Hechizos(Hechizo).Revivir = val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
    
        Hechizos(Hechizo).Ceguera = val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
        Hechizos(Hechizo).Estupidez = val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
    
        Hechizos(Hechizo).Invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
        Hechizos(Hechizo).NumNpc = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
        Hechizos(Hechizo).cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
    
        Hechizos(Hechizo).GolpeCertero = val(Leer.GetValue("Hechizo" & Hechizo, "GolpeCertero"))
    
        '    Hechizos(Hechizo).Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
        '    Hechizos(Hechizo).ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
    
        Hechizos(Hechizo).MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
        Hechizos(Hechizo).ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
        Hechizos(Hechizo).RequiredHP = val(Leer.GetValue("Hechizo" & Hechizo, "RequiredHP"))
    
        Hechizos(Hechizo).Duration = val(Leer.GetValue("Hechizo" & Hechizo, "Duration"))
    
        'Barrin 30/9/03
        Hechizos(Hechizo).StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
    
        Hechizos(Hechizo).Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
        frmCargando.cargar.Value = frmCargando.cargar.Value + 1
    
        Hechizos(Hechizo).NeedStaff = val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
        Hechizos(Hechizo).StaffAffected = CBool(val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))
    
    Next Hechizo

    Set Leer = Nothing
    Exit Sub

Errhandler:
    MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.description
 
End Sub

Sub LoadMotd()
        
        On Error GoTo LoadMotd_Err
        

        Dim i As Integer

100     MaxLines = val(GetVar(DatPath & "Motd.ini", "INIT", "NumLines"))

102     ReDim MOTD(1 To MaxLines)

104     For i = 1 To MaxLines
106         MOTD(i).texto = GetVar(DatPath & "Motd.ini", "Motd", "Line" & i)
108         MOTD(i).Formato = vbNullString
110     Next i

        
        Exit Sub

LoadMotd_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LoadMotd", Erl)
        Resume Next
        
End Sub

Public Sub DoBackUp()
    'Call LogTarea("Sub DoBackUp")
    haciendoBK = True

    Dim i As Integer

    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

    'Call WorldSave
    'Call modGuilds.v_RutinaElecciones
    
    'Reseteamos al centinela
    Call ResetCentinelaInfo

    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

    'Call EstadisticasWeb.Informar(EVENTO_NUEVO_CLAN, 0)

    haciendoBK = False

    'Log
    On Error Resume Next

    Dim nfile As Integer: nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time
    Close #nfile

End Sub

Public Sub GrabarMapa(ByVal Map As Long, ByVal MAPFILE As String)
    Debug.Print "Empezamos a grabar"

    On Error GoTo ErrorHandler

    Dim MapRoute As String: MapRoute = MAPFILE & ".csm"

    Dim fh           As Integer
    Dim MH           As tMapHeader

    Dim Blqs()       As tDatosBloqueados

    Dim L1()         As tDatosGrh
    Dim L2()         As tDatosGrh
    Dim L3()         As tDatosGrh
    Dim L4()         As tDatosGrh

    Dim Triggers()   As tDatosTrigger
    Dim Luces()      As tDatosLuces
    Dim Particulas() As tDatosParticulas
    Dim Objetos()    As tDatosObjs
    Dim NPCs()       As tDatosNPC
    Dim TEs()        As tDatosTE

    Dim i            As Long
    Dim j            As Integer

    Dim tmpLng       As Long

    For j = 1 To 100
        For i = 1 To 100

            With MapData(Map, i, j)
            
                If .Blocked Then
                    MH.NumeroBloqueados = MH.NumeroBloqueados + 1
                    ReDim Preserve Blqs(1 To MH.NumeroBloqueados)
                    Blqs(MH.NumeroBloqueados).x = i
                    Blqs(MH.NumeroBloqueados).Y = j

                End If
            
                Rem L1(i, j) = .Graphic(1).grhindex
  
                If .Graphic(1) > 0 Then
                    MH.NumeroLayers(1) = MH.NumeroLayers(1) + 1
                    ReDim Preserve L1(1 To MH.NumeroLayers(1))
                    L1(MH.NumeroLayers(1)).x = i
                    L1(MH.NumeroLayers(1)).Y = j
                    L1(MH.NumeroLayers(1)).GrhIndex = .Graphic(1)

                End If
            
                If .Graphic(2) > 0 Then
                    MH.NumeroLayers(2) = MH.NumeroLayers(2) + 1
                    ReDim Preserve L2(1 To MH.NumeroLayers(2))
                    L2(MH.NumeroLayers(2)).x = i
                    L2(MH.NumeroLayers(2)).Y = j
                    L2(MH.NumeroLayers(2)).GrhIndex = .Graphic(2)

                End If
            
                If .Graphic(3) > 0 Then
                    MH.NumeroLayers(3) = MH.NumeroLayers(3) + 1
                    ReDim Preserve L3(1 To MH.NumeroLayers(3))
                    L3(MH.NumeroLayers(3)).x = i
                    L3(MH.NumeroLayers(3)).Y = j
                    L3(MH.NumeroLayers(3)).GrhIndex = .Graphic(3)

                End If
            
                If .Graphic(4) > 0 Then
                    MH.NumeroLayers(4) = MH.NumeroLayers(4) + 1
                    ReDim Preserve L4(1 To MH.NumeroLayers(4))
                    L4(MH.NumeroLayers(4)).x = i
                    L4(MH.NumeroLayers(4)).Y = j
                    L4(MH.NumeroLayers(4)).GrhIndex = .Graphic(4)

                End If
            
                If .trigger > 0 Then
                    MH.NumeroTriggers = MH.NumeroTriggers + 1
                    ReDim Preserve Triggers(1 To MH.NumeroTriggers)
                    Triggers(MH.NumeroTriggers).x = i
                    Triggers(MH.NumeroTriggers).Y = j
                    Triggers(MH.NumeroTriggers).trigger = .trigger

                End If
            
                If .ParticulaIndex > 0 Then
                    MH.NumeroParticulas = MH.NumeroParticulas + 1
                    ReDim Preserve Particulas(1 To MH.NumeroParticulas)
                    Particulas(MH.NumeroParticulas).x = i
                    Particulas(MH.NumeroParticulas).Y = j
                    Particulas(MH.NumeroParticulas).Particula = .ParticulaIndex

                End If
            
                Rem   If MapData(i, j).luz.Rango > 0 Then
                Rem      MH.NumeroLuces = MH.NumeroLuces + 1
                Rem       ReDim Preserve Luces(1 To MH.NumeroLuces)
                Rem       Luces(MH.NumeroLuces).X = i
                Rem       Luces(MH.NumeroLuces).Y = j
                Rem      Luces(MH.NumeroLuces).color = .luz.color
                Rem       Luces(MH.NumeroLuces).Rango = .luz.Rango
                Rem  End If
            
                If .ObjInfo.ObjIndex > 0 Then
                    MH.NumeroOBJs = MH.NumeroOBJs + 1
                    ReDim Preserve Objetos(1 To MH.NumeroOBJs)
                    Objetos(MH.NumeroOBJs).ObjIndex = .ObjInfo.ObjIndex
                    Objetos(MH.NumeroOBJs).ObjAmmount = .ObjInfo.Amount
               
                    Objetos(MH.NumeroOBJs).x = i
                    Objetos(MH.NumeroOBJs).Y = j
                
                End If
            
                If .NpcIndex > 0 Then
                    MH.NumeroNPCs = MH.NumeroNPCs + 1
                    ReDim Preserve NPCs(1 To MH.NumeroNPCs)
                    NPCs(MH.NumeroNPCs).NpcIndex = .NpcIndex
                    NPCs(MH.NumeroNPCs).x = i
                    NPCs(MH.NumeroNPCs).Y = j

                End If
            
                If .TileExit.Map > 0 Then
                    MH.NumeroTE = MH.NumeroTE + 1
                    ReDim Preserve TEs(1 To MH.NumeroTE)
                    TEs(MH.NumeroTE).DestM = .TileExit.Map
                    TEs(MH.NumeroTE).DestX = .TileExit.x
                    TEs(MH.NumeroTE).DestY = .TileExit.Y
                    TEs(MH.NumeroTE).x = i
                    TEs(MH.NumeroTE).Y = j

                End If

            End With

        Next i
    Next j
          
    fh = FreeFile
    Open MapRoute For Binary As fh
    
    Put #fh, , MH
    Put #fh, , MapSize
    Put #fh, , MapDat
    Rem   Put #fh, , L1
    
    With MH

        If .NumeroBloqueados > 0 Then Put #fh, , Blqs

        If .NumeroLayers(1) > 0 Then Put #fh, , L1

        If .NumeroLayers(2) > 0 Then Put #fh, , L2

        If .NumeroLayers(3) > 0 Then Put #fh, , L3

        If .NumeroLayers(4) > 0 Then Put #fh, , L4

        If .NumeroTriggers > 0 Then Put #fh, , Triggers

        If .NumeroParticulas > 0 Then Put #fh, , Particulas

        If .NumeroLuces > 0 Then Put #fh, , Luces

        If .NumeroOBJs > 0 Then Put #fh, , Objetos

        If .NumeroNPCs > 0 Then Put #fh, , NPCs

        If .NumeroTE > 0 Then Put #fh, , TEs

    End With

    Close fh

    Rem MsgBox "Mapa grabado"

    Debug.Print "Mapa grabado"

ErrorHandler:

    If fh <> 0 Then Close fh

End Sub

Sub LoadArmasHerreria()
        
        On Error GoTo LoadArmasHerreria_Err
        

        Dim n As Integer, lc As Integer
    
100     n = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))
    
102     If n = 0 Then
104         ReDim ArmasHerrero(0) As Integer
            Exit Sub

        End If
    
106     ReDim Preserve ArmasHerrero(1 To n) As Integer
    
108     For lc = 1 To n
110         ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
112     Next lc

        
        Exit Sub

LoadArmasHerreria_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LoadArmasHerreria", Erl)
        Resume Next
        
End Sub

Sub LoadArmadurasHerreria()
        
        On Error GoTo LoadArmadurasHerreria_Err
        

        Dim n As Integer, lc As Integer
    
100     n = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))
    
102     If n = 0 Then
104         ReDim ArmadurasHerrero(0) As Integer
            Exit Sub

        End If
    
106     ReDim Preserve ArmadurasHerrero(1 To n) As Integer
    
108     For lc = 1 To n
110         ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
112     Next lc

        
        Exit Sub

LoadArmadurasHerreria_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LoadArmadurasHerreria", Erl)
        Resume Next
        
End Sub

Sub LoadBalance()
        
        On Error GoTo LoadBalance_Err
        

        Dim BalanceIni As clsIniReader

100     Set BalanceIni = New clsIniReader
    
102     BalanceIni.Initialize DatPath & "Balance.dat"
    
        Dim i         As Long

        Dim SearchVar As String

        'Modificadores de Clase
104     For i = 1 To NUMCLASES
            SearchVar = Replace$(ListaClases(i), " ", vbNullString)

106         With ModClase(i)
108             .Evasion = val(BalanceIni.GetValue("MODEVASION", SearchVar))
110             .AtaqueArmas = val(BalanceIni.GetValue("MODATAQUEARMAS", SearchVar))
112             .AtaqueProyectiles = val(BalanceIni.GetValue("MODATAQUEPROYECTILES", SearchVar))
                '.DañoWrestling = val(BalanceIni.GetValue("MODATAQUEWRESTLING", SearchVar))
114             .DañoArmas = val(BalanceIni.GetValue("MODDANOARMAS", SearchVar))
116             .DañoProyectiles = val(BalanceIni.GetValue("MODDANOPROYECTILES", SearchVar))
118             .DañoWrestling = val(BalanceIni.GetValue("MODDANOWRESTLING", SearchVar))
120             .Escudo = val(BalanceIni.GetValue("MODESCUDO", SearchVar))

                'Modificadores de Vida
                ModVida(i) = val(BalanceIni.GetValue("MODVIDA", SearchVar))

            End With

122     Next i
    
        'Modificadores de Raza
124     For i = 1 To NUMRAZAS
            SearchVar = Replace$(ListaRazas(i), " ", vbNullString)

126         With ModRaza(i)
128             .Fuerza = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Fuerza"))
130             .Agilidad = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Agilidad"))
132             .Inteligencia = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Inteligencia"))
                .Carisma = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Carisma"))
134             .Constitucion = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Constitucion"))
            End With

136     Next i
    
        'Distribucion de Vida
144     For i = 1 To 5
146         DistribucionEnteraVida(i) = val(BalanceIni.GetValue("DISTRIBUCION", "E" + CStr(i)))
148     Next i

150     For i = 1 To 4
152         DistribucionSemienteraVida(i) = val(BalanceIni.GetValue("DISTRIBUCION", "S" + CStr(i)))
154     Next i
    
            'Experiencia por nivel
        For i = 1 To STAT_MAXELV
            ExpByLevel(i) = val(BalanceIni.GetValue("EXPBYLEVEL", i))
        Next i

        'Extra
156     PorcentajeRecuperoMana = val(BalanceIni.GetValue("EXTRA", "PorcentajeRecuperoMana"))
        DificultadSubirSkill = val(BalanceIni.GetValue("EXTRA", "DificultadSubirSkill"))
    
158     Set BalanceIni = Nothing
    
160     AgregarAConsola "Se cargó el balance (Balance.dat)"

        
        Exit Sub

LoadBalance_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LoadBalance", Erl)
        Resume Next
        
End Sub

Sub LoadObjCarpintero()
        
        On Error GoTo LoadObjCarpintero_Err
        

        Dim n As Integer, lc As Integer
    
100     n = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))
    
102     If n = 0 Then
104         ReDim ObjCarpintero(0) As Integer
            Exit Sub

        End If
    
106     ReDim Preserve ObjCarpintero(1 To n) As Integer
    
108     For lc = 1 To n
110         ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
112     Next lc

        
        Exit Sub

LoadObjCarpintero_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LoadObjCarpintero", Erl)
        Resume Next
        
End Sub

Sub LoadObjAlquimista()
        
        On Error GoTo LoadObjAlquimista_Err
        
    
        Dim n As Integer, lc As Integer
    
100     n = val(GetVar(DatPath & "ObjAlquimista.dat", "INIT", "NumObjs"))
    
102     If n = 0 Then
104         ReDim ObjAlquimista(0) As Integer
            Exit Sub

        End If
    
106     ReDim Preserve ObjAlquimista(1 To n) As Integer
    
108     For lc = 1 To n
110         ObjAlquimista(lc) = val(GetVar(DatPath & "ObjAlquimista.dat", "Obj" & lc, "Index"))
112     Next lc

        
        Exit Sub

LoadObjAlquimista_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LoadObjAlquimista", Erl)
        Resume Next
        
End Sub

Sub LoadObjSastre()
        
        On Error GoTo LoadObjSastre_Err
        

        Dim n As Integer, lc As Integer
    
100     n = val(GetVar(DatPath & "ObjSastre.dat", "INIT", "NumObjs"))
    
102     If n = 0 Then
104         ReDim ObjSastre(0) As Integer
            Exit Sub

        End If
    
106     ReDim Preserve ObjSastre(1 To n) As Integer
    
108     For lc = 1 To n
110         ObjSastre(lc) = val(GetVar(DatPath & "ObjSastre.dat", "Obj" & lc, "Index"))
112     Next lc

        
        Exit Sub

LoadObjSastre_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LoadObjSastre", Erl)
        Resume Next
        
End Sub

Sub LoadObjDonador()
        
        On Error GoTo LoadObjDonador_Err
        

        Dim n As Integer, lc As Integer

100     n = val(GetVar(DatPath & "ObjDonador.dat", "INIT", "NumObjs"))

102     ReDim Preserve ObjDonador(1 To n) As tObjDonador

104     For lc = 1 To n
106         ObjDonador(lc).ObjIndex = val(GetVar(DatPath & "ObjDonador.dat", "Obj" & lc, "Index"))
108         ObjDonador(lc).Cantidad = val(GetVar(DatPath & "ObjDonador.dat", "Obj" & lc, "Cant"))
110         ObjDonador(lc).Valor = val(GetVar(DatPath & "ObjDonador.dat", "Obj" & lc, "Valor"))
112     Next lc

        
        Exit Sub

LoadObjDonador_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LoadObjDonador", Erl)
        Resume Next
        
End Sub

Sub LoadOBJData()

    '###################################################
    '#               ATENCION PELIGRO                  #
    '###################################################
    '
    '¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
    '
    'El que ose desafiar esta LEY, se las tendrá que ver
    'con migo. Para leer desde el OBJ.DAT se deberá usar
    'la nueva clase clsLeerInis.
    '
    'Alejo
    '
    '###################################################

    'Call LogTarea("Sub LoadOBJData")

    On Error GoTo Errhandler

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."

    '*****************************************************************
    'Carga la lista de objetos
    '*****************************************************************
    Dim Object As Integer

    Dim Leer   As clsIniReader
    Set Leer = New clsIniReader
    Call Leer.Initialize(DatPath & "Obj.dat")

    'obtiene el numero de obj
    NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))

    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumObjDatas
    frmCargando.cargar.Value = 0

    ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
    
    Dim str As String, Field() As String
  
    'Llena la lista
    For Object = 1 To NumObjDatas
        
        ObjData(Object).name = Leer.GetValue("OBJ" & Object, "Name")
    
        ' If ObjData(Object).Name = "" Then
        '   Call LogError("Objeto libre:" & Object)
        ' End If
    
        ' If ObjData(Object).name = "" Then
        ' Debug.Print Object
        ' End If
    
        'Pablo (ToxicWaste) Log de Objetos.
        ObjData(Object).Log = val(Leer.GetValue("OBJ" & Object, "Log"))
        ObjData(Object).NoLog = val(Leer.GetValue("OBJ" & Object, "NoLog"))
        '07/09/07
    
        ObjData(Object).GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex"))

        If ObjData(Object).GrhIndex = 0 Then
            ObjData(Object).GrhIndex = ObjData(Object).GrhIndex

        End If
    
        ObjData(Object).OBJType = val(Leer.GetValue("OBJ" & Object, "ObjType"))
    
        ObjData(Object).Newbie = val(Leer.GetValue("OBJ" & Object, "Newbie"))
        'Propiedades by Lader 05-05-08
        ObjData(Object).Instransferible = val(Leer.GetValue("OBJ" & Object, "Instransferible"))
        ObjData(Object).Destruye = val(Leer.GetValue("OBJ" & Object, "Destruye"))
        ObjData(Object).Intirable = val(Leer.GetValue("OBJ" & Object, "Intirable"))
    
        ObjData(Object).CantidadSkill = val(Leer.GetValue("OBJ" & Object, "CantidadSkill"))
        ObjData(Object).QueSkill = val(Leer.GetValue("OBJ" & Object, "QueSkill"))
        ObjData(Object).QueAtributo = val(Leer.GetValue("OBJ" & Object, "queatributo"))
        ObjData(Object).CuantoAumento = val(Leer.GetValue("OBJ" & Object, "cuantoaumento"))
        ObjData(Object).MinELV = val(Leer.GetValue("OBJ" & Object, "MinELV"))
        ObjData(Object).Subtipo = val(Leer.GetValue("OBJ" & Object, "Subtipo"))
        ObjData(Object).VidaUtil = val(Leer.GetValue("OBJ" & Object, "VidaUtil"))
        ObjData(Object).TiempoRegenerar = val(Leer.GetValue("OBJ" & Object, "TiempoRegenerar"))
    
        ObjData(Object).donador = val(Leer.GetValue("OBJ" & Object, "donador"))
    
        Dim i As Integer

        'Propiedades by Lader 05-05-08
        Select Case ObjData(Object).OBJType

            Case eOBJType.otHerramientas
                ObjData(Object).WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                ObjData(Object).Power = val(Leer.GetValue("OBJ" & Object, "Poder"))
            
            Case eOBJType.otArmadura
                ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                ObjData(Object).ResistenciaMagica = val(Leer.GetValue("OBJ" & Object, "ResistenciaMagica"))
        
            Case eOBJType.otESCUDO
                ObjData(Object).ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                ObjData(Object).ResistenciaMagica = val(Leer.GetValue("OBJ" & Object, "ResistenciaMagica"))
        
            Case eOBJType.otCASCO
                ObjData(Object).CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                ObjData(Object).ResistenciaMagica = val(Leer.GetValue("OBJ" & Object, "ResistenciaMagica"))
        
            Case eOBJType.otWeapon
                ObjData(Object).WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                ObjData(Object).Apuñala = val(Leer.GetValue("OBJ" & Object, "Apuñala"))
                ObjData(Object).Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
                ObjData(Object).Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
                ObjData(Object).Estupidiza = val(Leer.GetValue("OBJ" & Object, "Estupidiza"))
        
                ObjData(Object).MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                ObjData(Object).proyectil = val(Leer.GetValue("OBJ" & Object, "Proyectil"))
                ObjData(Object).Municion = val(Leer.GetValue("OBJ" & Object, "Municiones"))
                ObjData(Object).Power = val(Leer.GetValue("OBJ" & Object, "StaffPower"))
                ObjData(Object).StaffDamageBonus = val(Leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
                ObjData(Object).Refuerzo = val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
            
                ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                ObjData(Object).EfectoMagico = val(Leer.GetValue("OBJ" & Object, "efectomagico"))
        
            Case eOBJType.otInstrumentos
        
                'Pablo (ToxicWaste)
                ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
            Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
                ObjData(Object).IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
                ObjData(Object).IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
                ObjData(Object).IndexCerradaLlave = val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
        
            Case otPociones
                ObjData(Object).TipoPocion = val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
                ObjData(Object).MaxModificador = val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
                ObjData(Object).MinModificador = val(Leer.GetValue("OBJ" & Object, "MinModificador"))
            
                ObjData(Object).DuracionEfecto = val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
                ObjData(Object).Raices = val(Leer.GetValue("OBJ" & Object, "Raices"))
                ObjData(Object).SkPociones = val(Leer.GetValue("OBJ" & Object, "SkPociones"))
                ObjData(Object).Porcentaje = val(Leer.GetValue("OBJ" & Object, "Porcentaje"))
        
            Case eOBJType.otBarcos
                ObjData(Object).MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                ObjData(Object).Subtipo = val(Leer.GetValue("OBJ" & Object, "Subtipo"))
                ObjData(Object).Velocidad = val(Leer.GetValue("OBJ" & Object, "Velocidad"))

            Case eOBJType.otMonturas
                ObjData(Object).MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                ObjData(Object).MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
                ObjData(Object).MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
                ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
            Case eOBJType.otFlechas
                ObjData(Object).MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                ObjData(Object).Subtipo = val(Leer.GetValue("OBJ" & Object, "Subtipo"))
            
                ObjData(Object).Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
                ObjData(Object).Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
            
                'Case eOBJType.otAnillo 'Pablo (ToxicWaste)
                '  ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                '  ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                '  ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                '  ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            
                'Pasajes Ladder 05-05-08
            Case eOBJType.otpasajes
                ObjData(Object).DesdeMap = val(Leer.GetValue("OBJ" & Object, "DesdeMap"))
                ObjData(Object).HastaMap = val(Leer.GetValue("OBJ" & Object, "Map"))
                ObjData(Object).HastaX = val(Leer.GetValue("OBJ" & Object, "X"))
                ObjData(Object).HastaY = val(Leer.GetValue("OBJ" & Object, "Y"))
                ObjData(Object).NecesitaNave = val(Leer.GetValue("OBJ" & Object, "NecesitaNave"))
            
            Case eOBJType.OtDonador
                ObjData(Object).Subtipo = val(Leer.GetValue("OBJ" & Object, "Subtipo"))
                ObjData(Object).HastaMap = val(Leer.GetValue("OBJ" & Object, "Map"))
                ObjData(Object).HastaX = val(Leer.GetValue("OBJ" & Object, "X"))
                ObjData(Object).HastaY = val(Leer.GetValue("OBJ" & Object, "Y"))
        
            Case eOBJType.otmagicos
                ObjData(Object).EfectoMagico = val(Leer.GetValue("OBJ" & Object, "efectomagico"))

                If ObjData(Object).EfectoMagico = 15 Then
                    PENDIENTE = Object

                End If
            
            Case eOBJType.otRunas
                ObjData(Object).TipoRuna = val(Leer.GetValue("OBJ" & Object, "TipoRuna"))
                ObjData(Object).DesdeMap = val(Leer.GetValue("OBJ" & Object, "DesdeMap"))
                ObjData(Object).HastaMap = val(Leer.GetValue("OBJ" & Object, "Map"))
                ObjData(Object).HastaX = val(Leer.GetValue("OBJ" & Object, "X"))
                ObjData(Object).HastaY = val(Leer.GetValue("OBJ" & Object, "Y"))
                    
            Case eOBJType.otNUDILLOS
                ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                ObjData(Object).MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHit"))
                ObjData(Object).Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
                ObjData(Object).Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
                ObjData(Object).Estupidiza = val(Leer.GetValue("OBJ" & Object, "Estupidiza"))
                ObjData(Object).WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            
            Case eOBJType.otPergaminos
        
                ' ObjData(Object).ClasePermitida = Leer.GetValue("OBJ" & Object, "CP")
        
            Case eOBJType.OtCofre
                ObjData(Object).CantItem = val(Leer.GetValue("OBJ" & Object, "CantItem"))
                ObjData(Object).Subtipo = val(Leer.GetValue("OBJ" & Object, "SubTipo"))

                If ObjData(Object).Subtipo = 1 Then
                    ReDim ObjData(Object).Item(1 To ObjData(Object).CantItem)
                
                    For i = 1 To ObjData(Object).CantItem
                        ObjData(Object).Item(i).ObjIndex = val(Leer.GetValue("OBJ" & Object, "Item" & i))
                        ObjData(Object).Item(i).Amount = val(Leer.GetValue("OBJ" & Object, "Cantidad" & i))
                    Next i

                Else
                    ReDim ObjData(Object).Item(1 To ObjData(Object).CantItem)
                
                    ObjData(Object).CantEntrega = val(Leer.GetValue("OBJ" & Object, "CantEntrega"))

                    For i = 1 To ObjData(Object).CantItem
                        ObjData(Object).Item(i).ObjIndex = val(Leer.GetValue("OBJ" & Object, "Item" & i))
                        ObjData(Object).Item(i).Amount = val(Leer.GetValue("OBJ" & Object, "Cantidad" & i))
                    Next i

                End If
            
            Case eOBJType.otYacimiento
                ' Drop gemas yacimientos
                ObjData(Object).CantItem = val(Leer.GetValue("OBJ" & Object, "Gemas"))
            
                If ObjData(Object).CantItem > 0 Then
                    ReDim ObjData(Object).Item(1 To ObjData(Object).CantItem)

                    For i = 1 To ObjData(Object).CantItem
                        str = Leer.GetValue("OBJ" & Object, "Gema" & i)
                        Field = Split(str, "-")
                        ObjData(Object).Item(i).ObjIndex = val(Field(0))    ' ObjIndex
                        ObjData(Object).Item(i).Amount = val(Field(1))      ' Probabilidad de drop (1 en X)
                    Next i

                End If
            
        End Select
    
        ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))

        ObjData(Object).Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
        ObjData(Object).Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
        ObjData(Object).Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
        'DELETE
        ObjData(Object).SndAura = val(Leer.GetValue("OBJ" & Object, "SndAura"))
        '
    
        ObjData(Object).NoSeLimpia = val(Leer.GetValue("OBJ" & Object, "NoSeLimpia"))
        ObjData(Object).Subastable = val(Leer.GetValue("OBJ" & Object, "Subastable"))
    
        ObjData(Object).ParticulaGolpe = val(Leer.GetValue("OBJ" & Object, "ParticulaGolpe"))
        ObjData(Object).ParticulaViaje = val(Leer.GetValue("OBJ" & Object, "ParticulaViaje"))
        ObjData(Object).ParticulaGolpeTime = val(Leer.GetValue("OBJ" & Object, "ParticulaGolpeTime"))
    
        ObjData(Object).Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
        ObjData(Object).HechizoIndex = val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
    
        ObjData(Object).LingoteIndex = val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
    
        ObjData(Object).MineralIndex = val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
    
        ObjData(Object).MaxHp = val(Leer.GetValue("OBJ" & Object, "MaxHP"))
        ObjData(Object).MinHp = val(Leer.GetValue("OBJ" & Object, "MinHP"))
    
        ObjData(Object).Mujer = val(Leer.GetValue("OBJ" & Object, "Mujer"))
        ObjData(Object).Hombre = val(Leer.GetValue("OBJ" & Object, "Hombre"))
    
        ObjData(Object).PielLobo = val(Leer.GetValue("OBJ" & Object, "PielLobo"))
        ObjData(Object).PielOsoPardo = val(Leer.GetValue("OBJ" & Object, "PielOsoPardo"))
        ObjData(Object).PielOsoPolaR = val(Leer.GetValue("OBJ" & Object, "PielOsoPolaR"))
        ObjData(Object).SkMAGOria = val(Leer.GetValue("OBJ" & Object, "SKSastreria"))
    
        ObjData(Object).CreaParticula = Leer.GetValue("OBJ" & Object, "CreaParticula")
    
        ObjData(Object).CreaFX = val(Leer.GetValue("OBJ" & Object, "CreaFX"))
  
        'DELETE
        ObjData(Object).CreaParticulaPiso = val(Leer.GetValue("OBJ" & Object, "CreaParticulaPiso"))
        '
    
        ObjData(Object).CreaGRH = Leer.GetValue("OBJ" & Object, "CreaGRH")
        ObjData(Object).CreaLuz = Leer.GetValue("OBJ" & Object, "CreaLuz")
    
        ObjData(Object).MinHam = val(Leer.GetValue("OBJ" & Object, "MinHam"))
        ObjData(Object).MinSed = val(Leer.GetValue("OBJ" & Object, "MinAgu"))
    
        ObjData(Object).MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
        ObjData(Object).MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
        ObjData(Object).def = (ObjData(Object).MinDef + ObjData(Object).MaxDef) / 2
    
        ObjData(Object).ClaseTipo = val(Leer.GetValue("OBJ" & Object, "ClaseTipo"))
        ObjData(Object).RazaTipo = val(Leer.GetValue("OBJ" & Object, "RazaTipo"))

        ObjData(Object).RazaEnana = val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
        ObjData(Object).RazaDrow = val(Leer.GetValue("OBJ" & Object, "RazaDrow"))
        ObjData(Object).RazaElfa = val(Leer.GetValue("OBJ" & Object, "RazaElfa"))
        ObjData(Object).RazaGnoma = val(Leer.GetValue("OBJ" & Object, "RazaGnoma"))
    
        ObjData(Object).RazaOrca = val(Leer.GetValue("OBJ" & Object, "RazaOrca"))
    
        ObjData(Object).RazaHumana = val(Leer.GetValue("OBJ" & Object, "RazaHumana"))
    
        ObjData(Object).Valor = val(Leer.GetValue("OBJ" & Object, "Valor"))
    
        ObjData(Object).Crucial = val(Leer.GetValue("OBJ" & Object, "Crucial"))
    
        ObjData(Object).Cerrada = val(Leer.GetValue("OBJ" & Object, "abierta"))

        If ObjData(Object).Cerrada = 1 Then
            ObjData(Object).Llave = val(Leer.GetValue("OBJ" & Object, "Llave"))
            ObjData(Object).clave = val(Leer.GetValue("OBJ" & Object, "Clave"))

        End If
    
        'Puertas y llaves
        ObjData(Object).clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
    
        ObjData(Object).texto = Leer.GetValue("OBJ" & Object, "Texto")
        ObjData(Object).GrhSecundario = val(Leer.GetValue("OBJ" & Object, "VGrande"))
    
        ObjData(Object).Agarrable = val(Leer.GetValue("OBJ" & Object, "Agarrable"))
        ObjData(Object).ForoID = Leer.GetValue("OBJ" & Object, "ID")
    
        'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico
        Dim n As Integer
        Dim S As String

        For i = 1 To NUMCLASES
            S = UCase$(Leer.GetValue("OBJ" & Object, "CP" & i))
            n = 1

            Do While LenB(S) > 0 And UCase$(ListaClases(n)) <> S
                n = n + 1
            Loop
            ObjData(Object).ClaseProhibida(i) = IIf(LenB(S) > 0, n, 0)
        Next i
        
        ' Skill requerido
        str = Leer.GetValue("OBJ" & Object, "SkillRequerido")

        If Len(str) > 0 Then
            Field = Split(str, "-")
            
            n = 1
            Do While LenB(Field(0)) > 0 And UCase$(Tilde(SkillsNames(n))) <> UCase$(Tilde(Field(0)))
                n = n + 1
            Loop
    
            ObjData(Object).SkillIndex = IIf(LenB(Field(0)) > 0, n, 0)
            ObjData(Object).SkillRequerido = val(Field(1))
        End If
        ' -----------------
    
        ObjData(Object).DefensaMagicaMax = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
        ObjData(Object).DefensaMagicaMin = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
    
        ObjData(Object).SkCarpinteria = val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
    
        'If ObjData(Object).SkCarpinteria > 0 Then
        ObjData(Object).Madera = val(Leer.GetValue("OBJ" & Object, "Madera"))
    
        'Bebidas
        ObjData(Object).MinSta = val(Leer.GetValue("OBJ" & Object, "MinST"))
    
        ObjData(Object).NoSeCae = val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
    
        frmCargando.cargar.Value = frmCargando.cargar.Value + 1
    Next Object

    Set Leer = Nothing

    Exit Sub

Errhandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.description & ". Error producido al cargar el objeto: " & Object

End Sub

Sub LoadUserStats(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)
        
        On Error GoTo LoadUserStats_Err
        

        Dim LoopC As Long

100     For LoopC = 1 To NUMATRIBUTOS
102         UserList(UserIndex).Stats.UserAtributos(LoopC) = CInt(UserFile.GetValue("ATRIBUTOS", "AT" & LoopC))
104         UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) = UserList(UserIndex).Stats.UserAtributos(LoopC)
106     Next LoopC

108     For LoopC = 1 To NUMSKILLS
110         UserList(UserIndex).Stats.UserSkills(LoopC) = CInt(UserFile.GetValue("SKILLS", "SK" & LoopC))
112     Next LoopC

114     For LoopC = 1 To MAXUSERHECHIZOS
116         UserList(UserIndex).Stats.UserHechizos(LoopC) = CInt(UserFile.GetValue("Hechizos", "H" & LoopC))
118     Next LoopC

120     UserList(UserIndex).Stats.GLD = CLng(UserFile.GetValue("STATS", "GLD"))
122     UserList(UserIndex).Stats.Banco = CLng(UserFile.GetValue("STATS", "BANCO"))

124     UserList(UserIndex).Stats.MaxHp = CInt(UserFile.GetValue("STATS", "MaxHP"))
126     UserList(UserIndex).Stats.MinHp = CInt(UserFile.GetValue("STATS", "MinHP"))

128     UserList(UserIndex).Stats.MinSta = CInt(UserFile.GetValue("STATS", "MinSTA"))
130     UserList(UserIndex).Stats.MaxSta = CInt(UserFile.GetValue("STATS", "MaxSTA"))

132     UserList(UserIndex).Stats.MaxMAN = CInt(UserFile.GetValue("STATS", "MaxMAN"))
134     UserList(UserIndex).Stats.MinMAN = CInt(UserFile.GetValue("STATS", "MinMAN"))

136     UserList(UserIndex).Stats.MaxHit = CInt(UserFile.GetValue("STATS", "MaxHIT"))
138     UserList(UserIndex).Stats.MinHIT = CInt(UserFile.GetValue("STATS", "MinHIT"))

140     UserList(UserIndex).Stats.MaxAGU = CByte(UserFile.GetValue("STATS", "MaxAGU"))
142     UserList(UserIndex).Stats.MinAGU = CByte(UserFile.GetValue("STATS", "MinAGU"))

144     UserList(UserIndex).Stats.MaxHam = CByte(UserFile.GetValue("STATS", "MaxHAM"))
146     UserList(UserIndex).Stats.MinHam = CByte(UserFile.GetValue("STATS", "MinHAM"))

148     UserList(UserIndex).Stats.SkillPts = CInt(UserFile.GetValue("STATS", "SkillPtsLibres"))

150     UserList(UserIndex).Stats.Exp = CDbl(UserFile.GetValue("STATS", "EXP"))
152     UserList(UserIndex).Stats.ELU = CLng(UserFile.GetValue("STATS", "ELU"))
154     UserList(UserIndex).Stats.ELV = CByte(UserFile.GetValue("STATS", "ELV"))

156     UserList(UserIndex).flags.Envenena = CByte(UserFile.GetValue("MAGIA", "ENVENENA"))
158     UserList(UserIndex).flags.Paraliza = CByte(UserFile.GetValue("MAGIA", "PARALIZA"))
160     UserList(UserIndex).flags.incinera = CByte(UserFile.GetValue("MAGIA", "INCINERA")) 'Estupidiza
162     UserList(UserIndex).flags.Estupidiza = CByte(UserFile.GetValue("MAGIA", "Estupidiza"))

164     UserList(UserIndex).flags.PendienteDelSacrificio = CByte(UserFile.GetValue("MAGIA", "PENDIENTE"))
166     UserList(UserIndex).flags.CarroMineria = CByte(UserFile.GetValue("MAGIA", "CarroMineria"))
168     UserList(UserIndex).flags.NoPalabrasMagicas = CByte(UserFile.GetValue("MAGIA", "NOPALABRASMAGICAS"))

170     If UserList(UserIndex).flags.Muerto = 0 Then
172         UserList(UserIndex).Char.Otra_Aura = CStr(UserFile.GetValue("MAGIA", "OTRA_AURA"))

        End If

174     UserList(UserIndex).flags.DañoMagico = CByte(UserFile.GetValue("MAGIA", "DañoMagico"))
176     UserList(UserIndex).flags.ResistenciaMagica = CByte(UserFile.GetValue("MAGIA", "ResistenciaMagica"))

        'Nuevos
178     UserList(UserIndex).flags.RegeneracionMana = CByte(UserFile.GetValue("MAGIA", "RegeneracionMana"))
180     UserList(UserIndex).flags.AnilloOcultismo = CByte(UserFile.GetValue("MAGIA", "AnilloOcultismo"))
182     UserList(UserIndex).flags.NoDetectable = CByte(UserFile.GetValue("MAGIA", "NoDetectable"))
184     UserList(UserIndex).flags.NoMagiaEfeceto = CByte(UserFile.GetValue("MAGIA", "NoMagiaEfeceto"))
186     UserList(UserIndex).flags.RegeneracionHP = CByte(UserFile.GetValue("MAGIA", "RegeneracionHP"))
188     UserList(UserIndex).flags.RegeneracionSta = CByte(UserFile.GetValue("MAGIA", "RegeneracionSta"))

190     UserList(UserIndex).Stats.UsuariosMatados = CLng(UserFile.GetValue("MUERTES", "UserMuertes"))
192     UserList(UserIndex).Stats.NPCsMuertos = CInt(UserFile.GetValue("MUERTES", "NpcsMuertes"))

194     UserList(UserIndex).Stats.InventLevel = CInt(UserFile.GetValue("STATS", "InventLevel"))

196     If CByte(UserFile.GetValue("CONSEJO", "PERTENECE")) Then UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.RoyalCouncil

198     If CByte(UserFile.GetValue("CONSEJO", "PERTENECECAOS")) Then UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.ChaosCouncil

        
        Exit Sub

LoadUserStats_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LoadUserStats", Erl)
        Resume Next
        
End Sub

Sub LoadUserInit(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)
        
        On Error GoTo LoadUserInit_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 19/11/2006
        'Loads the Users records
        '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
        '23/01/2007 Pablo (ToxicWaste) - Quito CriminalesMatados de Stats porque era redundante.
        '*************************************************
        Dim LoopC As Long

        Dim ln    As String

100     UserList(UserIndex).Faccion.Status = CByte(UserFile.GetValue("FACCIONES", "Status"))
102     UserList(UserIndex).Faccion.ArmadaReal = CByte(UserFile.GetValue("FACCIONES", "EjercitoReal"))
104     UserList(UserIndex).Faccion.FuerzasCaos = CByte(UserFile.GetValue("FACCIONES", "EjercitoCaos"))
106     UserList(UserIndex).Faccion.CiudadanosMatados = CLng(UserFile.GetValue("FACCIONES", "CiudMatados"))
108     UserList(UserIndex).Faccion.CriminalesMatados = CLng(UserFile.GetValue("FACCIONES", "CrimMatados"))
110     UserList(UserIndex).Faccion.RecibioArmaduraCaos = CByte(UserFile.GetValue("FACCIONES", "rArCaos"))
112     UserList(UserIndex).Faccion.RecibioArmaduraReal = CByte(UserFile.GetValue("FACCIONES", "rArReal"))
114     UserList(UserIndex).Faccion.RecibioExpInicialCaos = CByte(UserFile.GetValue("FACCIONES", "rExCaos"))
116     UserList(UserIndex).Faccion.RecibioExpInicialReal = CByte(UserFile.GetValue("FACCIONES", "rExReal"))
118     UserList(UserIndex).Faccion.RecompensasCaos = CLng(UserFile.GetValue("FACCIONES", "recCaos"))
120     UserList(UserIndex).Faccion.RecompensasReal = CLng(UserFile.GetValue("FACCIONES", "recReal"))
122     UserList(UserIndex).Faccion.Reenlistadas = CByte(UserFile.GetValue("FACCIONES", "Reenlistadas"))
124     UserList(UserIndex).Faccion.NivelIngreso = CInt(UserFile.GetValue("FACCIONES", "NivelIngreso"))
126     UserList(UserIndex).Faccion.FechaIngreso = UserFile.GetValue("FACCIONES", "FechaIngreso")
128     UserList(UserIndex).Faccion.MatadosIngreso = CInt(UserFile.GetValue("FACCIONES", "MatadosIngreso"))
130     UserList(UserIndex).Faccion.NextRecompensa = CInt(UserFile.GetValue("FACCIONES", "NextRecompensa"))

132     UserList(UserIndex).flags.Muerto = CByte(UserFile.GetValue("FLAGS", "Muerto"))
134     UserList(UserIndex).flags.Escondido = CByte(UserFile.GetValue("FLAGS", "Escondido"))

136     UserList(UserIndex).flags.Hambre = CByte(UserFile.GetValue("FLAGS", "Hambre"))
138     UserList(UserIndex).flags.Sed = CByte(UserFile.GetValue("FLAGS", "Sed"))
140     UserList(UserIndex).flags.Desnudo = CByte(UserFile.GetValue("FLAGS", "Desnudo"))
142     UserList(UserIndex).flags.Navegando = CByte(UserFile.GetValue("FLAGS", "Navegando"))
144     UserList(UserIndex).flags.Envenenado = CByte(UserFile.GetValue("FLAGS", "Envenenado"))
146     UserList(UserIndex).flags.Paralizado = CByte(UserFile.GetValue("FLAGS", "Paralizado"))
148     UserList(UserIndex).flags.Incinerado = CByte(UserFile.GetValue("FLAGS", "Incinerado"))
150     UserList(UserIndex).flags.Inmovilizado = CByte(UserFile.GetValue("FLAGS", "Inmovilizado"))

152     UserList(UserIndex).flags.ScrollExp = CSng(UserFile.GetValue("FLAGS", "ScrollExp"))
154     UserList(UserIndex).flags.ScrollOro = CSng(UserFile.GetValue("FLAGS", "ScrollOro"))

156     If UserList(UserIndex).flags.Paralizado = 1 Then
158         UserList(UserIndex).Counters.Paralisis = IntervaloParalizado

        End If

160     UserList(UserIndex).flags.BattlePuntos = CLng(UserFile.GetValue("Battle", "Puntos"))

162     If UserList(UserIndex).flags.Inmovilizado = 1 Then
164         UserList(UserIndex).Counters.Inmovilizado = 20

        End If

166     UserList(UserIndex).Counters.Pena = CLng(UserFile.GetValue("COUNTERS", "Pena"))

168     UserList(UserIndex).Counters.ScrollExperiencia = CLng(UserFile.GetValue("COUNTERS", "ScrollExperiencia"))
170     UserList(UserIndex).Counters.ScrollOro = CLng(UserFile.GetValue("COUNTERS", "ScrollOro"))

172     UserList(UserIndex).Counters.Oxigeno = CLng(UserFile.GetValue("COUNTERS", "Oxigeno"))

174     UserList(UserIndex).MENSAJEINFORMACION = UserFile.GetValue("INIT", "MENSAJEINFORMACION")

176     UserList(UserIndex).genero = UserFile.GetValue("INIT", "Genero")
178     UserList(UserIndex).clase = UserFile.GetValue("INIT", "Clase")
180     UserList(UserIndex).raza = UserFile.GetValue("INIT", "Raza")
182     UserList(UserIndex).Hogar = UserFile.GetValue("INIT", "Hogar")
184     UserList(UserIndex).Char.heading = CInt(UserFile.GetValue("INIT", "Heading"))

186     UserList(UserIndex).OrigChar.Head = CInt(UserFile.GetValue("INIT", "Head"))
188     UserList(UserIndex).OrigChar.Body = CInt(UserFile.GetValue("INIT", "Body"))
190     UserList(UserIndex).OrigChar.WeaponAnim = CInt(UserFile.GetValue("INIT", "Arma"))
192     UserList(UserIndex).OrigChar.ShieldAnim = CInt(UserFile.GetValue("INIT", "Escudo"))
194     UserList(UserIndex).OrigChar.CascoAnim = CInt(UserFile.GetValue("INIT", "Casco"))

        #If ConUpTime Then
196         UserList(UserIndex).UpTime = CLng(UserFile.GetValue("INIT", "UpTime"))
        #End If

198     UserList(UserIndex).OrigChar.heading = UserList(UserIndex).Char.heading

200     If UserList(UserIndex).flags.Muerto = 0 Then
202         UserList(UserIndex).Char = UserList(UserIndex).OrigChar
        Else
204         UserList(UserIndex).Char.Body = iCuerpoMuerto
206         UserList(UserIndex).Char.Head = iCabezaMuerto
208         UserList(UserIndex).Char.WeaponAnim = NingunArma
210         UserList(UserIndex).Char.ShieldAnim = NingunEscudo
212         UserList(UserIndex).Char.CascoAnim = NingunCasco

        End If

214     UserList(UserIndex).Desc = UserFile.GetValue("INIT", "Desc")

216     UserList(UserIndex).flags.BanMotivo = UserFile.GetValue("BAN", "BanMotivo")
218     UserList(UserIndex).flags.Montado = CByte(UserFile.GetValue("FLAGS", "Montado"))
220     UserList(UserIndex).flags.VecesQueMoriste = CLng(UserFile.GetValue("FLAGS", "VecesQueMoriste"))

222     UserList(UserIndex).flags.MinutosRestantes = CLng(UserFile.GetValue("FLAGS", "MinutosRestantes"))
224     UserList(UserIndex).flags.Silenciado = CLng(UserFile.GetValue("FLAGS", "Silenciado"))
226     UserList(UserIndex).flags.SegundosPasados = CLng(UserFile.GetValue("FLAGS", "SegundosPasados"))

        'CASAMIENTO LADDER
228     UserList(UserIndex).flags.Casado = CInt(UserFile.GetValue("FLAGS", "CASADO"))
230     UserList(UserIndex).flags.Pareja = UserFile.GetValue("FLAGS", "PAREJA")

232     UserList(UserIndex).Pos.Map = CInt(ReadField(1, UserFile.GetValue("INIT", "Position"), 45))
234     UserList(UserIndex).Pos.x = CInt(ReadField(2, UserFile.GetValue("INIT", "Position"), 45))
236     UserList(UserIndex).Pos.Y = CInt(ReadField(3, UserFile.GetValue("INIT", "Position"), 45))

238     UserList(UserIndex).Invent.NroItems = CInt(UserFile.GetValue("Inventory", "CantidadItems"))

        '[KEVIN]--------------------------------------------------------------------
        '***********************************************************************************
240     UserList(UserIndex).BancoInvent.NroItems = CInt(UserFile.GetValue("BancoInventory", "CantidadItems"))

        'Lista de objetos del banco
242     For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
244         ln = UserFile.GetValue("BancoInventory", "Obj" & LoopC)
246         UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
248         UserList(UserIndex).BancoInvent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
250     Next LoopC

        '------------------------------------------------------------------------------------
        '[/KEVIN]*****************************************************************************

        'Lista de objetos
252     For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots
254         ln = UserFile.GetValue("Inventory", "Obj" & LoopC)
256         UserList(UserIndex).Invent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
258         UserList(UserIndex).Invent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
260         UserList(UserIndex).Invent.Object(LoopC).Equipped = CByte(ReadField(3, ln, 45))
262     Next LoopC

264     UserList(UserIndex).Invent.WeaponEqpSlot = CByte(UserFile.GetValue("Inventory", "WeaponEqpSlot"))
266     UserList(UserIndex).Invent.HerramientaEqpSlot = CByte(UserFile.GetValue("Inventory", "HerramientaEqpSlot"))
268     UserList(UserIndex).Invent.ArmourEqpSlot = CByte(UserFile.GetValue("Inventory", "ArmourEqpSlot"))
270     UserList(UserIndex).Invent.EscudoEqpSlot = CByte(UserFile.GetValue("Inventory", "EscudoEqpSlot"))
272     UserList(UserIndex).Invent.CascoEqpSlot = CByte(UserFile.GetValue("Inventory", "CascoEqpSlot"))
274     UserList(UserIndex).Invent.BarcoSlot = CByte(UserFile.GetValue("Inventory", "BarcoSlot"))
276     UserList(UserIndex).Invent.MonturaSlot = CByte(UserFile.GetValue("Inventory", "MonturaSlot"))
278     UserList(UserIndex).Invent.MunicionEqpSlot = CByte(UserFile.GetValue("Inventory", "MunicionSlot"))
280     UserList(UserIndex).Invent.AnilloEqpSlot = CByte(UserFile.GetValue("Inventory", "AnilloSlot"))
282     UserList(UserIndex).Invent.MagicoSlot = CByte(UserFile.GetValue("Inventory", "MagicoSlot"))
284     UserList(UserIndex).Invent.NudilloSlot = CByte(UserFile.GetValue("Inventory", "NudilloEqpSlot"))

286     UserList(UserIndex).ChatCombate = CByte(UserFile.GetValue("BINDKEYS", "ChatCombate"))
288     UserList(UserIndex).ChatGlobal = CByte(UserFile.GetValue("BINDKEYS", "ChatGlobal"))

290     UserList(UserIndex).Correo.CantCorreo = CByte(UserFile.GetValue("CORREO", "CantCorreo"))
292     UserList(UserIndex).Correo.NoLeidos = CByte(UserFile.GetValue("CORREO", "NoLeidos"))

294     For LoopC = 1 To UserList(UserIndex).Correo.CantCorreo
296         UserList(UserIndex).Correo.Mensaje(LoopC).Remitente = UserFile.GetValue("CORREO", "REMITENTE" & LoopC)
298         UserList(UserIndex).Correo.Mensaje(LoopC).Mensaje = UserFile.GetValue("CORREO", "MENSAJE" & LoopC)
300         UserList(UserIndex).Correo.Mensaje(LoopC).Item = UserFile.GetValue("CORREO", "Item" & LoopC)
302         UserList(UserIndex).Correo.Mensaje(LoopC).ItemCount = CByte(UserFile.GetValue("CORREO", "ItemCount" & LoopC))
304         UserList(UserIndex).Correo.Mensaje(LoopC).Fecha = UserFile.GetValue("CORREO", "DATE" & LoopC)
306         UserList(UserIndex).Correo.Mensaje(LoopC).Leido = CByte(UserFile.GetValue("CORREO", "LEIDO" & LoopC))
308     Next LoopC

        'Logros Ladder
310     UserList(UserIndex).UserLogros = UserFile.GetValue("LOGROS", "UserLogros")
312     UserList(UserIndex).NPcLogros = UserFile.GetValue("LOGROS", "NPcLogros")
314     UserList(UserIndex).LevelLogros = UserFile.GetValue("LOGROS", "LevelLogros")
        'Logros Ladder

316     ln = UserFile.GetValue("Guild", "GUILDINDEX")

318     If IsNumeric(ln) Then
320         UserList(UserIndex).GuildIndex = CInt(ln)
        Else
322         UserList(UserIndex).GuildIndex = 0

        End If

        
        Exit Sub

LoadUserInit_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LoadUserInit", Erl)
        Resume Next
        
End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String
        
        On Error GoTo GetVar_Err
        

        Dim sSpaces  As String ' This will hold the input that the program will retrieve

        Dim szReturn As String ' This will be the defaul value if the string is not found
  
100     szReturn = vbNullString
  
102     sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
  
104     GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, File
  
106     GetVar = RTrim$(sSpaces)
108     GetVar = Left$(GetVar, Len(GetVar) - 1)
  
        
        Exit Function

GetVar_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.GetVar", Erl)
        Resume Next
        
End Function

Sub CargarBackUp()
        
        On Error GoTo CargarBackUp_Err
        

100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."

        Dim Map       As Integer

        Dim TempInt   As Integer

        Dim tFileName As String

        Dim npcfile   As String
    
102     NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
104     Call InitAreas
    
106     frmCargando.cargar.min = 0
108     frmCargando.cargar.max = NumMaps
110     frmCargando.cargar.Value = 0
112     frmCargando.ToMapLbl.Visible = True
        ' MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
114     ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

116     ReDim MapInfo(1 To NumMaps) As MapInfo
      
118     For Map = 1 To NumMaps
120         frmCargando.ToMapLbl = Map & "/" & NumMaps
            Rem If val(GetVar(App.Path & MapPath & "Mapa" & map & ".Dat", "Mapa" & map, "BackUp")) <> 0 Then
122         tFileName = App.Path & "\WorldBackUp\Mapa" & Map & ".csm"
            Rem Else
            Rem     tFileName = App.Path & MapPath & "Mapa" & map
            Rem End If

124         Call CargarMapaFormatoCSM(Map, tFileName)
126         frmCargando.cargar.Value = frmCargando.cargar.Value + 1
128         DoEvents
130     Next Map

132     frmCargando.ToMapLbl.Visible = False
        Exit Sub

        
        Exit Sub

CargarBackUp_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.CargarBackUp", Erl)
        Resume Next
        
End Sub

Sub LoadMapData()

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas..."

    Dim Map       As Integer

    Dim TempInt   As Integer

    Dim tFileName As String

    Dim npcfile   As String

    On Error GoTo man
    
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.Value = 0
    frmCargando.ToMapLbl.Visible = True
    
    'MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

    ReDim MapInfo(1 To NumMaps) As MapInfo

    For Map = 1 To NumMaps
        frmCargando.ToMapLbl = Map & "/" & NumMaps
        tFileName = MapPath & "Mapa" & Map & ".csm"
 
        Call CargarMapaFormatoCSM(Map, tFileName)
        Rem Call Load_Map_Data_CSM(map, tFileName)
        frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        DoEvents
    Next Map

    frmCargando.ToMapLbl.Visible = False
    Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)

End Sub

Public Sub CargarMapaFormatoCSM(ByVal Map As Long, ByVal MAPFl As String)

    On Error GoTo errh:

    Dim npcfile      As String

    Dim fh           As Integer

    Dim MH           As tMapHeader

    Dim Blqs()       As tDatosBloqueados

    Dim L1()         As tDatosGrh

    Dim L2()         As tDatosGrh

    Dim L3()         As tDatosGrh

    Dim L4()         As tDatosGrh

    Dim Triggers()   As tDatosTrigger

    Dim Luces()      As tDatosLuces

    Dim Particulas() As tDatosParticulas

    Dim Objetos()    As tDatosObjs

    Dim NPCs()       As tDatosNPC

    Dim TEs()        As tDatosTE

    Dim Body         As Integer

    Dim Head         As Integer

    Dim heading      As Byte

    Dim i            As Long

    Dim j            As Long

    fh = FreeFile
    Open MAPFl For Binary As fh
    Get #fh, , MH
    Get #fh, , MapSize
    Get #fh, , MapDat

    Rem Get #fh, , L1

    With MH

        'Cargamos Bloqueos
        
        If .NumeroBloqueados > 0 Then

            ReDim Blqs(1 To .NumeroBloqueados)
            Get #fh, , Blqs

            For i = 1 To .NumeroBloqueados

                MapData(Map, Blqs(i).x, Blqs(i).Y).Blocked = 1
            Next i

        End If
        
        'Cargamos Layer 1
        
        If .NumeroLayers(1) > 0 Then
        
            ReDim L1(1 To .NumeroLayers(1))
            Get #fh, , L1

            For i = 1 To .NumeroLayers(1)
                        
                MapData(Map, L1(i).x, L1(i).Y).Graphic(1) = L1(i).GrhIndex
            
                'InitGrh MapData(L1(i).X, L1(i).Y).Graphic(1), MapData(L1(i).X, L1(i).Y).Graphic(1).GrhIndex
                ' Call Map_Grh_Set(L2(i).X, L2(i).Y, L2(i).GrhIndex, 2)
            Next i

        End If
        
        'Cargamos Layer 2
        If .NumeroLayers(2) > 0 Then
            ReDim L2(1 To .NumeroLayers(2))
            Get #fh, , L2

            For i = 1 To .NumeroLayers(2)
                MapData(Map, L2(i).x, L2(i).Y).Graphic(2) = L2(i).GrhIndex
            Next i

        End If
                
        If .NumeroLayers(3) > 0 Then
            ReDim L3(1 To .NumeroLayers(3))
            Get #fh, , L3

            For i = 1 To .NumeroLayers(3)
                MapData(Map, L3(i).x, L3(i).Y).Graphic(3) = L3(i).GrhIndex
            Next i

        End If
        
        If .NumeroLayers(4) > 0 Then
            ReDim L4(1 To .NumeroLayers(4))
            Get #fh, , L4

            For i = 1 To .NumeroLayers(4)
                MapData(Map, L4(i).x, L4(i).Y).Graphic(4) = L4(i).GrhIndex
            Next i

        End If

        If .NumeroTriggers > 0 Then
            ReDim Triggers(1 To .NumeroTriggers)
            Get #fh, , Triggers

            For i = 1 To .NumeroTriggers
                MapData(Map, Triggers(i).x, Triggers(i).Y).trigger = Triggers(i).trigger
            Next i

        End If

        If .NumeroParticulas > 0 Then
            ReDim Particulas(1 To .NumeroParticulas)
            Get #fh, , Particulas

            For i = 1 To .NumeroParticulas
                MapData(Map, Particulas(i).x, Particulas(i).Y).ParticulaIndex = Particulas(i).Particula
                MapData(Map, Particulas(i).x, Particulas(i).Y).ParticulaIndex = 0
            Next i

        End If

        If .NumeroLuces > 0 Then
            ReDim Luces(1 To .NumeroLuces)
            Get #fh, , Luces

            For i = 1 To .NumeroLuces
                MapData(Map, Luces(i).x, Luces(i).Y).Luz.Color = Luces(i).Color
                MapData(Map, Luces(i).x, Luces(i).Y).Luz.Rango = Luces(i).Rango
                MapData(Map, Luces(i).x, Luces(i).Y).Luz.Color = 0
                MapData(Map, Luces(i).x, Luces(i).Y).Luz.Rango = 0
            Next i

        End If
            
        If .NumeroOBJs > 0 Then
            ReDim Objetos(1 To .NumeroOBJs)
            Get #fh, , Objetos

            For i = 1 To .NumeroOBJs
                MapData(Map, Objetos(i).x, Objetos(i).Y).ObjInfo.ObjIndex = Objetos(i).ObjIndex

                Select Case ObjData(Objetos(i).ObjIndex).OBJType

                    Case eOBJType.otYacimiento, eOBJType.otArboles
                        MapData(Map, Objetos(i).x, Objetos(i).Y).ObjInfo.Amount = ObjData(Objetos(i).ObjIndex).VidaUtil
                        MapData(Map, Objetos(i).x, Objetos(i).Y).ObjInfo.data = &H7FFFFFFF ' Ultimo uso = Max Long

                    Case Else
                        MapData(Map, Objetos(i).x, Objetos(i).Y).ObjInfo.Amount = Objetos(i).ObjAmmount

                End Select

            Next i

        End If

        If .NumeroNPCs > 0 Then
            ReDim NPCs(1 To .NumeroNPCs)
            Get #fh, , NPCs
                 
            For i = 1 To .NumeroNPCs

                MapData(Map, NPCs(i).x, NPCs(i).Y).NpcIndex = NPCs(i).NpcIndex
                    
                If MapData(Map, NPCs(i).x, NPCs(i).Y).NpcIndex > 0 Then
                    npcfile = DatPath & "NPCs.dat"

                    'Si el npc debe hacer respawn en la pos
                    'original la guardamos
                    If val(GetVar(npcfile, "NPC" & MapData(Map, NPCs(i).x, NPCs(i).Y).NpcIndex, "PosOrig")) = 1 Then
                        MapData(Map, NPCs(i).x, NPCs(i).Y).NpcIndex = OpenNPC(MapData(Map, NPCs(i).x, NPCs(i).Y).NpcIndex)
                        Npclist(MapData(Map, NPCs(i).x, NPCs(i).Y).NpcIndex).Orig.Map = Map
                        Npclist(MapData(Map, NPCs(i).x, NPCs(i).Y).NpcIndex).Orig.x = NPCs(i).x
                        Npclist(MapData(Map, NPCs(i).x, NPCs(i).Y).NpcIndex).Orig.Y = NPCs(i).Y
                    Else
                        MapData(Map, NPCs(i).x, NPCs(i).Y).NpcIndex = OpenNPC(MapData(Map, NPCs(i).x, NPCs(i).Y).NpcIndex)

                    End If

                    Npclist(MapData(Map, NPCs(i).x, NPCs(i).Y).NpcIndex).Pos.Map = Map
                    Npclist(MapData(Map, NPCs(i).x, NPCs(i).Y).NpcIndex).Pos.x = NPCs(i).x
                    Npclist(MapData(Map, NPCs(i).x, NPCs(i).Y).NpcIndex).Pos.Y = NPCs(i).Y
                        
                    '        If NPCs(i).NpcIndex > 499 Then
                                            
                    '                                           Dim nfile As Integer
                    '  nfile = FreeFile ' obtenemos un canal
                    '  Open App.Path & "\logs\npcs.log" For Append Shared As #nfile
                    ' Print #nfile, NPCs(i).NpcIndex & "(" & Npclist(MapData(Map, NPCs(i).x, NPCs(i).y).NpcIndex).Name & ") "
                    ' Close #nfile
                                            
                    ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "Nombre", Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Name
                    ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "MaxHp", Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Stats.MaxHp
                    ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "GiveEXP", Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).GiveEXP
                    ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "GiveGLD", Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).GiveGLD
                    ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "MinHIT", Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Stats.MinHIT
                    ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "MaxHit", Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Stats.MaxHit
                    ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "def", Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Stats.def
                    ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "defM", Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Stats.defM
                    ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "PoderAtaque", Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).PoderAtaque
                    ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "PoderEvasion", Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).PoderEvasion
                    ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "Posicion" & i, Map & "-" & NPCs(i).X & "-" & NPCs(i).Y
                    'End If
                            
                    If Npclist(MapData(Map, NPCs(i).x, NPCs(i).Y).NpcIndex).name = "" Then
                       
                        MapData(Map, NPCs(i).x, NPCs(i).Y).NpcIndex = 0
                    Else
                        
                        Call MakeNPCChar(True, 0, MapData(Map, NPCs(i).x, NPCs(i).Y).NpcIndex, Map, NPCs(i).x, NPCs(i).Y)
                        
                    End If

                End If

            Next i
                
        End If
            
        If .NumeroTE > 0 Then
            ReDim TEs(1 To .NumeroTE)
            Get #fh, , TEs

            For i = 1 To .NumeroTE
                MapData(Map, TEs(i).x, TEs(i).Y).TileExit.Map = TEs(i).DestM
                MapData(Map, TEs(i).x, TEs(i).Y).TileExit.x = TEs(i).DestX
                MapData(Map, TEs(i).x, TEs(i).Y).TileExit.Y = TEs(i).DestY
            Next i

        End If
        
    End With

    Close fh
    
    MapInfo(Map).map_name = MapDat.map_name
    MapInfo(Map).ambient = MapDat.ambient
    MapInfo(Map).backup_mode = MapDat.backup_mode
    MapInfo(Map).base_light = MapDat.base_light
    MapInfo(Map).extra1 = MapDat.extra1
    'MapInfo(Map).extra2 = MapDat.extra2
    MapInfo(Map).extra2 = val(GetVar(DatPath & "Map.dat", "RESUCIUDAD", Map))
    
    MapInfo(Map).extra3 = MapDat.extra3
    MapInfo(Map).letter_grh = MapDat.letter_grh
    MapInfo(Map).lluvia = MapDat.lluvia
    MapInfo(Map).music_numberHi = MapDat.music_numberHi
    MapInfo(Map).music_numberLow = MapDat.music_numberLow
    MapInfo(Map).niebla = MapDat.niebla
    MapInfo(Map).Nieve = MapDat.Nieve
    MapInfo(Map).restrict_mode = MapDat.restrict_mode
    
    MapInfo(Map).Seguro = MapDat.Seguro

    MapInfo(Map).terrain = MapDat.terrain
    MapInfo(Map).zone = MapDat.zone
 
    Exit Sub
    Rem Dim N

    Rem N = FreeFile
    Rem Open App.Path & "\NameMapas.ini" For Binary Access Write As N

    Rem Put N, , "[Mapas]" & vbCrLf
    Rem Put N, , map & "=" & MapInfo(map).map_name & vbCrLf
    Rem Put N, , vbCrLf
    Rem Close #N
errh:
    MsgBox ("Error cargando mapa: " & Map & "." & Err.description & " - ")

    Rem    Call LogError("Error cargando mapa: " & map & "." & Err.description)
End Sub

Sub LoadSini()
        
        On Error GoTo LoadSini_Err
        

        Dim Lector   As clsIniReader

        Dim Temporal As Long
    
100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."
    
102     Set Lector = New clsIniReader
104     Call Lector.Initialize(IniPath & "Server.ini")
    
        'Misc
106     BootDelBackUp = val(Lector.GetValue("INIT", "IniciarDesdeBackUp"))
    
        'Directorios
108     DatPath = Lector.GetValue("DIRECTORIOS", "DatPath")
110     MapPath = Lector.GetValue("DIRECTORIOS", "MapPath")
112     CharPath = Lector.GetValue("DIRECTORIOS", "CharPath")
114     DeletePath = Lector.GetValue("DIRECTORIOS", "DeletePath")
116     CuentasPath = Lector.GetValue("DIRECTORIOS", "CuentasPath")
118     DeleteCuentasPath = Lector.GetValue("DIRECTORIOS", "DeleteCuentasPath")
        'Directorios
    
120     Puerto = val(Lector.GetValue("INIT", "StartPort"))
122     LastSockListen = val(Lector.GetValue("INIT", "LastSockListen"))
124     HideMe = val(Lector.GetValue("INIT", "Hide"))
126     MaxConexionesIP = val(Lector.GetValue("INIT", "MaxConexionesIP"))
128     MaxUsersPorCuenta = val(Lector.GetValue("INIT", "MaxUsersPorCuenta"))
130     IdleLimit = val(Lector.GetValue("INIT", "IdleLimit"))
        'Lee la version correcta del cliente
132     ULTIMAVERSION = Lector.GetValue("INIT", "Version")
    
134     PuedeCrearPersonajes = val(Lector.GetValue("INIT", "PuedeCrearPersonajes"))
136     ServerSoloGMs = val(Lector.GetValue("init", "ServerSoloGMs"))
    
138     DiceMinimum = val(Lector.GetValue("INIT", "MinDados"))
140     DiceMaximum = val(Lector.GetValue("INIT", "MaxDados"))
    
142     EnTesting = val(Lector.GetValue("INIT", "Testing"))
    
        ' Database
144     Database_Enabled = CBool(val(Lector.GetValue("DATABASE", "Enabled")))
146     Database_DataSource = Lector.GetValue("DATABASE", "DSN")
148     Database_Host = Lector.GetValue("DATABASE", "Host")
150     Database_Name = Lector.GetValue("DATABASE", "Name")
152     Database_Username = Lector.GetValue("DATABASE", "Username")
154     Database_Password = Lector.GetValue("DATABASE", "Password")
    
        'Ressurect pos
156     ResPos.Map = val(ReadField(1, Lector.GetValue("INIT", "ResPos"), 45))
158     ResPos.x = val(ReadField(2, Lector.GetValue("INIT", "ResPos"), 45))
160     ResPos.Y = val(ReadField(3, Lector.GetValue("INIT", "ResPos"), 45))
      
        If Not Database_Enabled Then
162         RecordUsuarios = val(Lector.GetValue("INIT", "Record"))
        End If
      
        'Max users
164     Temporal = val(Lector.GetValue("INIT", "MaxUsers"))

166     If MaxUsers = 0 Then
168         MaxUsers = Temporal
170         ReDim UserList(1 To MaxUsers) As user

        End If

172     NumCuentas = val(Lector.GetValue("INIT", "NumCuentas"))
174     frmMain.cuentas.Caption = NumCuentas
        #If DEBUGGING Then
            'Shell App.Path & "\estadisticas.exe" & " " & "NUEVACUENTALADDER" & "*" & NumCuentas & "*" & MaxUsers
        #End If
    
        '&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
        'Se agregó en LoadBalance y en el Balance.dat
        'PorcentajeRecuperoMana = val(Lector.GetValue("BALANCE", "PorcentajeRecuperoMana"))
    
        ''&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&
        'Call Statistics.Initialize
    
176     Nix.Map = GetVar(DatPath & "Ciudades.dat", "NIX", "Mapa")
178     Nix.x = GetVar(DatPath & "Ciudades.dat", "NIX", "X")
180     Nix.Y = GetVar(DatPath & "Ciudades.dat", "NIX", "Y")
    
182     Ullathorpe.Map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
184     Ullathorpe.x = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
186     Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")
    
188     Banderbill.Map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
190     Banderbill.x = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
192     Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")
    
194     Lindos.Map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
196     Lindos.x = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
198     Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")
    
200     Arghal.Map = GetVar(DatPath & "Ciudades.dat", "Arghal", "Mapa")
202     Arghal.x = GetVar(DatPath & "Ciudades.dat", "Arghal", "X")
204     Arghal.Y = GetVar(DatPath & "Ciudades.dat", "Arghal", "Y")
    
206     Hillidan.Map = GetVar(DatPath & "Ciudades.dat", "Hillidan", "Mapa")
208     Hillidan.x = GetVar(DatPath & "Ciudades.dat", "Hillidan", "X")
210     Hillidan.Y = GetVar(DatPath & "Ciudades.dat", "Hillidan", "Y")

212     CityNix.Map = GetVar(DatPath & "Ciudades.dat", "NIX", "Mapa")
214     CityNix.x = GetVar(DatPath & "Ciudades.dat", "NIX", "X")
216     CityNix.Y = GetVar(DatPath & "Ciudades.dat", "NIX", "Y")
218     CityNix.MapaViaje = GetVar(DatPath & "Ciudades.dat", "NIX", "MapaViaje")
220     CityNix.ViajeX = GetVar(DatPath & "Ciudades.dat", "NIX", "ViajeX")
222     CityNix.ViajeY = GetVar(DatPath & "Ciudades.dat", "NIX", "ViajeY")
224     CityNix.MapaResu = GetVar(DatPath & "Ciudades.dat", "NIX", "MapaResu")
226     CityNix.ResuX = GetVar(DatPath & "Ciudades.dat", "NIX", "ResuX")
228     CityNix.ResuY = GetVar(DatPath & "Ciudades.dat", "NIX", "ResuY")
230     CityNix.NecesitaNave = GetVar(DatPath & "Ciudades.dat", "NIX", "NecesitaNave")

232     CityUllathorpe.Map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
234     CityUllathorpe.x = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
236     CityUllathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")
238     CityUllathorpe.MapaViaje = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "MapaViaje")
240     CityUllathorpe.ViajeX = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "ViajeX")
242     CityUllathorpe.ViajeY = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "ViajeY")
244     CityUllathorpe.MapaResu = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "MapaResu")
246     CityUllathorpe.ResuX = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "ResuX")
248     CityUllathorpe.ResuY = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "ResuY")
250     CityUllathorpe.NecesitaNave = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "NecesitaNave")
    
252     CityBanderbill.Map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
254     CityBanderbill.x = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
256     CityBanderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")
258     CityBanderbill.MapaViaje = GetVar(DatPath & "Ciudades.dat", "Banderbill", "MapaViaje")
260     CityBanderbill.ViajeX = GetVar(DatPath & "Ciudades.dat", "Banderbill", "ViajeX")
262     CityBanderbill.ViajeY = GetVar(DatPath & "Ciudades.dat", "Banderbill", "ViajeY")
264     CityBanderbill.MapaResu = GetVar(DatPath & "Ciudades.dat", "Banderbill", "MapaResu")
266     CityBanderbill.ResuX = GetVar(DatPath & "Ciudades.dat", "Banderbill", "ResuX")
268     CityBanderbill.ResuY = GetVar(DatPath & "Ciudades.dat", "Banderbill", "ResuY")
270     CityBanderbill.NecesitaNave = GetVar(DatPath & "Ciudades.dat", "Banderbill", "NecesitaNave")
    
272     CityLindos.Map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
274     CityLindos.x = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
276     CityLindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")
278     CityLindos.MapaViaje = GetVar(DatPath & "Ciudades.dat", "Lindos", "MapaViaje")
280     CityLindos.ViajeX = GetVar(DatPath & "Ciudades.dat", "Lindos", "ViajeX")
282     CityLindos.ViajeY = GetVar(DatPath & "Ciudades.dat", "Lindos", "ViajeY")
284     CityLindos.MapaResu = GetVar(DatPath & "Ciudades.dat", "Lindos", "MapaResu")
286     CityLindos.ResuX = GetVar(DatPath & "Ciudades.dat", "Lindos", "ResuX")
288     CityLindos.ResuY = GetVar(DatPath & "Ciudades.dat", "Lindos", "ResuY")
290     CityLindos.NecesitaNave = GetVar(DatPath & "Ciudades.dat", "Lindos", "NecesitaNave")
    
292     CityArghal.Map = GetVar(DatPath & "Ciudades.dat", "Arghal", "Mapa")
294     CityArghal.x = GetVar(DatPath & "Ciudades.dat", "Arghal", "X")
296     CityArghal.Y = GetVar(DatPath & "Ciudades.dat", "Arghal", "Y")
298     CityArghal.MapaViaje = GetVar(DatPath & "Ciudades.dat", "Arghal", "MapaViaje")
300     CityArghal.ViajeX = GetVar(DatPath & "Ciudades.dat", "Arghal", "ViajeX")
302     CityArghal.ViajeY = GetVar(DatPath & "Ciudades.dat", "Arghal", "ViajeY")
304     CityArghal.MapaResu = GetVar(DatPath & "Ciudades.dat", "Arghal", "MapaResu")
306     CityArghal.ResuX = GetVar(DatPath & "Ciudades.dat", "Arghal", "ResuX")
308     CityArghal.ResuY = GetVar(DatPath & "Ciudades.dat", "Arghal", "ResuY")
310     CityArghal.NecesitaNave = GetVar(DatPath & "Ciudades.dat", "Arghal", "NecesitaNave")
    
312     CityHillidan.Map = GetVar(DatPath & "Ciudades.dat", "Hillidan", "Mapa")
314     CityHillidan.x = GetVar(DatPath & "Ciudades.dat", "Hillidan", "X")
316     CityHillidan.Y = GetVar(DatPath & "Ciudades.dat", "Hillidan", "Y")
318     CityHillidan.MapaViaje = GetVar(DatPath & "Ciudades.dat", "Hillidan", "MapaViaje")
320     CityHillidan.ViajeX = GetVar(DatPath & "Ciudades.dat", "Hillidan", "ViajeX")
322     CityHillidan.ViajeY = GetVar(DatPath & "Ciudades.dat", "Hillidan", "ViajeY")
324     CityHillidan.MapaResu = GetVar(DatPath & "Ciudades.dat", "Hillidan", "MapaResu")
326     CityHillidan.ResuX = GetVar(DatPath & "Ciudades.dat", "Hillidan", "ResuX")
328     CityHillidan.ResuY = GetVar(DatPath & "Ciudades.dat", "Hillidan", "ResuY")
330     CityHillidan.NecesitaNave = GetVar(DatPath & "Ciudades.dat", "Hillidan", "NecesitaNave")
    
332     Call MD5sCarga
    
334     Call ConsultaPopular.LoadData
    
336     Set Lector = Nothing

        
        Exit Sub

LoadSini_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LoadSini", Erl)
        Resume Next
        
End Sub

Sub LoadIntervalos()
        
        On Error GoTo LoadIntervalos_Err
        

        Dim Lector As clsIniReader

100     Set Lector = New clsIniReader
102     Call Lector.Initialize(IniPath & "intervalos.ini")
    
        'Intervalos
104     SanaIntervaloSinDescansar = val(Lector.GetValue("INTERVALOS", "SanaIntervaloSinDescansar"))
106     FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar
    
108     StaminaIntervaloSinDescansar = val(Lector.GetValue("INTERVALOS", "StaminaIntervaloSinDescansar"))
110     FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar
    
112     SanaIntervaloDescansar = val(Lector.GetValue("INTERVALOS", "SanaIntervaloDescansar"))
114     FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar
    
116     StaminaIntervaloDescansar = val(Lector.GetValue("INTERVALOS", "StaminaIntervaloDescansar"))
118     FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar
    
120     IntervaloSed = val(Lector.GetValue("INTERVALOS", "IntervaloSed"))
122     FrmInterv.txtIntervaloSed.Text = IntervaloSed
    
124     IntervaloHambre = val(Lector.GetValue("INTERVALOS", "IntervaloHambre"))
126     FrmInterv.txtIntervaloHambre.Text = IntervaloHambre
    
128     IntervaloVeneno = val(Lector.GetValue("INTERVALOS", "IntervaloVeneno"))
130     FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno
    
132     IntervaloParalizado = val(Lector.GetValue("INTERVALOS", "IntervaloParalizado"))
134     FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado
    
136     IntervaloInmovilizado = val(Lector.GetValue("INTERVALOS", "IntervaloInmovilizado"))
138     FrmInterv.txtIntervaloInmovilizado.Text = IntervaloInmovilizado
    
140     IntervaloInvisible = val(Lector.GetValue("INTERVALOS", "IntervaloInvisible"))
142     FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible
    
144     IntervaloFrio = val(Lector.GetValue("INTERVALOS", "IntervaloFrio"))
146     FrmInterv.txtIntervaloFrio.Text = IntervaloFrio
    
148     IntervaloWavFx = val(Lector.GetValue("INTERVALOS", "IntervaloWAVFX"))
150     FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx
    
152     IntervaloInvocacion = val(Lector.GetValue("INTERVALOS", "IntervaloInvocacion"))
154     FrmInterv.txtInvocacion.Text = IntervaloInvocacion
    
156     TimeoutPrimerPaquete = val(Lector.GetValue("INTERVALOS", "TimeoutPrimerPaquete"))
158     FrmInterv.txtTimeoutPrimerPaquete.Text = TimeoutPrimerPaquete
    
160     TimeoutEsperandoLoggear = val(Lector.GetValue("INTERVALOS", "TimeoutEsperandoLoggear"))
162     FrmInterv.txtTimeoutEsperandoLoggear.Text = TimeoutEsperandoLoggear
    
164     IntervaloIncineracion = val(Lector.GetValue("INTERVALOS", "IntervaloFuego"))
166     FrmInterv.txtintervalofuego.Text = IntervaloIncineracion
    
168     IntervaloTirar = val(Lector.GetValue("INTERVALOS", "IntervaloTirar"))
170     FrmInterv.txtintervalotirar.Text = IntervaloTirar
    
172     IntervaloCaminar = val(Lector.GetValue("INTERVALOS", "IntervaloCaminar"))
174     FrmInterv.txtintervalocaminar.Text = IntervaloCaminar
        'Ladder
    
        '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    
176     IntervaloUserPuedeCastear = val(Lector.GetValue("INTERVALOS", "IntervaloLanzaHechizo"))
178     FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear
    
180     frmMain.TIMER_AI.Interval = val(Lector.GetValue("INTERVALOS", "IntervaloNpcAI"))
182     FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval
    
184     frmMain.npcataca.Interval = val(Lector.GetValue("INTERVALOS", "IntervaloNpcPuedeAtacar"))
186     FrmInterv.txtNPCPuedeAtacar.Text = frmMain.npcataca.Interval
    
188     IntervaloUserPuedeTrabajar = val(Lector.GetValue("INTERVALOS", "IntervaloTrabajo"))
190     FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar
    
192     IntervaloUserPuedeAtacar = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeAtacar"))
194     FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar
    
        'TODO : Agregar estos intervalos al form!!!
196     IntervaloMagiaGolpe = val(Lector.GetValue("INTERVALOS", "IntervaloMagiaGolpe"))
198     IntervaloGolpeMagia = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeMagia"))
    
        'frmMain.tLluvia.Interval = val(Lector.GetValue("INTERVALOS", "IntervaloPerdidaStaminaLluvia"))
        'FrmInterv.txtIntervaloPerdidaStaminaLluvia.Text = frmMain.tLluvia.Interval
    
200     MinutosWs = val(Lector.GetValue("INTERVALOS", "IntervaloWS"))

202     If MinutosWs < 1 Then MinutosWs = 10
    
204     IntervaloCerrarConexion = val(Lector.GetValue("INTERVALOS", "IntervaloCerrarConexion"))
206     IntervaloUserPuedeUsarU = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeUsarU"))
207     IntervaloUserPuedeUsarClic = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeUsarClic"))
208     IntervaloFlechasCazadores = val(Lector.GetValue("INTERVALOS", "IntervaloFlechasCazadores"))
209     IntervaloGolpeUsar = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeUsar"))
    
210     IntervaloOculto = val(Lector.GetValue("INTERVALOS", "IntervaloOculto"))

211     MargenDeIntervaloPorPing = val(Lector.GetValue("INTERVALOS", "MargenDeIntervaloPorPing"))
    
212     IntervaloPuedeSerAtacado = val(Lector.GetValue("INTERVALOS", "IntervaloPuedeSerAtacado"))
    
        '&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    
214     Set Lector = Nothing

        
        Exit Sub

LoadIntervalos_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LoadIntervalos", Erl)
        Resume Next
        
End Sub

Sub LoadConfiguraciones()
        
        On Error GoTo LoadConfiguraciones_Err
        
100     ExpMult = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "ExpMult"))
102     OroMult = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "OroMult"))
104     OroAutoEquipable = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "OroAutoEquipable"))
106     DropMult = val(GetVar(IniPath & "Configuracion.ini", "DROPEO", "DropMult"))
108     DropActive = val(GetVar(IniPath & "Configuracion.ini", "DROPEO", "DropActive"))
110     RecoleccionMult = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "RecoleccionMult"))

112     TimerLimpiarObjetos = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "TimerLimpiarObjetos"))
114     OroPorNivel = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "OroPorNivel"))

116     TimerHoraFantasia = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "TimerHoraFantasia"))

118     BattleActivado = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "BattleActivado"))
120     BattleMinNivel = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "BattleMinNivel"))

122     frmMain.HoraFantasia.Interval = TimerHoraFantasia

126     frmMain.lblLimpieza.Caption = "Limpieza de objetos cada: " & TimerLimpiarObjetos & " minutos."

        
        Exit Sub

LoadConfiguraciones_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LoadConfiguraciones", Erl)
        Resume Next
        
End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
        '*****************************************************************
        'Escribe VAR en un archivo
        '*****************************************************************
        
        On Error GoTo WriteVar_Err
        

100     writeprivateprofilestring Main, Var, Value, File
    
        
        Exit Sub

WriteVar_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.WriteVar", Erl)
        Resume Next
        
End Sub

Sub LoadUser(ByVal UserIndex As Integer)

    On Error GoTo ErrorHandler
    
    If Database_Enabled Then
        Call LoadUserDatabase(UserIndex)
    Else
        Call LoadUserBinary(UserIndex)

    End If
    
    With UserList(UserIndex)

        If .flags.Paralizado = 1 Then
            .Counters.Paralisis = IntervaloParalizado

        End If

        'Obtiene el indice-objeto del arma
        If .Invent.WeaponEqpSlot > 0 Then
            .Invent.WeaponEqpObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).ObjIndex
            
            If .flags.Muerto = 0 Then
                .Char.Arma_Aura = ObjData(.Invent.WeaponEqpObjIndex).CreaGRH

            End If

        End If

        'Obtiene el indice-objeto del armadura
        If .Invent.ArmourEqpSlot > 0 Then
            .Invent.ArmourEqpObjIndex = .Invent.Object(.Invent.ArmourEqpSlot).ObjIndex
            
            If .flags.Muerto = 0 Then
                .Char.Body_Aura = ObjData(.Invent.ArmourEqpObjIndex).CreaGRH

            End If

            .flags.Desnudo = 0
        Else
            .flags.Desnudo = 1

        End If

        'Obtiene el indice-objeto del escudo
        If .Invent.EscudoEqpSlot > 0 Then
            .Invent.EscudoEqpObjIndex = .Invent.Object(.Invent.EscudoEqpSlot).ObjIndex
            
            If .flags.Muerto = 0 Then
                .Char.Escudo_Aura = ObjData(.Invent.EscudoEqpObjIndex).CreaGRH

            End If

        End If
        
        'Obtiene el indice-objeto del casco
        If .Invent.CascoEqpSlot > 0 Then
            .Invent.CascoEqpObjIndex = .Invent.Object(.Invent.CascoEqpSlot).ObjIndex
            
            If .flags.Muerto = 0 Then
                .Char.Head_Aura = ObjData(.Invent.CascoEqpObjIndex).CreaGRH

            End If

        End If

        'Obtiene el indice-objeto barco
        If .Invent.BarcoSlot > 0 Then
            .Invent.BarcoObjIndex = .Invent.Object(.Invent.BarcoSlot).ObjIndex

        End If

        'Obtiene el indice-objeto municion
        If .Invent.MunicionEqpSlot > 0 Then
            .Invent.MunicionEqpObjIndex = .Invent.Object(.Invent.MunicionEqpSlot).ObjIndex

        End If

        '[Alejo]
        'Obtiene el indice-objeto anilo
        If .Invent.AnilloEqpSlot > 0 Then
            .Invent.AnilloEqpObjIndex = .Invent.Object(.Invent.AnilloEqpSlot).ObjIndex

        End If

        If .Invent.MonturaSlot > 0 Then
            .Invent.MonturaObjIndex = .Invent.Object(.Invent.MonturaSlot).ObjIndex

        End If
        
        If .Invent.HerramientaEqpSlot > 0 Then
            .Invent.HerramientaEqpObjIndex = .Invent.Object(.Invent.HerramientaEqpSlot).ObjIndex

        End If
        
        If .Invent.NudilloSlot > 0 Then
            .Invent.NudilloObjIndex = .Invent.Object(.Invent.NudilloSlot).ObjIndex
            
            If .flags.Muerto = 0 Then
                .Char.Arma_Aura = ObjData(.Invent.NudilloObjIndex).CreaGRH

            End If

        End If
        
        If .Invent.MagicoSlot > 0 Then
            .Invent.MagicoObjIndex = .Invent.Object(.Invent.MagicoSlot).ObjIndex

        End If

        If .flags.Muerto = 0 Then

            .Char = .OrigChar
            
            If .Char.Body = 0 Then
                Call DarCuerpoDesnudo(Userindex)
            End If
            
            If .Char.Head = 0 Then
                .Char.Head = 1
            End If
            
        Else
            .Char.Body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.WeaponAnim = NingunArma
            .Char.ShieldAnim = NingunEscudo
            .Char.CascoAnim = NingunCasco
            .Char.heading = eHeading.SOUTH

        End If

    End With

    Exit Sub

ErrorHandler:
    Call LogError("Error en LoadUser: " & UserList(UserIndex).name & " - " & Err.Number & " - " & Err.description)
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer, Optional ByVal Logout As Boolean = False)
        
        On Error GoTo SaveUser_Err
        

100     If Database_Enabled Then
102         Call SaveUserDatabase(UserIndex, Logout)
        Else
104         Call SaveUserBinary(UserIndex, Logout)

        End If

        
        Exit Sub

SaveUser_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.SaveUser", Erl)
        Resume Next
        
End Sub

Sub LoadUserBinary(ByVal UserIndex As Integer)
        
        On Error GoTo LoadUserBinary_Err
        

        'Cargamos el personaje
        Dim Leer As New clsIniReader
    
100     Call Leer.Initialize(CharPath & UCase$(UserList(UserIndex).name) & ".chr")
    
        'Cargamos los datos del personaje

102     Call LoadUserInit(UserIndex, Leer)
    
104     Call LoadUserStats(UserIndex, Leer)
    
106     Call LoadQuestStats(UserIndex, Leer)
    
108     Set Leer = Nothing

        
        Exit Sub

LoadUserBinary_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LoadUserBinary", Erl)
        Resume Next
        
End Sub

Sub SaveUserBinary(ByVal UserIndex As Integer, Optional ByVal Logout As Boolean)
    '*************************************************
    'Author: Unknown
    'Last modified: 23/01/2007
    'Saves the Users records
    '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
    '*************************************************
    
    On Error GoTo Errhandler
    
    Dim UserFile    As String

    Dim OldUserHead As Long
    
    UserFile = CharPath & UCase$(UserList(UserIndex).name) & ".chr"
    
    'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
    'clase=0 es el error, porq el enum empieza de 1!!
    If UserList(UserIndex).clase = 0 Or UserList(UserIndex).Stats.ELV = 0 Then
        Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & UserList(UserIndex).name)
        Exit Sub

    End If
    
    Debug.Print UserFile
    
    If FileExist(UserFile, vbNormal) Then
        If UserList(UserIndex).flags.Muerto = 1 Then
            OldUserHead = UserList(UserIndex).Char.Head
            UserList(UserIndex).Char.Head = GetVar(UserFile, "INIT", "Head")

        End If

        '       Kill UserFile
    End If
    
    Dim LoopC As Integer
    
    If FileExist(UserFile, vbNormal) Then Kill UserFile

    Dim File As String

    File = UserFile

    Dim n

    Dim Datos$

    n = FreeFile
    Open File For Binary Access Write As n
        
    'INIT
    Put n, , "[INIT]" & vbCrLf & "Cuenta=" & UserList(UserIndex).Cuenta & vbCrLf
    Put n, , "Genero=" & UserList(UserIndex).genero & vbCrLf
    Put n, , "Raza=" & UserList(UserIndex).raza & vbCrLf
    Put n, , "Hogar=" & UserList(UserIndex).Hogar & vbCrLf
    Put n, , "Clase=" & UserList(UserIndex).clase & vbCrLf
    Put n, , "Desc=" & UserList(UserIndex).Desc & vbCrLf
    Put n, , "Heading=" & CStr(UserList(UserIndex).Char.heading) & vbCrLf

    If UserList(UserIndex).Char.Head = 0 Then
        Put n, , "Head=" & CStr(UserList(UserIndex).OrigChar.Head) & vbCrLf
    Else
        Put n, , "Head=" & CStr(UserList(UserIndex).Char.Head) & vbCrLf

    End If

    Put n, , "Arma=" & CStr(UserList(UserIndex).Char.WeaponAnim) & vbCrLf
    Put n, , "Escudo=" & CStr(UserList(UserIndex).Char.ShieldAnim) & vbCrLf
    Put n, , "Casco=" & CStr(UserList(UserIndex).Char.CascoAnim) & vbCrLf
    Put n, , "Position=" & UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.x & "-" & UserList(UserIndex).Pos.Y & vbCrLf
    'If UserList(UserIndex).flags.Muerto = 0 Then
    Put n, , "Body=" & CStr(UserList(UserIndex).Char.Body) & vbCrLf
    'End If
    #If ConUpTime Then

        Dim TempDate As Date

        TempDate = Now - UserList(UserIndex).LogOnTime
        UserList(UserIndex).LogOnTime = Now
        UserList(UserIndex).UpTime = UserList(UserIndex).UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
        UserList(UserIndex).UpTime = UserList(UserIndex).UpTime
        Put n, , "UpTime=" & UserList(UserIndex).UpTime & vbCrLf
    #End If
        
    'Call WriteVar(CharPath & UserList(UserIndex).Name & ".chr", "INIT", "Logged", "0")
        
    If Logout Then
        Put n, , "Logged=0" & vbCrLf
    Else
        Put n, , "Logged=1" & vbCrLf

    End If

    Put n, , "MENSAJEINFORMACION=" & UserList(UserIndex).MENSAJEINFORMACION & vbCrLf

    Put n, , vbCrLf
        
    'baneo
    Put n, , "[BAN]" & vbCrLf & "Baneado=" & CStr(UserList(UserIndex).flags.Ban) & vbCrLf
    Put n, , "BanMotivo=" & CStr(UserList(UserIndex).flags.BanMotivo) & vbCrLf
        
    Put n, , vbCrLf
        
    'STATS
    Put n, , "[STATS]" & vbCrLf & "GLD=" & CStr(UserList(UserIndex).Stats.GLD) & vbCrLf
    Put n, , "BANCO=" & CStr(UserList(UserIndex).Stats.Banco) & vbCrLf
    Put n, , "MaxHP=" & CStr(UserList(UserIndex).Stats.MaxHp) & vbCrLf
    Put n, , "MinHP=" & CStr(UserList(UserIndex).Stats.MinHp) & vbCrLf
    Put n, , "MaxSTA=" & CStr(UserList(UserIndex).Stats.MaxSta) & vbCrLf
    Put n, , "MinSTA=" & CStr(UserList(UserIndex).Stats.MinSta) & vbCrLf
    Put n, , "MaxMAN=" & CStr(UserList(UserIndex).Stats.MaxMAN) & vbCrLf
    Put n, , "MinMAN=" & CStr(UserList(UserIndex).Stats.MinMAN) & vbCrLf
    Put n, , "MaxHIT=" & CStr(UserList(UserIndex).Stats.MaxHit) & vbCrLf
    Put n, , "MinHIT=" & CStr(UserList(UserIndex).Stats.MinHIT) & vbCrLf
    Put n, , "MaxAGU=" & CStr(UserList(UserIndex).Stats.MaxAGU) & vbCrLf
    Put n, , "MinAGU=" & CStr(UserList(UserIndex).Stats.MinAGU) & vbCrLf
    Put n, , "MaxHAM=" & CStr(UserList(UserIndex).Stats.MaxHam) & vbCrLf
    Put n, , "MinHAM=" & CStr(UserList(UserIndex).Stats.MinHam) & vbCrLf
    Put n, , "SkillPtsLibres=" & CStr(UserList(UserIndex).Stats.SkillPts) & vbCrLf
    Put n, , "EXP=" & CStr(UserList(UserIndex).Stats.Exp) & vbCrLf
    Put n, , "ELV=" & CStr(UserList(UserIndex).Stats.ELV) & vbCrLf
    Put n, , "ELU=" & CStr(UserList(UserIndex).Stats.ELU) & vbCrLf
    Put n, , "InventLevel=" & CByte(UserList(UserIndex).Stats.InventLevel) & vbCrLf
        
    Put n, , vbCrLf
        
    'FLAGS
    Put n, , "[FLAGS]" & vbCrLf & "CASADO=" & CStr(UserList(UserIndex).flags.Casado) & vbCrLf
    Put n, , "PAREJA=" & CStr(UserList(UserIndex).flags.Pareja) & vbCrLf
    Put n, , "Muerto=" & CStr(UserList(UserIndex).flags.Muerto) & vbCrLf
    Put n, , "Escondido=" & CStr(UserList(UserIndex).flags.Escondido) & vbCrLf
    Put n, , "Hambre=" & CStr(UserList(UserIndex).flags.Hambre) & vbCrLf
    Put n, , "Sed=" & CStr(UserList(UserIndex).flags.Sed) & vbCrLf
    Put n, , "Desnudo=" & CStr(UserList(UserIndex).flags.Desnudo) & vbCrLf
    Put n, , "Navegando=" & CStr(UserList(UserIndex).flags.Navegando) & vbCrLf
    Put n, , "Envenenado=" & CStr(UserList(UserIndex).flags.Envenenado) & vbCrLf
    Put n, , "Paralizado=" & CStr(UserList(UserIndex).flags.Paralizado) & vbCrLf
    Put n, , "Inmovilizado=" & CStr(UserList(UserIndex).flags.Inmovilizado) & vbCrLf
    Put n, , "Incinerado=" & CStr(UserList(UserIndex).flags.Incinerado) & vbCrLf
    Put n, , "VecesQueMoriste=" & CStr(UserList(UserIndex).flags.VecesQueMoriste) & vbCrLf
    Put n, , "ScrollExp=" & CStr(UserList(UserIndex).flags.ScrollExp) & vbCrLf
    Put n, , "ScrollOro=" & CStr(UserList(UserIndex).flags.ScrollOro) & vbCrLf
    Put n, , "MinutosRestantes=" & CStr(UserList(UserIndex).flags.MinutosRestantes) & vbCrLf
    Put n, , "SegundosPasados=" & CStr(UserList(UserIndex).flags.SegundosPasados) & vbCrLf
    Put n, , "Silenciado=" & CStr(UserList(UserIndex).flags.Silenciado) & vbCrLf
    Put n, , "Montado=" & CStr(UserList(UserIndex).flags.Montado) & vbCrLf
        
    Put n, , vbCrLf
        
    'GRABADO DE CLAN
    Put n, , "[GUILD]" & vbCrLf & "GUILDINDEX=" & CInt(UserList(UserIndex).GuildIndex) & vbCrLf
        
    Put n, , vbCrLf
        
    Put n, , "[CONSEJO]" & vbCrLf
        
    Dim PERTENECEb As Byte

    PERTENECEb = IIf(UserList(UserIndex).flags.Privilegios And PlayerType.RoyalCouncil, "1", "0")
        
    Dim PERTENECECAOSb As Byte

    PERTENECECAOSb = IIf(UserList(UserIndex).flags.Privilegios And PlayerType.ChaosCouncil, "1", "0")

    Put n, , "PERTENECE=" & PERTENECEb & vbCrLf
    Put n, , "PERTENECECAOS=" & PERTENECECAOSb & vbCrLf
        
    Put n, , vbCrLf
    Put n, , "[FACCIONES]" & vbCrLf & "EjercitoReal=" & CStr(UserList(UserIndex).Faccion.ArmadaReal) & vbCrLf
    Put n, , "Status=" & CStr(UserList(UserIndex).Faccion.Status) & vbCrLf
    Put n, , "EjercitoCaos=" & CStr(UserList(UserIndex).Faccion.FuerzasCaos) & vbCrLf
    Put n, , "CiudMatados=" & CStr(UserList(UserIndex).Faccion.CiudadanosMatados) & vbCrLf
    Put n, , "CrimMatados=" & CStr(UserList(UserIndex).Faccion.CriminalesMatados) & vbCrLf
    Put n, , "rArCaos=" & CStr(UserList(UserIndex).Faccion.RecibioArmaduraCaos) & vbCrLf
    Put n, , "rArReal=" & CStr(UserList(UserIndex).Faccion.RecibioArmaduraReal) & vbCrLf
    Put n, , "rExCaos=" & CStr(UserList(UserIndex).Faccion.RecibioExpInicialCaos) & vbCrLf
    Put n, , "rExReal=" & CStr(UserList(UserIndex).Faccion.RecibioExpInicialReal) & vbCrLf
    Put n, , "recCaos=" & CStr(UserList(UserIndex).Faccion.RecompensasCaos) & vbCrLf
    Put n, , "recReal=" & CStr(UserList(UserIndex).Faccion.RecompensasReal) & vbCrLf
    Put n, , "Reenlistadas=" & CStr(UserList(UserIndex).Faccion.Reenlistadas) & vbCrLf
    Put n, , "NivelIngreso=" & CStr(UserList(UserIndex).Faccion.NivelIngreso) & vbCrLf
    Put n, , "FechaIngreso=" & CStr(UserList(UserIndex).Faccion.FechaIngreso) & vbCrLf
    Put n, , "MatadosIngreso=" & CStr(UserList(UserIndex).Faccion.MatadosIngreso) & vbCrLf
    Put n, , "NextRecompensa=" & CStr(UserList(UserIndex).Faccion.NextRecompensa) & vbCrLf

    Put n, , vbCrLf
        
    'MAHIA ESTUPIDIZA
    Put n, , "[MAGIA]" & vbCrLf & "ENVENENA=" & CByte(UserList(UserIndex).flags.Envenena) & vbCrLf
    Put n, , "PARALIZA=" & CByte(UserList(UserIndex).flags.Paraliza) & vbCrLf
    Put n, , "AnilloOcultismo=" & CByte(UserList(UserIndex).flags.AnilloOcultismo) & vbCrLf
    Put n, , "incinera=" & CByte(UserList(UserIndex).flags.incinera) & vbCrLf
    Put n, , "Estupidiza=" & CByte(UserList(UserIndex).flags.Estupidiza) & vbCrLf
    Put n, , "Pendiente=" & CByte(UserList(UserIndex).flags.PendienteDelSacrificio) & vbCrLf
    Put n, , "CarroMineria=" & CByte(UserList(UserIndex).flags.CarroMineria) & vbCrLf
    Put n, , "NoPalabrasMagicas=" & CByte(UserList(UserIndex).flags.NoPalabrasMagicas) & vbCrLf
    Put n, , "NoDetectable=" & CByte(UserList(UserIndex).flags.NoDetectable) & vbCrLf
    Put n, , "Otra_Aura=" & CStr(UserList(UserIndex).Char.Otra_Aura) & vbCrLf
    Put n, , "DañoMagico=" & CByte(UserList(UserIndex).flags.DañoMagico) & vbCrLf
    Put n, , "ResistenciaMagica=" & CByte(UserList(UserIndex).flags.ResistenciaMagica) & vbCrLf
    Put n, , "RegeneracionMana=" & CByte(UserList(UserIndex).flags.RegeneracionMana) & vbCrLf
    Put n, , "NoMagiaEfeceto=" & CByte(UserList(UserIndex).flags.NoMagiaEfeceto) & vbCrLf
    Put n, , "RegeneracionHP=" & CByte(UserList(UserIndex).flags.RegeneracionHP) & vbCrLf
    Put n, , "RegeneracionSta=" & CByte(UserList(UserIndex).flags.RegeneracionSta) & vbCrLf

    Put n, , vbCrLf
    'SKILLS
    Put n, , "[SKILLS]" & vbCrLf

    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserSkills)
        Put n, , "SK" & LoopC & "=" & CStr(UserList(UserIndex).Stats.UserSkills(LoopC)) & vbCrLf
    Next

    Put n, , vbCrLf

    'INVENTARIO
    Put n, , "[Inventory]" & vbCrLf & "CantidadItems=" & val(UserList(UserIndex).Invent.NroItems) & vbCrLf

    For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots
        Put n, , "Obj" & LoopC & "=" & UserList(UserIndex).Invent.Object(LoopC).ObjIndex & "-" & UserList(UserIndex).Invent.Object(LoopC).Amount & "-" & UserList(UserIndex).Invent.Object(LoopC).Equipped & vbCrLf
    Next
    Put n, , "WeaponEqpSlot=" & CStr(UserList(UserIndex).Invent.WeaponEqpSlot) & vbCrLf
    Put n, , "HerramientaEqpSlot=" & CStr(UserList(UserIndex).Invent.HerramientaEqpSlot) & vbCrLf
    Put n, , "ArmourEqpSlot=" & CStr(UserList(UserIndex).Invent.ArmourEqpSlot) & vbCrLf
    Put n, , "CascoEqpSlot=" & CStr(UserList(UserIndex).Invent.CascoEqpSlot) & vbCrLf
    Put n, , "EscudoEqpSlot=" & CStr(UserList(UserIndex).Invent.EscudoEqpSlot) & vbCrLf
    Put n, , "BarcoSlot=" & CStr(UserList(UserIndex).Invent.BarcoSlot) & vbCrLf
    Put n, , "MonturaSlot=" & CStr(UserList(UserIndex).Invent.MonturaSlot) & vbCrLf
    Put n, , "MunicionSlot=" & CStr(UserList(UserIndex).Invent.MunicionEqpSlot) & vbCrLf
    Put n, , "AnilloSlot=" & CStr(UserList(UserIndex).Invent.AnilloEqpSlot) & vbCrLf
    Put n, , "MagicoSlot=" & CStr(UserList(UserIndex).Invent.MagicoSlot) & vbCrLf
    Put n, , "NudilloEqpSlot=" & CStr(UserList(UserIndex).Invent.NudilloSlot) & vbCrLf
        
    Put n, , vbCrLf

    Put n, , "[ATRIBUTOS]" & vbCrLf

    '¿Fueron modificados los atributos del usuario?
    If Not UserList(UserIndex).flags.TomoPocion Then

        For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
            Put n, , "AT" & LoopC & "=" & CStr(UserList(UserIndex).Stats.UserAtributos(LoopC)) & vbCrLf
        Next
    Else

        For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
            'UserList(UserIndex).Stats.UserAtributos(LoopC) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)
            Put n, , "AT" & LoopC & "=" & CStr(UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)) & vbCrLf
        Next

    End If

    Put n, , vbCrLf
            
    'COUNTERS
    Put n, , "[COUNTERS]" & vbCrLf & "Pena=" & CStr(UserList(UserIndex).Counters.Pena) & vbCrLf
    Put n, , "ScrollOro=" & CStr(UserList(UserIndex).Counters.ScrollOro) & vbCrLf
    Put n, , "ScrollExperiencia=" & CStr(UserList(UserIndex).Counters.ScrollExperiencia) & vbCrLf
    Put n, , "Oxigeno=" & CStr(UserList(UserIndex).Counters.Oxigeno) & vbCrLf
        
    Put n, , vbCrLf

    Put n, , "[MUERTES]" & vbCrLf & "UserMuertes=" & CStr(UserList(UserIndex).Stats.UsuariosMatados) & vbCrLf
    Put n, , "NpcsMuertes=" & CStr(UserList(UserIndex).Stats.NPCsMuertos) & vbCrLf
        
    Put n, , vbCrLf
    'BANCO
    Put n, , "[BancoInventory]" & vbCrLf & "CantidadItems=" & val(UserList(UserIndex).BancoInvent.NroItems) & vbCrLf

    Dim loopd As Integer

    For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
        Put n, , "Obj" & loopd & "=" & UserList(UserIndex).BancoInvent.Object(loopd).ObjIndex & "-" & UserList(UserIndex).BancoInvent.Object(loopd).Amount & vbCrLf
    Next loopd
        
    Put n, , vbCrLf
        
    Put n, , "[LOGROS]" & vbCrLf & "UserLogros=" & CByte(UserList(UserIndex).UserLogros) & vbCrLf
    Put n, , "NPcLogros=" & CByte(UserList(UserIndex).NPcLogros) & vbCrLf
    Put n, , "LevelLogros=" & CByte(UserList(UserIndex).LevelLogros) & vbCrLf
        
    Put n, , vbCrLf
        
    Put n, , "[BINDKEYS]" & vbCrLf
    Put n, , "ChatCombate=" & CByte(UserList(UserIndex).ChatCombate) & vbCrLf
    Put n, , "ChatGlobal=" & CByte(UserList(UserIndex).ChatGlobal) & vbCrLf
        
    Put n, , vbCrLf

    'HECHIZOS
    Put n, , "[HECHIZOS]" & vbCrLf

    Dim cad As String

    For LoopC = 1 To MAXUSERHECHIZOS
        cad = UserList(UserIndex).Stats.UserHechizos(LoopC)
        Put n, , "H" & LoopC & "=" & cad & vbCrLf
    Next
        
    Put n, , vbCrLf
        
    'BATTLE
    Put n, , "[Battle]" & vbCrLf & "Puntos=" & CStr(UserList(UserIndex).flags.BattlePuntos) & vbCrLf
        
    Put n, , vbCrLf
        
    Put n, , "[CORREO]" & vbCrLf & "NoLeidos=" & CByte(UserList(UserIndex).Correo.NoLeidos) & vbCrLf
    Put n, , "CANTCORREO=" & CByte(UserList(UserIndex).Correo.CantCorreo) & vbCrLf
        
    Put n, , vbCrLf
    'Correo Ladder
        
    For LoopC = 1 To UserList(UserIndex).Correo.CantCorreo
        
        Put n, , "REMITENTE" & LoopC & "=" & UserList(UserIndex).Correo.Mensaje(LoopC).Remitente & vbCrLf
        Put n, , "MENSAJE" & LoopC & "=" & UserList(UserIndex).Correo.Mensaje(LoopC).Mensaje & vbCrLf
        Put n, , "Item" & LoopC & "=" & UserList(UserIndex).Correo.Mensaje(LoopC).Item & vbCrLf
        Put n, , "ItemCount" & LoopC & "=" & UserList(UserIndex).Correo.Mensaje(LoopC).ItemCount & vbCrLf
        Put n, , "DATE" & LoopC & "=" & UserList(UserIndex).Correo.Mensaje(LoopC).Fecha & vbCrLf
        Put n, , "LEIDO" & LoopC & "=" & UserList(UserIndex).Correo.Mensaje(LoopC).Leido & vbCrLf
            
    Next LoopC
        
    Close #n
        
    Call SaveQuestStats(UserIndex, UserFile)

    'Devuelve el head de muerto
    If UserList(UserIndex).flags.Muerto = 1 Then
        UserList(UserIndex).Char.Head = iCabezaMuerto

    End If
        
    Exit Sub

Errhandler:
    Call LogError("Error en SaveUserBinary")
    Close #n

End Sub

Sub SaveNewUser(ByVal UserIndex As Integer)
        
        On Error GoTo SaveNewUser_Err
        
    
100     If Database_Enabled Then
102         Call SaveNewUserDatabase(UserIndex)
        Else
104         Call SaveNewUserCharfile(UserIndex)

        End If
    
        
        Exit Sub

SaveNewUser_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.SaveNewUser", Erl)
        Resume Next
        
End Sub

Sub SaveNewUserCharfile(ByVal UserIndex As Integer)
    '*************************************************
    'Author: Unknown
    'Last modified: 23/01/2007
    'Saves the Users records
    '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
    '*************************************************
    
    On Error GoTo Errhandler
    
    Dim UserFile    As String

    Dim OldUserHead As Long
    
    UserFile = CharPath & UCase$(UserList(UserIndex).name) & ".chr"
    
    'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
    'clase=0 es el error, porq el enum empieza de 1!!
    If UserList(UserIndex).clase = 0 Or UserList(UserIndex).Stats.ELV = 0 Then
        Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & UserList(UserIndex).name)
        Exit Sub

    End If
    
    If FileExist(UserFile, vbNormal) Then
        If UserList(UserIndex).flags.Muerto = 1 Then
            OldUserHead = UserList(UserIndex).Char.Head
            UserList(UserIndex).Char.Head = GetVar(UserFile, "INIT", "Head")

        End If

        '       Kill UserFile
    End If
    
    Dim LoopC As Integer

    Dim n

    Dim Datos$

    n = FreeFile
    Open UserFile For Binary Access Write As n
    
    'BATTLE
    Put n, , "[Battle]" & vbCrLf & "Puntos=" & CStr(UserList(UserIndex).flags.BattlePuntos) & vbCrLf
    
    Put n, , vbCrLf
    
    'FLAGS
    Put n, , "[FLAGS]" & vbCrLf & "CASADO=" & CStr(UserList(UserIndex).flags.Casado) & vbCrLf
    Put n, , "PAREJA=" & vbCrLf
    Put n, , "Muerto=0" & vbCrLf
    Put n, , "Escondido=0" & vbCrLf
    Put n, , "Hambre=0" & vbCrLf
    Put n, , "Sed=0" & vbCrLf
    Put n, , "Desnudo=0" & vbCrLf
    Put n, , "Navegando=0" & vbCrLf
    Put n, , "Envenenado=0" & vbCrLf
    Put n, , "Paralizado=0" & vbCrLf
    Put n, , "Inmovilizado=0" & vbCrLf
    Put n, , "Incinerado=0" & vbCrLf
    Put n, , "VecesQueMoriste=0" & vbCrLf
    Put n, , "ScrollExp=" & CStr(UserList(UserIndex).flags.ScrollExp) & vbCrLf
    Put n, , "ScrollOro=" & CStr(UserList(UserIndex).flags.ScrollOro) & vbCrLf
    Put n, , "MinutosRestantes=0" & vbCrLf
    Put n, , "SegundosPasados=0" & vbCrLf
    Put n, , "Silenciado=0" & vbCrLf
    Put n, , "Montado=0" & vbCrLf
    
    Put n, , "InventLevel=0" & vbCrLf
    
    Put n, , vbCrLf
    
    Put n, , "[CONSEJO]" & vbCrLf
    Put n, , "PERTENECE=0" & vbCrLf
    Put n, , "PERTENECECAOS=0" & vbCrLf
    
    Put n, , "[FACCIONES]" & vbCrLf & "EjercitoReal=" & CStr(UserList(UserIndex).Faccion.ArmadaReal) & vbCrLf
    Put n, , "Status=" & CStr(UserList(UserIndex).Faccion.Status) & vbCrLf
    Put n, , "EjercitoCaos=" & CStr(UserList(UserIndex).Faccion.FuerzasCaos) & vbCrLf
    Put n, , "CiudMatados=" & CStr(UserList(UserIndex).Faccion.CiudadanosMatados) & vbCrLf
    Put n, , "CrimMatados=" & CStr(UserList(UserIndex).Faccion.CriminalesMatados) & vbCrLf
    Put n, , "rArCaos=" & CStr(UserList(UserIndex).Faccion.RecibioArmaduraCaos) & vbCrLf
    Put n, , "rArReal=" & CStr(UserList(UserIndex).Faccion.RecibioArmaduraReal) & vbCrLf
    Put n, , "rExCaos=" & CStr(UserList(UserIndex).Faccion.RecibioExpInicialCaos) & vbCrLf
    Put n, , "rExReal=" & CStr(UserList(UserIndex).Faccion.RecibioExpInicialReal) & vbCrLf
    Put n, , "recCaos=" & CStr(UserList(UserIndex).Faccion.RecompensasCaos) & vbCrLf
    Put n, , "recReal=" & CStr(UserList(UserIndex).Faccion.RecompensasReal) & vbCrLf
    Put n, , "Reenlistadas=" & CStr(UserList(UserIndex).Faccion.Reenlistadas) & vbCrLf
    Put n, , "NivelIngreso=" & CStr(UserList(UserIndex).Faccion.NivelIngreso) & vbCrLf
    Put n, , "FechaIngreso=" & CStr(UserList(UserIndex).Faccion.FechaIngreso) & vbCrLf
    Put n, , "MatadosIngreso=" & CStr(UserList(UserIndex).Faccion.MatadosIngreso) & vbCrLf
    Put n, , "NextRecompensa=" & CStr(UserList(UserIndex).Faccion.NextRecompensa) & vbCrLf
    
    Put n, , vbCrLf
    
    'STATS
    Put n, , "[STATS]" & vbCrLf & "GLD=0" & vbCrLf
    Put n, , "BANCO=0" & vbCrLf
    Put n, , "MaxHP=" & CStr(UserList(UserIndex).Stats.MaxHp) & vbCrLf
    Put n, , "MinHP=" & CStr(UserList(UserIndex).Stats.MinHp) & vbCrLf
    Put n, , "MaxSTA=" & CStr(UserList(UserIndex).Stats.MaxSta) & vbCrLf
    Put n, , "MinSTA=" & CStr(UserList(UserIndex).Stats.MinSta) & vbCrLf
    Put n, , "MaxMAN=" & CStr(UserList(UserIndex).Stats.MaxMAN) & vbCrLf
    Put n, , "MinMAN=" & CStr(UserList(UserIndex).Stats.MinMAN) & vbCrLf
    Put n, , "MaxHIT=" & CStr(UserList(UserIndex).Stats.MaxHit) & vbCrLf
    Put n, , "MinHIT=" & CStr(UserList(UserIndex).Stats.MinHIT) & vbCrLf
    Put n, , "MaxAGU=" & CStr(UserList(UserIndex).Stats.MaxAGU) & vbCrLf
    Put n, , "MinAGU=" & CStr(UserList(UserIndex).Stats.MinAGU) & vbCrLf
    Put n, , "MaxHAM=" & CStr(UserList(UserIndex).Stats.MaxHam) & vbCrLf
    Put n, , "MinHAM=" & CStr(UserList(UserIndex).Stats.MinHam) & vbCrLf
    Put n, , "SkillPtsLibres=" & CStr(UserList(UserIndex).Stats.SkillPts) & vbCrLf
    Put n, , "EXP=" & CStr(UserList(UserIndex).Stats.Exp) & vbCrLf
    Put n, , "ELV=" & CStr(UserList(UserIndex).Stats.ELV) & vbCrLf
    Put n, , "ELU=" & CStr(UserList(UserIndex).Stats.ELU) & vbCrLf
    
    Put n, , vbCrLf
    
    'MAHIA
    Put n, , "[MAGIA]" & vbCrLf & "ENVENENA=0" & vbCrLf
    Put n, , "PARALIZA=0" & vbCrLf
    Put n, , "INCINERA=0" & vbCrLf
    Put n, , "Estupidiza=0" & vbCrLf
    Put n, , "PENDIENTE=0" & vbCrLf
    Put n, , "CARROMINERIA=0" & vbCrLf
    Put n, , "NOPALABRASMAGICAS=0" & vbCrLf
    Put n, , "OTRA_AURA=0" & vbCrLf
    Put n, , "DAÑOMAGICO=0" & vbCrLf
    Put n, , "ResistenciaMagica=0" & vbCrLf
    Put n, , "NoDetectable=0" & vbCrLf
    Put n, , "AnilloOcultismo=0" & vbCrLf
    Put n, , "RegeneracionMana=0" & vbCrLf
    Put n, , "NoMagiaEfeceto=0" & vbCrLf
    Put n, , "RegeneracionHP=0" & vbCrLf
    Put n, , "RegeneracionSta=0" & vbCrLf
    
    Put n, , vbCrLf
    
    'SKILLS
    Put n, , "[SKILLS]" & vbCrLf

    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserSkills)
        Put n, , "SK" & LoopC & "=0" & vbCrLf
    Next
    
    Put n, , vbCrLf
    
    'INVENTARIO
    Put n, , "[Inventory]" & vbCrLf & "CantidadItems=" & val(UserList(UserIndex).Invent.NroItems) & vbCrLf

    For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots
        Put n, , "Obj" & LoopC & "=" & UserList(UserIndex).Invent.Object(LoopC).ObjIndex & "-" & UserList(UserIndex).Invent.Object(LoopC).Amount & "-" & UserList(UserIndex).Invent.Object(LoopC).Equipped & vbCrLf
    Next
    Put n, , "WeaponEqpSlot=" & CStr(UserList(UserIndex).Invent.WeaponEqpSlot) & vbCrLf
    Put n, , "HerramientaEqpSlot=" & CStr(UserList(UserIndex).Invent.HerramientaEqpSlot) & vbCrLf
    Put n, , "ArmourEqpSlot=" & CStr(UserList(UserIndex).Invent.ArmourEqpSlot) & vbCrLf
    Put n, , "CascoEqpSlot=" & CStr(UserList(UserIndex).Invent.CascoEqpSlot) & vbCrLf
    Put n, , "EscudoEqpSlot=" & CStr(UserList(UserIndex).Invent.EscudoEqpSlot) & vbCrLf
    Put n, , "BarcoSlot=" & CStr(UserList(UserIndex).Invent.BarcoSlot) & vbCrLf
    Put n, , "MonturaSlot=" & CStr(UserList(UserIndex).Invent.MonturaSlot) & vbCrLf
    Put n, , "MunicionSlot=" & CStr(UserList(UserIndex).Invent.MunicionEqpSlot) & vbCrLf
    Put n, , "AnilloSlot=" & CStr(UserList(UserIndex).Invent.AnilloEqpSlot) & vbCrLf
    Put n, , "MagicoSlot=" & CStr(UserList(UserIndex).Invent.MagicoSlot) & vbCrLf
    Put n, , "NudilloEqpSlot=" & CStr(UserList(UserIndex).Invent.NudilloSlot) & vbCrLf
    
    Put n, , vbCrLf
    
    'INIT
    Put n, , "[INIT]" & vbCrLf & "Cuenta=" & UserList(UserIndex).Cuenta & vbCrLf
    Put n, , "Genero=" & UserList(UserIndex).genero & vbCrLf
    Put n, , "Raza=" & UserList(UserIndex).raza & vbCrLf
    Put n, , "Hogar=" & UserList(UserIndex).Hogar & vbCrLf
    Put n, , "Clase=" & UserList(UserIndex).clase & vbCrLf
    Put n, , "Desc=" & UserList(UserIndex).Desc & vbCrLf
    Put n, , "Heading=" & CStr(UserList(UserIndex).Char.heading) & vbCrLf
    Put n, , "Head=" & CStr(UserList(UserIndex).Char.Head) & vbCrLf
    Put n, , "Arma=" & CStr(UserList(UserIndex).Char.WeaponAnim) & vbCrLf
    Put n, , "Escudo=" & CStr(UserList(UserIndex).Char.ShieldAnim) & vbCrLf
    Put n, , "Casco=" & CStr(UserList(UserIndex).Char.CascoAnim) & vbCrLf
    Put n, , "Position=" & UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.x & "-" & UserList(UserIndex).Pos.Y & vbCrLf
    ' If UserList(UserIndex).flags.Muerto = 0 Then
    Put n, , "Body=" & CStr(UserList(UserIndex).Char.Body) & vbCrLf
    'Else
    '   Put N, , "Body=" & iCuerpoMuerto & vbCrLf 'poner body muerto
    '  End If
    #If ConUpTime Then

        Dim TempDate As Date

        TempDate = Now - UserList(UserIndex).LogOnTime
        UserList(UserIndex).LogOnTime = Now
        UserList(UserIndex).UpTime = UserList(UserIndex).UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
        UserList(UserIndex).UpTime = UserList(UserIndex).UpTime
        Put n, , "UpTime=" & UserList(UserIndex).UpTime & vbCrLf
    #End If
    
    Put n, , vbCrLf
    
    Put n, , "[ATRIBUTOS]" & vbCrLf

    '¿Fueron modificados los atributos del usuario?
    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        Put n, , "AT" & LoopC & "=" & CStr(UserList(UserIndex).Stats.UserAtributos(LoopC)) & vbCrLf
    Next
    Put n, , vbCrLf
    
    'baneo
    Put n, , "[BAN]" & vbCrLf & "Baneado=" & CStr(UserList(UserIndex).flags.Ban) & vbCrLf
    Put n, , "BanMotivo=" & CStr(UserList(UserIndex).flags.BanMotivo) & vbCrLf
    
    Put n, , vbCrLf
    
    'COUNTERS
    Put n, , "[COUNTERS]" & vbCrLf & "Pena=" & CStr(UserList(UserIndex).Counters.Pena) & vbCrLf
    Put n, , "ScrollOro=" & CStr(UserList(UserIndex).Counters.ScrollOro) & vbCrLf
    Put n, , "ScrollExperiencia=" & CStr(UserList(UserIndex).Counters.ScrollExperiencia) & vbCrLf
    Put n, , "Oxigeno=" & CStr(UserList(UserIndex).Counters.Oxigeno) & vbCrLf
    
    Put n, , vbCrLf
    
    Put n, , "[MUERTES]" & vbCrLf & "UserMuertes=0" & vbCrLf
    Put n, , "NpcsMuertes=0" & vbCrLf
    
    Put n, , vbCrLf
    
    'BANCO
    Put n, , "[BancoInventory]" & vbCrLf & "CantidadItems=0" & vbCrLf

    Dim loopd As Integer

    For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
        Put n, , "Obj" & loopd & "=" & UserList(UserIndex).BancoInvent.Object(loopd).ObjIndex & "-" & UserList(UserIndex).BancoInvent.Object(loopd).Amount & vbCrLf
    Next loopd
    
    Put n, , vbCrLf
    
    Put n, , "[LOGROS]" & vbCrLf & "UserLogros=" & CByte(UserList(UserIndex).UserLogros) & vbCrLf
    Put n, , "NPcLogros=" & CByte(UserList(UserIndex).NPcLogros) & vbCrLf
    Put n, , "LevelLogros=" & CByte(UserList(UserIndex).LevelLogros) & vbCrLf
    
    Put n, , vbCrLf
    
    Put n, , "[BINDKEYS]" & vbCrLf
    Put n, , "ChatCombate=" & CByte(UserList(UserIndex).ChatCombate) & vbCrLf
    Put n, , "ChatGlobal=" & CByte(UserList(UserIndex).ChatGlobal) & vbCrLf
    
    Put n, , vbCrLf
    
    'HECHIZOS
    Put n, , "[HECHIZOS]" & vbCrLf

    Dim cad As String

    For LoopC = 1 To MAXUSERHECHIZOS
        cad = UserList(UserIndex).Stats.UserHechizos(LoopC)
        Put n, , "H" & LoopC & "=" & cad & vbCrLf
    Next
      
    Put n, , vbCrLf
    
    Put n, , "[CORREO]" & vbCrLf & "NoLeidos=0" & vbCrLf
    Put n, , "CANTCORREO=0" & vbCrLf
    
    'Correo Ladder
    
    For LoopC = 1 To UserList(UserIndex).Correo.CantCorreo
    
        Put n, , "REMITENTE" & LoopC & "=" & UserList(UserIndex).Correo.Mensaje(LoopC).Remitente & vbCrLf
        Put n, , "MENSAJE" & LoopC & "=" & UserList(UserIndex).Correo.Mensaje(LoopC).Mensaje & vbCrLf
        Put n, , "Item" & LoopC & "=" & UserList(UserIndex).Correo.Mensaje(LoopC).Item & vbCrLf
        Put n, , "ItemCount" & LoopC & "=" & UserList(UserIndex).Correo.Mensaje(LoopC).ItemCount & vbCrLf
        Put n, , "DATE" & LoopC & "=" & UserList(UserIndex).Correo.Mensaje(LoopC).Fecha & vbCrLf
        Put n, , "LEIDO" & LoopC & "=" & UserList(UserIndex).Correo.Mensaje(LoopC).Leido & vbCrLf
        
    Next LoopC
    
    Close #n
    
    'Devuelve el head de muerto
    If UserList(UserIndex).flags.Muerto = 1 Then
        UserList(UserIndex).Char.Head = iCabezaMuerto

    End If
    
    Exit Sub
    
Errhandler:
    Call LogError("Error en SaveNewUserCharfile")
    Close #n

End Sub

Sub SetUserLogged(ByVal UserIndex As Integer)
        
        On Error GoTo SetUserLogged_Err
        

100     If Database_Enabled Then
102         Call SetUserLoggedDatabase(UserList(UserIndex).Id, UserList(UserIndex).AccountID)
        Else
104         Call WriteVar(CharPath & UCase$(UserList(UserIndex).name) & ".chr", "INIT", "Logged", 1)
106         Call WriteVar(CuentasPath & UCase$(UserList(UserIndex).Cuenta) & ".act", "INIT", "LOGEADA", 1)

        End If

        
        Exit Sub

SetUserLogged_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.SetUserLogged", Erl)
        Resume Next
        
End Sub

Sub SaveBattlePoints(ByVal UserIndex As Integer)
        
        On Error GoTo SaveBattlePoints_Err
        
    
100     If Database_Enabled Then
102         Call SaveBattlePointsDatabase(UserList(UserIndex).Id, UserList(UserIndex).flags.BattlePuntos)
        Else
104         Call WriteVar(CharPath & UserList(UserIndex).name & ".chr", "Battle", "Puntos", UserList(UserIndex).flags.BattlePuntos)

        End If
    
        
        Exit Sub

SaveBattlePoints_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.SaveBattlePoints", Erl)
        Resume Next
        
End Sub

Function Status(ByVal UserIndex As Integer) As Byte
        
        On Error GoTo Status_Err
        

100     Status = UserList(UserIndex).Faccion.Status

        
        Exit Function

Status_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.Status", Erl)
        Resume Next
        
End Function

Sub BackUPnPc(NpcIndex As Integer)
        
        On Error GoTo BackUPnPc_Err
        

        Dim NpcNumero As Integer

        Dim npcfile   As String

        Dim LoopC     As Integer

100     NpcNumero = Npclist(NpcIndex).Numero

        'If NpcNumero > 499 Then
        '    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
        'Else
102     npcfile = DatPath & "bkNPCs.dat"
        'End If

        'General
104     Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", Npclist(NpcIndex).name)
106     Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", Npclist(NpcIndex).Desc)
108     Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(Npclist(NpcIndex).Char.Head))
110     Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(Npclist(NpcIndex).Char.Body))
112     Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(Npclist(NpcIndex).Char.heading))
114     Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(Npclist(NpcIndex).Movement))
116     Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(Npclist(NpcIndex).Attackable))
118     Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(Npclist(NpcIndex).Comercia))
120     Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(Npclist(NpcIndex).TipoItems))
122     Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
124     Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(Npclist(NpcIndex).GiveEXP))
126     Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", val(Npclist(NpcIndex).GiveGLD))
128     Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
130     Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(Npclist(NpcIndex).InvReSpawn))
132     Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(Npclist(NpcIndex).NPCtype))

        'Stats
134     Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(Npclist(NpcIndex).Stats.Alineacion))
136     Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.def))
138     Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(Npclist(NpcIndex).Stats.MaxHit))
140     Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(Npclist(NpcIndex).Stats.MaxHp))
142     Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(Npclist(NpcIndex).Stats.MinHIT))
144     Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(Npclist(NpcIndex).Stats.MinHp))
146     Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.UsuariosMatados)) 'Que es ESTO?!!

        'Flags
148     Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(Npclist(NpcIndex).flags.Respawn))
150     Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(Npclist(NpcIndex).flags.backup))
152     Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(Npclist(NpcIndex).flags.Domable))

        'Inventario
154     Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(Npclist(NpcIndex).Invent.NroItems))

156     If Npclist(NpcIndex).Invent.NroItems > 0 Then

158         For LoopC = 1 To MAX_INVENTORY_SLOTS
160             Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & LoopC, Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex & "-" & Npclist(NpcIndex).Invent.Object(LoopC).Amount)
            Next

        End If

        
        Exit Sub

BackUPnPc_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.BackUPnPc", Erl)
        Resume Next
        
End Sub

Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NpcNumber As Integer)
        
        On Error GoTo CargarNpcBackUp_Err
        

        'Status
100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup Npc"

        Dim npcfile As String

        'If NpcNumber > 499 Then
        '    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
        'Else
102     npcfile = DatPath & "bkNPCs.dat"
        'End If

104     Npclist(NpcIndex).Numero = NpcNumber
106     Npclist(NpcIndex).name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
108     Npclist(NpcIndex).Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
110     Npclist(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
112     Npclist(NpcIndex).NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))

114     Npclist(NpcIndex).Char.Body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
116     Npclist(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
118     Npclist(NpcIndex).Char.heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))

120     Npclist(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
122     Npclist(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
124     Npclist(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
126     Npclist(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))

128     Npclist(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))

130     Npclist(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))

132     Npclist(NpcIndex).Stats.MaxHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
134     Npclist(NpcIndex).Stats.MinHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
136     Npclist(NpcIndex).Stats.MaxHit = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
138     Npclist(NpcIndex).Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
140     Npclist(NpcIndex).Stats.def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
142     Npclist(NpcIndex).Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))

        Dim LoopC As Integer

        Dim ln    As String

144     Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))

146     If Npclist(NpcIndex).Invent.NroItems > 0 Then

148         For LoopC = 1 To MAX_INVENTORY_SLOTS
150             ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
152             Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
154             Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
       
156         Next LoopC

        Else

158         For LoopC = 1 To MAX_INVENTORY_SLOTS
160             Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = 0
162             Npclist(NpcIndex).Invent.Object(LoopC).Amount = 0
164         Next LoopC

        End If

166     Npclist(NpcIndex).flags.NPCActive = True
168     Npclist(NpcIndex).flags.UseAINow = False
170     Npclist(NpcIndex).flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
172     Npclist(NpcIndex).flags.backup = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
174     Npclist(NpcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
176     Npclist(NpcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))

        'Tipo de items con los que comercia
178     Npclist(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

        
        Exit Sub

CargarNpcBackUp_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.CargarNpcBackUp", Erl)
        Resume Next
        
End Sub

Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal motivo As String)
        
        On Error GoTo LogBan_Err
        

100     Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).name, "BannedBy", UserList(UserIndex).name)
102     Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).name, "Reason", motivo)

        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim mifile As Integer

104     mifile = FreeFile
106     Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
108     Print #mifile, UserList(BannedIndex).name
110     Close #mifile

        
        Exit Sub

LogBan_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LogBan", Erl)
        Resume Next
        
End Sub

Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal motivo As String)
        
        On Error GoTo LogBanFromName_Err
        

100     Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).name)
102     Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)

        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim mifile As Integer

104     mifile = FreeFile
106     Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
108     Print #mifile, BannedName
110     Close #mifile

        
        Exit Sub

LogBanFromName_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LogBanFromName", Erl)
        Resume Next
        
End Sub

Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal motivo As String)
        
        On Error GoTo Ban_Err
        

100     Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
102     Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)

        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim mifile As Integer

104     mifile = FreeFile
106     Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
108     Print #mifile, BannedName
110     Close #mifile

        
        Exit Sub

Ban_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.Ban", Erl)
        Resume Next
        
End Sub

Public Sub CargaApuestas()
        
        On Error GoTo CargaApuestas_Err
        

100     Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
102     Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
104     Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

        
        Exit Sub

CargaApuestas_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.CargaApuestas", Erl)
        Resume Next
        
End Sub

Public Sub LoadRecursosEspeciales()
        
        On Error GoTo LoadRecursosEspeciales_Err
        

100     If Not FileExist(DatPath & "RecursosEspeciales.dat", vbArchive) Then
102         ReDim EspecialesTala(0) As obj
104         ReDim EspecialesPesca(0) As obj
            Exit Sub

        End If

        Dim IniFile As clsIniReader

106     Set IniFile = New clsIniReader
    
108     Call IniFile.Initialize(DatPath & "RecursosEspeciales.dat")
    
        Dim Count As Long, i As Long, str As String, Field() As String
    
        ' Tala
110     Count = val(IniFile.GetValue("Tala", "Items"))

112     If Count > 0 Then
114         ReDim EspecialesTala(1 To Count) As obj

116         For i = 1 To Count
118             str = IniFile.GetValue("Tala", "Item" & i)
120             Field = Split(str, "-")
            
122             EspecialesTala(i).ObjIndex = val(Field(0))
124             EspecialesTala(i).data = val(Field(1))      ' Probabilidad
            Next
        Else
126         ReDim EspecialesTala(0) As obj

        End If
    
        ' Pesca
128     Count = val(IniFile.GetValue("Pesca", "Items"))

130     If Count > 0 Then
132         ReDim EspecialesPesca(1 To Count) As obj

134         For i = 1 To Count
136             str = IniFile.GetValue("Pesca", "Item" & i)
138             Field = Split(str, "-")
            
140             EspecialesPesca(i).ObjIndex = val(Field(0))
142             EspecialesPesca(i).data = val(Field(1))     ' Probabilidad
            Next
        Else
144         ReDim EspecialesPesca(0) As obj

        End If
    
146     Set IniFile = Nothing

        
        Exit Sub

LoadRecursosEspeciales_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LoadRecursosEspeciales", Erl)
        Resume Next
        
End Sub

Public Sub LoadPesca()
        
        On Error GoTo LoadPesca_Err
        

100     If Not FileExist(DatPath & "pesca.dat", vbArchive) Then
102         ReDim Peces(0) As obj
104         ReDim PesoPeces(0) As Long
            Exit Sub

        End If

        Dim IniFile As clsIniReader

106     Set IniFile = New clsIniReader
    
108     Call IniFile.Initialize(DatPath & "pesca.dat")
    
        Dim Count As Long, i As Long, str As String, Field() As String, nivel As Integer, MaxLvlCania As Long

110     Count = val(IniFile.GetValue("PECES", "NumPeces"))
112     MaxLvlCania = val(IniFile.GetValue("PECES", "Maxlvlcaña"))
    
114     ReDim PesoPeces(0 To MaxLvlCania) As Long
    
116     If Count > 0 Then
118         ReDim Peces(1 To Count) As obj

            ' Cargo todos los peces
120         For i = 1 To Count
122             str = IniFile.GetValue("PECES", "Pez" & i)
124             Field = Split(str, "-")
            
126             Peces(i).ObjIndex = val(Field(0))
128             Peces(i).data = val(Field(1))       ' Peso

130             nivel = val(Field(2))               ' Nivel de caña

132             If (nivel > MaxLvlCania) Then nivel = MaxLvlCania
134             Peces(i).Amount = nivel
            Next

            ' Los ordeno segun nivel de caña (quick sort)
136         Call QuickSortPeces(1, Count)

            ' Sumo los pesos
138         For i = 1 To Count
140             PesoPeces(Peces(i).Amount) = PesoPeces(Peces(i).Amount) + Peces(i).data
142             Peces(i).data = PesoPeces(Peces(i).Amount)
            Next
        Else
144         ReDim Peces(0) As obj

        End If
    
146     For i = 1 To MaxLvlCania
148         PesoPeces(i) = PesoPeces(i) + PesoPeces(i - 1)
        Next
    
150     Set IniFile = Nothing

        
        Exit Sub

LoadPesca_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LoadPesca", Erl)
        Resume Next
        
End Sub

' Adaptado de https://www.vbforums.com/showthread.php?231925-VB-Quick-Sort-algorithm-(very-fast-sorting-algorithm)
Private Sub QuickSortPeces(ByVal First As Long, ByVal Last As Long)
        
        On Error GoTo QuickSortPeces_Err
        

        Dim low      As Long, high As Long

        Dim MidValue As String

        Dim aux      As obj
    
100     low = First
102     high = Last
104     MidValue = Peces((First + Last) \ 2).Amount
    
        Do

106         While Peces(low).Amount < MidValue

108             low = low + 1
            Wend

110         While Peces(high).Amount > MidValue

112             high = high - 1
            Wend

114         If low <= high Then
116             aux = Peces(low)
118             Peces(low) = Peces(high)
120             Peces(high) = aux
122             low = low + 1
124             high = high - 1

            End If

126     Loop While low <= high
    
128     If First < high Then QuickSortPeces First, high
130     If low < Last Then QuickSortPeces low, Last

        
        Exit Sub

QuickSortPeces_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.QuickSortPeces", Erl)
        Resume Next
        
End Sub

' Adaptado de https://www.freevbcode.com/ShowCode.asp?ID=9416
Public Function BinarySearchPeces(ByVal Value As Long) As Long
        
        On Error GoTo BinarySearchPeces_Err
        

        Dim low  As Long

        Dim high As Long

100     low = 1
102     high = UBound(Peces)

        Dim i              As Long

        Dim valor_anterior As Long
    
104     Do While low <= high
106         i = (low + high) \ 2

108         If i > 1 Then
110             valor_anterior = Peces(i - 1).data
            Else
112             valor_anterior = 0
            End If

114         If Value >= valor_anterior And Value < Peces(i).data Then
116             BinarySearchPeces = i
                Exit Do
            
118         ElseIf Value < valor_anterior Then
120             high = (i - 1)
            
            Else
122             low = (i + 1)

            End If

        Loop

        
        Exit Function

BinarySearchPeces_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.BinarySearchPeces", Erl)
        Resume Next
        
End Function

Public Sub LoadUserIntervals(ByVal UserIndex As Integer)
        
        On Error GoTo LoadUserIntervals_Err
        

100     With UserList(UserIndex).Intervals
102         .Arco = IntervaloFlechasCazadores
104         .Caminar = IntervaloCaminar
106         .Golpe = IntervaloUserPuedeAtacar
108         .magia = IntervaloUserPuedeCastear
110         .GolpeMagia = IntervaloGolpeMagia
112         .MagiaGolpe = IntervaloMagiaGolpe
113         .GolpeUsar = IntervaloGolpeUsar
114         .Trabajar = IntervaloUserPuedeTrabajar
116         .UsarU = IntervaloUserPuedeUsarU
            .UsarClic = IntervaloUserPuedeUsarClic

        End With

        
        Exit Sub

LoadUserIntervals_Err:
        Call RegistrarError(Err.Number, Err.description, "ES.LoadUserIntervals", Erl)
        Resume Next
        
End Sub

Public Sub RegistrarError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
'**********************************************************
'Author: Jopi
'Guarda una descripcion detallada del error en Errores.log
'**********************************************************
        
    'Si lo del parametro Componente es ES IGUAL, al Componente del anterior error...
    If Componente = HistorialError.Componente And _
       Numero = HistorialError.ErrorCode Then
        
        'Agregamos el error al historial.
        HistorialError.Contador = HistorialError.Contador + 1
        HistorialError.Componente = Componente
        HistorialError.ErrorCode = Numero
        
    Else 'Si NO es igual, reestablecemos el contador.

        HistorialError.Contador = 0
        HistorialError.ErrorCode = 0
        HistorialError.Componente = vbNullString
            
    End If
    
    'Si ya recibimos error en el mismo componente 10 veces, es bastante probable que estemos en un bucle
    'x lo que no hace falta registrar el error.
    If HistorialError.Contador = 10 Then Exit Sub
    
    'Registramos el error en Errores.log
    Dim File As Integer: File = FreeFile
        
    Open App.Path & "\logs\Errores.log" For Append As #File
    
        Print #File, "Error: " & Numero
        Print #File, "Descripcion: " & Descripcion
        
        If LenB(Linea) <> 0 Then
            Print #File, "Linea: " & Linea
        End If
        
        Print #File, "Componente: " & Componente
        Print #File, "Fecha y Hora: " & Date$ & "-" & Time$
        
        Print #File, vbNullString
        
    Close #File
    
    Debug.Print "Error: " & Numero & vbNewLine & _
                "Descripcion: " & Descripcion & vbNewLine & _
                "Componente: " & Componente & vbNewLine & _
                "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine
End Sub

