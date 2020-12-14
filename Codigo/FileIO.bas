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

    X As Integer
    Y As Integer

End Type

'Item type
Private Type tItem

    ObjIndex As Integer
    Amount As Integer

End Type

Private Type tWorldPos

    Map As Integer
    X As Byte
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

    X As Integer
    Y As Integer
    Lados As Byte

End Type

Private Type tDatosGrh

    X As Integer
    Y As Integer
    GrhIndex As Long

End Type

Private Type tDatosTrigger

    X As Integer
    Y As Integer
    trigger As Integer

End Type

Private Type tDatosLuces

    X As Integer
    Y As Integer
    Color As Long
    Rango As Byte

End Type

Private Type tDatosParticulas

    X As Integer
    Y As Integer
    Particula As Long

End Type

Private Type tDatosNPC

    X As Integer
    Y As Integer
    NpcIndex As Integer

End Type

Private Type tDatosObjs

    X As Integer
    Y As Integer
    ObjIndex As Integer
    ObjAmmount As Integer

End Type

Private Type tDatosTE

    X As Integer
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
102     Call RegistrarError(Err.Number, Err.description, "ES.EsAdmin", Erl)
104     Resume Next
        
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
102     Call RegistrarError(Err.Number, Err.description, "ES.EsDios", Erl)
104     Resume Next
        
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
102     Call RegistrarError(Err.Number, Err.description, "ES.EsSemiDios", Erl)
104     Resume Next
        
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
102     Call RegistrarError(Err.Number, Err.description, "ES.EsConsejero", Erl)
104     Resume Next
        
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
102     Call RegistrarError(Err.Number, Err.description, "ES.EsRolesMaster", Erl)
104     Resume Next
        
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
110     Call RegistrarError(Err.Number, Err.description, "ES.EsGmChar", Erl)
112     Resume Next
        
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
168     Call RegistrarError(Err.Number, Err.description, "ES.loadAdministrativeUsers", Erl)
170     Resume Next
        
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
120     Call RegistrarError(Err.Number, Err.description, "ES.GetCharPrivs", Erl)
122     Resume Next
        
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
116     Call RegistrarError(Err.Number, Err.description, "ES.TxtDimension", Erl)
118     Resume Next
        
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
120     Call RegistrarError(Err.Number, Err.description, "ES.CargarForbidenWords", Erl)
122     Resume Next
        
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

        On Error GoTo ErrHandler

100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."

        Dim Hechizo As Integer

        Dim Leer    As New clsIniReader

102     Call Leer.Initialize(DatPath & "Hechizos.dat")

        'obtiene el numero de hechizos
104     NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))

106     ReDim Hechizos(1 To NumeroHechizos) As tHechizo

108     frmCargando.cargar.min = 0
110     frmCargando.cargar.max = NumeroHechizos
112     frmCargando.cargar.Value = 0

        'Llena la lista
114     For Hechizo = 1 To NumeroHechizos

116         Hechizos(Hechizo).Velocidad = val(Leer.GetValue("Hechizo" & Hechizo, "Velocidad"))
    
            'Materializacion
118         Hechizos(Hechizo).MaterializaObj = val(Leer.GetValue("Hechizo" & Hechizo, "MaterializaObj"))
120         Hechizos(Hechizo).MaterializaCant = val(Leer.GetValue("Hechizo" & Hechizo, "MaterializaCant"))
            'Materializacion
    
            'Screen Efecto
122         Hechizos(Hechizo).ScreenColor = val(Leer.GetValue("Hechizo" & Hechizo, "ScreenColor"))
124         Hechizos(Hechizo).TimeEfect = val(Leer.GetValue("Hechizo" & Hechizo, "TimeEfect"))
            'Screen Efecto

126         Hechizos(Hechizo).TeleportX = val(Leer.GetValue("Hechizo" & Hechizo, "Teleport"))
128         Hechizos(Hechizo).TeleportXMap = val(Leer.GetValue("Hechizo" & Hechizo, "TeleportMap"))
130         Hechizos(Hechizo).TeleportXX = val(Leer.GetValue("Hechizo" & Hechizo, "TeleportX"))
132         Hechizos(Hechizo).TeleportXY = val(Leer.GetValue("Hechizo" & Hechizo, "TeleportY"))

134         Hechizos(Hechizo).nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
136         Hechizos(Hechizo).Desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
138         Hechizos(Hechizo).PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
    
140         Hechizos(Hechizo).HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
142         Hechizos(Hechizo).TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
144         Hechizos(Hechizo).PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
    
146         Hechizos(Hechizo).NecesitaObj = val(Leer.GetValue("Hechizo" & Hechizo, "NecesitaObj"))
148         Hechizos(Hechizo).NecesitaObj2 = val(Leer.GetValue("Hechizo" & Hechizo, "NecesitaObj2"))
    
150         Hechizos(Hechizo).Tipo = val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
152         Hechizos(Hechizo).wav = val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
154         Hechizos(Hechizo).FXgrh = val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
156         Hechizos(Hechizo).Particle = val(Leer.GetValue("Hechizo" & Hechizo, "Particle"))
158         Hechizos(Hechizo).ParticleViaje = val(Leer.GetValue("Hechizo" & Hechizo, "ParticleViaje"))
160         Hechizos(Hechizo).TimeParticula = val(Leer.GetValue("Hechizo" & Hechizo, "TimeParticula"))
162         Hechizos(Hechizo).desencantar = val(Leer.GetValue("Hechizo" & Hechizo, "desencantar"))
164         Hechizos(Hechizo).Sanacion = val(Leer.GetValue("Hechizo" & Hechizo, "Sanacion"))
166         Hechizos(Hechizo).AntiRm = val(Leer.GetValue("Hechizo" & Hechizo, "AntiRm"))
            'Hechizos de area
168         Hechizos(Hechizo).AreaRadio = val(Leer.GetValue("Hechizo" & Hechizo, "AreaRadio"))
170         Hechizos(Hechizo).AreaAfecta = val(Leer.GetValue("Hechizo" & Hechizo, "AreaAfecta"))
            'Hechizos de area
    
172         Hechizos(Hechizo).incinera = val(Leer.GetValue("Hechizo" & Hechizo, "Incinera"))
    
174         Hechizos(Hechizo).AutoLanzar = val(Leer.GetValue("Hechizo" & Hechizo, "AutoLanzar"))
    
176         Hechizos(Hechizo).CoolDown = val(Leer.GetValue("Hechizo" & Hechizo, "CoolDown"))
    
178         Hechizos(Hechizo).loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
    
            '    Hechizos(Hechizo).Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
    
180         Hechizos(Hechizo).SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
182         Hechizos(Hechizo).MinHp = val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
184         Hechizos(Hechizo).MaxHp = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
    
186         Hechizos(Hechizo).SubeMana = val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
188         Hechizos(Hechizo).MiMana = val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
190         Hechizos(Hechizo).MaMana = val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
    
192         Hechizos(Hechizo).SubeSta = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
194         Hechizos(Hechizo).MinSta = val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
196         Hechizos(Hechizo).MaxSta = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
    
198         Hechizos(Hechizo).SubeHam = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
200         Hechizos(Hechizo).MinHam = val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
202         Hechizos(Hechizo).MaxHam = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
    
204         Hechizos(Hechizo).SubeSed = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
206         Hechizos(Hechizo).MinSed = val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
208         Hechizos(Hechizo).MaxSed = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
    
210         Hechizos(Hechizo).SubeAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
212         Hechizos(Hechizo).MinAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
214         Hechizos(Hechizo).MaxAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
    
216         Hechizos(Hechizo).SubeFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
218         Hechizos(Hechizo).MinFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
220         Hechizos(Hechizo).MaxFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
    
222         Hechizos(Hechizo).SubeCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
224         Hechizos(Hechizo).MinCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
226         Hechizos(Hechizo).MaxCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
    
228         Hechizos(Hechizo).Invisibilidad = val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
230         Hechizos(Hechizo).Paraliza = val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
232         Hechizos(Hechizo).Inmoviliza = val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
    
234         Hechizos(Hechizo).RemoverParalisis = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
236         Hechizos(Hechizo).RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
238         Hechizos(Hechizo).RemueveInvisibilidadParcial = val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
    
240         Hechizos(Hechizo).CuraVeneno = val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
242         Hechizos(Hechizo).Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
244         Hechizos(Hechizo).Maldicion = val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
246         Hechizos(Hechizo).RemoverMaldicion = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
248         Hechizos(Hechizo).Bendicion = val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
250         Hechizos(Hechizo).Revivir = val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
    
252         Hechizos(Hechizo).Ceguera = val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
254         Hechizos(Hechizo).Estupidez = val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
    
256         Hechizos(Hechizo).Invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
258         Hechizos(Hechizo).NumNpc = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
260         Hechizos(Hechizo).cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
262         Hechizos(Hechizo).Mimetiza = val(Leer.GetValue("Hechizo" & Hechizo, "Mimetiza"))
    
264         Hechizos(Hechizo).GolpeCertero = val(Leer.GetValue("Hechizo" & Hechizo, "GolpeCertero"))
    
            '    Hechizos(Hechizo).Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
            '    Hechizos(Hechizo).ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
    
266         Hechizos(Hechizo).MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
268         Hechizos(Hechizo).ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
270         Hechizos(Hechizo).RequiredHP = val(Leer.GetValue("Hechizo" & Hechizo, "RequiredHP"))
    
272         Hechizos(Hechizo).Duration = val(Leer.GetValue("Hechizo" & Hechizo, "Duration"))
    
            'Barrin 30/9/03
274         Hechizos(Hechizo).StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
    
276         Hechizos(Hechizo).Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
278         frmCargando.cargar.Value = frmCargando.cargar.Value + 1
    
280         Hechizos(Hechizo).NeedStaff = val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
282         Hechizos(Hechizo).StaffAffected = CBool(val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))
    
284     Next Hechizo

286     Set Leer = Nothing
        Exit Sub

ErrHandler:
288     MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.description
 
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
112     Call RegistrarError(Err.Number, Err.description, "ES.LoadMotd", Erl)
114     Resume Next
        
End Sub

Public Sub DoBackUp()
        'Call LogTarea("Sub DoBackUp")
100     haciendoBK = True

        Dim i As Integer

102     Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

        'Call WorldSave
        'Call modGuilds.v_RutinaElecciones
    
        'Reseteamos al centinela
104     Call ResetCentinelaInfo

106     Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

        'Call EstadisticasWeb.Informar(EVENTO_NUEVO_CLAN, 0)

108     haciendoBK = False

        'Log
        On Error Resume Next

110     Dim nfile As Integer: nfile = FreeFile ' obtenemos un canal
112     Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
114     Print #nfile, Date & " " & Time
116     Close #nfile

End Sub

Public Sub GrabarMapa(ByVal Map As Long, ByVal MAPFILE As String)
100     Debug.Print "Empezamos a grabar"

        On Error GoTo ErrorHandler

102     Dim MapRoute As String: MapRoute = MAPFILE & ".csm"

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

104     For j = 1 To 100
106         For i = 1 To 100

108             With MapData(Map, i, j)
            
110                 If (.Blocked And eBlock.ALL_SIDES) <> 0 Then
112                     MH.NumeroBloqueados = MH.NumeroBloqueados + 1
114                     ReDim Preserve Blqs(1 To MH.NumeroBloqueados)
116                     Blqs(MH.NumeroBloqueados).X = i
118                     Blqs(MH.NumeroBloqueados).Y = j
120                     Blqs(MH.NumeroBloqueados).Lados = .Blocked And eBlock.ALL_SIDES
                    End If
            
                    Rem L1(i, j) = .Graphic(1).grhindex
  
122                 If .Graphic(1) > 0 Then
124                     MH.NumeroLayers(1) = MH.NumeroLayers(1) + 1
126                     ReDim Preserve L1(1 To MH.NumeroLayers(1))
128                     L1(MH.NumeroLayers(1)).X = i
130                     L1(MH.NumeroLayers(1)).Y = j
132                     L1(MH.NumeroLayers(1)).GrhIndex = .Graphic(1)

                    End If
            
134                 If .Graphic(2) > 0 Then
136                     MH.NumeroLayers(2) = MH.NumeroLayers(2) + 1
138                     ReDim Preserve L2(1 To MH.NumeroLayers(2))
140                     L2(MH.NumeroLayers(2)).X = i
142                     L2(MH.NumeroLayers(2)).Y = j
144                     L2(MH.NumeroLayers(2)).GrhIndex = .Graphic(2)

                    End If
            
146                 If .Graphic(3) > 0 Then
148                     MH.NumeroLayers(3) = MH.NumeroLayers(3) + 1
150                     ReDim Preserve L3(1 To MH.NumeroLayers(3))
152                     L3(MH.NumeroLayers(3)).X = i
154                     L3(MH.NumeroLayers(3)).Y = j
156                     L3(MH.NumeroLayers(3)).GrhIndex = .Graphic(3)

                    End If
            
158                 If .Graphic(4) > 0 Then
160                     MH.NumeroLayers(4) = MH.NumeroLayers(4) + 1
162                     ReDim Preserve L4(1 To MH.NumeroLayers(4))
164                     L4(MH.NumeroLayers(4)).X = i
166                     L4(MH.NumeroLayers(4)).Y = j
168                     L4(MH.NumeroLayers(4)).GrhIndex = .Graphic(4)

                    End If
            
170                 If .trigger > 0 Then
172                     MH.NumeroTriggers = MH.NumeroTriggers + 1
174                     ReDim Preserve Triggers(1 To MH.NumeroTriggers)
176                     Triggers(MH.NumeroTriggers).X = i
178                     Triggers(MH.NumeroTriggers).Y = j
180                     Triggers(MH.NumeroTriggers).trigger = .trigger

                    End If
            
182                 If .ParticulaIndex > 0 Then
184                     MH.NumeroParticulas = MH.NumeroParticulas + 1
186                     ReDim Preserve Particulas(1 To MH.NumeroParticulas)
188                     Particulas(MH.NumeroParticulas).X = i
190                     Particulas(MH.NumeroParticulas).Y = j
192                     Particulas(MH.NumeroParticulas).Particula = .ParticulaIndex

                    End If
            
                    Rem   If MapData(i, j).luz.Rango > 0 Then
                    Rem      MH.NumeroLuces = MH.NumeroLuces + 1
                    Rem       ReDim Preserve Luces(1 To MH.NumeroLuces)
                    Rem       Luces(MH.NumeroLuces).X = i
                    Rem       Luces(MH.NumeroLuces).Y = j
                    Rem      Luces(MH.NumeroLuces).color = .luz.color
                    Rem       Luces(MH.NumeroLuces).Rango = .luz.Rango
                    Rem  End If
            
194                 If .ObjInfo.ObjIndex > 0 Then
196                     MH.NumeroOBJs = MH.NumeroOBJs + 1
198                     ReDim Preserve Objetos(1 To MH.NumeroOBJs)
200                     Objetos(MH.NumeroOBJs).ObjIndex = .ObjInfo.ObjIndex
202                     Objetos(MH.NumeroOBJs).ObjAmmount = .ObjInfo.Amount
               
204                     Objetos(MH.NumeroOBJs).X = i
206                     Objetos(MH.NumeroOBJs).Y = j
                
                    End If
            
208                 If .NpcIndex > 0 Then
210                     MH.NumeroNPCs = MH.NumeroNPCs + 1
212                     ReDim Preserve NPCs(1 To MH.NumeroNPCs)
214                     NPCs(MH.NumeroNPCs).NpcIndex = .NpcIndex
216                     NPCs(MH.NumeroNPCs).X = i
218                     NPCs(MH.NumeroNPCs).Y = j

                    End If
            
220                 If .TileExit.Map > 0 Then
222                     MH.NumeroTE = MH.NumeroTE + 1
224                     ReDim Preserve TEs(1 To MH.NumeroTE)
226                     TEs(MH.NumeroTE).DestM = .TileExit.Map
228                     TEs(MH.NumeroTE).DestX = .TileExit.X
230                     TEs(MH.NumeroTE).DestY = .TileExit.Y
232                     TEs(MH.NumeroTE).X = i
234                     TEs(MH.NumeroTE).Y = j

                    End If

                End With

236         Next i
238     Next j
          
240     fh = FreeFile
242     Open MapRoute For Binary As fh
    
244     Put #fh, , MH
246     Put #fh, , MapSize
248     Put #fh, , MapDat
        Rem   Put #fh, , L1
    
250     With MH

252         If .NumeroBloqueados > 0 Then Put #fh, , Blqs

254         If .NumeroLayers(1) > 0 Then Put #fh, , L1

256         If .NumeroLayers(2) > 0 Then Put #fh, , L2

258         If .NumeroLayers(3) > 0 Then Put #fh, , L3

260         If .NumeroLayers(4) > 0 Then Put #fh, , L4

262         If .NumeroTriggers > 0 Then Put #fh, , Triggers

264         If .NumeroParticulas > 0 Then Put #fh, , Particulas

266         If .NumeroLuces > 0 Then Put #fh, , Luces

268         If .NumeroOBJs > 0 Then Put #fh, , Objetos

270         If .NumeroNPCs > 0 Then Put #fh, , NPCs

272         If .NumeroTE > 0 Then Put #fh, , TEs

        End With

274     Close fh

        Rem MsgBox "Mapa grabado"

276     Debug.Print "Mapa grabado"

ErrorHandler:

278     If fh <> 0 Then Close fh

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
114     Call RegistrarError(Err.Number, Err.description, "ES.LoadArmasHerreria", Erl)
116     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.description, "ES.LoadArmadurasHerreria", Erl)
116     Resume Next
        
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
106         SearchVar = Replace$(Tilde(ListaClases(i)), " ", vbNullString)

108         With ModClase(i)
110             .Evasion = val(BalanceIni.GetValue("MODEVASION", SearchVar))
112             .AtaqueArmas = val(BalanceIni.GetValue("MODATAQUEARMAS", SearchVar))
114             .AtaqueProyectiles = val(BalanceIni.GetValue("MODATAQUEPROYECTILES", SearchVar))
                '.DañoWrestling = val(BalanceIni.GetValue("MODATAQUEWRESTLING", SearchVar))
116             .DañoArmas = val(BalanceIni.GetValue("MODDANOARMAS", SearchVar))
118             .DañoProyectiles = val(BalanceIni.GetValue("MODDANOPROYECTILES", SearchVar))
120             .DañoWrestling = val(BalanceIni.GetValue("MODDANOWRESTLING", SearchVar))
122             .Escudo = val(BalanceIni.GetValue("MODESCUDO", SearchVar))
124             .ModApuñalar = val(BalanceIni.GetValue("MODAPUÑALAR", SearchVar))
                'Modificadores de Vida
126             ModVida(i) = val(BalanceIni.GetValue("MODVIDA", SearchVar))

            End With

128     Next i
    
        'Modificadores de Raza
130     For i = 1 To NUMRAZAS
132         SearchVar = Replace$(Tilde(ListaRazas(i)), " ", vbNullString)

134         With ModRaza(i)
136             .Fuerza = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Fuerza"))
138             .Agilidad = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Agilidad"))
140             .Inteligencia = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Inteligencia"))
142             .Carisma = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Carisma"))
144             .Constitucion = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Constitucion"))
            End With

146     Next i
    
        'Distribucion de Vida
148     For i = 1 To 5
150         DistribucionEnteraVida(i) = val(BalanceIni.GetValue("DISTRIBUCION", "E" + CStr(i)))
152     Next i

154     For i = 1 To 4
156         DistribucionSemienteraVida(i) = val(BalanceIni.GetValue("DISTRIBUCION", "S" + CStr(i)))
158     Next i
    
            'Experiencia por nivel
160     For i = 1 To STAT_MAXELV
162         ExpByLevel(i) = val(BalanceIni.GetValue("EXPBYLEVEL", i))
164     Next i

        'Extra
166     PorcentajeRecuperoMana = val(BalanceIni.GetValue("EXTRA", "PorcentajeRecuperoMana"))
168     DificultadSubirSkill = val(BalanceIni.GetValue("EXTRA", "DificultadSubirSkill"))
    
170     Set BalanceIni = Nothing
    
172     AgregarAConsola "Se cargó el balance (Balance.dat)"

        
        Exit Sub

LoadBalance_Err:
174     Call RegistrarError(Err.Number, Err.description, "ES.LoadBalance", Erl)
176     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.description, "ES.LoadObjCarpintero", Erl)
116     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.description, "ES.LoadObjAlquimista", Erl)
116     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.description, "ES.LoadObjSastre", Erl)
116     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.description, "ES.LoadObjDonador", Erl)
116     Resume Next
        
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

        On Error GoTo ErrHandler

100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."

        '*****************************************************************
        'Carga la lista de objetos
        '*****************************************************************
        Dim Object As Integer

        Dim Leer   As clsIniReader
102     Set Leer = New clsIniReader
104     Call Leer.Initialize(DatPath & "Obj.dat")

        'obtiene el numero de obj
106     NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))

108     frmCargando.cargar.min = 0
110     frmCargando.cargar.max = NumObjDatas
112     frmCargando.cargar.Value = 0

114     ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
    
        Dim str As String, Field() As String
  
        'Llena la lista
116     For Object = 1 To NumObjDatas
        
118         ObjData(Object).name = Leer.GetValue("OBJ" & Object, "Name")
    
            ' If ObjData(Object).Name = "" Then
            '   Call LogError("Objeto libre:" & Object)
            ' End If
    
            ' If ObjData(Object).name = "" Then
            ' Debug.Print Object
            ' End If
    
            'Pablo (ToxicWaste) Log de Objetos.
120         ObjData(Object).Log = val(Leer.GetValue("OBJ" & Object, "Log"))
122         ObjData(Object).NoLog = val(Leer.GetValue("OBJ" & Object, "NoLog"))
            '07/09/07
    
124         ObjData(Object).GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex"))

126         If ObjData(Object).GrhIndex = 0 Then
128             ObjData(Object).GrhIndex = ObjData(Object).GrhIndex

            End If
    
130         ObjData(Object).OBJType = val(Leer.GetValue("OBJ" & Object, "ObjType"))
    
132         ObjData(Object).Newbie = val(Leer.GetValue("OBJ" & Object, "Newbie"))
            'Propiedades by Lader 05-05-08
134         ObjData(Object).Instransferible = val(Leer.GetValue("OBJ" & Object, "Instransferible"))
136         ObjData(Object).Destruye = val(Leer.GetValue("OBJ" & Object, "Destruye"))
138         ObjData(Object).Intirable = val(Leer.GetValue("OBJ" & Object, "Intirable"))
    
140         ObjData(Object).CantidadSkill = val(Leer.GetValue("OBJ" & Object, "CantidadSkill"))
142         ObjData(Object).QueSkill = val(Leer.GetValue("OBJ" & Object, "QueSkill"))
144         ObjData(Object).QueAtributo = val(Leer.GetValue("OBJ" & Object, "queatributo"))
146         ObjData(Object).CuantoAumento = val(Leer.GetValue("OBJ" & Object, "cuantoaumento"))
148         ObjData(Object).MinELV = val(Leer.GetValue("OBJ" & Object, "MinELV"))
150         ObjData(Object).Subtipo = val(Leer.GetValue("OBJ" & Object, "Subtipo"))
152         ObjData(Object).Dorada = val(Leer.GetValue("OBJ" & Object, "Dorada"))
154         ObjData(Object).VidaUtil = val(Leer.GetValue("OBJ" & Object, "VidaUtil"))
156         ObjData(Object).TiempoRegenerar = val(Leer.GetValue("OBJ" & Object, "TiempoRegenerar"))
    
158         ObjData(Object).donador = val(Leer.GetValue("OBJ" & Object, "donador"))
    
            Dim i As Integer

            'Propiedades by Lader 05-05-08
160         Select Case ObjData(Object).OBJType

                Case eOBJType.otHerramientas
162                 ObjData(Object).WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
164                 ObjData(Object).Power = val(Leer.GetValue("OBJ" & Object, "Poder"))
            
166             Case eOBJType.otArmadura
168                 ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
170                 ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
172                 ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
174                 ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
176                 ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
178                 ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
180                 ObjData(Object).ResistenciaMagica = val(Leer.GetValue("OBJ" & Object, "ResistenciaMagica"))
        
182             Case eOBJType.otESCUDO
184                 ObjData(Object).ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
186                 ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
188                 ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
190                 ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
192                 ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
194                 ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
196                 ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
198                 ObjData(Object).ResistenciaMagica = val(Leer.GetValue("OBJ" & Object, "ResistenciaMagica"))
        
200             Case eOBJType.otCASCO
202                 ObjData(Object).CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
204                 ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
206                 ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
208                 ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
210                 ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
212                 ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
214                 ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
216                 ObjData(Object).ResistenciaMagica = val(Leer.GetValue("OBJ" & Object, "ResistenciaMagica"))
        
218             Case eOBJType.otWeapon
220                 ObjData(Object).WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
222                 ObjData(Object).Apuñala = val(Leer.GetValue("OBJ" & Object, "Apuñala"))
224                 ObjData(Object).Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
226                 ObjData(Object).Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
228                 ObjData(Object).Estupidiza = val(Leer.GetValue("OBJ" & Object, "Estupidiza"))
        
230                 ObjData(Object).MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
232                 ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
234                 ObjData(Object).proyectil = val(Leer.GetValue("OBJ" & Object, "Proyectil"))
236                 ObjData(Object).Municion = val(Leer.GetValue("OBJ" & Object, "Municiones"))
238                 ObjData(Object).Power = val(Leer.GetValue("OBJ" & Object, "StaffPower"))
240                 ObjData(Object).MagicDamageBonus = val(Leer.GetValue("OBJ" & Object, "MagicDamageBonus"))
242                 ObjData(Object).Refuerzo = val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
            
244                 ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
246                 ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
248                 ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
250                 ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
252                 ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
254                 ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
256                 ObjData(Object).EfectoMagico = val(Leer.GetValue("OBJ" & Object, "efectomagico"))
        
258             Case eOBJType.otInstrumentos
        
                    'Pablo (ToxicWaste)
260                 ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
262                 ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
264             Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
266                 ObjData(Object).IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
268                 ObjData(Object).IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
270                 ObjData(Object).IndexCerradaLlave = val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
        
272             Case otPociones
274                 ObjData(Object).TipoPocion = val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
276                 ObjData(Object).MaxModificador = val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
278                 ObjData(Object).MinModificador = val(Leer.GetValue("OBJ" & Object, "MinModificador"))
            
280                 ObjData(Object).DuracionEfecto = val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
282                 ObjData(Object).Raices = val(Leer.GetValue("OBJ" & Object, "Raices"))
284                 ObjData(Object).SkPociones = val(Leer.GetValue("OBJ" & Object, "SkPociones"))
286                 ObjData(Object).Porcentaje = val(Leer.GetValue("OBJ" & Object, "Porcentaje"))
        
288             Case eOBJType.otBarcos
290                 ObjData(Object).MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
292                 ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
294                 ObjData(Object).Subtipo = val(Leer.GetValue("OBJ" & Object, "Subtipo"))
296                 ObjData(Object).Velocidad = val(Leer.GetValue("OBJ" & Object, "Velocidad"))

298             Case eOBJType.otMonturas
300                 ObjData(Object).MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
302                 ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
304                 ObjData(Object).MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
306                 ObjData(Object).MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
308                 ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
310                 ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
312             Case eOBJType.otFlechas
314                 ObjData(Object).MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
316                 ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
318                 ObjData(Object).Subtipo = val(Leer.GetValue("OBJ" & Object, "Subtipo"))
            
320                 ObjData(Object).Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
322                 ObjData(Object).Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
            
                    'Case eOBJType.otAnillos 'Pablo (ToxicWaste)
                    '  ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    '  ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    '  ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    '  ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            
                    'Pasajes Ladder 05-05-08
324             Case eOBJType.otpasajes
326                 ObjData(Object).DesdeMap = val(Leer.GetValue("OBJ" & Object, "DesdeMap"))
328                 ObjData(Object).HastaMap = val(Leer.GetValue("OBJ" & Object, "Map"))
330                 ObjData(Object).HastaX = val(Leer.GetValue("OBJ" & Object, "X"))
332                 ObjData(Object).HastaY = val(Leer.GetValue("OBJ" & Object, "Y"))
334                 ObjData(Object).NecesitaNave = val(Leer.GetValue("OBJ" & Object, "NecesitaNave"))
            
336             Case eOBJType.OtDonador
338                 ObjData(Object).Subtipo = val(Leer.GetValue("OBJ" & Object, "Subtipo"))
340                 ObjData(Object).HastaMap = val(Leer.GetValue("OBJ" & Object, "Map"))
342                 ObjData(Object).HastaX = val(Leer.GetValue("OBJ" & Object, "X"))
344                 ObjData(Object).HastaY = val(Leer.GetValue("OBJ" & Object, "Y"))
        
346             Case eOBJType.otmagicos
348                 ObjData(Object).EfectoMagico = val(Leer.GetValue("OBJ" & Object, "efectomagico"))

350                 If ObjData(Object).EfectoMagico = 15 Then
352                     PENDIENTE = Object

                    End If
            
354             Case eOBJType.otRunas
356                 ObjData(Object).TipoRuna = val(Leer.GetValue("OBJ" & Object, "TipoRuna"))
358                 ObjData(Object).DesdeMap = val(Leer.GetValue("OBJ" & Object, "DesdeMap"))
360                 ObjData(Object).HastaMap = val(Leer.GetValue("OBJ" & Object, "Map"))
362                 ObjData(Object).HastaX = val(Leer.GetValue("OBJ" & Object, "X"))
364                 ObjData(Object).HastaY = val(Leer.GetValue("OBJ" & Object, "Y"))
                    
366             Case eOBJType.otNUDILLOS
368                 ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
370                 ObjData(Object).MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHit"))
372                 ObjData(Object).Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
374                 ObjData(Object).Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
376                 ObjData(Object).Estupidiza = val(Leer.GetValue("OBJ" & Object, "Estupidiza"))
378                 ObjData(Object).WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
380                 ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            
382             Case eOBJType.otPergaminos
        
                    ' ObjData(Object).ClasePermitida = Leer.GetValue("OBJ" & Object, "CP")
        
384             Case eOBJType.OtCofre
386                 ObjData(Object).CantItem = val(Leer.GetValue("OBJ" & Object, "CantItem"))
388                 ObjData(Object).Subtipo = val(Leer.GetValue("OBJ" & Object, "SubTipo"))

390                 If ObjData(Object).Subtipo = 1 Then
392                     ReDim ObjData(Object).Item(1 To ObjData(Object).CantItem)
                
394                     For i = 1 To ObjData(Object).CantItem
396                         ObjData(Object).Item(i).ObjIndex = val(Leer.GetValue("OBJ" & Object, "Item" & i))
398                         ObjData(Object).Item(i).Amount = val(Leer.GetValue("OBJ" & Object, "Cantidad" & i))
400                     Next i

                    Else
402                     ReDim ObjData(Object).Item(1 To ObjData(Object).CantItem)
                
404                     ObjData(Object).CantEntrega = val(Leer.GetValue("OBJ" & Object, "CantEntrega"))

406                     For i = 1 To ObjData(Object).CantItem
408                         ObjData(Object).Item(i).ObjIndex = val(Leer.GetValue("OBJ" & Object, "Item" & i))
410                         ObjData(Object).Item(i).Amount = val(Leer.GetValue("OBJ" & Object, "Cantidad" & i))
412                     Next i

                    End If
            
414             Case eOBJType.otYacimiento
                    ' Drop gemas yacimientos
416                 ObjData(Object).CantItem = val(Leer.GetValue("OBJ" & Object, "Gemas"))
            
418                 If ObjData(Object).CantItem > 0 Then
420                     ReDim ObjData(Object).Item(1 To ObjData(Object).CantItem)

422                     For i = 1 To ObjData(Object).CantItem
424                         str = Leer.GetValue("OBJ" & Object, "Gema" & i)
426                         Field = Split(str, "-")
428                         ObjData(Object).Item(i).ObjIndex = val(Field(0))    ' ObjIndex
430                         ObjData(Object).Item(i).Amount = val(Field(1))      ' Probabilidad de drop (1 en X)
432                     Next i

                    End If
                
434             Case eOBJType.otAnillos
436                 ObjData(Object).MagicDamageBonus = val(Leer.GetValue("OBJ" & Object, "MagicDamageBonus"))
438                 ObjData(Object).ResistenciaMagica = val(Leer.GetValue("OBJ" & Object, "ResistenciaMagica"))
            
            End Select
    
440         ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))

442         ObjData(Object).Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
444         ObjData(Object).Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
446         ObjData(Object).Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
            'DELETE
448         ObjData(Object).SndAura = val(Leer.GetValue("OBJ" & Object, "SndAura"))
            '
    
450         ObjData(Object).NoSeLimpia = val(Leer.GetValue("OBJ" & Object, "NoSeLimpia"))
452         ObjData(Object).Subastable = val(Leer.GetValue("OBJ" & Object, "Subastable"))
    
454         ObjData(Object).ParticulaGolpe = val(Leer.GetValue("OBJ" & Object, "ParticulaGolpe"))
456         ObjData(Object).ParticulaViaje = val(Leer.GetValue("OBJ" & Object, "ParticulaViaje"))
458         ObjData(Object).ParticulaGolpeTime = val(Leer.GetValue("OBJ" & Object, "ParticulaGolpeTime"))
    
460         ObjData(Object).Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
462         ObjData(Object).HechizoIndex = val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
    
464         ObjData(Object).LingoteIndex = val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
    
466         ObjData(Object).MineralIndex = val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
    
468         ObjData(Object).MaxHp = val(Leer.GetValue("OBJ" & Object, "MaxHP"))
470         ObjData(Object).MinHp = val(Leer.GetValue("OBJ" & Object, "MinHP"))
    
472         ObjData(Object).Mujer = val(Leer.GetValue("OBJ" & Object, "Mujer"))
474         ObjData(Object).Hombre = val(Leer.GetValue("OBJ" & Object, "Hombre"))
    
476         ObjData(Object).PielLobo = val(Leer.GetValue("OBJ" & Object, "PielLobo"))
478         ObjData(Object).PielOsoPardo = val(Leer.GetValue("OBJ" & Object, "PielOsoPardo"))
480         ObjData(Object).PielOsoPolaR = val(Leer.GetValue("OBJ" & Object, "PielOsoPolaR"))
482         ObjData(Object).SkMAGOria = val(Leer.GetValue("OBJ" & Object, "SKSastreria"))
    
484         ObjData(Object).CreaParticula = Leer.GetValue("OBJ" & Object, "CreaParticula")
    
486         ObjData(Object).CreaFX = val(Leer.GetValue("OBJ" & Object, "CreaFX"))
  
            'DELETE
488         ObjData(Object).CreaParticulaPiso = val(Leer.GetValue("OBJ" & Object, "CreaParticulaPiso"))
            '
    
490         ObjData(Object).CreaGRH = Leer.GetValue("OBJ" & Object, "CreaGRH")
492         ObjData(Object).CreaLuz = Leer.GetValue("OBJ" & Object, "CreaLuz")
    
494         ObjData(Object).MinHam = val(Leer.GetValue("OBJ" & Object, "MinHam"))
496         ObjData(Object).MinSed = val(Leer.GetValue("OBJ" & Object, "MinAgu"))
    
498         ObjData(Object).MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
500         ObjData(Object).MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
502         ObjData(Object).def = (ObjData(Object).MinDef + ObjData(Object).MaxDef) / 2
    
504         ObjData(Object).ClaseTipo = val(Leer.GetValue("OBJ" & Object, "ClaseTipo"))
506         ObjData(Object).RazaTipo = val(Leer.GetValue("OBJ" & Object, "RazaTipo"))

508         ObjData(Object).RazaEnana = val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
510         ObjData(Object).RazaDrow = val(Leer.GetValue("OBJ" & Object, "RazaDrow"))
512         ObjData(Object).RazaElfa = val(Leer.GetValue("OBJ" & Object, "RazaElfa"))
514         ObjData(Object).RazaGnoma = val(Leer.GetValue("OBJ" & Object, "RazaGnoma"))
    
516         ObjData(Object).RazaOrca = val(Leer.GetValue("OBJ" & Object, "RazaOrca"))
    
518         ObjData(Object).RazaHumana = val(Leer.GetValue("OBJ" & Object, "RazaHumana"))
    
520         ObjData(Object).Valor = val(Leer.GetValue("OBJ" & Object, "Valor"))
    
522         ObjData(Object).Crucial = val(Leer.GetValue("OBJ" & Object, "Crucial"))
    
            'ObjData(Object).Cerrada = val(Leer.GetValue("OBJ" & Object, "abierta")) cerrada = abierta??? WTF???????
524         ObjData(Object).Cerrada = val(Leer.GetValue("OBJ" & Object, "Cerrada"))

526         If ObjData(Object).Cerrada = 1 Then
528             ObjData(Object).Llave = val(Leer.GetValue("OBJ" & Object, "Llave"))
530             ObjData(Object).clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
            End If
    
            'Puertas y llaves
532         ObjData(Object).clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
    
534         ObjData(Object).texto = Leer.GetValue("OBJ" & Object, "Texto")
536         ObjData(Object).GrhSecundario = val(Leer.GetValue("OBJ" & Object, "VGrande"))
    
538         ObjData(Object).Agarrable = val(Leer.GetValue("OBJ" & Object, "Agarrable"))
540         ObjData(Object).ForoID = Leer.GetValue("OBJ" & Object, "ID")
    
            'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico  -  Nunca más papu
            Dim n As Integer
            Dim S As String

542         For i = 1 To NUMCLASES
544             S = UCase$(Leer.GetValue("OBJ" & Object, "CP" & i))
546             n = 1

548             Do While LenB(S) > 0 And Tilde(ListaClases(n)) <> Trim$(S)
550                 n = n + 1
                Loop
552             ObjData(Object).ClaseProhibida(i) = IIf(LenB(S) > 0, n, 0)
554         Next i
        
            ' Skill requerido
556         str = Leer.GetValue("OBJ" & Object, "SkillRequerido")

558         If Len(str) > 0 Then
560             Field = Split(str, "-")
            
562             n = 1
564             Do While LenB(Field(0)) > 0 And Tilde(SkillsNames(n)) <> Tilde(Field(0))
566                 n = n + 1
                Loop
    
568             ObjData(Object).SkillIndex = IIf(LenB(Field(0)) > 0, n, 0)
570             ObjData(Object).SkillRequerido = val(Field(1))
            End If
            ' -----------------
    
572         ObjData(Object).DefensaMagicaMax = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
574         ObjData(Object).DefensaMagicaMin = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
    
576         ObjData(Object).SkCarpinteria = val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
    
            'If ObjData(Object).SkCarpinteria > 0 Then
578         ObjData(Object).Madera = val(Leer.GetValue("OBJ" & Object, "Madera"))
    
            'Bebidas
580         ObjData(Object).MinSta = val(Leer.GetValue("OBJ" & Object, "MinST"))
    
582         ObjData(Object).NoSeCae = val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
    
584         frmCargando.cargar.Value = frmCargando.cargar.Value + 1
586     Next Object

588     Set Leer = Nothing

        Exit Sub

ErrHandler:
590     MsgBox "error cargando objetos " & Err.Number & ": " & Err.description & ". Error producido al cargar el objeto: " & Object

End Sub

Sub LoadUserStats(ByVal Userindex As Integer, ByRef UserFile As clsIniReader)
        
        On Error GoTo LoadUserStats_Err
        

        Dim LoopC As Long

100     For LoopC = 1 To NUMATRIBUTOS
102         UserList(Userindex).Stats.UserAtributos(LoopC) = CInt(UserFile.GetValue("ATRIBUTOS", "AT" & LoopC))
104         UserList(Userindex).Stats.UserAtributosBackUP(LoopC) = UserList(Userindex).Stats.UserAtributos(LoopC)
106     Next LoopC

108     For LoopC = 1 To NUMSKILLS
110         UserList(Userindex).Stats.UserSkills(LoopC) = CInt(UserFile.GetValue("SKILLS", "SK" & LoopC))
112     Next LoopC

114     For LoopC = 1 To MAXUSERHECHIZOS
116         UserList(Userindex).Stats.UserHechizos(LoopC) = CInt(UserFile.GetValue("Hechizos", "H" & LoopC))
118     Next LoopC

120     UserList(Userindex).Stats.GLD = CLng(UserFile.GetValue("STATS", "GLD"))
122     UserList(Userindex).Stats.Banco = CLng(UserFile.GetValue("STATS", "BANCO"))

124     UserList(Userindex).Stats.MaxHp = CInt(UserFile.GetValue("STATS", "MaxHP"))
126     UserList(Userindex).Stats.MinHp = CInt(UserFile.GetValue("STATS", "MinHP"))

128     UserList(Userindex).Stats.MinSta = CInt(UserFile.GetValue("STATS", "MinSTA"))
130     UserList(Userindex).Stats.MaxSta = CInt(UserFile.GetValue("STATS", "MaxSTA"))

132     UserList(Userindex).Stats.MaxMAN = CInt(UserFile.GetValue("STATS", "MaxMAN"))
134     UserList(Userindex).Stats.MinMAN = CInt(UserFile.GetValue("STATS", "MinMAN"))

136     UserList(Userindex).Stats.MaxHit = CInt(UserFile.GetValue("STATS", "MaxHIT"))
138     UserList(Userindex).Stats.MinHIT = CInt(UserFile.GetValue("STATS", "MinHIT"))

140     UserList(Userindex).Stats.MaxAGU = CByte(UserFile.GetValue("STATS", "MaxAGU"))
142     UserList(Userindex).Stats.MinAGU = CByte(UserFile.GetValue("STATS", "MinAGU"))

144     UserList(Userindex).Stats.MaxHam = CByte(UserFile.GetValue("STATS", "MaxHAM"))
146     UserList(Userindex).Stats.MinHam = CByte(UserFile.GetValue("STATS", "MinHAM"))

148     UserList(Userindex).Stats.SkillPts = CInt(UserFile.GetValue("STATS", "SkillPtsLibres"))

150     UserList(Userindex).Stats.Exp = CDbl(UserFile.GetValue("STATS", "EXP"))
152     UserList(Userindex).Stats.ELU = CLng(UserFile.GetValue("STATS", "ELU"))
154     UserList(Userindex).Stats.ELV = CByte(UserFile.GetValue("STATS", "ELV"))

156     UserList(Userindex).flags.Envenena = CByte(UserFile.GetValue("MAGIA", "ENVENENA"))
158     UserList(Userindex).flags.Paraliza = CByte(UserFile.GetValue("MAGIA", "PARALIZA"))
160     UserList(Userindex).flags.incinera = CByte(UserFile.GetValue("MAGIA", "INCINERA")) 'Estupidiza
162     UserList(Userindex).flags.Estupidiza = CByte(UserFile.GetValue("MAGIA", "Estupidiza"))

164     UserList(Userindex).flags.PendienteDelSacrificio = CByte(UserFile.GetValue("MAGIA", "PENDIENTE"))
166     UserList(Userindex).flags.CarroMineria = CByte(UserFile.GetValue("MAGIA", "CarroMineria"))
168     UserList(Userindex).flags.NoPalabrasMagicas = CByte(UserFile.GetValue("MAGIA", "NOPALABRASMAGICAS"))

170     If UserList(Userindex).flags.Muerto = 0 Then
172         UserList(Userindex).Char.Otra_Aura = CStr(UserFile.GetValue("MAGIA", "OTRA_AURA"))

        End If

        'UserList(UserIndex).flags.DañoMagico = CByte(UserFile.GetValue("MAGIA", "DañoMagico"))
        'UserList(UserIndex).flags.ResistenciaMagica = CByte(UserFile.GetValue("MAGIA", "ResistenciaMagica"))

        'Nuevos
174     UserList(Userindex).flags.RegeneracionMana = CByte(UserFile.GetValue("MAGIA", "RegeneracionMana"))
176     UserList(Userindex).flags.AnilloOcultismo = CByte(UserFile.GetValue("MAGIA", "AnilloOcultismo"))
178     UserList(Userindex).flags.NoDetectable = CByte(UserFile.GetValue("MAGIA", "NoDetectable"))
180     UserList(Userindex).flags.NoMagiaEfeceto = CByte(UserFile.GetValue("MAGIA", "NoMagiaEfeceto"))
182     UserList(Userindex).flags.RegeneracionHP = CByte(UserFile.GetValue("MAGIA", "RegeneracionHP"))
184     UserList(Userindex).flags.RegeneracionSta = CByte(UserFile.GetValue("MAGIA", "RegeneracionSta"))

186     UserList(Userindex).Stats.UsuariosMatados = CLng(UserFile.GetValue("MUERTES", "UserMuertes"))
188     UserList(Userindex).Stats.NPCsMuertos = CInt(UserFile.GetValue("MUERTES", "NpcsMuertes"))

190     UserList(Userindex).Stats.InventLevel = CInt(UserFile.GetValue("STATS", "InventLevel"))

192     If CByte(UserFile.GetValue("CONSEJO", "PERTENECE")) Then UserList(Userindex).flags.Privilegios = UserList(Userindex).flags.Privilegios Or PlayerType.RoyalCouncil

194     If CByte(UserFile.GetValue("CONSEJO", "PERTENECECAOS")) Then UserList(Userindex).flags.Privilegios = UserList(Userindex).flags.Privilegios Or PlayerType.ChaosCouncil

        
        Exit Sub

LoadUserStats_Err:
196     Call RegistrarError(Err.Number, Err.description, "ES.LoadUserStats", Erl)
198     Resume Next
        
End Sub

Sub LoadUserInit(ByVal Userindex As Integer, ByRef UserFile As clsIniReader)
        
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

100     UserList(Userindex).Faccion.Status = CByte(UserFile.GetValue("FACCIONES", "Status"))
102     UserList(Userindex).Faccion.ArmadaReal = CByte(UserFile.GetValue("FACCIONES", "EjercitoReal"))
104     UserList(Userindex).Faccion.FuerzasCaos = CByte(UserFile.GetValue("FACCIONES", "EjercitoCaos"))
106     UserList(Userindex).Faccion.CiudadanosMatados = CLng(UserFile.GetValue("FACCIONES", "CiudMatados"))
108     UserList(Userindex).Faccion.CriminalesMatados = CLng(UserFile.GetValue("FACCIONES", "CrimMatados"))
110     UserList(Userindex).Faccion.RecibioArmaduraCaos = CByte(UserFile.GetValue("FACCIONES", "rArCaos"))
112     UserList(Userindex).Faccion.RecibioArmaduraReal = CByte(UserFile.GetValue("FACCIONES", "rArReal"))
114     UserList(Userindex).Faccion.RecibioExpInicialCaos = CByte(UserFile.GetValue("FACCIONES", "rExCaos"))
116     UserList(Userindex).Faccion.RecibioExpInicialReal = CByte(UserFile.GetValue("FACCIONES", "rExReal"))
118     UserList(Userindex).Faccion.RecompensasCaos = CLng(UserFile.GetValue("FACCIONES", "recCaos"))
120     UserList(Userindex).Faccion.RecompensasReal = CLng(UserFile.GetValue("FACCIONES", "recReal"))
122     UserList(Userindex).Faccion.Reenlistadas = CByte(UserFile.GetValue("FACCIONES", "Reenlistadas"))
124     UserList(Userindex).Faccion.NivelIngreso = CInt(UserFile.GetValue("FACCIONES", "NivelIngreso"))
126     UserList(Userindex).Faccion.FechaIngreso = UserFile.GetValue("FACCIONES", "FechaIngreso")
128     UserList(Userindex).Faccion.MatadosIngreso = CInt(UserFile.GetValue("FACCIONES", "MatadosIngreso"))
130     UserList(Userindex).Faccion.NextRecompensa = CInt(UserFile.GetValue("FACCIONES", "NextRecompensa"))

132     UserList(Userindex).flags.Muerto = CByte(UserFile.GetValue("FLAGS", "Muerto"))
134     UserList(Userindex).flags.Escondido = CByte(UserFile.GetValue("FLAGS", "Escondido"))

136     UserList(Userindex).flags.Hambre = CByte(UserFile.GetValue("FLAGS", "Hambre"))
138     UserList(Userindex).flags.Sed = CByte(UserFile.GetValue("FLAGS", "Sed"))
140     UserList(Userindex).flags.Desnudo = CByte(UserFile.GetValue("FLAGS", "Desnudo"))
142     UserList(Userindex).flags.Navegando = CByte(UserFile.GetValue("FLAGS", "Navegando"))
144     UserList(Userindex).flags.Envenenado = CByte(UserFile.GetValue("FLAGS", "Envenenado"))
146     UserList(Userindex).flags.Paralizado = CByte(UserFile.GetValue("FLAGS", "Paralizado"))
148     UserList(Userindex).flags.Incinerado = CByte(UserFile.GetValue("FLAGS", "Incinerado"))
150     UserList(Userindex).flags.Inmovilizado = CByte(UserFile.GetValue("FLAGS", "Inmovilizado"))

152     UserList(Userindex).flags.ScrollExp = CSng(UserFile.GetValue("FLAGS", "ScrollExp"))
154     UserList(Userindex).flags.ScrollOro = CSng(UserFile.GetValue("FLAGS", "ScrollOro"))

156     If UserList(Userindex).flags.Paralizado = 1 Then
158         UserList(Userindex).Counters.Paralisis = IntervaloParalizado

        End If

160     UserList(Userindex).flags.BattlePuntos = CLng(UserFile.GetValue("Battle", "Puntos"))

162     If UserList(Userindex).flags.Inmovilizado = 1 Then
164         UserList(Userindex).Counters.Inmovilizado = 20

        End If

166     UserList(Userindex).Counters.Pena = CLng(UserFile.GetValue("COUNTERS", "Pena"))

168     UserList(Userindex).Counters.ScrollExperiencia = CLng(UserFile.GetValue("COUNTERS", "ScrollExperiencia"))
170     UserList(Userindex).Counters.ScrollOro = CLng(UserFile.GetValue("COUNTERS", "ScrollOro"))

172     UserList(Userindex).Counters.Oxigeno = CLng(UserFile.GetValue("COUNTERS", "Oxigeno"))

174     UserList(Userindex).MENSAJEINFORMACION = UserFile.GetValue("INIT", "MENSAJEINFORMACION")

176     UserList(Userindex).genero = UserFile.GetValue("INIT", "Genero")
178     UserList(Userindex).clase = UserFile.GetValue("INIT", "Clase")
180     UserList(Userindex).raza = UserFile.GetValue("INIT", "Raza")
182     UserList(Userindex).Hogar = UserFile.GetValue("INIT", "Hogar")
184     UserList(Userindex).Char.Heading = CInt(UserFile.GetValue("INIT", "Heading"))

186     UserList(Userindex).OrigChar.Head = CInt(UserFile.GetValue("INIT", "Head"))
188     UserList(Userindex).OrigChar.Body = CInt(UserFile.GetValue("INIT", "Body"))
190     UserList(Userindex).OrigChar.WeaponAnim = CInt(UserFile.GetValue("INIT", "Arma"))
192     UserList(Userindex).OrigChar.ShieldAnim = CInt(UserFile.GetValue("INIT", "Escudo"))
194     UserList(Userindex).OrigChar.CascoAnim = CInt(UserFile.GetValue("INIT", "Casco"))

        #If ConUpTime Then
196         UserList(Userindex).UpTime = CLng(UserFile.GetValue("INIT", "UpTime"))
        #End If

198     UserList(Userindex).OrigChar.Heading = UserList(Userindex).Char.Heading

200     If UserList(Userindex).flags.Muerto = 0 Then
202         UserList(Userindex).Char = UserList(Userindex).OrigChar
        Else
204         UserList(Userindex).Char.Body = iCuerpoMuerto
206         UserList(Userindex).Char.Head = iCabezaMuerto
208         UserList(Userindex).Char.WeaponAnim = NingunArma
210         UserList(Userindex).Char.ShieldAnim = NingunEscudo
212         UserList(Userindex).Char.CascoAnim = NingunCasco

        End If

214     UserList(Userindex).Desc = UserFile.GetValue("INIT", "Desc")

216     UserList(Userindex).flags.BanMotivo = UserFile.GetValue("BAN", "BanMotivo")
218     UserList(Userindex).flags.Montado = CByte(UserFile.GetValue("FLAGS", "Montado"))
220     UserList(Userindex).flags.VecesQueMoriste = CLng(UserFile.GetValue("FLAGS", "VecesQueMoriste"))

222     UserList(Userindex).flags.MinutosRestantes = CLng(UserFile.GetValue("FLAGS", "MinutosRestantes"))
224     UserList(Userindex).flags.Silenciado = CLng(UserFile.GetValue("FLAGS", "Silenciado"))
226     UserList(Userindex).flags.SegundosPasados = CLng(UserFile.GetValue("FLAGS", "SegundosPasados"))

        'CASAMIENTO LADDER
228     UserList(Userindex).flags.Casado = CInt(UserFile.GetValue("FLAGS", "CASADO"))
230     UserList(Userindex).flags.Pareja = UserFile.GetValue("FLAGS", "PAREJA")

232     UserList(Userindex).Pos.Map = CInt(ReadField(1, UserFile.GetValue("INIT", "Position"), 45))
234     UserList(Userindex).Pos.X = CInt(ReadField(2, UserFile.GetValue("INIT", "Position"), 45))
236     UserList(Userindex).Pos.Y = CInt(ReadField(3, UserFile.GetValue("INIT", "Position"), 45))

238     UserList(Userindex).Invent.NroItems = CInt(UserFile.GetValue("Inventory", "CantidadItems"))

        '[KEVIN]--------------------------------------------------------------------
        '***********************************************************************************
240     UserList(Userindex).BancoInvent.NroItems = CInt(UserFile.GetValue("BancoInventory", "CantidadItems"))

        'Lista de objetos del banco
242     For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
244         ln = UserFile.GetValue("BancoInventory", "Obj" & LoopC)
246         UserList(Userindex).BancoInvent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
248         UserList(Userindex).BancoInvent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
250     Next LoopC

        '------------------------------------------------------------------------------------
        '[/KEVIN]*****************************************************************************

        'Lista de objetos
252     For LoopC = 1 To UserList(Userindex).CurrentInventorySlots
254         ln = UserFile.GetValue("Inventory", "Obj" & LoopC)
256         UserList(Userindex).Invent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
258         UserList(Userindex).Invent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
260         UserList(Userindex).Invent.Object(LoopC).Equipped = CByte(ReadField(3, ln, 45))
262     Next LoopC

264     UserList(Userindex).Invent.WeaponEqpSlot = CByte(UserFile.GetValue("Inventory", "WeaponEqpSlot"))
266     UserList(Userindex).Invent.HerramientaEqpSlot = CByte(UserFile.GetValue("Inventory", "HerramientaEqpSlot"))
268     UserList(Userindex).Invent.ArmourEqpSlot = CByte(UserFile.GetValue("Inventory", "ArmourEqpSlot"))
270     UserList(Userindex).Invent.EscudoEqpSlot = CByte(UserFile.GetValue("Inventory", "EscudoEqpSlot"))
272     UserList(Userindex).Invent.CascoEqpSlot = CByte(UserFile.GetValue("Inventory", "CascoEqpSlot"))
274     UserList(Userindex).Invent.BarcoSlot = CByte(UserFile.GetValue("Inventory", "BarcoSlot"))
276     UserList(Userindex).Invent.MonturaSlot = CByte(UserFile.GetValue("Inventory", "MonturaSlot"))
278     UserList(Userindex).Invent.MunicionEqpSlot = CByte(UserFile.GetValue("Inventory", "MunicionSlot"))
280     UserList(Userindex).Invent.AnilloEqpSlot = CByte(UserFile.GetValue("Inventory", "AnilloSlot"))
282     UserList(Userindex).Invent.MagicoSlot = CByte(UserFile.GetValue("Inventory", "MagicoSlot"))
284     UserList(Userindex).Invent.NudilloSlot = CByte(UserFile.GetValue("Inventory", "NudilloEqpSlot"))

286     UserList(Userindex).ChatCombate = CByte(UserFile.GetValue("BINDKEYS", "ChatCombate"))
288     UserList(Userindex).ChatGlobal = CByte(UserFile.GetValue("BINDKEYS", "ChatGlobal"))

290     UserList(Userindex).Correo.CantCorreo = CByte(UserFile.GetValue("CORREO", "CantCorreo"))
292     UserList(Userindex).Correo.NoLeidos = CByte(UserFile.GetValue("CORREO", "NoLeidos"))

294     For LoopC = 1 To UserList(Userindex).Correo.CantCorreo
296         UserList(Userindex).Correo.Mensaje(LoopC).Remitente = UserFile.GetValue("CORREO", "REMITENTE" & LoopC)
298         UserList(Userindex).Correo.Mensaje(LoopC).Mensaje = UserFile.GetValue("CORREO", "MENSAJE" & LoopC)
300         UserList(Userindex).Correo.Mensaje(LoopC).Item = UserFile.GetValue("CORREO", "Item" & LoopC)
302         UserList(Userindex).Correo.Mensaje(LoopC).ItemCount = CByte(UserFile.GetValue("CORREO", "ItemCount" & LoopC))
304         UserList(Userindex).Correo.Mensaje(LoopC).Fecha = UserFile.GetValue("CORREO", "DATE" & LoopC)
306         UserList(Userindex).Correo.Mensaje(LoopC).Leido = CByte(UserFile.GetValue("CORREO", "LEIDO" & LoopC))
308     Next LoopC

        'Logros Ladder
310     UserList(Userindex).UserLogros = UserFile.GetValue("LOGROS", "UserLogros")
312     UserList(Userindex).NPcLogros = UserFile.GetValue("LOGROS", "NPcLogros")
314     UserList(Userindex).LevelLogros = UserFile.GetValue("LOGROS", "LevelLogros")
        'Logros Ladder

316     ln = UserFile.GetValue("Guild", "GUILDINDEX")

318     If IsNumeric(ln) Then
320         UserList(Userindex).GuildIndex = CInt(ln)
        Else
322         UserList(Userindex).GuildIndex = 0

        End If

        
        Exit Sub

LoadUserInit_Err:
324     Call RegistrarError(Err.Number, Err.description, "ES.LoadUserInit", Erl)
326     Resume Next
        
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
110     Call RegistrarError(Err.Number, Err.description, "ES.GetVar", Erl)
112     Resume Next
        
End Function

Sub CargarBackUp()
        
        On Error GoTo CargarBackUp_Err
        

100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."

        Dim Map       As Integer
        Dim TempInt   As Integer
        Dim npcfile   As String
    
102     NumMaps = CountFiles(MapPath, "*.csm")
104     NumMaps = NumMaps - 1
    
106     frmCargando.cargar.min = 0
108     frmCargando.cargar.max = NumMaps
110     frmCargando.cargar.Value = 0
112     frmCargando.ToMapLbl.Visible = True
    
114     ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

116     ReDim MapInfo(1 To NumMaps) As MapInfo
      
118     For Map = 1 To NumMaps
120         frmCargando.ToMapLbl = Map & "/" & NumMaps

122         Call CargarMapaFormatoCSM(Map, App.Path & "\WorldBackUp\Mapa" & Map & ".csm")

124         frmCargando.cargar.Value = frmCargando.cargar.Value + 1

126         DoEvents
128     Next Map

130     Call InitAreas

132     frmCargando.ToMapLbl.Visible = False

        Exit Sub

CargarBackUp_Err:
134     Call RegistrarError(Err.Number, Err.description, "ES.CargarBackUp", Erl)
136     Resume Next
        
End Sub

Sub LoadMapData()

100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas..."

        Dim Map       As Integer
        Dim TempInt   As Integer
        Dim npcfile   As String

        On Error GoTo man
    
102     NumMaps = CountFiles(MapPath, "*.csm")
    
104     NumMaps = NumMaps - 1
    
106     frmCargando.cargar.min = 0
108     frmCargando.cargar.max = NumMaps
110     frmCargando.cargar.Value = 0
112     frmCargando.ToMapLbl.Visible = True

114     ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

116     ReDim MapInfo(1 To NumMaps) As MapInfo

118     For Map = 1 To NumMaps
    
120         frmCargando.ToMapLbl = Map & "/" & NumMaps

122         Call CargarMapaFormatoCSM(Map, MapPath & "Mapa" & Map & ".csm")

124         frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        
126         DoEvents
        
128     Next Map
    
130     Call InitAreas

132     frmCargando.ToMapLbl.Visible = False
    
        Exit Sub

man:
134     Call MsgBox("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
136     Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)

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
        Dim Heading      As Byte

        Dim i            As Long
        Dim j            As Long
    
        Dim X As Integer, Y As Integer
    
100     If FileLen(MAPFl) = 0 Then
102         Call RegistrarError(4333, "Se trato de cargar un mapa corrupto o mal generado" & vbNewLine & "Mapa: " & MAPFl, "ES.CargarMapaFormatoCSM")
            Exit Sub
        End If
    
104     fh = FreeFile
106     Open MAPFl For Binary As fh
    
108     Get #fh, , MH
110     Get #fh, , MapSize
112     Get #fh, , MapDat

        Rem Get #fh, , L1

114     With MH

            'Cargamos Bloqueos
        
116         If .NumeroBloqueados > 0 Then

118             ReDim Blqs(1 To .NumeroBloqueados)
120             Get #fh, , Blqs

122             For i = 1 To .NumeroBloqueados
124                 MapData(Map, Blqs(i).X, Blqs(i).Y).Blocked = Blqs(i).Lados
126             Next i

            End If
        
            'Cargamos Layer 1
        
128         If .NumeroLayers(1) > 0 Then
        
130             ReDim L1(1 To .NumeroLayers(1))
132             Get #fh, , L1

134             For i = 1 To .NumeroLayers(1)
                
136                 X = L1(i).X
138                 Y = L1(i).Y
                        
140                 MapData(Map, X, Y).Graphic(1) = L1(i).GrhIndex
            
                    'InitGrh MapData(L1(i).X, L1(i).Y).Graphic(1), MapData(L1(i).X, L1(i).Y).Graphic(1).GrhIndex
                    ' Call Map_Grh_Set(L2(i).X, L2(i).Y, L2(i).GrhIndex, 2)
142                 If HayAgua(Map, X, Y) Then
144                     MapData(Map, X, Y).Blocked = MapData(Map, X, Y).Blocked Or FLAG_AGUA
                    End If
                
146             Next i

            End If
        
            'Cargamos Layer 2
148         If .NumeroLayers(2) > 0 Then
150             ReDim L2(1 To .NumeroLayers(2))
152             Get #fh, , L2

154             For i = 1 To .NumeroLayers(2)
                
156                 X = L2(i).X
158                 Y = L2(i).Y

160                 MapData(Map, X, Y).Graphic(2) = L2(i).GrhIndex
                
162                 MapData(Map, X, Y).Blocked = MapData(Map, X, Y).Blocked And Not FLAG_AGUA
                
164             Next i

            End If
                
166         If .NumeroLayers(3) > 0 Then
168             ReDim L3(1 To .NumeroLayers(3))
170             Get #fh, , L3

172             For i = 1 To .NumeroLayers(3)
174                 MapData(Map, L3(i).X, L3(i).Y).Graphic(3) = L3(i).GrhIndex
                
176                 If EsArbol(L3(i).GrhIndex) Then
178                     MapData(Map, X, Y).Blocked = MapData(Map, X, Y).Blocked Or FLAG_ARBOL
                    End If
180             Next i

            End If
        
182         If .NumeroLayers(4) > 0 Then
184             ReDim L4(1 To .NumeroLayers(4))
186             Get #fh, , L4

188             For i = 1 To .NumeroLayers(4)
190                 MapData(Map, L4(i).X, L4(i).Y).Graphic(4) = L4(i).GrhIndex
192             Next i

            End If

194         If .NumeroTriggers > 0 Then
196             ReDim Triggers(1 To .NumeroTriggers)
198             Get #fh, , Triggers

200             For i = 1 To .NumeroTriggers
202                 MapData(Map, Triggers(i).X, Triggers(i).Y).trigger = Triggers(i).trigger
204             Next i

            End If

206         If .NumeroParticulas > 0 Then
208             ReDim Particulas(1 To .NumeroParticulas)
210             Get #fh, , Particulas

212             For i = 1 To .NumeroParticulas
214                 MapData(Map, Particulas(i).X, Particulas(i).Y).ParticulaIndex = Particulas(i).Particula
216                 MapData(Map, Particulas(i).X, Particulas(i).Y).ParticulaIndex = 0
218             Next i

            End If

220         If .NumeroLuces > 0 Then
222             ReDim Luces(1 To .NumeroLuces)
224             Get #fh, , Luces

226             For i = 1 To .NumeroLuces
228                 MapData(Map, Luces(i).X, Luces(i).Y).Luz.Color = Luces(i).Color
230                 MapData(Map, Luces(i).X, Luces(i).Y).Luz.Rango = Luces(i).Rango
232                 MapData(Map, Luces(i).X, Luces(i).Y).Luz.Color = 0
234                 MapData(Map, Luces(i).X, Luces(i).Y).Luz.Rango = 0
236             Next i

            End If
            
238         If .NumeroOBJs > 0 Then
240             ReDim Objetos(1 To .NumeroOBJs)
242             Get #fh, , Objetos

244             For i = 1 To .NumeroOBJs
246                 MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.ObjIndex = Objetos(i).ObjIndex

248                 Select Case ObjData(Objetos(i).ObjIndex).OBJType

                        Case eOBJType.otYacimiento, eOBJType.otArboles
250                         MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.Amount = ObjData(Objetos(i).ObjIndex).VidaUtil
252                         MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.data = &H7FFFFFFF ' Ultimo uso = Max Long

254                     Case Else
256                         MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.Amount = Objetos(i).ObjAmmount

                    End Select

258             Next i

            End If

260         If .NumeroNPCs > 0 Then
262             ReDim NPCs(1 To .NumeroNPCs)
264             Get #fh, , NPCs
                 
266             For i = 1 To .NumeroNPCs

268                 MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex = NPCs(i).NpcIndex
                    
270                 If MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex > 0 Then
272                     npcfile = DatPath & "NPCs.dat"

                        'Si el npc debe hacer respawn en la pos
                        'original la guardamos
274                     If val(GetVar(npcfile, "NPC" & MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex, "PosOrig")) = 1 Then
276                         MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex = OpenNPC(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex)
278                         Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Orig.Map = Map
280                         Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Orig.X = NPCs(i).X
282                         Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Orig.Y = NPCs(i).Y
                        Else
284                         MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex = OpenNPC(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex)

                        End If

286                     Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Pos.Map = Map
288                     Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Pos.X = NPCs(i).X
290                     Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Pos.Y = NPCs(i).Y
                        
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
                            
292                     If Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).name = "" Then
                       
294                         MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex = 0
                        Else
                        
296                         Call MakeNPCChar(True, 0, MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex, Map, NPCs(i).X, NPCs(i).Y)
                        
                        End If

                    End If

298             Next i
                
            End If
            
300         If .NumeroTE > 0 Then
302             ReDim TEs(1 To .NumeroTE)
304             Get #fh, , TEs

306             For i = 1 To .NumeroTE
308                 MapData(Map, TEs(i).X, TEs(i).Y).TileExit.Map = TEs(i).DestM
310                 MapData(Map, TEs(i).X, TEs(i).Y).TileExit.X = TEs(i).DestX
312                 MapData(Map, TEs(i).X, TEs(i).Y).TileExit.Y = TEs(i).DestY
314             Next i

            End If
        
        End With

316     Close fh
    
318     MapInfo(Map).map_name = MapDat.map_name
320     MapInfo(Map).ambient = MapDat.ambient
322     MapInfo(Map).backup_mode = MapDat.backup_mode
324     MapInfo(Map).base_light = MapDat.base_light
326     MapInfo(Map).extra1 = MapDat.extra1
        'MapInfo(Map).extra2 = MapDat.extra2
328     MapInfo(Map).extra2 = val(GetVar(DatPath & "Map.dat", "RESUCIUDAD", Map))
    
330     MapInfo(Map).extra3 = MapDat.extra3
332     MapInfo(Map).letter_grh = MapDat.letter_grh
334     MapInfo(Map).lluvia = MapDat.lluvia
336     MapInfo(Map).music_numberHi = MapDat.music_numberHi
338     MapInfo(Map).music_numberLow = MapDat.music_numberLow
340     MapInfo(Map).niebla = MapDat.niebla
342     MapInfo(Map).Nieve = MapDat.Nieve
344     MapInfo(Map).restrict_mode = MapDat.restrict_mode
    
346     MapInfo(Map).Seguro = MapDat.Seguro

348     MapInfo(Map).terrain = MapDat.terrain
350     MapInfo(Map).zone = MapDat.zone
 
        Exit Sub

errh:
352     Close fh
354     Call MsgBox("Error cargando mapa: " & Map & ". " & Err.Number & " - " & Err.description & " - ")
    
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
158     ResPos.X = val(ReadField(2, Lector.GetValue("INIT", "ResPos"), 45))
160     ResPos.Y = val(ReadField(3, Lector.GetValue("INIT", "ResPos"), 45))
      
162     If Not Database_Enabled Then
164         RecordUsuarios = val(Lector.GetValue("INIT", "Record"))
        End If
      
        'Max users
166     Temporal = val(Lector.GetValue("INIT", "MaxUsers"))

168     If MaxUsers = 0 Then
170         MaxUsers = Temporal
172         ReDim UserList(1 To MaxUsers) As user

        End If

174     NumCuentas = val(Lector.GetValue("INIT", "NumCuentas"))
176     frmMain.cuentas.Caption = NumCuentas
        #If DEBUGGING Then
            'Shell App.Path & "\estadisticas.exe" & " " & "NUEVACUENTALADDER" & "*" & NumCuentas & "*" & MaxUsers
        #End If
    
        '&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
        'Se agregó en LoadBalance y en el Balance.dat
        'PorcentajeRecuperoMana = val(Lector.GetValue("BALANCE", "PorcentajeRecuperoMana"))
    
        ''&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&
        'Call Statistics.Initialize
    
178     Call CargarCiudades
            
180     Call MD5sCarga
    
182     Call ConsultaPopular.LoadData
    
184     Set Lector = Nothing

        
        Exit Sub

LoadSini_Err:
186     Call RegistrarError(Err.Number, Err.description, "ES.LoadSini", Erl)
188     Resume Next
        
End Sub

Sub CargarCiudades()

        Dim Lector As clsIniReader
100     Set Lector = New clsIniReader
102     Call Lector.Initialize(DatPath & "Ciudades.dat")
    
104     With CityNix
106         .Map = val(Lector.GetValue("NIX", "Mapa"))
108         .X = val(Lector.GetValue("NIX", "X"))
110         .Y = val(Lector.GetValue("NIX", "Y"))
112         .MapaViaje = val(Lector.GetValue("NIX", "MapaViaje"))
114         .ViajeX = val(Lector.GetValue("NIX", "ViajeX"))
116         .ViajeY = val(Lector.GetValue("NIX", "ViajeY"))
118         .MapaResu = val(Lector.GetValue("NIX", "MapaResu"))
120         .ResuX = val(Lector.GetValue("NIX", "ResuX"))
122         .ResuY = val(Lector.GetValue("NIX", "ResuY"))
124         .NecesitaNave = val(Lector.GetValue("NIX", "NecesitaNave"))
        End With
    
126     With CityUllathorpe
128         .Map = val(Lector.GetValue("Ullathorpe", "Mapa"))
130         .X = val(Lector.GetValue("Ullathorpe", "X"))
132         .Y = val(Lector.GetValue("Ullathorpe", "Y"))
134         .MapaViaje = val(Lector.GetValue("Ullathorpe", "MapaViaje"))
136         .ViajeX = val(Lector.GetValue("Ullathorpe", "ViajeX"))
138         .ViajeY = val(Lector.GetValue("Ullathorpe", "ViajeY"))
140         .MapaResu = val(Lector.GetValue("Ullathorpe", "MapaResu"))
142         .ResuX = val(Lector.GetValue("Ullathorpe", "ResuX"))
144         .ResuY = val(Lector.GetValue("Ullathorpe", "ResuY"))
146         .NecesitaNave = val(Lector.GetValue("Ullathorpe", "NecesitaNave"))
        End With
    
148     With CityBanderbill
150         .Map = val(Lector.GetValue("Banderbill", "Mapa"))
152         .X = val(Lector.GetValue("Banderbill", "X"))
154         .Y = val(Lector.GetValue("Banderbill", "Y"))
156         .MapaViaje = val(Lector.GetValue("Banderbill", "MapaViaje"))
158         .ViajeX = val(Lector.GetValue("Banderbill", "ViajeX"))
160         .ViajeY = val(Lector.GetValue("Banderbill", "ViajeY"))
162         .MapaResu = val(Lector.GetValue("Banderbill", "MapaResu"))
164         .ResuX = val(Lector.GetValue("Banderbill", "ResuX"))
166         .ResuY = val(Lector.GetValue("Banderbill", "ResuY"))
168         .NecesitaNave = val(Lector.GetValue("Banderbill", "NecesitaNave"))
        End With
    
170     With CityLindos
172         .Map = val(Lector.GetValue("Lindos", "Mapa"))
174         .X = val(Lector.GetValue("Lindos", "X"))
176         .Y = val(Lector.GetValue("Lindos", "Y"))
178         .MapaViaje = val(Lector.GetValue("Lindos", "MapaViaje"))
180         .ViajeX = val(Lector.GetValue("Lindos", "ViajeX"))
182         .ViajeY = val(Lector.GetValue("Lindos", "ViajeY"))
184         .MapaResu = val(Lector.GetValue("Lindos", "MapaResu"))
186         .ResuX = val(Lector.GetValue("Lindos", "ResuX"))
188         .ResuY = val(Lector.GetValue("Lindos", "ResuY"))
190         .NecesitaNave = val(Lector.GetValue("Lindos", "NecesitaNave"))
        End With
    
192     With CityArghal
194         .Map = val(Lector.GetValue("Arghal", "Mapa"))
196         .X = val(Lector.GetValue("Arghal", "X"))
198         .Y = val(Lector.GetValue("Arghal", "Y"))
200         .MapaViaje = val(Lector.GetValue("Arghal", "MapaViaje"))
202         .ViajeX = val(Lector.GetValue("Arghal", "ViajeX"))
204         .ViajeY = val(Lector.GetValue("Arghal", "ViajeY"))
206         .MapaResu = val(Lector.GetValue("Arghal", "MapaResu"))
208         .ResuX = val(Lector.GetValue("Arghal", "ResuX"))
210         .ResuY = val(Lector.GetValue("Arghal", "ResuY"))
212         .NecesitaNave = val(Lector.GetValue("Arghal", "NecesitaNave"))
        End With
    
214     With CityHillidan
216         .Map = val(Lector.GetValue("Hillidan", "Mapa"))
218         .X = val(Lector.GetValue("Hillidan", "X"))
220         .Y = val(Lector.GetValue("Hillidan", "Y"))
222         .MapaViaje = val(Lector.GetValue("Hillidan", "MapaViaje"))
224         .ViajeX = val(Lector.GetValue("Hillidan", "ViajeX"))
226         .ViajeY = val(Lector.GetValue("Hillidan", "ViajeY"))
228         .MapaResu = val(Lector.GetValue("Hillidan", "MapaResu"))
230         .ResuX = val(Lector.GetValue("Hillidan", "ResuX"))
232         .ResuY = val(Lector.GetValue("Hillidan", "ResuY"))
234         .NecesitaNave = val(Lector.GetValue("Hillidan", "NecesitaNave"))
        End With
    
236     With Prision
238         .Map = val(Lector.GetValue("Prision", "Mapa"))
240         .X = val(Lector.GetValue("Prision", "X"))
242         .Y = val(Lector.GetValue("Prision", "Y"))
        End With
    
244     With Libertad
246         .Map = val(Lector.GetValue("Libertad", "Mapa"))
248         .X = val(Lector.GetValue("Libertad", "X"))
250         .Y = val(Lector.GetValue("Libertad", "Y"))
        End With
    
252     Set Lector = Nothing
    
254     Nix.Map = CityNix.Map
256     Nix.X = CityNix.X
258     Nix.Y = CityNix.Y
    
260     Ullathorpe.Map = CityUllathorpe.Map
262     Ullathorpe.X = CityUllathorpe.X
264     Ullathorpe.Y = CityUllathorpe.Y
    
266     Banderbill.Map = CityBanderbill.Map
268     Banderbill.X = CityBanderbill.X
270     Banderbill.Y = CityBanderbill.Y
    
272     Lindos.Map = CityLindos.Map
274     Lindos.X = CityLindos.X
276     Lindos.Y = CityLindos.Y
    
278     Arghal.Map = CityArghal.Map
280     Arghal.X = CityArghal.X
282     Arghal.Y = CityArghal.Y
    
284     Hillidan.Map = CityHillidan.Map
286     Hillidan.X = CityHillidan.X
288     Hillidan.Y = CityHillidan.Y
    
        'Esto es para el /HOGAR
290     Ciudades(eCiudad.cNix) = Nix
292     Ciudades(eCiudad.cUllathorpe) = Ullathorpe
294     Ciudades(eCiudad.cBanderbill) = Banderbill
296     Ciudades(eCiudad.cLindos) = Lindos
298     Ciudades(eCiudad.cArghal) = Arghal
300     Ciudades(eCiudad.CHillidan) = Hillidan
    
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
208     IntervaloUserPuedeUsarClic = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeUsarClic"))
210     IntervaloFlechasCazadores = val(Lector.GetValue("INTERVALOS", "IntervaloFlechasCazadores"))
212     IntervaloGolpeUsar = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeUsar"))
    
214     IntervaloOculto = val(Lector.GetValue("INTERVALOS", "IntervaloOculto"))

216     MargenDeIntervaloPorPing = val(Lector.GetValue("INTERVALOS", "MargenDeIntervaloPorPing"))
    
218     IntervaloPuedeSerAtacado = val(Lector.GetValue("INTERVALOS", "IntervaloPuedeSerAtacado"))

220     IntervaloGuardarUsuarios = val(Lector.GetValue("INTERVALOS", "IntervaloGuardarUsuarios"))
        
222     LimiteSaveUserPorMinuto = val(Lector.GetValue("INTERVALOS", "LimiteSaveUserPorMinuto"))
    
        '&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    
224     Set Lector = Nothing

        
        Exit Sub

LoadIntervalos_Err:
226     Call RegistrarError(Err.Number, Err.description, "ES.LoadIntervalos", Erl)
228     Resume Next
        
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

116     DuracionDia = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "DuracionDia")) * 60 * 1000 ' De minutos a milisegundos

118     BattleActivado = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "BattleActivado"))
120     BattleMinNivel = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "BattleMinNivel"))

122     frmMain.lblLimpieza.Caption = "Limpieza de objetos cada: " & TimerLimpiarObjetos & " minutos."

        
        Exit Sub

LoadConfiguraciones_Err:
124     Call RegistrarError(Err.Number, Err.description, "ES.LoadConfiguraciones", Erl)
126     Resume Next
        
End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
        '*****************************************************************
        'Escribe VAR en un archivo
        '*****************************************************************
        
        On Error GoTo WriteVar_Err
        

100     writeprivateprofilestring Main, Var, Value, File
    
        
        Exit Sub

WriteVar_Err:
102     Call RegistrarError(Err.Number, Err.description, "ES.WriteVar", Erl)
104     Resume Next
        
End Sub

Sub LoadUser(ByVal Userindex As Integer)

        On Error GoTo ErrorHandler
    
100     If Database_Enabled Then
102         Call LoadUserDatabase(Userindex)
        Else
104         Call LoadUserBinary(Userindex)
        End If
    
106     With UserList(Userindex)

108         If .flags.Paralizado = 1 Then
110             .Counters.Paralisis = IntervaloParalizado
            End If

112         If .flags.Muerto = 0 Then
114             .Char = .OrigChar
            
116             If .Char.Body = 0 Then
118                 Call DarCuerpoDesnudo(Userindex)
                End If
            
120             If .Char.Head = 0 Then
122                 .Char.Head = 1
                End If
            Else
124             .Char.Body = iCuerpoMuerto
126             .Char.Head = iCabezaMuerto
128             .Char.WeaponAnim = NingunArma
130             .Char.ShieldAnim = NingunEscudo
132             .Char.CascoAnim = NingunCasco
134             .Char.Heading = eHeading.SOUTH
            End If
        
            'Obtiene el indice-objeto del arma
136         If .Invent.WeaponEqpSlot > 0 Then
138             .Invent.WeaponEqpObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).ObjIndex
            
140             If .flags.Muerto = 0 Then
142                 .Char.Arma_Aura = ObjData(.Invent.WeaponEqpObjIndex).CreaGRH
                End If
            End If

            'Obtiene el indice-objeto del armadura
144         If .Invent.ArmourEqpSlot > 0 Then
146             .Invent.ArmourEqpObjIndex = .Invent.Object(.Invent.ArmourEqpSlot).ObjIndex
            
148             If .flags.Muerto = 0 Then
150                 .Char.Body_Aura = ObjData(.Invent.ArmourEqpObjIndex).CreaGRH
                End If

152             .flags.Desnudo = 0
            Else
154             .flags.Desnudo = 1
            End If

            'Obtiene el indice-objeto del escudo
156         If .Invent.EscudoEqpSlot > 0 Then
158             .Invent.EscudoEqpObjIndex = .Invent.Object(.Invent.EscudoEqpSlot).ObjIndex
            
160             If .flags.Muerto = 0 Then
162                 .Char.Escudo_Aura = ObjData(.Invent.EscudoEqpObjIndex).CreaGRH
                End If
            End If
        
            'Obtiene el indice-objeto del casco
164         If .Invent.CascoEqpSlot > 0 Then
166             .Invent.CascoEqpObjIndex = .Invent.Object(.Invent.CascoEqpSlot).ObjIndex
            
168             If .flags.Muerto = 0 Then
170                 .Char.Head_Aura = ObjData(.Invent.CascoEqpObjIndex).CreaGRH
                End If
            End If

            'Obtiene el indice-objeto barco
172         If .Invent.BarcoSlot > 0 Then
174             .Invent.BarcoObjIndex = .Invent.Object(.Invent.BarcoSlot).ObjIndex
            End If

            'Obtiene el indice-objeto municion
176         If .Invent.MunicionEqpSlot > 0 Then
178             .Invent.MunicionEqpObjIndex = .Invent.Object(.Invent.MunicionEqpSlot).ObjIndex
            End If

            'Obtiene el indice-objeto anilo
180         If .Invent.AnilloEqpSlot > 0 Then
182             .Invent.AnilloEqpObjIndex = .Invent.Object(.Invent.AnilloEqpSlot).ObjIndex
            
184             If .flags.Muerto = 0 Then
186                 .Char.Anillo_Aura = ObjData(.Invent.AnilloEqpObjIndex).CreaGRH
                End If
            End If

188         If .Invent.MonturaSlot > 0 Then
190             .Invent.MonturaObjIndex = .Invent.Object(.Invent.MonturaSlot).ObjIndex
            End If
        
192         If .Invent.HerramientaEqpSlot > 0 Then
194             .Invent.HerramientaEqpObjIndex = .Invent.Object(.Invent.HerramientaEqpSlot).ObjIndex
            End If
        
196         If .Invent.NudilloSlot > 0 Then
198             .Invent.NudilloObjIndex = .Invent.Object(.Invent.NudilloSlot).ObjIndex
            
200             If .flags.Muerto = 0 Then
202                 .Char.Arma_Aura = ObjData(.Invent.NudilloObjIndex).CreaGRH
                End If
            End If
        
204         If .Invent.MagicoSlot > 0 Then
206             .Invent.MagicoObjIndex = .Invent.Object(.Invent.MagicoSlot).ObjIndex

208             If .flags.Muerto = 0 Then
210                 .Char.Otra_Aura = ObjData(.Invent.MagicoObjIndex).CreaGRH
                End If
            End If

        End With

        Exit Sub

ErrorHandler:
212     Call LogError("Error en LoadUser: " & UserList(Userindex).name & " - " & Err.Number & " - " & Err.description)
    
End Sub

Sub SaveUser(ByVal Userindex As Integer, Optional ByVal Logout As Boolean = False)
        
        On Error GoTo SaveUser_Err
        

100     If Database_Enabled Then
102         Call SaveUserDatabase(Userindex, Logout)
        Else
104         Call SaveUserCharfile(Userindex, Logout)
        End If

106     UserList(Userindex).Counters.LastSave = GetTickCount

        
        Exit Sub

SaveUser_Err:
108     Call RegistrarError(Err.Number, Err.description, "ES.SaveUser", Erl)
110     Resume Next
        
End Sub

Sub LoadUserBinary(ByVal Userindex As Integer)
        
        On Error GoTo LoadUserBinary_Err
        

        'Cargamos el personaje
        Dim Leer As New clsIniReader
100     Call Leer.Initialize(CharPath & UCase$(UserList(Userindex).name) & ".chr")
    
        'Cargamos los datos del personaje

102     Call LoadUserInit(Userindex, Leer)
    
104     Call LoadUserStats(Userindex, Leer)
    
106     Call LoadQuestStats(Userindex, Leer)
    
108     Set Leer = Nothing

        
        Exit Sub

LoadUserBinary_Err:
110     Call RegistrarError(Err.Number, Err.description, "ES.LoadUserBinary", Erl)
112     Resume Next
        
End Sub

Sub SaveUserCharfile(ByVal Userindex As Integer, Optional ByVal Logout As Boolean)
    '*************************************************
    'Author: Unknown
    'Last modified: 23/01/2007
    'Saves the Users records
    '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
    '*************************************************
    
    On Error GoTo ErrHandler
    
    Dim UserFile    As String
    Dim OldUserHead As Long
    
    With UserList(Userindex)
    
        UserFile = CharPath & UCase$(.name) & ".chr"
    
        'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
        'clase=0 es el error, porq el enum empieza de 1!!
        If .clase = 0 Or .Stats.ELV = 0 Then
            Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & .name)
            Exit Sub
        End If
    
        Debug.Print UserFile
    
        If FileExist(UserFile, vbNormal) Then
        
            If .flags.Muerto = 1 Then
                OldUserHead = .Char.Head
                .Char.Head = GetVar(UserFile, "INIT", "Head")

            End If

        End If
    
        Dim LoopC As Integer
    
        If FileExist(UserFile, vbNormal) Then Kill UserFile

        Dim File As String: File = UserFile
        Dim n As Integer: n = FreeFile
        
        Open File For Output Access Write As n
        
        'INIT
        Print n, , "[INIT]" & vbCrLf
        Print n, , "Cuenta=" & .Cuenta & vbCrLf
        Print n, , "Genero=" & .genero & vbCrLf
        Print n, , "Raza=" & .raza & vbCrLf
        Print n, , "Hogar=" & .Hogar & vbCrLf
        Print n, , "Clase=" & .clase & vbCrLf
        Print n, , "Desc=" & .Desc & vbCrLf
        Print n, , "Heading=" & CStr(.Char.Heading) & vbCrLf

        If .Char.Head = 0 Then
            Print n, , "Head=" & CStr(.OrigChar.Head) & vbCrLf
        Else
            Print n, , "Head=" & CStr(.Char.Head) & vbCrLf
        End If

        Print n, , "Arma=" & CStr(.Char.WeaponAnim) & vbCrLf
        Print n, , "Escudo=" & CStr(.Char.ShieldAnim) & vbCrLf
        Print n, , "Casco=" & CStr(.Char.CascoAnim) & vbCrLf
        Print n, , "Position=" & .Pos.Map & "-" & .Pos.X & "-" & .Pos.Y & vbCrLf
       
       'If .flags.Muerto = 0 Then
        Print n, , "Body=" & CStr(.Char.Body) & vbCrLf
        'End If
        
        #If ConUpTime Then
            Dim TempDate As Date: TempDate = Now - .LogOnTime
            
            .LogOnTime = Now
            .UpTime = .UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
            
            Print n, , "UpTime=" & .UpTime & vbCrLf
        #End If

        If Logout Then
            Print n, , "Logged=0" & vbCrLf
        Else
            Print n, , "Logged=1" & vbCrLf
        End If

        Print n, , "MENSAJEINFORMACION=" & .MENSAJEINFORMACION & vbCrLf

        Print n, , vbCrLf
        
        'baneo
        Print n, , "[BAN]" & vbCrLf
        Print n, , "Baneado=" & CStr(.flags.Ban) & vbCrLf
        Print n, , "BanMotivo=" & CStr(.flags.BanMotivo) & vbCrLf
        
        Print n, , vbCrLf
        
        'STATS
        With .Stats
            Print n, , "[STATS]" & vbCrLf
            Print n, , "GLD=" & CStr(.GLD) & vbCrLf
            Print n, , "BANCO=" & CStr(.Banco) & vbCrLf
            Print n, , "MaxHP=" & CStr(.MaxHp) & vbCrLf
            Print n, , "MinHP=" & CStr(.MinHp) & vbCrLf
            Print n, , "MaxSTA=" & CStr(.MaxSta) & vbCrLf
            Print n, , "MinSTA=" & CStr(.MinSta) & vbCrLf
            Print n, , "MaxMAN=" & CStr(.MaxMAN) & vbCrLf
            Print n, , "MinMAN=" & CStr(.MinMAN) & vbCrLf
            Print n, , "MaxHIT=" & CStr(.MaxHit) & vbCrLf
            Print n, , "MinHIT=" & CStr(.MinHIT) & vbCrLf
            Print n, , "MaxAGU=" & CStr(.MaxAGU) & vbCrLf
            Print n, , "MinAGU=" & CStr(.MinAGU) & vbCrLf
            Print n, , "MaxHAM=" & CStr(.MaxHam) & vbCrLf
            Print n, , "MinHAM=" & CStr(.MinHam) & vbCrLf
            Print n, , "SkillPtsLibres=" & CStr(.SkillPts) & vbCrLf
            Print n, , "EXP=" & CStr(.Exp) & vbCrLf
            Print n, , "ELV=" & CStr(.ELV) & vbCrLf
            Print n, , "ELU=" & CStr(.ELU) & vbCrLf
            Print n, , "InventLevel=" & CByte(.InventLevel) & vbCrLf
        End With
        
        Print n, , vbCrLf
        
        'FLAGS
        With .flags
            Print n, , "[FLAGS]" & vbCrLf
            Print n, , "CASADO=" & CStr(.Casado) & vbCrLf
            Print n, , "PAREJA=" & CStr(.Pareja) & vbCrLf
            Print n, , "Muerto=" & CStr(.Muerto) & vbCrLf
            Print n, , "Escondido=" & CStr(.Escondido) & vbCrLf
            Print n, , "Hambre=" & CStr(.Hambre) & vbCrLf
            Print n, , "Sed=" & CStr(.Sed) & vbCrLf
            Print n, , "Desnudo=" & CStr(.Desnudo) & vbCrLf
            Print n, , "Navegando=" & CStr(.Navegando) & vbCrLf
            Print n, , "Envenenado=" & CStr(.Envenenado) & vbCrLf
            Print n, , "Paralizado=" & CStr(.Paralizado) & vbCrLf
            Print n, , "Inmovilizado=" & CStr(.Inmovilizado) & vbCrLf
            Print n, , "Incinerado=" & CStr(.Incinerado) & vbCrLf
            Print n, , "VecesQueMoriste=" & CStr(.VecesQueMoriste) & vbCrLf
            Print n, , "ScrollExp=" & CStr(.ScrollExp) & vbCrLf
            Print n, , "ScrollOro=" & CStr(.ScrollOro) & vbCrLf
            Print n, , "MinutosRestantes=" & CStr(.MinutosRestantes) & vbCrLf
            Print n, , "SegundosPasados=" & CStr(.SegundosPasados) & vbCrLf
            Print n, , "Silenciado=" & CStr(.Silenciado) & vbCrLf
            Print n, , "Montado=" & CStr(.Montado) & vbCrLf
        End With

        Print n, , vbCrLf
        
        'GRABADO DE CLAN
        Print n, , "[GUILD]" & vbCrLf
        Print n, , "GUILDINDEX=" & CInt(.GuildIndex) & vbCrLf
        
        Print n, , vbCrLf
        
        Print n, , "[CONSEJO]" & vbCrLf
        Print n, , "PERTENECE=" & IIf(.flags.Privilegios And PlayerType.RoyalCouncil, "1", "0") & vbCrLf
        Print n, , "PERTENECECAOS=" & IIf(.flags.Privilegios And PlayerType.ChaosCouncil, "1", "0") & vbCrLf
        
        Print n, , vbCrLf
        
        With .Faccion
            Print n, , "[FACCIONES]" & vbCrLf
            Print n, , "EjercitoReal=" & CStr(.ArmadaReal) & vbCrLf
            Print n, , "Status=" & CStr(.Status) & vbCrLf
            Print n, , "EjercitoCaos=" & CStr(.FuerzasCaos) & vbCrLf
            Print n, , "CiudMatados=" & CStr(.CiudadanosMatados) & vbCrLf
            Print n, , "CrimMatados=" & CStr(.CriminalesMatados) & vbCrLf
            Print n, , "rArCaos=" & CStr(.RecibioArmaduraCaos) & vbCrLf
            Print n, , "rArReal=" & CStr(.RecibioArmaduraReal) & vbCrLf
            Print n, , "rExCaos=" & CStr(.RecibioExpInicialCaos) & vbCrLf
            Print n, , "rExReal=" & CStr(.RecibioExpInicialReal) & vbCrLf
            Print n, , "recCaos=" & CStr(.RecompensasCaos) & vbCrLf
            Print n, , "recReal=" & CStr(.RecompensasReal) & vbCrLf
            Print n, , "Reenlistadas=" & CStr(.Reenlistadas) & vbCrLf
            Print n, , "NivelIngreso=" & CStr(.NivelIngreso) & vbCrLf
            Print n, , "FechaIngreso=" & CStr(.FechaIngreso) & vbCrLf
            Print n, , "MatadosIngreso=" & CStr(.MatadosIngreso) & vbCrLf
            Print n, , "NextRecompensa=" & CStr(.NextRecompensa) & vbCrLf
        End With
        
        Print n, , vbCrLf
        
        'MAHIA ESTUPIDIZA
        Print n, , "[MAGIA]" & vbCrLf
        Print n, , "ENVENENA=" & CByte(.flags.Envenena) & vbCrLf
        Print n, , "PARALIZA=" & CByte(.flags.Paraliza) & vbCrLf
        Print n, , "AnilloOcultismo=" & CByte(.flags.AnilloOcultismo) & vbCrLf
        Print n, , "incinera=" & CByte(.flags.incinera) & vbCrLf
        Print n, , "Estupidiza=" & CByte(.flags.Estupidiza) & vbCrLf
        Print n, , "Pendiente=" & CByte(.flags.PendienteDelSacrificio) & vbCrLf
        Print n, , "CarroMineria=" & CByte(.flags.CarroMineria) & vbCrLf
        Print n, , "NoPalabrasMagicas=" & CByte(.flags.NoPalabrasMagicas) & vbCrLf
        Print n, , "NoDetectable=" & CByte(.flags.NoDetectable) & vbCrLf
        Print n, , "Otra_Aura=" & CStr(.Char.Otra_Aura) & vbCrLf
        'Print n, , "DañoMagico=" & CByte(.flags.DañoMagico) & vbCrLf
        'Print n, , "ResistenciaMagica=" & CByte(.flags.ResistenciaMagica) & vbCrLf
        Print n, , "RegeneracionMana=" & CByte(.flags.RegeneracionMana) & vbCrLf
        Print n, , "NoMagiaEfeceto=" & CByte(.flags.NoMagiaEfeceto) & vbCrLf
        Print n, , "RegeneracionHP=" & CByte(.flags.RegeneracionHP) & vbCrLf
        Print n, , "RegeneracionSta=" & CByte(.flags.RegeneracionSta) & vbCrLf

        Print n, , vbCrLf
        
        'SKILLS
        Print n, , "[SKILLS]" & vbCrLf

        For LoopC = 1 To UBound(.Stats.UserSkills)
            Print n, , "SK" & LoopC & "=" & CStr(.Stats.UserSkills(LoopC)) & vbCrLf
        Next

        Print n, , vbCrLf

        'INVENTARIO
        With .Invent
            Print n, , "[Inventory]" & vbCrLf
            Print n, , "CantidadItems=" & val(.NroItems) & vbCrLf
    
            For LoopC = 1 To .CurrentInventorySlots
                Print n, , "Obj" & LoopC & "=" & .Object(LoopC).ObjIndex & "-" & .Object(LoopC).Amount & "-" & .Object(LoopC).Equipped & vbCrLf
            Next
            
            Print n, , "WeaponEqpSlot=" & CStr(.WeaponEqpSlot) & vbCrLf
            Print n, , "HerramientaEqpSlot=" & CStr(.HerramientaEqpSlot) & vbCrLf
            Print n, , "ArmourEqpSlot=" & CStr(.ArmourEqpSlot) & vbCrLf
            Print n, , "CascoEqpSlot=" & CStr(.CascoEqpSlot) & vbCrLf
            Print n, , "EscudoEqpSlot=" & CStr(.EscudoEqpSlot) & vbCrLf
            Print n, , "BarcoSlot=" & CStr(.BarcoSlot) & vbCrLf
            Print n, , "MonturaSlot=" & CStr(.MonturaSlot) & vbCrLf
            Print n, , "MunicionSlot=" & CStr(.MunicionEqpSlot) & vbCrLf
            Print n, , "AnilloSlot=" & CStr(.AnilloEqpSlot) & vbCrLf
            Print n, , "MagicoSlot=" & CStr(.MagicoSlot) & vbCrLf
            Print n, , "NudilloEqpSlot=" & CStr(.NudilloSlot) & vbCrLf
        End With
        
        Print n, , vbCrLf

        Print n, , "[ATRIBUTOS]" & vbCrLf

        '¿Fueron modificados los atributos del usuario?
        If Not .flags.TomoPocion Then

            For LoopC = 1 To UBound(.Stats.UserAtributos)
                Print n, , "AT" & LoopC & "=" & CStr(.Stats.UserAtributos(LoopC)) & vbCrLf
            Next

        Else

            For LoopC = 1 To UBound(.Stats.UserAtributos)
                '.Stats.UserAtributos(LoopC) = .Stats.UserAtributosBackUP(LoopC)
                Print n, , "AT" & LoopC & "=" & CStr(.Stats.UserAtributosBackUP(LoopC)) & vbCrLf
            Next

        End If

        Print n, , vbCrLf
            
        'COUNTERS
        Print n, , "[COUNTERS]" & vbCrLf
        Print n, , "Pena=" & CStr(.Counters.Pena) & vbCrLf
        Print n, , "ScrollOro=" & CStr(.Counters.ScrollOro) & vbCrLf
        Print n, , "ScrollExperiencia=" & CStr(.Counters.ScrollExperiencia) & vbCrLf
        Print n, , "Oxigeno=" & CStr(.Counters.Oxigeno) & vbCrLf
        
        Print n, , vbCrLf

        Print n, , "[MUERTES]" & vbCrLf
        Print n, , "UserMuertes=" & CStr(.Stats.UsuariosMatados) & vbCrLf
        Print n, , "NpcsMuertes=" & CStr(.Stats.NPCsMuertos) & vbCrLf
        
        Print n, , vbCrLf
        
        'BANCO
        Print n, , "[BancoInventory]" & vbCrLf
        Print n, , "CantidadItems=" & val(.BancoInvent.NroItems) & vbCrLf

        Dim loopd As Long
        For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
            Print n, , "Obj" & loopd & "=" & .BancoInvent.Object(loopd).ObjIndex & "-" & .BancoInvent.Object(loopd).Amount & vbCrLf
        Next loopd
        
        Print n, , vbCrLf
        
        Print n, , "[LOGROS]" & vbCrLf
        Print n, , "UserLogros=" & CByte(.UserLogros) & vbCrLf
        Print n, , "NPcLogros=" & CByte(.NPcLogros) & vbCrLf
        Print n, , "LevelLogros=" & CByte(.LevelLogros) & vbCrLf
        
        Print n, , vbCrLf
        
        Print n, , "[BINDKEYS]" & vbCrLf
        Print n, , "ChatCombate=" & CByte(.ChatCombate) & vbCrLf
        Print n, , "ChatGlobal=" & CByte(.ChatGlobal) & vbCrLf
        
        Print n, , vbCrLf

        'HECHIZOS
        Print n, , "[HECHIZOS]" & vbCrLf
        
        Dim cad As String
        For LoopC = 1 To MAXUSERHECHIZOS
            cad = .Stats.UserHechizos(LoopC)
            Print n, , "H" & LoopC & "=" & cad & vbCrLf
        Next
        
        Print n, , vbCrLf
        
        'BATTLE
        Print n, , "[Battle]" & vbCrLf
        Print n, , "Puntos=" & CStr(.flags.BattlePuntos) & vbCrLf
        
        Print n, , vbCrLf
        
        Print n, , "[CORREO]" & vbCrLf & "NoLeidos=" & CByte(.Correo.NoLeidos) & vbCrLf
        Print n, , "CANTCORREO=" & CByte(.Correo.CantCorreo) & vbCrLf
        
        Print n, , vbCrLf
        
        'Correo Ladder
        With .Correo
        
            For LoopC = 1 To .CantCorreo
        
                Print n, , "REMITENTE" & LoopC & "=" & .Mensaje(LoopC).Remitente & vbCrLf
                Print n, , "MENSAJE" & LoopC & "=" & .Mensaje(LoopC).Mensaje & vbCrLf
                Print n, , "Item" & LoopC & "=" & .Mensaje(LoopC).Item & vbCrLf
                Print n, , "ItemCount" & LoopC & "=" & .Mensaje(LoopC).ItemCount & vbCrLf
                Print n, , "DATE" & LoopC & "=" & .Mensaje(LoopC).Fecha & vbCrLf
                Print n, , "LEIDO" & LoopC & "=" & .Mensaje(LoopC).Leido & vbCrLf
                
            Next LoopC
        
        End With
        
        Close #n
        
        Call SaveQuestStats(Userindex, UserFile)

        'Devuelve el head de muerto
        If .flags.Muerto = 1 Then
            .Char.Head = iCabezaMuerto
        End If
    
    End With
        
    Exit Sub

ErrHandler:
    Call LogError("Error en SaveUserCharfile")
    Close #n

End Sub

Sub SaveNewUser(ByVal Userindex As Integer)
        
        On Error GoTo SaveNewUser_Err
        
    
100     If Database_Enabled Then
102         Call SaveNewUserDatabase(Userindex)
        Else
104         Call SaveNewUserCharfile(Userindex)

        End If
    
        
        Exit Sub

SaveNewUser_Err:
106     Call RegistrarError(Err.Number, Err.description, "ES.SaveNewUser", Erl)
108     Resume Next
        
End Sub

Sub SaveNewUserCharfile(ByVal Userindex As Integer)
        '*************************************************
        'Author: Unknown
        'Last modified: 23/01/2007
        'Saves the Users records
        '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
        '*************************************************
    
        On Error GoTo ErrHandler
    
        Dim UserFile    As String

        Dim OldUserHead As Long
    
100     UserFile = CharPath & UCase$(UserList(Userindex).name) & ".chr"
    
        'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
        'clase=0 es el error, porq el enum empieza de 1!!
102     If UserList(Userindex).clase = 0 Or UserList(Userindex).Stats.ELV = 0 Then
104         Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & UserList(Userindex).name)
            Exit Sub

        End If
    
106     If FileExist(UserFile, vbNormal) Then
108         If UserList(Userindex).flags.Muerto = 1 Then
110             OldUserHead = UserList(Userindex).Char.Head
112             UserList(Userindex).Char.Head = GetVar(UserFile, "INIT", "Head")

            End If

            '       Kill UserFile
        End If
    
        Dim LoopC As Integer

        Dim n

        Dim Datos$

114     n = FreeFile
116     Open UserFile For Binary Access Write As n
    
        'BATTLE
118     Put n, , "[Battle]" & vbCrLf & "Puntos=" & CStr(UserList(Userindex).flags.BattlePuntos) & vbCrLf
    
120     Put n, , vbCrLf
    
        'FLAGS
122     Put n, , "[FLAGS]" & vbCrLf & "CASADO=" & CStr(UserList(Userindex).flags.Casado) & vbCrLf
124     Put n, , "PAREJA=" & vbCrLf
126     Put n, , "Muerto=0" & vbCrLf
128     Put n, , "Escondido=0" & vbCrLf
130     Put n, , "Hambre=0" & vbCrLf
132     Put n, , "Sed=0" & vbCrLf
134     Put n, , "Desnudo=0" & vbCrLf
136     Put n, , "Navegando=0" & vbCrLf
138     Put n, , "Envenenado=0" & vbCrLf
140     Put n, , "Paralizado=0" & vbCrLf
142     Put n, , "Inmovilizado=0" & vbCrLf
144     Put n, , "Incinerado=0" & vbCrLf
146     Put n, , "VecesQueMoriste=0" & vbCrLf
148     Put n, , "ScrollExp=" & CStr(UserList(Userindex).flags.ScrollExp) & vbCrLf
150     Put n, , "ScrollOro=" & CStr(UserList(Userindex).flags.ScrollOro) & vbCrLf
152     Put n, , "MinutosRestantes=0" & vbCrLf
154     Put n, , "SegundosPasados=0" & vbCrLf
156     Put n, , "Silenciado=0" & vbCrLf
158     Put n, , "Montado=0" & vbCrLf
    
160     Put n, , "InventLevel=0" & vbCrLf
    
162     Put n, , vbCrLf
    
164     Put n, , "[CONSEJO]" & vbCrLf
166     Put n, , "PERTENECE=0" & vbCrLf
168     Put n, , "PERTENECECAOS=0" & vbCrLf
    
170     Put n, , "[FACCIONES]" & vbCrLf & "EjercitoReal=" & CStr(UserList(Userindex).Faccion.ArmadaReal) & vbCrLf
172     Put n, , "Status=" & CStr(UserList(Userindex).Faccion.Status) & vbCrLf
174     Put n, , "EjercitoCaos=" & CStr(UserList(Userindex).Faccion.FuerzasCaos) & vbCrLf
176     Put n, , "CiudMatados=" & CStr(UserList(Userindex).Faccion.CiudadanosMatados) & vbCrLf
178     Put n, , "CrimMatados=" & CStr(UserList(Userindex).Faccion.CriminalesMatados) & vbCrLf
180     Put n, , "rArCaos=" & CStr(UserList(Userindex).Faccion.RecibioArmaduraCaos) & vbCrLf
182     Put n, , "rArReal=" & CStr(UserList(Userindex).Faccion.RecibioArmaduraReal) & vbCrLf
184     Put n, , "rExCaos=" & CStr(UserList(Userindex).Faccion.RecibioExpInicialCaos) & vbCrLf
186     Put n, , "rExReal=" & CStr(UserList(Userindex).Faccion.RecibioExpInicialReal) & vbCrLf
188     Put n, , "recCaos=" & CStr(UserList(Userindex).Faccion.RecompensasCaos) & vbCrLf
190     Put n, , "recReal=" & CStr(UserList(Userindex).Faccion.RecompensasReal) & vbCrLf
192     Put n, , "Reenlistadas=" & CStr(UserList(Userindex).Faccion.Reenlistadas) & vbCrLf
194     Put n, , "NivelIngreso=" & CStr(UserList(Userindex).Faccion.NivelIngreso) & vbCrLf
196     Put n, , "FechaIngreso=" & CStr(UserList(Userindex).Faccion.FechaIngreso) & vbCrLf
198     Put n, , "MatadosIngreso=" & CStr(UserList(Userindex).Faccion.MatadosIngreso) & vbCrLf
200     Put n, , "NextRecompensa=" & CStr(UserList(Userindex).Faccion.NextRecompensa) & vbCrLf
    
202     Put n, , vbCrLf
    
        'STATS
204     Put n, , "[STATS]" & vbCrLf & "GLD=0" & vbCrLf
206     Put n, , "BANCO=0" & vbCrLf
208     Put n, , "MaxHP=" & CStr(UserList(Userindex).Stats.MaxHp) & vbCrLf
210     Put n, , "MinHP=" & CStr(UserList(Userindex).Stats.MinHp) & vbCrLf
212     Put n, , "MaxSTA=" & CStr(UserList(Userindex).Stats.MaxSta) & vbCrLf
214     Put n, , "MinSTA=" & CStr(UserList(Userindex).Stats.MinSta) & vbCrLf
216     Put n, , "MaxMAN=" & CStr(UserList(Userindex).Stats.MaxMAN) & vbCrLf
218     Put n, , "MinMAN=" & CStr(UserList(Userindex).Stats.MinMAN) & vbCrLf
220     Put n, , "MaxHIT=" & CStr(UserList(Userindex).Stats.MaxHit) & vbCrLf
222     Put n, , "MinHIT=" & CStr(UserList(Userindex).Stats.MinHIT) & vbCrLf
224     Put n, , "MaxAGU=" & CStr(UserList(Userindex).Stats.MaxAGU) & vbCrLf
226     Put n, , "MinAGU=" & CStr(UserList(Userindex).Stats.MinAGU) & vbCrLf
228     Put n, , "MaxHAM=" & CStr(UserList(Userindex).Stats.MaxHam) & vbCrLf
230     Put n, , "MinHAM=" & CStr(UserList(Userindex).Stats.MinHam) & vbCrLf
232     Put n, , "SkillPtsLibres=" & CStr(UserList(Userindex).Stats.SkillPts) & vbCrLf
234     Put n, , "EXP=" & CStr(UserList(Userindex).Stats.Exp) & vbCrLf
236     Put n, , "ELV=" & CStr(UserList(Userindex).Stats.ELV) & vbCrLf
238     Put n, , "ELU=" & CStr(UserList(Userindex).Stats.ELU) & vbCrLf
    
240     Put n, , vbCrLf
    
        'MAHIA
242     Put n, , "[MAGIA]" & vbCrLf & "ENVENENA=0" & vbCrLf
244     Put n, , "PARALIZA=0" & vbCrLf
246     Put n, , "INCINERA=0" & vbCrLf
248     Put n, , "Estupidiza=0" & vbCrLf
250     Put n, , "PENDIENTE=0" & vbCrLf
252     Put n, , "CARROMINERIA=0" & vbCrLf
254     Put n, , "NOPALABRASMAGICAS=0" & vbCrLf
256     Put n, , "OTRA_AURA=0" & vbCrLf
258     Put n, , "DAÑOMAGICO=0" & vbCrLf
260     Put n, , "ResistenciaMagica=0" & vbCrLf
262     Put n, , "NoDetectable=0" & vbCrLf
264     Put n, , "AnilloOcultismo=0" & vbCrLf
266     Put n, , "RegeneracionMana=0" & vbCrLf
268     Put n, , "NoMagiaEfeceto=0" & vbCrLf
270     Put n, , "RegeneracionHP=0" & vbCrLf
272     Put n, , "RegeneracionSta=0" & vbCrLf
    
274     Put n, , vbCrLf
    
        'SKILLS
276     Put n, , "[SKILLS]" & vbCrLf

278     For LoopC = 1 To UBound(UserList(Userindex).Stats.UserSkills)
280         Put n, , "SK" & LoopC & "=0" & vbCrLf
        Next
    
282     Put n, , vbCrLf
    
        'INVENTARIO
284     Put n, , "[Inventory]" & vbCrLf & "CantidadItems=" & val(UserList(Userindex).Invent.NroItems) & vbCrLf

286     For LoopC = 1 To UserList(Userindex).CurrentInventorySlots
288         Put n, , "Obj" & LoopC & "=" & UserList(Userindex).Invent.Object(LoopC).ObjIndex & "-" & UserList(Userindex).Invent.Object(LoopC).Amount & "-" & UserList(Userindex).Invent.Object(LoopC).Equipped & vbCrLf
        Next
290     Put n, , "WeaponEqpSlot=" & CStr(UserList(Userindex).Invent.WeaponEqpSlot) & vbCrLf
292     Put n, , "HerramientaEqpSlot=" & CStr(UserList(Userindex).Invent.HerramientaEqpSlot) & vbCrLf
294     Put n, , "ArmourEqpSlot=" & CStr(UserList(Userindex).Invent.ArmourEqpSlot) & vbCrLf
296     Put n, , "CascoEqpSlot=" & CStr(UserList(Userindex).Invent.CascoEqpSlot) & vbCrLf
298     Put n, , "EscudoEqpSlot=" & CStr(UserList(Userindex).Invent.EscudoEqpSlot) & vbCrLf
300     Put n, , "BarcoSlot=" & CStr(UserList(Userindex).Invent.BarcoSlot) & vbCrLf
302     Put n, , "MonturaSlot=" & CStr(UserList(Userindex).Invent.MonturaSlot) & vbCrLf
304     Put n, , "MunicionSlot=" & CStr(UserList(Userindex).Invent.MunicionEqpSlot) & vbCrLf
306     Put n, , "AnilloSlot=" & CStr(UserList(Userindex).Invent.AnilloEqpSlot) & vbCrLf
308     Put n, , "MagicoSlot=" & CStr(UserList(Userindex).Invent.MagicoSlot) & vbCrLf
310     Put n, , "NudilloEqpSlot=" & CStr(UserList(Userindex).Invent.NudilloSlot) & vbCrLf
    
312     Put n, , vbCrLf
    
        'INIT
314     Put n, , "[INIT]" & vbCrLf & "Cuenta=" & UserList(Userindex).Cuenta & vbCrLf
316     Put n, , "Genero=" & UserList(Userindex).genero & vbCrLf
318     Put n, , "Raza=" & UserList(Userindex).raza & vbCrLf
320     Put n, , "Hogar=" & UserList(Userindex).Hogar & vbCrLf
322     Put n, , "Clase=" & UserList(Userindex).clase & vbCrLf
324     Put n, , "Desc=" & UserList(Userindex).Desc & vbCrLf
326     Put n, , "Heading=" & CStr(UserList(Userindex).Char.Heading) & vbCrLf
328     Put n, , "Head=" & CStr(UserList(Userindex).Char.Head) & vbCrLf
330     Put n, , "Arma=" & CStr(UserList(Userindex).Char.WeaponAnim) & vbCrLf
332     Put n, , "Escudo=" & CStr(UserList(Userindex).Char.ShieldAnim) & vbCrLf
334     Put n, , "Casco=" & CStr(UserList(Userindex).Char.CascoAnim) & vbCrLf
336     Put n, , "Position=" & UserList(Userindex).Pos.Map & "-" & UserList(Userindex).Pos.X & "-" & UserList(Userindex).Pos.Y & vbCrLf
        ' If UserList(UserIndex).flags.Muerto = 0 Then
338     Put n, , "Body=" & CStr(UserList(Userindex).Char.Body) & vbCrLf
        'Else
        '   Put N, , "Body=" & iCuerpoMuerto & vbCrLf 'poner body muerto
        '  End If
        #If ConUpTime Then

            Dim TempDate As Date

340         TempDate = Now - UserList(Userindex).LogOnTime
342         UserList(Userindex).LogOnTime = Now
344         UserList(Userindex).UpTime = UserList(Userindex).UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
346         UserList(Userindex).UpTime = UserList(Userindex).UpTime
348         Put n, , "UpTime=" & UserList(Userindex).UpTime & vbCrLf
        #End If
    
350     Put n, , vbCrLf
    
352     Put n, , "[ATRIBUTOS]" & vbCrLf

        '¿Fueron modificados los atributos del usuario?
354     For LoopC = 1 To UBound(UserList(Userindex).Stats.UserAtributos)
356         Put n, , "AT" & LoopC & "=" & CStr(UserList(Userindex).Stats.UserAtributos(LoopC)) & vbCrLf
        Next
358     Put n, , vbCrLf
    
        'baneo
360     Put n, , "[BAN]" & vbCrLf & "Baneado=" & CStr(UserList(Userindex).flags.Ban) & vbCrLf
362     Put n, , "BanMotivo=" & CStr(UserList(Userindex).flags.BanMotivo) & vbCrLf
    
364     Put n, , vbCrLf
    
        'COUNTERS
366     Put n, , "[COUNTERS]" & vbCrLf & "Pena=" & CStr(UserList(Userindex).Counters.Pena) & vbCrLf
368     Put n, , "ScrollOro=" & CStr(UserList(Userindex).Counters.ScrollOro) & vbCrLf
370     Put n, , "ScrollExperiencia=" & CStr(UserList(Userindex).Counters.ScrollExperiencia) & vbCrLf
372     Put n, , "Oxigeno=" & CStr(UserList(Userindex).Counters.Oxigeno) & vbCrLf
    
374     Put n, , vbCrLf
    
376     Put n, , "[MUERTES]" & vbCrLf & "UserMuertes=0" & vbCrLf
378     Put n, , "NpcsMuertes=0" & vbCrLf
    
380     Put n, , vbCrLf
    
        'BANCO
382     Put n, , "[BancoInventory]" & vbCrLf & "CantidadItems=0" & vbCrLf

        Dim loopd As Integer

384     For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
386         Put n, , "Obj" & loopd & "=" & UserList(Userindex).BancoInvent.Object(loopd).ObjIndex & "-" & UserList(Userindex).BancoInvent.Object(loopd).Amount & vbCrLf
388     Next loopd
    
390     Put n, , vbCrLf
    
392     Put n, , "[LOGROS]" & vbCrLf & "UserLogros=" & CByte(UserList(Userindex).UserLogros) & vbCrLf
394     Put n, , "NPcLogros=" & CByte(UserList(Userindex).NPcLogros) & vbCrLf
396     Put n, , "LevelLogros=" & CByte(UserList(Userindex).LevelLogros) & vbCrLf
    
398     Put n, , vbCrLf
    
400     Put n, , "[BINDKEYS]" & vbCrLf
402     Put n, , "ChatCombate=" & CByte(UserList(Userindex).ChatCombate) & vbCrLf
404     Put n, , "ChatGlobal=" & CByte(UserList(Userindex).ChatGlobal) & vbCrLf
    
406     Put n, , vbCrLf
    
        'HECHIZOS
408     Put n, , "[HECHIZOS]" & vbCrLf

        Dim cad As String

410     For LoopC = 1 To MAXUSERHECHIZOS
412         cad = UserList(Userindex).Stats.UserHechizos(LoopC)
414         Put n, , "H" & LoopC & "=" & cad & vbCrLf
        Next
      
416     Put n, , vbCrLf
    
418     Put n, , "[CORREO]" & vbCrLf & "NoLeidos=0" & vbCrLf
420     Put n, , "CANTCORREO=0" & vbCrLf
    
        'Correo Ladder
    
422     For LoopC = 1 To UserList(Userindex).Correo.CantCorreo
    
424         Put n, , "REMITENTE" & LoopC & "=" & UserList(Userindex).Correo.Mensaje(LoopC).Remitente & vbCrLf
426         Put n, , "MENSAJE" & LoopC & "=" & UserList(Userindex).Correo.Mensaje(LoopC).Mensaje & vbCrLf
428         Put n, , "Item" & LoopC & "=" & UserList(Userindex).Correo.Mensaje(LoopC).Item & vbCrLf
430         Put n, , "ItemCount" & LoopC & "=" & UserList(Userindex).Correo.Mensaje(LoopC).ItemCount & vbCrLf
432         Put n, , "DATE" & LoopC & "=" & UserList(Userindex).Correo.Mensaje(LoopC).Fecha & vbCrLf
434         Put n, , "LEIDO" & LoopC & "=" & UserList(Userindex).Correo.Mensaje(LoopC).Leido & vbCrLf
        
436     Next LoopC
    
438     Close #n
    
        'Devuelve el head de muerto
440     If UserList(Userindex).flags.Muerto = 1 Then
442         UserList(Userindex).Char.Head = iCabezaMuerto

        End If
    
        Exit Sub
    
ErrHandler:
444     Call LogError("Error en SaveNewUserCharfile")
446     Close #n

End Sub

Sub SetUserLogged(ByVal Userindex As Integer)
        
        On Error GoTo SetUserLogged_Err
        

100     If Database_Enabled Then
102         Call SetUserLoggedDatabase(UserList(Userindex).Id, UserList(Userindex).AccountID)
        Else
104         Call WriteVar(CharPath & UCase$(UserList(Userindex).name) & ".chr", "INIT", "Logged", 1)
106         Call WriteVar(CuentasPath & UCase$(UserList(Userindex).Cuenta) & ".act", "INIT", "LOGEADA", 1)

        End If

        
        Exit Sub

SetUserLogged_Err:
108     Call RegistrarError(Err.Number, Err.description, "ES.SetUserLogged", Erl)
110     Resume Next
        
End Sub

Sub SaveBattlePoints(ByVal Userindex As Integer)
        
        On Error GoTo SaveBattlePoints_Err
        
    
100     If Database_Enabled Then
102         Call SaveBattlePointsDatabase(UserList(Userindex).Id, UserList(Userindex).flags.BattlePuntos)
        Else
104         Call WriteVar(CharPath & UserList(Userindex).name & ".chr", "Battle", "Puntos", UserList(Userindex).flags.BattlePuntos)

        End If
    
        
        Exit Sub

SaveBattlePoints_Err:
106     Call RegistrarError(Err.Number, Err.description, "ES.SaveBattlePoints", Erl)
108     Resume Next
        
End Sub

Function Status(ByVal Userindex As Integer) As Byte
        
        On Error GoTo Status_Err
        

100     Status = UserList(Userindex).Faccion.Status

        
        Exit Function

Status_Err:
102     Call RegistrarError(Err.Number, Err.description, "ES.Status", Erl)
104     Resume Next
        
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
112     Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(Npclist(NpcIndex).Char.Heading))
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
162     Call RegistrarError(Err.Number, Err.description, "ES.BackUPnPc", Erl)
164     Resume Next
        
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
118     Npclist(NpcIndex).Char.Heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))

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
180     Call RegistrarError(Err.Number, Err.description, "ES.CargarNpcBackUp", Erl)
182     Resume Next
        
End Sub

Sub LogBan(ByVal BannedIndex As Integer, ByVal Userindex As Integer, ByVal motivo As String)
        
        On Error GoTo LogBan_Err
        

100     Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).name, "BannedBy", UserList(Userindex).name)
102     Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).name, "Reason", motivo)

        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim mifile As Integer

104     mifile = FreeFile
106     Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
108     Print #mifile, UserList(BannedIndex).name
110     Close #mifile

        
        Exit Sub

LogBan_Err:
112     Call RegistrarError(Err.Number, Err.description, "ES.LogBan", Erl)
114     Resume Next
        
End Sub

Sub LogBanFromName(ByVal BannedName As String, ByVal Userindex As Integer, ByVal motivo As String)
        
        On Error GoTo LogBanFromName_Err
        

100     Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(Userindex).name)
102     Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)

        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim mifile As Integer

104     mifile = FreeFile
106     Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
108     Print #mifile, BannedName
110     Close #mifile

        
        Exit Sub

LogBanFromName_Err:
112     Call RegistrarError(Err.Number, Err.description, "ES.LogBanFromName", Erl)
114     Resume Next
        
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
112     Call RegistrarError(Err.Number, Err.description, "ES.Ban", Erl)
114     Resume Next
        
End Sub

Public Sub CargaApuestas()
        
        On Error GoTo CargaApuestas_Err
        

100     Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
102     Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
104     Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

        
        Exit Sub

CargaApuestas_Err:
106     Call RegistrarError(Err.Number, Err.description, "ES.CargaApuestas", Erl)
108     Resume Next
        
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
148     Call RegistrarError(Err.Number, Err.description, "ES.LoadRecursosEspeciales", Erl)
150     Resume Next
        
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
    
        Dim Count As Long, i As Long, j As Long, str As String, Field() As String, nivel As Integer, MaxLvlCania As Long

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
140             For j = Peces(i).Amount To MaxLvlCania
142                 PesoPeces(j) = PesoPeces(j) + Peces(i).data
144             Next j

146             Peces(i).data = PesoPeces(Peces(i).Amount)
148         Next i
        Else
150         ReDim Peces(0) As obj

        End If
    
152     Set IniFile = Nothing

        
        Exit Sub

LoadPesca_Err:
154     Call RegistrarError(Err.Number, Err.description, "ES.LoadPesca", Erl)
156     Resume Next
        
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
132     Call RegistrarError(Err.Number, Err.description, "ES.QuickSortPeces", Erl)
134     Resume Next
        
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
124     Call RegistrarError(Err.Number, Err.description, "ES.BinarySearchPeces", Erl)
126     Resume Next
        
End Function

Public Sub LoadUserIntervals(ByVal Userindex As Integer)
        
        On Error GoTo LoadUserIntervals_Err
        

100     With UserList(Userindex).Intervals
102         .Arco = IntervaloFlechasCazadores
104         .Caminar = IntervaloCaminar
106         .Golpe = IntervaloUserPuedeAtacar
108         .magia = IntervaloUserPuedeCastear
110         .GolpeMagia = IntervaloGolpeMagia
112         .MagiaGolpe = IntervaloMagiaGolpe
114         .GolpeUsar = IntervaloGolpeUsar
116         .Trabajar = IntervaloUserPuedeTrabajar
118         .UsarU = IntervaloUserPuedeUsarU
120         .UsarClic = IntervaloUserPuedeUsarClic

        End With

        
        Exit Sub

LoadUserIntervals_Err:
122     Call RegistrarError(Err.Number, Err.description, "ES.LoadUserIntervals", Erl)
124     Resume Next
        
End Sub

Public Sub RegistrarError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
    '**********************************************************
    'Author: Jopi
    'Guarda una descripcion detallada del error en Errores.log
    '**********************************************************
        
        'Si lo del parametro Componente es ES IGUAL, al Componente del anterior error...
100     If Componente = HistorialError.Componente And _
           Numero = HistorialError.ErrorCode Then
       
           'Si ya recibimos error en el mismo componente 10 veces, es bastante probable que estemos en un bucle
            'x lo que no hace falta registrar el error.
102         If HistorialError.Contador = 10 Then Exit Sub
        
            'Agregamos el error al historial.
104         HistorialError.Contador = HistorialError.Contador + 1
        
        Else 'Si NO es igual, reestablecemos el contador.

106         HistorialError.Contador = 0
108         HistorialError.ErrorCode = Numero
110         HistorialError.Componente = Componente
            
        End If
    
        'Registramos el error en Errores.log
112     Dim File As Integer: File = FreeFile
        
114     Open App.Path & "\logs\Errores.log" For Append As #File
    
116         Print #File, "Error: " & Numero
118         Print #File, "Descripcion: " & Descripcion
        
120         If LenB(Linea) <> 0 Then
122             Print #File, "Linea: " & Linea
            End If
        
124         Print #File, "Componente: " & Componente
126         Print #File, "Fecha y Hora: " & Date$ & "-" & Time$
        
128         Print #File, vbNullString
        
130     Close #File
    
132     Debug.Print "Error: " & Numero & vbNewLine & _
                    "Descripcion: " & Descripcion & vbNewLine & _
                    "Componente: " & Componente & vbNewLine & _
                    "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine
End Sub

Function CountFiles(strFolder As String, strPattern As String) As Integer
   
        Dim strFile As String
100         strFile = dir$(strFolder & "\" & strPattern)
    
102     Do Until Len(strFile) = 0
104         CountFiles = CountFiles + 1
106         strFile = dir$()
        Loop
    
108     If CountFiles <> 0 Then CountFiles = CountFiles + 1
    
End Function
