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
    level As Long
    extra2 As Long
    Salida As String
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
116     Call RegistrarError(Err.Number, Err.Description, "ES.CargarSpawnList", Erl)
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
102     Call RegistrarError(Err.Number, Err.Description, "ES.EsAdmin", Erl)
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
102     Call RegistrarError(Err.Number, Err.Description, "ES.EsDios", Erl)
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
102     Call RegistrarError(Err.Number, Err.Description, "ES.EsSemiDios", Erl)
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
102     Call RegistrarError(Err.Number, Err.Description, "ES.EsConsejero", Erl)
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
102     Call RegistrarError(Err.Number, Err.Description, "ES.EsRolesMaster", Erl)
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
110     Call RegistrarError(Err.Number, Err.Description, "ES.EsGmChar", Erl)
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
        
        ' Anti-choreo de GM's
100     Set AdministratorAccounts = New Dictionary
        Dim TempName() As String
       
        ' Public container
102     Set Administradores = New clsIniReader
    
        ' Server ini info file
        Dim ServerIni As clsIniReader
104     Set ServerIni = New clsIniReader
106     Call ServerIni.Initialize(IniPath & "Server.ini")
       
        ' Admines
108     buf = val(ServerIni.GetValue("INIT", "Admines"))
    
110     For i = 1 To buf
112         name = UCase$(ServerIni.GetValue("Admines", "Admin" & i))
114         TempName = Split(name, "|", , vbTextCompare)
        
            ' Si NO declara el mail de la cuenta en el Server.ini, NO le doy privilegios.
116         If UBound(TempName()) > 0 Then
                ' AdministratorAccounts("Email") = Nick
118             AdministratorAccounts(TempName(1)) = TempName(0)
            
                ' Add key
120             Call Administradores.ChangeValue("Admin", TempName(0), "1")
            End If
        
122     Next i
    
        ' Dioses
124     buf = val(ServerIni.GetValue("INIT", "Dioses"))
    
126     For i = 1 To buf
128         name = UCase$(ServerIni.GetValue("Dioses", "Dios" & i))
130         TempName = Split(name, "|", , vbTextCompare)
        
            ' Si NO declara el mail de la cuenta en el Server.ini, NO le doy privilegios.
132         If UBound(TempName()) > 0 Then
                ' AdministratorAccounts("Email") = Nick
134             AdministratorAccounts(TempName(1)) = TempName(0)
            
                ' Add key
136             Call Administradores.ChangeValue("Dios", TempName(0), "1")
            End If
        
138     Next i
        
        ' SemiDioses
140     buf = val(ServerIni.GetValue("INIT", "SemiDioses"))
    
142     For i = 1 To buf
144         name = UCase$(ServerIni.GetValue("SemiDioses", "SemiDios" & i))
146         TempName = Split(name, "|", , vbTextCompare)
        
            ' Si NO declara el mail de la cuenta en el Server.ini, NO le doy privilegios.
148         If UBound(TempName()) > 0 Then
                ' AdministratorAccounts("Email") = Nick
150             AdministratorAccounts(TempName(1)) = TempName(0)
            
                ' Add key
152             Call Administradores.ChangeValue("SemiDios", TempName(0), "1")
            End If
        
154     Next i
    
        ' Consejeros
156     buf = val(ServerIni.GetValue("INIT", "Consejeros"))
        
158     For i = 1 To buf
160         name = UCase$(ServerIni.GetValue("Consejeros", "Consejero" & i))
162         TempName = Split(name, "|", , vbTextCompare)
        
            ' Si NO declara el mail de la cuenta en el Server.ini, NO le doy privilegios.
164         If UBound(TempName()) > 0 Then
                ' AdministratorAccounts("Email") = Nick
166             AdministratorAccounts(TempName(1)) = TempName(0)
            
                ' Add key
168             Call Administradores.ChangeValue("Consejero", TempName(0), "1")
            End If
        
170     Next i
    
        ' RolesMasters
172     buf = val(ServerIni.GetValue("INIT", "RolesMasters"))
        
174     For i = 1 To buf
176         name = UCase$(ServerIni.GetValue("RolesMasters", "RM" & i))
178         TempName = Split(name, "|", , vbTextCompare)
        
            ' Si NO declara el mail de la cuenta en el Server.ini, NO le doy privilegios.
180         If UBound(TempName()) > 0 Then
                ' AdministratorAccounts("Email") = Nick
182             AdministratorAccounts(TempName(1)) = TempName(0)
            
                ' Add key
184             Call Administradores.ChangeValue("RM", TempName(0), "1")
            End If
        
186     Next i

188     Set ServerIni = Nothing

        'If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & Time & " - Los Administradores/Dioses/Gms se han cargado correctamente."

        
        Exit Sub

loadAdministrativeUsers_Err:
190     Call RegistrarError(Err.Number, Err.Description, "ES.loadAdministrativeUsers", Erl)
192     Resume Next
        
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
120     Call RegistrarError(Err.Number, Err.Description, "ES.GetCharPrivs", Erl)
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
116     Call RegistrarError(Err.Number, Err.Description, "ES.TxtDimension", Erl)
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
120     Call RegistrarError(Err.Number, Err.Description, "ES.CargarForbidenWords", Erl)
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
260         Hechizos(Hechizo).Cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
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
288     MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.Description
 
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
112     Call RegistrarError(Err.Number, Err.Description, "ES.LoadMotd", Erl)
114     Resume Next
        
End Sub

Public Sub DoBackUp()
        'Call LogTarea("Sub DoBackUp")
        
        On Error GoTo DoBackUp_Err
    
        
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
        

110     Dim nfile As Integer: nfile = FreeFile ' obtenemos un canal
112     Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
114     Print #nfile, Date & " " & Time
116     Close #nfile

        
        Exit Sub

DoBackUp_Err:
118     Call RegistrarError(Err.Number, Err.Description, "ES.DoBackUp", Erl)

        
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
114     Call RegistrarError(Err.Number, Err.Description, "ES.LoadArmasHerreria", Erl)
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
114     Call RegistrarError(Err.Number, Err.Description, "ES.LoadArmadurasHerreria", Erl)
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
126             .Vida = val(BalanceIni.GetValue("MODVIDA", SearchVar))
128             .ManaInicial = val(BalanceIni.GetValue("MANA_INICIAL", SearchVar))
130             .MultMana = val(BalanceIni.GetValue("MULT_MANA", SearchVar))
132             .AumentoSta = val(BalanceIni.GetValue("AUMENTO_STA", SearchVar))
134             .HitPre36 = val(BalanceIni.GetValue("GOLPE_PRE_36", SearchVar))
136             .HitPost36 = val(BalanceIni.GetValue("GOLPE_POST_36", SearchVar))
            End With

138     Next i
    
        'Modificadores de Raza
140     For i = 1 To NUMRAZAS
142         SearchVar = Replace$(Tilde(ListaRazas(i)), " ", vbNullString)

144         With ModRaza(i)
146             .Fuerza = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Fuerza"))
148             .Agilidad = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Agilidad"))
150             .Inteligencia = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Inteligencia"))
152             .Carisma = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Carisma"))
154             .Constitucion = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Constitucion"))
            End With

156     Next i

        'Extra
158     PorcentajeRecuperoMana = val(BalanceIni.GetValue("EXTRA", "PorcentajeRecuperoMana"))
160     DificultadSubirSkill = val(BalanceIni.GetValue("EXTRA", "DificultadSubirSkill"))
162     InfluenciaPromedioVidas = val(BalanceIni.GetValue("EXTRA", "InfluenciaPromedioVidas"))
164     DesbalancePromedioVidas = val(BalanceIni.GetValue("EXTRA", "DesbalancePromedioVidas"))
166     RangoVidas = val(BalanceIni.GetValue("EXTRA", "RangoVidas"))
168     ModDañoGolpeCritico = val(BalanceIni.GetValue("EXTRA", "ModDañoGolpeCritico"))

        ' Exp
        For i = 1 To STAT_MAXELV
170         ExpLevelUp(i) = val(BalanceIni.GetValue("EXP", i))
        Next
    
172     Set BalanceIni = Nothing
    
174     AgregarAConsola "Se cargó el balance (Balance.dat)"

        
        Exit Sub

LoadBalance_Err:
176     Call RegistrarError(Err.Number, Err.Description, "ES.LoadBalance", Erl)
178     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.Description, "ES.LoadObjCarpintero", Erl)
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
114     Call RegistrarError(Err.Number, Err.Description, "ES.LoadObjAlquimista", Erl)
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
114     Call RegistrarError(Err.Number, Err.Description, "ES.LoadObjSastre", Erl)
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
114     Call RegistrarError(Err.Number, Err.Description, "ES.LoadObjDonador", Erl)
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
    
108     With frmCargando.cargar
110         .min = 0
112         .max = NumObjDatas
114         .Value = 0
        End With
    
116     ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
    
        Dim str As String, Field() As String
  
        'Llena la lista
118     For Object = 1 To NumObjDatas
        
120         With ObjData(Object)
        
122             .name = Leer.GetValue("OBJ" & Object, "Name")
    
                ' If .Name = "" Then
                '   Call LogError("Objeto libre:" & Object)
                ' End If
    
                ' If .name = "" Then
                ' Debug.Print Object
                ' End If
    
                'Pablo (ToxicWaste) Log de Objetos.
124             .Log = val(Leer.GetValue("OBJ" & Object, "Log"))
126             .NoLog = val(Leer.GetValue("OBJ" & Object, "NoLog"))
                '07/09/07
    
128             .GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex"))

130             If .GrhIndex = 0 Then
132                 .GrhIndex = .GrhIndex

                End If
    
134             .OBJType = val(Leer.GetValue("OBJ" & Object, "ObjType"))
    
136             .Newbie = val(Leer.GetValue("OBJ" & Object, "Newbie"))

                'Propiedades by Lader 05-05-08
138             .Instransferible = val(Leer.GetValue("OBJ" & Object, "Instransferible"))
140             .Destruye = val(Leer.GetValue("OBJ" & Object, "Destruye"))
142             .Intirable = val(Leer.GetValue("OBJ" & Object, "Intirable"))
    
144             .CantidadSkill = val(Leer.GetValue("OBJ" & Object, "CantidadSkill"))
146             .QueSkill = val(Leer.GetValue("OBJ" & Object, "QueSkill"))
148             .QueAtributo = val(Leer.GetValue("OBJ" & Object, "queatributo"))
150             .CuantoAumento = val(Leer.GetValue("OBJ" & Object, "cuantoaumento"))
152             .MinELV = val(Leer.GetValue("OBJ" & Object, "MinELV"))
154             .Subtipo = val(Leer.GetValue("OBJ" & Object, "Subtipo"))
156             .Dorada = val(Leer.GetValue("OBJ" & Object, "Dorada"))
158             .VidaUtil = val(Leer.GetValue("OBJ" & Object, "VidaUtil"))
160             .TiempoRegenerar = val(Leer.GetValue("OBJ" & Object, "TiempoRegenerar"))
    
162             .donador = val(Leer.GetValue("OBJ" & Object, "donador"))
    
                Dim i As Integer

                'Propiedades by Lader 05-05-08
164             Select Case .OBJType

                    Case eOBJType.otHerramientas
166                     .WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
168                     .Power = val(Leer.GetValue("OBJ" & Object, "Poder"))
            
170                 Case eOBJType.otArmadura
172                     .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
174                     .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
176                     .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
178                     .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
180                     .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
182                     .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
184                     .ResistenciaMagica = val(Leer.GetValue("OBJ" & Object, "ResistenciaMagica"))
        
186                 Case eOBJType.otESCUDO
188                     .ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
190                     .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
192                     .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
194                     .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
196                     .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
198                     .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
200                     .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
202                     .ResistenciaMagica = val(Leer.GetValue("OBJ" & Object, "ResistenciaMagica"))
        
204                 Case eOBJType.otCASCO
206                     .CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
208                     .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
210                     .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
212                     .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
214                     .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
216                     .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
218                     .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
220                     .ResistenciaMagica = val(Leer.GetValue("OBJ" & Object, "ResistenciaMagica"))
        
222                 Case eOBJType.otWeapon
224                     .WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
226                     .Apuñala = val(Leer.GetValue("OBJ" & Object, "Apuñala"))
228                     .Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
230                     .Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
232                     .Estupidiza = val(Leer.GetValue("OBJ" & Object, "Estupidiza"))
        
234                     .MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
236                     .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
238                     .Proyectil = val(Leer.GetValue("OBJ" & Object, "Proyectil"))
240                     .Municion = val(Leer.GetValue("OBJ" & Object, "Municiones"))
242                     .Power = val(Leer.GetValue("OBJ" & Object, "StaffPower"))
244                     .MagicDamageBonus = val(Leer.GetValue("OBJ" & Object, "MagicDamageBonus"))
246                     .Refuerzo = val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
            
248                     .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
250                     .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
252                     .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
254                     .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
256                     .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
258                     .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
260                     .EfectoMagico = val(Leer.GetValue("OBJ" & Object, "efectomagico"))
262                     .Revive = val(Leer.GetValue("OBJ" & Object, "Revive")) <> 0
        
264                 Case eOBJType.otInstrumentos
        
                        'Pablo (ToxicWaste)
266                     .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
268                     .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
270                 Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
272                     .IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
274                     .IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
276                     .IndexCerradaLlave = val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
        
278                 Case otPociones
280                     .TipoPocion = val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
282                     .MaxModificador = val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
284                     .MinModificador = val(Leer.GetValue("OBJ" & Object, "MinModificador"))
            
286                     .DuracionEfecto = val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
288                     .Raices = val(Leer.GetValue("OBJ" & Object, "Raices"))
290                     .SkPociones = val(Leer.GetValue("OBJ" & Object, "SkPociones"))
292                     .Porcentaje = val(Leer.GetValue("OBJ" & Object, "Porcentaje"))
        
294                 Case eOBJType.otBarcos
296                     .MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
298                     .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
300                     .Velocidad = val(Leer.GetValue("OBJ" & Object, "Velocidad"))

302                 Case eOBJType.otMonturas
304                     .MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
306                     .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
308                     .MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
310                     .MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
312                     .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
314                     .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
316                 Case eOBJType.otFlechas
318                     .MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
320                     .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
            
322                     .Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
324                     .Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
            
                        'Case eOBJType.otAnillos 'Pablo (ToxicWaste)
                        '  .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                        '  .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                        '  .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                        '  .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            
                        'Pasajes Ladder 05-05-08
326                 Case eOBJType.otpasajes
328                     .DesdeMap = val(Leer.GetValue("OBJ" & Object, "DesdeMap"))
330                     .HastaMap = val(Leer.GetValue("OBJ" & Object, "Map"))
332                     .HastaX = val(Leer.GetValue("OBJ" & Object, "X"))
334                     .HastaY = val(Leer.GetValue("OBJ" & Object, "Y"))
336                     .NecesitaNave = val(Leer.GetValue("OBJ" & Object, "NecesitaNave"))
            
338                 Case eOBJType.OtDonador
340                     .HastaMap = val(Leer.GetValue("OBJ" & Object, "Map"))
342                     .HastaX = val(Leer.GetValue("OBJ" & Object, "X"))
344                     .HastaY = val(Leer.GetValue("OBJ" & Object, "Y"))
        
346                 Case eOBJType.otMagicos
348                     .EfectoMagico = val(Leer.GetValue("OBJ" & Object, "efectomagico"))

350                     If .EfectoMagico = 15 Then
352                         PENDIENTE = Object

                        End If
            
354                 Case eOBJType.otRunas
356                     .TipoRuna = val(Leer.GetValue("OBJ" & Object, "TipoRuna"))
358                     .DesdeMap = val(Leer.GetValue("OBJ" & Object, "DesdeMap"))
360                     .HastaMap = val(Leer.GetValue("OBJ" & Object, "Map"))
362                     .HastaX = val(Leer.GetValue("OBJ" & Object, "X"))
364                     .HastaY = val(Leer.GetValue("OBJ" & Object, "Y"))
                    
366                 Case eOBJType.otNUDILLOS
368                     .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
370                     .MaxHit = val(Leer.GetValue("OBJ" & Object, "MaxHit"))
372                     .Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
374                     .Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
376                     .Estupidiza = val(Leer.GetValue("OBJ" & Object, "Estupidiza"))
378                     .WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
380                     .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            
382                 Case eOBJType.otPergaminos
        
                        ' .ClasePermitida = Leer.GetValue("OBJ" & Object, "CP")
        
384                 Case eOBJType.OtCofre
386                     .CantItem = val(Leer.GetValue("OBJ" & Object, "CantItem"))

388                     If .Subtipo = 1 Then
390                         ReDim .Item(1 To .CantItem)
                
392                         For i = 1 To .CantItem
394                             .Item(i).ObjIndex = val(Leer.GetValue("OBJ" & Object, "Item" & i))
396                             .Item(i).Amount = val(Leer.GetValue("OBJ" & Object, "Cantidad" & i))
398                         Next i

                        Else
400                         ReDim .Item(1 To .CantItem)
                
402                         .CantEntrega = val(Leer.GetValue("OBJ" & Object, "CantEntrega"))

404                         For i = 1 To .CantItem
406                             .Item(i).ObjIndex = val(Leer.GetValue("OBJ" & Object, "Item" & i))
408                             .Item(i).Amount = val(Leer.GetValue("OBJ" & Object, "Cantidad" & i))
410                         Next i

                        End If
            
412                 Case eOBJType.otYacimiento
                        ' Drop gemas yacimientos
414                     .CantItem = val(Leer.GetValue("OBJ" & Object, "Gemas"))
            
416                     If .CantItem > 0 Then
418                         ReDim .Item(1 To .CantItem)

420                         For i = 1 To .CantItem
422                             str = Leer.GetValue("OBJ" & Object, "Gema" & i)
424                             Field = Split(str, "-")
426                             .Item(i).ObjIndex = val(Field(0))    ' ObjIndex
428                             .Item(i).Amount = val(Field(1))      ' Probabilidad de drop (1 en X)
430                         Next i

                        End If
                
432                 Case eOBJType.otDañoMagico
434                     .MagicDamageBonus = val(Leer.GetValue("OBJ" & Object, "MagicDamageBonus"))
436                     .Revive = val(Leer.GetValue("OBJ" & Object, "Revive")) <> 0

438                 Case eOBJType.otResistencia
440                     .ResistenciaMagica = val(Leer.GetValue("OBJ" & Object, "ResistenciaMagica"))
            
                End Select
    
442             .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))

444             .Elfico = val(Leer.GetValue("OBJ" & Object, "Elfico"))

446             .Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
448             .Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
450             .Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
                'DELETE
452             .SndAura = val(Leer.GetValue("OBJ" & Object, "SndAura"))
                '
    
454             .NoSeLimpia = val(Leer.GetValue("OBJ" & Object, "NoSeLimpia"))
456             .Subastable = val(Leer.GetValue("OBJ" & Object, "Subastable"))
    
458             .ParticulaGolpe = val(Leer.GetValue("OBJ" & Object, "ParticulaGolpe"))
460             .ParticulaViaje = val(Leer.GetValue("OBJ" & Object, "ParticulaViaje"))
462             .ParticulaGolpeTime = val(Leer.GetValue("OBJ" & Object, "ParticulaGolpeTime"))
    
464             .Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
466             .HechizoIndex = val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
    
468             .LingoteIndex = val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
    
470             .MineralIndex = val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
    
472             .MaxHp = val(Leer.GetValue("OBJ" & Object, "MaxHP"))
474             .MinHp = val(Leer.GetValue("OBJ" & Object, "MinHP"))
    
476             .Mujer = val(Leer.GetValue("OBJ" & Object, "Mujer"))
478             .Hombre = val(Leer.GetValue("OBJ" & Object, "Hombre"))
    
480             .PielLobo = val(Leer.GetValue("OBJ" & Object, "PielLobo"))
482             .PielOsoPardo = val(Leer.GetValue("OBJ" & Object, "PielOsoPardo"))
484             .PielOsoPolaR = val(Leer.GetValue("OBJ" & Object, "PielOsoPolaR"))
486             .SkMAGOria = val(Leer.GetValue("OBJ" & Object, "SKSastreria"))
    
488             .CreaParticula = Leer.GetValue("OBJ" & Object, "CreaParticula")
    
490             .CreaFX = val(Leer.GetValue("OBJ" & Object, "CreaFX"))
  
                'DELETE
492             .CreaParticulaPiso = val(Leer.GetValue("OBJ" & Object, "CreaParticulaPiso"))
                '
    
494             .CreaGRH = Leer.GetValue("OBJ" & Object, "CreaGRH")
496             .CreaLuz = Leer.GetValue("OBJ" & Object, "CreaLuz")
    
498             .MinHam = val(Leer.GetValue("OBJ" & Object, "MinHam"))
500             .MinSed = val(Leer.GetValue("OBJ" & Object, "MinAgu"))
    
502             .MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
504             .MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
506             .def = (.MinDef + .MaxDef) / 2
    
508             .ClaseTipo = val(Leer.GetValue("OBJ" & Object, "ClaseTipo"))
510             .RazaTipo = val(Leer.GetValue("OBJ" & Object, "RazaTipo"))

512             .RazaEnana = val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
514             .RazaDrow = val(Leer.GetValue("OBJ" & Object, "RazaDrow"))
516             .RazaElfa = val(Leer.GetValue("OBJ" & Object, "RazaElfa"))
518             .RazaGnoma = val(Leer.GetValue("OBJ" & Object, "RazaGnoma"))
    
520             .RazaOrca = val(Leer.GetValue("OBJ" & Object, "RazaOrca"))
    
522             .RazaHumana = val(Leer.GetValue("OBJ" & Object, "RazaHumana"))
    
524             .Valor = val(Leer.GetValue("OBJ" & Object, "Valor"))
    
526             .Crucial = val(Leer.GetValue("OBJ" & Object, "Crucial"))
    
                '.Cerrada = val(Leer.GetValue("OBJ" & Object, "abierta")) cerrada = abierta??? WTF???????
528             .Cerrada = val(Leer.GetValue("OBJ" & Object, "Cerrada"))

530             If .Cerrada = 1 Then
532                 .Llave = val(Leer.GetValue("OBJ" & Object, "Llave"))
534                 .clave = val(Leer.GetValue("OBJ" & Object, "Clave"))

                End If
    
                'Puertas y llaves
536             .clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
    
538             .texto = Leer.GetValue("OBJ" & Object, "Texto")
540             .GrhSecundario = val(Leer.GetValue("OBJ" & Object, "VGrande"))
    
542             .Agarrable = val(Leer.GetValue("OBJ" & Object, "Agarrable"))
544             .ForoID = Leer.GetValue("OBJ" & Object, "ID")
    
                'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico  -  Nunca más papu
                Dim n As Integer
                Dim s As String

546             For i = 1 To NUMCLASES
548                 s = UCase$(Leer.GetValue("OBJ" & Object, "CP" & i))
550                 n = 1

552                 Do While LenB(s) > 0 And Tilde(ListaClases(n)) <> Trim$(s)
554                     n = n + 1
                    Loop
            
556                 .ClaseProhibida(i) = IIf(LenB(s) > 0, n, 0)
558             Next i
        
560             For i = 1 To NUMRAZAS
562                 s = UCase$(Leer.GetValue("OBJ" & Object, "RP" & i))
564                 n = 1

566                 Do While LenB(s) > 0 And Tilde(ListaRazas(n)) <> Trim$(s)
568                     n = n + 1
                    Loop
            
570                 .RazaProhibida(i) = IIf(LenB(s) > 0, n, 0)
572             Next i
        
                ' Skill requerido
574             str = Leer.GetValue("OBJ" & Object, "SkillRequerido")

576             If Len(str) > 0 Then
578                 Field = Split(str, "-")
            
580                 n = 1

582                 Do While LenB(Field(0)) > 0 And Tilde(SkillsNames(n)) <> Tilde(Field(0))
584                     n = n + 1
                    Loop
    
586                 .SkillIndex = IIf(LenB(Field(0)) > 0, n, 0)
588                 .SkillRequerido = val(Field(1))

                End If

                ' -----------------
    
590             .SkCarpinteria = val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
    
                'If .SkCarpinteria > 0 Then
592             .Madera = val(Leer.GetValue("OBJ" & Object, "Madera"))

594             .MaderaElfica = val(Leer.GetValue("OBJ" & Object, "MaderaElfica"))
    
                'Bebidas
596             .MinSta = val(Leer.GetValue("OBJ" & Object, "MinST"))
    
598             .NoSeCae = val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
    
600             frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        
            End With
        
602     Next Object

604     Set Leer = Nothing

        Exit Sub

ErrHandler:
606     MsgBox "error cargando objetos " & Err.Number & ": " & Err.Description & ". Error producido al cargar el objeto: " & Object

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

        'UserList(UserIndex).flags.DañoMagico = CByte(UserFile.GetValue("MAGIA", "DañoMagico"))
        'UserList(UserIndex).flags.ResistenciaMagica = CByte(UserFile.GetValue("MAGIA", "ResistenciaMagica"))

        'Nuevos
174     UserList(UserIndex).flags.RegeneracionMana = CByte(UserFile.GetValue("MAGIA", "RegeneracionMana"))
176     UserList(UserIndex).flags.AnilloOcultismo = CByte(UserFile.GetValue("MAGIA", "AnilloOcultismo"))
178     UserList(UserIndex).flags.NoDetectable = CByte(UserFile.GetValue("MAGIA", "NoDetectable"))
180     UserList(UserIndex).flags.NoMagiaEfeceto = CByte(UserFile.GetValue("MAGIA", "NoMagiaEfeceto"))
182     UserList(UserIndex).flags.RegeneracionHP = CByte(UserFile.GetValue("MAGIA", "RegeneracionHP"))
184     UserList(UserIndex).flags.RegeneracionSta = CByte(UserFile.GetValue("MAGIA", "RegeneracionSta"))

186     UserList(UserIndex).Stats.UsuariosMatados = CLng(UserFile.GetValue("MUERTES", "UserMuertes"))
188     UserList(UserIndex).Stats.NPCsMuertos = CInt(UserFile.GetValue("MUERTES", "NpcsMuertes"))

190     UserList(UserIndex).Stats.InventLevel = CInt(UserFile.GetValue("STATS", "InventLevel"))

192     If CByte(UserFile.GetValue("CONSEJO", "PERTENECE")) Then UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.RoyalCouncil

194     If CByte(UserFile.GetValue("CONSEJO", "PERTENECECAOS")) Then UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.ChaosCouncil

        
        Exit Sub

LoadUserStats_Err:
196     Call RegistrarError(Err.Number, Err.Description, "ES.LoadUserStats", Erl)
198     Resume Next
        
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
184     UserList(UserIndex).Char.Heading = CInt(UserFile.GetValue("INIT", "Heading"))

186     UserList(UserIndex).OrigChar.Head = CInt(UserFile.GetValue("INIT", "Head"))
188     UserList(UserIndex).OrigChar.Body = CInt(UserFile.GetValue("INIT", "Body"))
190     UserList(UserIndex).OrigChar.WeaponAnim = CInt(UserFile.GetValue("INIT", "Arma"))
192     UserList(UserIndex).OrigChar.ShieldAnim = CInt(UserFile.GetValue("INIT", "Escudo"))
194     UserList(UserIndex).OrigChar.CascoAnim = CInt(UserFile.GetValue("INIT", "Casco"))

        #If ConUpTime Then
196         UserList(UserIndex).UpTime = CLng(UserFile.GetValue("INIT", "UpTime"))
        #End If

198     UserList(UserIndex).OrigChar.Heading = UserList(UserIndex).Char.Heading

200     If UserList(UserIndex).flags.Muerto = 0 Then
202         UserList(UserIndex).Char = UserList(UserIndex).OrigChar
        Else
204         UserList(UserIndex).Char.Body = iCuerpoMuerto
206         UserList(UserIndex).Char.Head = 0
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
234     UserList(UserIndex).Pos.X = CInt(ReadField(2, UserFile.GetValue("INIT", "Position"), 45))
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
280     UserList(UserIndex).Invent.DañoMagicoEqpSlot = CByte(UserFile.GetValue("Inventory", "DMSlot"))
282     UserList(UserIndex).Invent.ResistenciaEqpSlot = CByte(UserFile.GetValue("Inventory", "RMSlot"))
284     UserList(UserIndex).Invent.MagicoSlot = CByte(UserFile.GetValue("Inventory", "MagicoSlot"))
286     UserList(UserIndex).Invent.NudilloSlot = CByte(UserFile.GetValue("Inventory", "NudilloEqpSlot"))

288     UserList(UserIndex).ChatCombate = CByte(UserFile.GetValue("BINDKEYS", "ChatCombate"))
290     UserList(UserIndex).ChatGlobal = CByte(UserFile.GetValue("BINDKEYS", "ChatGlobal"))

292     UserList(UserIndex).Correo.CantCorreo = CByte(UserFile.GetValue("CORREO", "CantCorreo"))
294     UserList(UserIndex).Correo.NoLeidos = CByte(UserFile.GetValue("CORREO", "NoLeidos"))

296     For LoopC = 1 To UserList(UserIndex).Correo.CantCorreo
298         UserList(UserIndex).Correo.Mensaje(LoopC).Remitente = UserFile.GetValue("CORREO", "REMITENTE" & LoopC)
300         UserList(UserIndex).Correo.Mensaje(LoopC).Mensaje = UserFile.GetValue("CORREO", "MENSAJE" & LoopC)
302         UserList(UserIndex).Correo.Mensaje(LoopC).Item = UserFile.GetValue("CORREO", "Item" & LoopC)
304         UserList(UserIndex).Correo.Mensaje(LoopC).ItemCount = CByte(UserFile.GetValue("CORREO", "ItemCount" & LoopC))
306         UserList(UserIndex).Correo.Mensaje(LoopC).Fecha = UserFile.GetValue("CORREO", "DATE" & LoopC)
308         UserList(UserIndex).Correo.Mensaje(LoopC).Leido = CByte(UserFile.GetValue("CORREO", "LEIDO" & LoopC))
310     Next LoopC

        'Logros Ladder
312     UserList(UserIndex).UserLogros = UserFile.GetValue("LOGROS", "UserLogros")
314     UserList(UserIndex).NPcLogros = UserFile.GetValue("LOGROS", "NPcLogros")
316     UserList(UserIndex).LevelLogros = UserFile.GetValue("LOGROS", "LevelLogros")
        'Logros Ladder

318     ln = UserFile.GetValue("Guild", "GUILDINDEX")

320     If IsNumeric(ln) Then
322         UserList(UserIndex).GuildIndex = CInt(ln)
        Else
324         UserList(UserIndex).GuildIndex = 0

        End If

        
        Exit Sub

LoadUserInit_Err:
326     Call RegistrarError(Err.Number, Err.Description, "ES.LoadUserInit", Erl)
328     Resume Next
        
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
110     Call RegistrarError(Err.Number, Err.Description, "ES.GetVar", Erl)
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
134     Call RegistrarError(Err.Number, Err.Description, "ES.CargarBackUp", Erl)
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
136     Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

End Sub

Public Sub CargarMapaFormatoCSM(ByVal Map As Long, ByVal MAPFl As String)

        On Error GoTo ErrorHandler:

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
        
100     If Not FileExist(MAPFl, vbNormal) Then
102         Call RegistrarError(404, "Estas tratando de cargar un MAPA que NO EXISTE" & vbNewLine & "Mapa: " & MAPFl, "ES.CargarMapaFormatoCSM")
            Exit Sub
        End If
        
104     If FileLen(MAPFl) = 0 Then
106         Call RegistrarError(500, "Se trato de cargar un mapa corrupto o mal generado" & vbNewLine & "Mapa: " & MAPFl, "ES.CargarMapaFormatoCSM")
            Exit Sub
        End If
    
108     fh = FreeFile
110     Open MAPFl For Binary As fh
    
112     Get #fh, , MH
114     Get #fh, , MapSize
116     Get #fh, , MapDat

        Rem Get #fh, , L1

118     With MH

            'Cargamos Bloqueos
        
120         If .NumeroBloqueados > 0 Then

122             ReDim Blqs(1 To .NumeroBloqueados)
124             Get #fh, , Blqs

126             For i = 1 To .NumeroBloqueados
128                 MapData(Map, Blqs(i).X, Blqs(i).Y).Blocked = Blqs(i).Lados
130             Next i

            End If
        
            'Cargamos Layer 1
        
132         If .NumeroLayers(1) > 0 Then
        
134             ReDim L1(1 To .NumeroLayers(1))
136             Get #fh, , L1

138             For i = 1 To .NumeroLayers(1)
                
140                 X = L1(i).X
142                 Y = L1(i).Y
                        
144                 MapData(Map, X, Y).Graphic(1) = L1(i).GrhIndex
            
                    'InitGrh MapData(L1(i).X, L1(i).Y).Graphic(1), MapData(L1(i).X, L1(i).Y).Graphic(1).GrhIndex
                    ' Call Map_Grh_Set(L2(i).X, L2(i).Y, L2(i).GrhIndex, 2)
146                 If HayAgua(Map, X, Y) Then
148                     MapData(Map, X, Y).Blocked = MapData(Map, X, Y).Blocked Or FLAG_AGUA
                    End If
                
150             Next i

            End If
        
            'Cargamos Layer 2
152         If .NumeroLayers(2) > 0 Then
154             ReDim L2(1 To .NumeroLayers(2))
156             Get #fh, , L2

158             For i = 1 To .NumeroLayers(2)
                
160                 X = L2(i).X
162                 Y = L2(i).Y

164                 MapData(Map, X, Y).Graphic(2) = L2(i).GrhIndex
                
166                 MapData(Map, X, Y).Blocked = MapData(Map, X, Y).Blocked And Not FLAG_AGUA
                
168             Next i

            End If
                
170         If .NumeroLayers(3) > 0 Then
172             ReDim L3(1 To .NumeroLayers(3))
174             Get #fh, , L3

176             For i = 1 To .NumeroLayers(3)
178                 X = L3(i).X
180                 Y = L3(i).Y

182                 MapData(Map, X, Y).Graphic(3) = L3(i).GrhIndex
                
184                 If EsArbol(L3(i).GrhIndex) Then
186                     MapData(Map, X, Y).Blocked = MapData(Map, X, Y).Blocked Or FLAG_ARBOL
                    End If
188             Next i

            End If
        
190         If .NumeroLayers(4) > 0 Then
192             ReDim L4(1 To .NumeroLayers(4))
194             Get #fh, , L4

196             For i = 1 To .NumeroLayers(4)
198                 MapData(Map, L4(i).X, L4(i).Y).Graphic(4) = L4(i).GrhIndex
200             Next i

            End If

202         If .NumeroTriggers > 0 Then
204             ReDim Triggers(1 To .NumeroTriggers)
206             Get #fh, , Triggers

208             For i = 1 To .NumeroTriggers
210                 X = Triggers(i).X
212                 Y = Triggers(i).Y

214                 MapData(Map, X, Y).trigger = Triggers(i).trigger

                    ' Trigger detalles en agua
216                 If Triggers(i).trigger = eTrigger.DETALLEAGUA Then
                        ' Vuelvo a poner flag agua
218                     MapData(Map, X, Y).Blocked = MapData(Map, X, Y).Blocked Or FLAG_AGUA
                    End If
220             Next i

            End If

222         If .NumeroParticulas > 0 Then
224             ReDim Particulas(1 To .NumeroParticulas)
226             Get #fh, , Particulas

228             For i = 1 To .NumeroParticulas
230                 MapData(Map, Particulas(i).X, Particulas(i).Y).ParticulaIndex = Particulas(i).Particula
232                 MapData(Map, Particulas(i).X, Particulas(i).Y).ParticulaIndex = 0
234             Next i

            End If

236         If .NumeroLuces > 0 Then
238             ReDim Luces(1 To .NumeroLuces)
240             Get #fh, , Luces

242             For i = 1 To .NumeroLuces
244                 MapData(Map, Luces(i).X, Luces(i).Y).Luz.Color = Luces(i).Color
246                 MapData(Map, Luces(i).X, Luces(i).Y).Luz.Rango = Luces(i).Rango
248                 MapData(Map, Luces(i).X, Luces(i).Y).Luz.Color = 0
250                 MapData(Map, Luces(i).X, Luces(i).Y).Luz.Rango = 0
252             Next i

            End If
            
254         If .NumeroOBJs > 0 Then
256             ReDim Objetos(1 To .NumeroOBJs)
258             Get #fh, , Objetos

260             For i = 1 To .NumeroOBJs
262                 MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.ObjIndex = Objetos(i).ObjIndex

264                 Select Case ObjData(Objetos(i).ObjIndex).OBJType

                        Case eOBJType.otYacimiento, eOBJType.otArboles
266                         MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.Amount = ObjData(Objetos(i).ObjIndex).VidaUtil
268                         MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.data = &H7FFFFFFF ' Ultimo uso = Max Long

270                     Case Else
272                         MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.Amount = Objetos(i).ObjAmmount

                    End Select

274             Next i

            End If

276         If .NumeroNPCs > 0 Then
278             ReDim NPCs(1 To .NumeroNPCs)
280             Get #fh, , NPCs

                Dim NumNpc As Integer
                 
282             For i = 1 To .NumeroNPCs

284                 NumNpc = NPCs(i).NpcIndex
                    
286                 If NumNpc > 0 Then
288                     npcfile = DatPath & "NPCs.dat"

                        'Si el npc debe hacer respawn en la pos
                        'original la guardamos
290                     If val(GetVar(npcfile, "NPC" & NumNpc, "PosOrig")) = 1 Then
292                         MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex = OpenNPC(NumNpc)
294                         Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Orig.Map = Map
296                         Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Orig.X = NPCs(i).X
298                         Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Orig.Y = NPCs(i).Y
                        Else
300                         MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex = OpenNPC(NumNpc)

                        End If

302                     Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Pos.Map = Map
304                     Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Pos.X = NPCs(i).X
306                     Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Pos.Y = NPCs(i).Y
                        
                        '        If NPCs(i).NpcIndex > 499 Then
                                            
                        '                                           Dim nfile As Integer
                        '  nfile = FreeFile ' obtenemos un canal
                        '  Open App.Path & "\logs\npcs.log" For Append Shared As #nfile
                        ' Print #nfile, NPCs(i).NpcIndex & "(" & NpcList(MapData(Map, NPCs(i).x, NPCs(i).y).NpcIndex).Name & ") "
                        ' Close #nfile
                                            
                        ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "Nombre", NpcList(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Name
                        ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "MaxHp", NpcList(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Stats.MaxHp
                        ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "GiveEXP", NpcList(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).GiveEXP
                        ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "GiveGLD", NpcList(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).GiveGLD
                        ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "MinHIT", NpcList(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Stats.MinHIT
                        ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "MaxHit", NpcList(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Stats.MaxHit
                        ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "def", NpcList(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Stats.def
                        ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "defM", NpcList(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Stats.defM
                        ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "PoderAtaque", NpcList(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).PoderAtaque
                        ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "PoderEvasion", NpcList(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).PoderEvasion
                        ' WriteVar App.Path & "\npcenuso.txt", NPCs(i).NpcIndex, "Posicion" & i, Map & "-" & NPCs(i).X & "-" & NPCs(i).Y
                        'End If
                            
308                     If Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).name = "" Then
                       
310                         MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex = 0
                        Else
                        
312                         Call MakeNPCChar(True, 0, MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex, Map, NPCs(i).X, NPCs(i).Y)
                        
                        End If

                    End If

314             Next i
                
            End If
            
316         If .NumeroTE > 0 Then
318             ReDim TEs(1 To .NumeroTE)
320             Get #fh, , TEs

322             For i = 1 To .NumeroTE
324                 MapData(Map, TEs(i).X, TEs(i).Y).TileExit.Map = TEs(i).DestM
326                 MapData(Map, TEs(i).X, TEs(i).Y).TileExit.X = TEs(i).DestX
328                 MapData(Map, TEs(i).X, TEs(i).Y).TileExit.Y = TEs(i).DestY
330             Next i

            End If
        
        End With

332     Close fh

        ' WyroX: Nuevo sistema de restricciones
334     If Not IsNumeric(MapDat.restrict_mode) Then
            ' Solo se usaba el "NEWBIE"
336         If UCase$(MapDat.restrict_mode) = "NEWBIE" Then
338             MapDat.restrict_mode = "1"
            Else
340             MapDat.restrict_mode = "0"
            End If
        End If
    
342     MapInfo(Map).map_name = MapDat.map_name
344     MapInfo(Map).ambient = MapDat.ambient
346     MapInfo(Map).backup_mode = MapDat.backup_mode
348     MapInfo(Map).base_light = MapDat.base_light
350     MapInfo(Map).Newbie = (val(MapDat.restrict_mode) And 1) <> 0
352     MapInfo(Map).SinMagia = (val(MapDat.restrict_mode) And 2) <> 0
354     MapInfo(Map).NoPKs = (val(MapDat.restrict_mode) And 4) <> 0
356     MapInfo(Map).NoCiudadanos = (val(MapDat.restrict_mode) And 8) <> 0
358     MapInfo(Map).SinInviOcul = (val(MapDat.restrict_mode) And 16) <> 0
359     MapInfo(Map).SoloClanes = (val(MapDat.restrict_mode) And 32) <> 0
360     MapInfo(Map).ResuCiudad = val(GetVar(DatPath & "Map.dat", "RESUCIUDAD", Map)) <> 0
362     MapInfo(Map).letter_grh = MapDat.letter_grh
364     MapInfo(Map).lluvia = MapDat.lluvia
366     MapInfo(Map).music_numberHi = MapDat.music_numberHi
368     MapInfo(Map).music_numberLow = MapDat.music_numberLow
370     MapInfo(Map).niebla = MapDat.niebla
372     MapInfo(Map).Nieve = MapDat.Nieve
373     MapInfo(Map).MinLevel = MapDat.level And &HFF
374     MapInfo(Map).MaxLevel = (MapDat.level And &HFF00) / &H100
    
375     MapInfo(Map).Seguro = MapDat.Seguro

376     MapInfo(Map).terrain = MapDat.terrain
378     MapInfo(Map).zone = MapDat.zone

        If LenB(MapDat.Salida) <> 0 Then
            Dim Fields() As String
            Fields = Split(MapDat.Salida, "-")
            MapInfo(Map).Salida.Map = val(Fields(0))
            MapInfo(Map).Salida.X = val(Fields(1))
            MapInfo(Map).Salida.Y = val(Fields(2))
        End If
 
        Exit Sub

ErrorHandler:
380     Close fh
382     Call RegistrarError(Err.Number, Err.Description, "ES.CargarMapaFormatoCSM", Erl)
    
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
107     Md5Cliente = Lector.GetValue("CHECKSUM", "Cliente")
    
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
        
        ' Configuracion de la API
        API_Enabled = CBool(Lector.GetValue("API_SOCKET", "Enabled"))
        API_HostName = Lector.GetValue("API_SOCKET", "HostName")
        API_Port = val(Lector.GetValue("API_SOCKET", "Port"))
        
        ' Manejo de el Socket reservado a la API al activar/desactivar
        If API_Enabled Then
        
            ' Si la API esta activada, activamos el timer.
            frmMain.t_ColaAPI.Enabled = API_Enabled
            
            Call frmAPISocket.Connect
            
        Else
            
            ' Cerramos el socket ya que no vamos a estar usándolo
            If Not frmAPISocket.Socket Is Nothing Then
                Call frmAPISocket.Socket.CloseSck
            End If
            
        End If
    
178     Call CargarCiudades
182     Call ConsultaPopular.LoadData

184     Set Lector = Nothing

        Exit Sub

LoadSini_Err:
        Set Lector = Nothing
186     Call RegistrarError(Err.Number, Err.Description, "ES.LoadSini", Erl)
188     Resume Next
        
End Sub

Sub CargarCiudades()
        
        On Error GoTo CargarCiudades_Err
    
        

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
    
        
        Exit Sub

CargarCiudades_Err:
302     Call RegistrarError(Err.Number, Err.Description, "ES.CargarCiudades", Erl)

        
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

120     IntervaloPerderStamina = val(Lector.GetValue("INTERVALOS", "IntervaloPerderStamina"))
122     FrmInterv.txtIntervaloPerderStamina.Text = IntervaloPerderStamina
    
124     IntervaloSed = val(Lector.GetValue("INTERVALOS", "IntervaloSed"))
126     FrmInterv.txtIntervaloSed.Text = IntervaloSed
    
128     IntervaloHambre = val(Lector.GetValue("INTERVALOS", "IntervaloHambre"))
130     FrmInterv.txtIntervaloHambre.Text = IntervaloHambre
    
132     IntervaloVeneno = val(Lector.GetValue("INTERVALOS", "IntervaloVeneno"))
134     FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno
    
136     IntervaloParalizado = val(Lector.GetValue("INTERVALOS", "IntervaloParalizado"))
138     FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado
    
140     IntervaloInmovilizado = val(Lector.GetValue("INTERVALOS", "IntervaloInmovilizado"))
142     FrmInterv.txtIntervaloInmovilizado.Text = IntervaloInmovilizado
    
144     IntervaloInvisible = val(Lector.GetValue("INTERVALOS", "IntervaloInvisible"))
146     FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible
    
148     IntervaloFrio = val(Lector.GetValue("INTERVALOS", "IntervaloFrio"))
150     FrmInterv.txtIntervaloFrio.Text = IntervaloFrio
    
152     IntervaloWavFx = val(Lector.GetValue("INTERVALOS", "IntervaloWAVFX"))
154     FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx
    
156     IntervaloInvocacion = val(Lector.GetValue("INTERVALOS", "IntervaloInvocacion"))
158     FrmInterv.txtInvocacion.Text = IntervaloInvocacion
    
160     TimeoutPrimerPaquete = val(Lector.GetValue("INTERVALOS", "TimeoutPrimerPaquete"))
162     FrmInterv.txtTimeoutPrimerPaquete.Text = TimeoutPrimerPaquete
    
164     TimeoutEsperandoLoggear = val(Lector.GetValue("INTERVALOS", "TimeoutEsperandoLoggear"))
166     FrmInterv.txtTimeoutEsperandoLoggear.Text = TimeoutEsperandoLoggear
    
168     IntervaloIncineracion = val(Lector.GetValue("INTERVALOS", "IntervaloFuego"))
170     FrmInterv.txtintervalofuego.Text = IntervaloIncineracion
    
172     IntervaloTirar = val(Lector.GetValue("INTERVALOS", "IntervaloTirar"))
174     FrmInterv.txtintervalotirar.Text = IntervaloTirar

176     IntervaloMeditar = val(Lector.GetValue("INTERVALOS", "IntervaloMeditar"))
178     FrmInterv.txtIntervaloMeditar.Text = IntervaloMeditar
    
180     IntervaloCaminar = val(Lector.GetValue("INTERVALOS", "IntervaloCaminar"))
182     FrmInterv.txtintervalocaminar.Text = IntervaloCaminar
        'Ladder
    
        '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    
184     IntervaloUserPuedeCastear = val(Lector.GetValue("INTERVALOS", "IntervaloLanzaHechizo"))
186     FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear
    
188     frmMain.TIMER_AI.Interval = val(Lector.GetValue("INTERVALOS", "IntervaloNpcAI"))
190     FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval
    
192     frmMain.npcataca.Interval = val(Lector.GetValue("INTERVALOS", "IntervaloNpcPuedeAtacar"))
194     FrmInterv.txtNPCPuedeAtacar.Text = frmMain.npcataca.Interval
    
196     IntervaloTrabajarExtraer = val(Lector.GetValue("INTERVALOS", "IntervaloTrabajarExtraer"))
198     FrmInterv.txtTrabajoExtraer.Text = IntervaloTrabajarExtraer

200     IntervaloTrabajarConstruir = val(Lector.GetValue("INTERVALOS", "IntervaloTrabajarConstruir"))
202     FrmInterv.txtTrabajoConstruir.Text = IntervaloTrabajarConstruir
    
204     IntervaloUserPuedeAtacar = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeAtacar"))
206     FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar
    
        'TODO : Agregar estos intervalos al form!!!
208     IntervaloMagiaGolpe = val(Lector.GetValue("INTERVALOS", "IntervaloMagiaGolpe"))
210     IntervaloGolpeMagia = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeMagia"))
    
        'frmMain.tLluvia.Interval = val(Lector.GetValue("INTERVALOS", "IntervaloPerdidaStaminaLluvia"))
        'FrmInterv.txtIntervaloPerdidaStaminaLluvia.Text = frmMain.tLluvia.Interval
    
212     MinutosWs = val(Lector.GetValue("INTERVALOS", "IntervaloWS"))

214     If MinutosWs < 1 Then MinutosWs = 10
    
216     IntervaloCerrarConexion = val(Lector.GetValue("INTERVALOS", "IntervaloCerrarConexion"))
218     IntervaloUserPuedeUsarU = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeUsarU"))
220     IntervaloUserPuedeUsarClic = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeUsarClic"))
222     IntervaloFlechasCazadores = val(Lector.GetValue("INTERVALOS", "IntervaloFlechasCazadores"))
224     IntervaloGolpeUsar = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeUsar"))
    
226     IntervaloOculto = val(Lector.GetValue("INTERVALOS", "IntervaloOculto"))

228     MargenDeIntervaloPorPing = val(Lector.GetValue("INTERVALOS", "MargenDeIntervaloPorPing"))
    
230     IntervaloPuedeSerAtacado = val(Lector.GetValue("INTERVALOS", "IntervaloPuedeSerAtacado"))

232     IntervaloGuardarUsuarios = val(Lector.GetValue("INTERVALOS", "IntervaloGuardarUsuarios"))
        
234     LimiteGuardarUsuarios = val(Lector.GetValue("INTERVALOS", "LimiteGuardarUsuarios"))

236     IntervaloTimerGuardarUsuarios = val(Lector.GetValue("INTERVALOS", "IntervaloTimerGuardarUsuarios"))

        IntervaloMensajeGlobal = val(Lector.GetValue("INTERVALOS", "IntervaloMensajeGlobal"))
    
        '&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    
238     Set Lector = Nothing

        
        Exit Sub

LoadIntervalos_Err:
240     Call RegistrarError(Err.Number, Err.Description, "ES.LoadIntervalos", Erl)
242     Resume Next
        
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

122     CostoPerdonPorCiudadano = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "CostoPerdonPorCiudadano"))

124     frmMain.lblLimpieza.Caption = "Limpieza de objetos cada: " & TimerLimpiarObjetos & " minutos."

        
        Exit Sub

LoadConfiguraciones_Err:
126     Call RegistrarError(Err.Number, Err.Description, "ES.LoadConfiguraciones", Erl)
128     Resume Next
        
End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
        '*****************************************************************
        'Escribe VAR en un archivo
        '*****************************************************************
        
        On Error GoTo WriteVar_Err
        

100     writeprivateprofilestring Main, Var, Value, File
    
        
        Exit Sub

WriteVar_Err:
102     Call RegistrarError(Err.Number, Err.Description, "ES.WriteVar", Erl)
104     Resume Next
        
End Sub

Sub LoadUser(ByVal UserIndex As Integer)

        On Error GoTo ErrorHandler
    
100     If Database_Enabled Then
102         Call LoadUserDatabase(UserIndex)
        Else
104         Call LoadUserBinary(UserIndex)
        End If
    
106     With UserList(UserIndex)

108         If .flags.Paralizado = 1 Then
110             .Counters.Paralisis = IntervaloParalizado
            End If

112         If .flags.Muerto = 0 Then
114             .Char = .OrigChar
            
116             If .Char.Body = 0 Then
118                 Call DarCuerpoDesnudo(UserIndex)
                End If
            
120             If .Char.Head = 0 Then
122                 .Char.Head = 1
                End If
            Else
124             .Char.Body = iCuerpoMuerto
126             .Char.Head = 0
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

            ' DM
180         If .Invent.DañoMagicoEqpSlot > 0 Then
182             .Invent.DañoMagicoEqpObjIndex = .Invent.Object(.Invent.DañoMagicoEqpSlot).ObjIndex
            
184             If .flags.Muerto = 0 Then
186                 .Char.DM_Aura = ObjData(.Invent.DañoMagicoEqpObjIndex).CreaGRH
                End If
            End If
            
            ' RM
188         If .Invent.ResistenciaEqpSlot > 0 Then
190            .Invent.ResistenciaEqpObjIndex = .Invent.Object(.Invent.ResistenciaEqpSlot).ObjIndex
             
192             If .flags.Muerto = 0 Then
194                 .Char.RM_Aura = ObjData(.Invent.ResistenciaEqpObjIndex).CreaGRH
                End If
            End If

196         If .Invent.MonturaSlot > 0 Then
198             .Invent.MonturaObjIndex = .Invent.Object(.Invent.MonturaSlot).ObjIndex
            End If
        
200         If .Invent.HerramientaEqpSlot > 0 Then
202             .Invent.HerramientaEqpObjIndex = .Invent.Object(.Invent.HerramientaEqpSlot).ObjIndex
            End If
        
204         If .Invent.NudilloSlot > 0 Then
206             .Invent.NudilloObjIndex = .Invent.Object(.Invent.NudilloSlot).ObjIndex
            
208             If .flags.Muerto = 0 Then
210                 .Char.Arma_Aura = ObjData(.Invent.NudilloObjIndex).CreaGRH
                End If
            End If
        
212         If .Invent.MagicoSlot > 0 Then
214             .Invent.MagicoObjIndex = .Invent.Object(.Invent.MagicoSlot).ObjIndex

216             If .flags.Muerto = 0 Then
218                 .Char.Otra_Aura = ObjData(.Invent.MagicoObjIndex).CreaGRH
                End If
            End If

        End With

        Exit Sub

ErrorHandler:
220     Call RegistrarError(Err.Number, Err.Description & " UserName: " & UserList(UserIndex).name, "ES.LoadUser", Erl)
222     Resume Next
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer, Optional ByVal Logout As Boolean = False)

    On Error GoTo SaveUser_Err
    
    #If DEBUGGING = 1 Then
        Call GetElapsedTime
    #End If
    
    If Not API_Enabled Then
    
        If Database_Enabled Then
            Call SaveUserDatabase(UserIndex, Logout)
        Else
            Call SaveUserCharfile(UserIndex, Logout)

        End If
        
    Else
    
        Call SaveUserAPI(UserIndex, Logout)
        
    End If
    
    UserList(UserIndex).Counters.LastSave = GetTickCount
    
    #If DEBUGGING = 1 Then
        Call LogPerformance("Guardado de Personaje " & IIf(API_Enabled, "(API)", "(ADO)") & " - Tiempo transcurrido: " & Round(GetElapsedTime(), 1) & " ms")
    #End If
    
    Exit Sub

SaveUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "ES.SaveUser", Erl)

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
110     Call RegistrarError(Err.Number, Err.Description, "ES.LoadUserBinary", Erl)
112     Resume Next
        
End Sub

Sub SaveUserCharfile(ByVal UserIndex As Integer, Optional ByVal Logout As Boolean)
        '*************************************************
        'Author: Unknown
        'Last modified: 23/01/2007
        'Saves the Users records
        '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
        '*************************************************
    
        On Error GoTo ErrHandler
    
        Dim UserFile    As String
        Dim OldUserHead As Long
    
100     With UserList(UserIndex)
    
102         UserFile = CharPath & UCase$(.name) & ".chr"
    
            'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
            'clase=0 es el error, porq el enum empieza de 1!!
104         If .clase = 0 Or .Stats.ELV = 0 Then
106             Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & .name)
                Exit Sub
            End If
    
108         Debug.Print UserFile
    
110         If FileExist(UserFile, vbNormal) Then
        
112             If .flags.Muerto = 1 Then
114                 OldUserHead = .Char.Head
116                 .Char.Head = GetVar(UserFile, "INIT", "Head")

                End If

            End If
    
            Dim LoopC As Integer
    
118         If FileExist(UserFile, vbNormal) Then Kill UserFile

120         Dim File As String: File = UserFile
122         Dim n As Integer: n = FreeFile
        
124         Open File For Output Access Write As n
        
            'INIT
126         Print #n, , "[INIT]" & vbCrLf
128         Print #n, , "Cuenta=" & .Cuenta & vbCrLf
130         Print #n, , "Genero=" & .genero & vbCrLf
132         Print #n, , "Raza=" & .raza & vbCrLf
134         Print #n, , "Hogar=" & .Hogar & vbCrLf
136         Print #n, , "Clase=" & .clase & vbCrLf
138         Print #n, , "Desc=" & .Desc & vbCrLf
140         Print #n, , "Heading=" & CStr(.Char.Heading) & vbCrLf

142         If .Char.Head = 0 Then
144             Print #n, , "Head=" & CStr(.OrigChar.Head) & vbCrLf
            Else
146             Print #n, , "Head=" & CStr(.Char.Head) & vbCrLf
            End If

148         Print #n, , "Arma=" & CStr(.Char.WeaponAnim) & vbCrLf
150         Print #n, , "Escudo=" & CStr(.Char.ShieldAnim) & vbCrLf
152         Print #n, , "Casco=" & CStr(.Char.CascoAnim) & vbCrLf
154         Print #n, , "Position=" & .Pos.Map & "-" & .Pos.X & "-" & .Pos.Y & vbCrLf
       
           'If .flags.Muerto = 0 Then
156         Print #n, , "Body=" & CStr(.Char.Body) & vbCrLf
            'End If
        
            #If ConUpTime Then
158             Dim TempDate As Date: TempDate = Now - .LogOnTime
            
160             .LogOnTime = Now
162             .UpTime = .UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
            
164             Print #n, , "UpTime=" & .UpTime & vbCrLf
            #End If

166         If Logout Then
168             Print #n, , "Logged=0" & vbCrLf
            Else
170             Print #n, , "Logged=1" & vbCrLf
            End If

172         Print #n, , "MENSAJEINFORMACION=" & .MENSAJEINFORMACION & vbCrLf

174         Print #n, , vbCrLf
        
            'baneo
176         Print #n, , "[BAN]" & vbCrLf
178         Print #n, , "Baneado=" & CStr(.flags.Ban) & vbCrLf
180         Print #n, , "BanMotivo=" & CStr(.flags.BanMotivo) & vbCrLf
        
182         Print #n, , vbCrLf
        
            'STATS
184         With .Stats
186             Print #n, , "[STATS]" & vbCrLf
188             Print #n, , "GLD=" & CStr(.GLD) & vbCrLf
190             Print #n, , "BANCO=" & CStr(.Banco) & vbCrLf
192             Print #n, , "MaxHP=" & CStr(.MaxHp) & vbCrLf
194             Print #n, , "MinHP=" & CStr(.MinHp) & vbCrLf
196             Print #n, , "MaxSTA=" & CStr(.MaxSta) & vbCrLf
198             Print #n, , "MinSTA=" & CStr(.MinSta) & vbCrLf
200             Print #n, , "MaxMAN=" & CStr(.MaxMAN) & vbCrLf
202             Print #n, , "MinMAN=" & CStr(.MinMAN) & vbCrLf
204             Print #n, , "MaxHIT=" & CStr(.MaxHit) & vbCrLf
206             Print #n, , "MinHIT=" & CStr(.MinHIT) & vbCrLf
208             Print #n, , "MaxAGU=" & CStr(.MaxAGU) & vbCrLf
210             Print #n, , "MinAGU=" & CStr(.MinAGU) & vbCrLf
212             Print #n, , "MaxHAM=" & CStr(.MaxHam) & vbCrLf
214             Print #n, , "MinHAM=" & CStr(.MinHam) & vbCrLf
216             Print #n, , "SkillPtsLibres=" & CStr(.SkillPts) & vbCrLf
218             Print #n, , "EXP=" & CStr(.Exp) & vbCrLf
220             Print #n, , "ELV=" & CStr(.ELV) & vbCrLf
222             Print #n, , "ELU=" & CStr(.ELU) & vbCrLf
224             Print #n, , "InventLevel=" & CByte(.InventLevel) & vbCrLf
            End With
        
226         Print #n, , vbCrLf
        
            'FLAGS
228         With .flags
230             Print #n, , "[FLAGS]" & vbCrLf
232             Print #n, , "CASADO=" & CStr(.Casado) & vbCrLf
234             Print #n, , "PAREJA=" & CStr(.Pareja) & vbCrLf
236             Print #n, , "Muerto=" & CStr(.Muerto) & vbCrLf
238             Print #n, , "Escondido=" & CStr(.Escondido) & vbCrLf
240             Print #n, , "Hambre=" & CStr(.Hambre) & vbCrLf
242             Print #n, , "Sed=" & CStr(.Sed) & vbCrLf
244             Print #n, , "Desnudo=" & CStr(.Desnudo) & vbCrLf
246             Print #n, , "Navegando=" & CStr(.Navegando) & vbCrLf
248             Print #n, , "Envenenado=" & CStr(.Envenenado) & vbCrLf
250             Print #n, , "Paralizado=" & CStr(.Paralizado) & vbCrLf
252             Print #n, , "Inmovilizado=" & CStr(.Inmovilizado) & vbCrLf
254             Print #n, , "Incinerado=" & CStr(.Incinerado) & vbCrLf
256             Print #n, , "VecesQueMoriste=" & CStr(.VecesQueMoriste) & vbCrLf
258             Print #n, , "ScrollExp=" & CStr(.ScrollExp) & vbCrLf
260             Print #n, , "ScrollOro=" & CStr(.ScrollOro) & vbCrLf
262             Print #n, , "MinutosRestantes=" & CStr(.MinutosRestantes) & vbCrLf
264             Print #n, , "SegundosPasados=" & CStr(.SegundosPasados) & vbCrLf
266             Print #n, , "Silenciado=" & CStr(.Silenciado) & vbCrLf
268             Print #n, , "Montado=" & CStr(.Montado) & vbCrLf
            End With

270         Print #n, , vbCrLf
        
            'GRABADO DE CLAN
272         Print #n, , "[GUILD]" & vbCrLf
274         Print #n, , "GUILDINDEX=" & CInt(.GuildIndex) & vbCrLf
        
276         Print #n, , vbCrLf
        
278         Print #n, , "[CONSEJO]" & vbCrLf
280         Print #n, , "PERTENECE=" & IIf(.flags.Privilegios And PlayerType.RoyalCouncil, "1", "0") & vbCrLf
282         Print #n, , "PERTENECECAOS=" & IIf(.flags.Privilegios And PlayerType.ChaosCouncil, "1", "0") & vbCrLf
        
284         Print #n, , vbCrLf
        
286         With .Faccion
288             Print #n, , "[FACCIONES]" & vbCrLf
290             Print #n, , "EjercitoReal=" & CStr(.ArmadaReal) & vbCrLf
292             Print #n, , "Status=" & CStr(.Status) & vbCrLf
294             Print #n, , "EjercitoCaos=" & CStr(.FuerzasCaos) & vbCrLf
296             Print #n, , "CiudMatados=" & CStr(.CiudadanosMatados) & vbCrLf
298             Print #n, , "CrimMatados=" & CStr(.CriminalesMatados) & vbCrLf
300             Print #n, , "rArCaos=" & CStr(.RecibioArmaduraCaos) & vbCrLf
302             Print #n, , "rArReal=" & CStr(.RecibioArmaduraReal) & vbCrLf
304             Print #n, , "rExCaos=" & CStr(.RecibioExpInicialCaos) & vbCrLf
306             Print #n, , "rExReal=" & CStr(.RecibioExpInicialReal) & vbCrLf
308             Print #n, , "recCaos=" & CStr(.RecompensasCaos) & vbCrLf
310             Print #n, , "recReal=" & CStr(.RecompensasReal) & vbCrLf
312             Print #n, , "Reenlistadas=" & CStr(.Reenlistadas) & vbCrLf
314             Print #n, , "NivelIngreso=" & CStr(.NivelIngreso) & vbCrLf
316             Print #n, , "FechaIngreso=" & CStr(.FechaIngreso) & vbCrLf
318             Print #n, , "MatadosIngreso=" & CStr(.MatadosIngreso) & vbCrLf
320             Print #n, , "NextRecompensa=" & CStr(.NextRecompensa) & vbCrLf
            End With
        
322         Print #n, , vbCrLf
        
            'MAHIA ESTUPIDIZA
324         Print #n, , "[MAGIA]" & vbCrLf
326         Print #n, , "ENVENENA=" & CByte(.flags.Envenena) & vbCrLf
328         Print #n, , "PARALIZA=" & CByte(.flags.Paraliza) & vbCrLf
330         Print #n, , "AnilloOcultismo=" & CByte(.flags.AnilloOcultismo) & vbCrLf
332         Print #n, , "incinera=" & CByte(.flags.incinera) & vbCrLf
334         Print #n, , "Estupidiza=" & CByte(.flags.Estupidiza) & vbCrLf
336         Print #n, , "Pendiente=" & CByte(.flags.PendienteDelSacrificio) & vbCrLf
338         Print #n, , "CarroMineria=" & CByte(.flags.CarroMineria) & vbCrLf
340         Print #n, , "NoPalabrasMagicas=" & CByte(.flags.NoPalabrasMagicas) & vbCrLf
342         Print #n, , "NoDetectable=" & CByte(.flags.NoDetectable) & vbCrLf
344         Print #n, , "Otra_Aura=" & CStr(.Char.Otra_Aura) & vbCrLf
            'Print #n, , "DañoMagico=" & CByte(.flags.DañoMagico) & vbCrLf
            'Print #n, , "ResistenciaMagica=" & CByte(.flags.ResistenciaMagica) & vbCrLf
346         Print #n, , "RegeneracionMana=" & CByte(.flags.RegeneracionMana) & vbCrLf
348         Print #n, , "NoMagiaEfeceto=" & CByte(.flags.NoMagiaEfeceto) & vbCrLf
350         Print #n, , "RegeneracionHP=" & CByte(.flags.RegeneracionHP) & vbCrLf
352         Print #n, , "RegeneracionSta=" & CByte(.flags.RegeneracionSta) & vbCrLf

354         Print #n, , vbCrLf
        
            'SKILLS
356         Print #n, , "[SKILLS]" & vbCrLf

358         For LoopC = 1 To UBound(.Stats.UserSkills)
360             Print #n, , "SK" & LoopC & "=" & CStr(.Stats.UserSkills(LoopC)) & vbCrLf
            Next

362         Print #n, , vbCrLf

            'INVENTARIO
364         With .Invent
366             Print #n, , "[Inventory]" & vbCrLf
368             Print #n, , "CantidadItems=" & val(.NroItems) & vbCrLf
    
370             For LoopC = 1 To MAX_INVENTORY_SLOTS
372                 Print #n, , "Obj" & LoopC & "=" & .Object(LoopC).ObjIndex & "-" & .Object(LoopC).Amount & "-" & .Object(LoopC).Equipped & vbCrLf
                Next
            
374             Print #n, , "WeaponEqpSlot=" & CStr(.WeaponEqpSlot) & vbCrLf
376             Print #n, , "HerramientaEqpSlot=" & CStr(.HerramientaEqpSlot) & vbCrLf
378             Print #n, , "ArmourEqpSlot=" & CStr(.ArmourEqpSlot) & vbCrLf
380             Print #n, , "CascoEqpSlot=" & CStr(.CascoEqpSlot) & vbCrLf
382             Print #n, , "EscudoEqpSlot=" & CStr(.EscudoEqpSlot) & vbCrLf
384             Print #n, , "BarcoSlot=" & CStr(.BarcoSlot) & vbCrLf
386             Print #n, , "MonturaSlot=" & CStr(.MonturaSlot) & vbCrLf
388             Print #n, , "MunicionSlot=" & CStr(.MunicionEqpSlot) & vbCrLf
390             Print #n, , "DMSlot=" & CStr(.DañoMagicoEqpSlot) & vbCrLf
392             Print #n, , "RMSlot=" & CStr(.ResistenciaEqpSlot) & vbCrLf
394             Print #n, , "MagicoSlot=" & CStr(.MagicoSlot) & vbCrLf
396             Print #n, , "NudilloEqpSlot=" & CStr(.NudilloSlot) & vbCrLf
            End With
        
398         Print #n, , vbCrLf

400         Print #n, , "[ATRIBUTOS]" & vbCrLf

            '¿Fueron modificados los atributos del usuario?
402         If Not .flags.TomoPocion Then

404             For LoopC = 1 To UBound(.Stats.UserAtributos)
406                 Print #n, , "AT" & LoopC & "=" & CStr(.Stats.UserAtributos(LoopC)) & vbCrLf
                Next

            Else

408             For LoopC = 1 To UBound(.Stats.UserAtributos)
                    '.Stats.UserAtributos(LoopC) = .Stats.UserAtributosBackUP(LoopC)
410                 Print #n, , "AT" & LoopC & "=" & CStr(.Stats.UserAtributosBackUP(LoopC)) & vbCrLf
                Next

            End If

412         Print #n, , vbCrLf
            
            'COUNTERS
414         Print #n, , "[COUNTERS]" & vbCrLf
416         Print #n, , "Pena=" & CStr(.Counters.Pena) & vbCrLf
418         Print #n, , "ScrollOro=" & CStr(.Counters.ScrollOro) & vbCrLf
420         Print #n, , "ScrollExperiencia=" & CStr(.Counters.ScrollExperiencia) & vbCrLf
422         Print #n, , "Oxigeno=" & CStr(.Counters.Oxigeno) & vbCrLf
        
424         Print #n, , vbCrLf

426         Print #n, , "[MUERTES]" & vbCrLf
428         Print #n, , "UserMuertes=" & CStr(.Stats.UsuariosMatados) & vbCrLf
430         Print #n, , "NpcsMuertes=" & CStr(.Stats.NPCsMuertos) & vbCrLf
        
432         Print #n, , vbCrLf
        
            'BANCO
434         Print #n, , "[BancoInventory]" & vbCrLf
436         Print #n, , "CantidadItems=" & val(.BancoInvent.NroItems) & vbCrLf

            Dim loopd As Long
438         For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
440             Print #n, , "Obj" & loopd & "=" & .BancoInvent.Object(loopd).ObjIndex & "-" & .BancoInvent.Object(loopd).Amount & vbCrLf
442         Next loopd
        
444         Print #n, , vbCrLf
        
446         Print #n, , "[LOGROS]" & vbCrLf
448         Print #n, , "UserLogros=" & CByte(.UserLogros) & vbCrLf
450         Print #n, , "NPcLogros=" & CByte(.NPcLogros) & vbCrLf
452         Print #n, , "LevelLogros=" & CByte(.LevelLogros) & vbCrLf
        
454         Print #n, , vbCrLf
        
456         Print #n, , "[BINDKEYS]" & vbCrLf
458         Print #n, , "ChatCombate=" & CByte(.ChatCombate) & vbCrLf
460         Print #n, , "ChatGlobal=" & CByte(.ChatGlobal) & vbCrLf
        
462         Print #n, , vbCrLf

            'HECHIZOS
464         Print #n, , "[HECHIZOS]" & vbCrLf
        
            Dim cad As String
466         For LoopC = 1 To MAXUSERHECHIZOS
468             cad = .Stats.UserHechizos(LoopC)
470             Print #n, , "H" & LoopC & "=" & cad & vbCrLf
            Next
        
472         Print #n, , vbCrLf
        
            'BATTLE
474         Print #n, , "[Battle]" & vbCrLf
476         Print #n, , "Puntos=" & CStr(.flags.BattlePuntos) & vbCrLf
        
478         Print #n, , vbCrLf
        
480         Print #n, , "[CORREO]" & vbCrLf & "NoLeidos=" & CByte(.Correo.NoLeidos) & vbCrLf
482         Print #n, , "CANTCORREO=" & CByte(.Correo.CantCorreo) & vbCrLf
        
484         Print #n, , vbCrLf
        
            'Correo Ladder
486         With .Correo
        
488             For LoopC = 1 To .CantCorreo
        
490                 Print #n, , "REMITENTE" & LoopC & "=" & .Mensaje(LoopC).Remitente & vbCrLf
492                 Print #n, , "MENSAJE" & LoopC & "=" & .Mensaje(LoopC).Mensaje & vbCrLf
494                 Print #n, , "Item" & LoopC & "=" & .Mensaje(LoopC).Item & vbCrLf
496                 Print #n, , "ItemCount" & LoopC & "=" & .Mensaje(LoopC).ItemCount & vbCrLf
498                 Print #n, , "DATE" & LoopC & "=" & .Mensaje(LoopC).Fecha & vbCrLf
500                 Print #n, , "LEIDO" & LoopC & "=" & .Mensaje(LoopC).Leido & vbCrLf
                
502             Next LoopC
        
            End With
        
504         Close #n
        
506         Call SaveQuestStats(UserIndex, UserFile)

            'Devuelve el head de muerto
508         If .flags.Muerto = 1 Then
510             .Char.Head = 0
            End If
    
        End With
        
        Exit Sub

ErrHandler:
512     Call LogError("Error en SaveUserCharfile")
514     Close #n

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
106     Call RegistrarError(Err.Number, Err.Description, "ES.SaveNewUser", Erl)
108     Resume Next
        
End Sub

Sub SaveNewUserCharfile(ByVal UserIndex As Integer)
        '*************************************************
        'Author: Unknown
        'Last modified: 23/01/2007
        'Saves the Users records
        '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
        '*************************************************
    
        On Error GoTo ErrHandler
    
        Dim UserFile    As String

        Dim OldUserHead As Long
    
100     UserFile = CharPath & UCase$(UserList(UserIndex).name) & ".chr"
    
        'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
        'clase=0 es el error, porq el enum empieza de 1!!
102     If UserList(UserIndex).clase = 0 Or UserList(UserIndex).Stats.ELV = 0 Then
104         Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & UserList(UserIndex).name)
            Exit Sub

        End If
    
106     If FileExist(UserFile, vbNormal) Then
108         If UserList(UserIndex).flags.Muerto = 1 Then
110             OldUserHead = UserList(UserIndex).Char.Head
112             UserList(UserIndex).Char.Head = GetVar(UserFile, "INIT", "Head")

            End If

            '       Kill UserFile
        End If
    
        Dim LoopC As Integer

        Dim n

        Dim Datos$

114     n = FreeFile
116     Open UserFile For Binary Access Write As n
    
        'BATTLE
118     Put n, , "[Battle]" & vbCrLf & "Puntos=" & CStr(UserList(UserIndex).flags.BattlePuntos) & vbCrLf
    
120     Put n, , vbCrLf
    
        'FLAGS
122     Put n, , "[FLAGS]" & vbCrLf & "CASADO=" & CStr(UserList(UserIndex).flags.Casado) & vbCrLf
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
148     Put n, , "ScrollExp=" & CStr(UserList(UserIndex).flags.ScrollExp) & vbCrLf
150     Put n, , "ScrollOro=" & CStr(UserList(UserIndex).flags.ScrollOro) & vbCrLf
152     Put n, , "MinutosRestantes=0" & vbCrLf
154     Put n, , "SegundosPasados=0" & vbCrLf
156     Put n, , "Silenciado=0" & vbCrLf
158     Put n, , "Montado=0" & vbCrLf
    
160     Put n, , "InventLevel=0" & vbCrLf
    
162     Put n, , vbCrLf
    
164     Put n, , "[CONSEJO]" & vbCrLf
166     Put n, , "PERTENECE=0" & vbCrLf
168     Put n, , "PERTENECECAOS=0" & vbCrLf
    
170     Put n, , "[FACCIONES]" & vbCrLf & "EjercitoReal=" & CStr(UserList(UserIndex).Faccion.ArmadaReal) & vbCrLf
172     Put n, , "Status=" & CStr(UserList(UserIndex).Faccion.Status) & vbCrLf
174     Put n, , "EjercitoCaos=" & CStr(UserList(UserIndex).Faccion.FuerzasCaos) & vbCrLf
176     Put n, , "CiudMatados=" & CStr(UserList(UserIndex).Faccion.CiudadanosMatados) & vbCrLf
178     Put n, , "CrimMatados=" & CStr(UserList(UserIndex).Faccion.CriminalesMatados) & vbCrLf
180     Put n, , "rArCaos=" & CStr(UserList(UserIndex).Faccion.RecibioArmaduraCaos) & vbCrLf
182     Put n, , "rArReal=" & CStr(UserList(UserIndex).Faccion.RecibioArmaduraReal) & vbCrLf
184     Put n, , "rExCaos=" & CStr(UserList(UserIndex).Faccion.RecibioExpInicialCaos) & vbCrLf
186     Put n, , "rExReal=" & CStr(UserList(UserIndex).Faccion.RecibioExpInicialReal) & vbCrLf
188     Put n, , "recCaos=" & CStr(UserList(UserIndex).Faccion.RecompensasCaos) & vbCrLf
190     Put n, , "recReal=" & CStr(UserList(UserIndex).Faccion.RecompensasReal) & vbCrLf
192     Put n, , "Reenlistadas=" & CStr(UserList(UserIndex).Faccion.Reenlistadas) & vbCrLf
194     Put n, , "NivelIngreso=" & CStr(UserList(UserIndex).Faccion.NivelIngreso) & vbCrLf
196     Put n, , "FechaIngreso=" & CStr(UserList(UserIndex).Faccion.FechaIngreso) & vbCrLf
198     Put n, , "MatadosIngreso=" & CStr(UserList(UserIndex).Faccion.MatadosIngreso) & vbCrLf
200     Put n, , "NextRecompensa=" & CStr(UserList(UserIndex).Faccion.NextRecompensa) & vbCrLf
    
202     Put n, , vbCrLf
    
        'STATS
204     Put n, , "[STATS]" & vbCrLf & "GLD=0" & vbCrLf
206     Put n, , "BANCO=0" & vbCrLf
208     Put n, , "MaxHP=" & CStr(UserList(UserIndex).Stats.MaxHp) & vbCrLf
210     Put n, , "MinHP=" & CStr(UserList(UserIndex).Stats.MinHp) & vbCrLf
212     Put n, , "MaxSTA=" & CStr(UserList(UserIndex).Stats.MaxSta) & vbCrLf
214     Put n, , "MinSTA=" & CStr(UserList(UserIndex).Stats.MinSta) & vbCrLf
216     Put n, , "MaxMAN=" & CStr(UserList(UserIndex).Stats.MaxMAN) & vbCrLf
218     Put n, , "MinMAN=" & CStr(UserList(UserIndex).Stats.MinMAN) & vbCrLf
220     Put n, , "MaxHIT=" & CStr(UserList(UserIndex).Stats.MaxHit) & vbCrLf
222     Put n, , "MinHIT=" & CStr(UserList(UserIndex).Stats.MinHIT) & vbCrLf
224     Put n, , "MaxAGU=" & CStr(UserList(UserIndex).Stats.MaxAGU) & vbCrLf
226     Put n, , "MinAGU=" & CStr(UserList(UserIndex).Stats.MinAGU) & vbCrLf
228     Put n, , "MaxHAM=" & CStr(UserList(UserIndex).Stats.MaxHam) & vbCrLf
230     Put n, , "MinHAM=" & CStr(UserList(UserIndex).Stats.MinHam) & vbCrLf
232     Put n, , "SkillPtsLibres=" & CStr(UserList(UserIndex).Stats.SkillPts) & vbCrLf
234     Put n, , "EXP=" & CStr(UserList(UserIndex).Stats.Exp) & vbCrLf
236     Put n, , "ELV=" & CStr(UserList(UserIndex).Stats.ELV) & vbCrLf
238     Put n, , "ELU=" & CStr(UserList(UserIndex).Stats.ELU) & vbCrLf
    
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

278     For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserSkills)
280         Put n, , "SK" & LoopC & "=0" & vbCrLf
        Next
    
282     Put n, , vbCrLf
    
        'INVENTARIO
284     Put n, , "[Inventory]" & vbCrLf & "CantidadItems=" & val(UserList(UserIndex).Invent.NroItems) & vbCrLf

286     For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots
288         Put n, , "Obj" & LoopC & "=" & UserList(UserIndex).Invent.Object(LoopC).ObjIndex & "-" & UserList(UserIndex).Invent.Object(LoopC).Amount & "-" & UserList(UserIndex).Invent.Object(LoopC).Equipped & vbCrLf
        Next
290     Put n, , "WeaponEqpSlot=" & CStr(UserList(UserIndex).Invent.WeaponEqpSlot) & vbCrLf
292     Put n, , "HerramientaEqpSlot=" & CStr(UserList(UserIndex).Invent.HerramientaEqpSlot) & vbCrLf
294     Put n, , "ArmourEqpSlot=" & CStr(UserList(UserIndex).Invent.ArmourEqpSlot) & vbCrLf
296     Put n, , "CascoEqpSlot=" & CStr(UserList(UserIndex).Invent.CascoEqpSlot) & vbCrLf
298     Put n, , "EscudoEqpSlot=" & CStr(UserList(UserIndex).Invent.EscudoEqpSlot) & vbCrLf
300     Put n, , "BarcoSlot=" & CStr(UserList(UserIndex).Invent.BarcoSlot) & vbCrLf
302     Put n, , "MonturaSlot=" & CStr(UserList(UserIndex).Invent.MonturaSlot) & vbCrLf
304     Put n, , "MunicionSlot=" & CStr(UserList(UserIndex).Invent.MunicionEqpSlot) & vbCrLf
306     Put n, , "DMSlot=" & CStr(UserList(UserIndex).Invent.DañoMagicoEqpSlot) & vbCrLf
308     Put n, , "RMSlot=" & CStr(UserList(UserIndex).Invent.ResistenciaEqpSlot) & vbCrLf
310     Put n, , "MagicoSlot=" & CStr(UserList(UserIndex).Invent.MagicoSlot) & vbCrLf
312     Put n, , "NudilloEqpSlot=" & CStr(UserList(UserIndex).Invent.NudilloSlot) & vbCrLf
    
314     Put n, , vbCrLf
    
        'INIT
316     Put n, , "[INIT]" & vbCrLf & "Cuenta=" & UserList(UserIndex).Cuenta & vbCrLf
318     Put n, , "Genero=" & UserList(UserIndex).genero & vbCrLf
320     Put n, , "Raza=" & UserList(UserIndex).raza & vbCrLf
322     Put n, , "Hogar=" & UserList(UserIndex).Hogar & vbCrLf
324     Put n, , "Clase=" & UserList(UserIndex).clase & vbCrLf
326     Put n, , "Desc=" & UserList(UserIndex).Desc & vbCrLf
328     Put n, , "Heading=" & CStr(UserList(UserIndex).Char.Heading) & vbCrLf
330     Put n, , "Head=" & CStr(UserList(UserIndex).Char.Head) & vbCrLf
332     Put n, , "Arma=" & CStr(UserList(UserIndex).Char.WeaponAnim) & vbCrLf
334     Put n, , "Escudo=" & CStr(UserList(UserIndex).Char.ShieldAnim) & vbCrLf
336     Put n, , "Casco=" & CStr(UserList(UserIndex).Char.CascoAnim) & vbCrLf
338     Put n, , "Position=" & UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.Y & vbCrLf
        ' If UserList(UserIndex).flags.Muerto = 0 Then
340     Put n, , "Body=" & CStr(UserList(UserIndex).Char.Body) & vbCrLf
        'Else
        '   Put N, , "Body=" & iCuerpoMuerto & vbCrLf 'poner body muerto
        '  End If
        #If ConUpTime Then

            Dim TempDate As Date

342         TempDate = Now - UserList(UserIndex).LogOnTime
344         UserList(UserIndex).LogOnTime = Now
346         UserList(UserIndex).UpTime = UserList(UserIndex).UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
348         UserList(UserIndex).UpTime = UserList(UserIndex).UpTime
350         Put n, , "UpTime=" & UserList(UserIndex).UpTime & vbCrLf
        #End If
    
352     Put n, , vbCrLf
    
354     Put n, , "[ATRIBUTOS]" & vbCrLf

        '¿Fueron modificados los atributos del usuario?
356     For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
358         Put n, , "AT" & LoopC & "=" & CStr(UserList(UserIndex).Stats.UserAtributos(LoopC)) & vbCrLf
        Next
360     Put n, , vbCrLf
    
        'baneo
362     Put n, , "[BAN]" & vbCrLf & "Baneado=" & CStr(UserList(UserIndex).flags.Ban) & vbCrLf
364     Put n, , "BanMotivo=" & CStr(UserList(UserIndex).flags.BanMotivo) & vbCrLf
    
366     Put n, , vbCrLf
    
        'COUNTERS
368     Put n, , "[COUNTERS]" & vbCrLf & "Pena=" & CStr(UserList(UserIndex).Counters.Pena) & vbCrLf
370     Put n, , "ScrollOro=" & CStr(UserList(UserIndex).Counters.ScrollOro) & vbCrLf
372     Put n, , "ScrollExperiencia=" & CStr(UserList(UserIndex).Counters.ScrollExperiencia) & vbCrLf
374     Put n, , "Oxigeno=" & CStr(UserList(UserIndex).Counters.Oxigeno) & vbCrLf
    
376     Put n, , vbCrLf
    
378     Put n, , "[MUERTES]" & vbCrLf & "UserMuertes=0" & vbCrLf
380     Put n, , "NpcsMuertes=0" & vbCrLf
    
382     Put n, , vbCrLf
    
        'BANCO
384     Put n, , "[BancoInventory]" & vbCrLf & "CantidadItems=0" & vbCrLf

        Dim loopd As Integer

386     For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
388         Put n, , "Obj" & loopd & "=" & UserList(UserIndex).BancoInvent.Object(loopd).ObjIndex & "-" & UserList(UserIndex).BancoInvent.Object(loopd).Amount & vbCrLf
390     Next loopd
    
392     Put n, , vbCrLf
    
394     Put n, , "[LOGROS]" & vbCrLf & "UserLogros=" & CByte(UserList(UserIndex).UserLogros) & vbCrLf
396     Put n, , "NPcLogros=" & CByte(UserList(UserIndex).NPcLogros) & vbCrLf
398     Put n, , "LevelLogros=" & CByte(UserList(UserIndex).LevelLogros) & vbCrLf
    
400     Put n, , vbCrLf
    
402     Put n, , "[BINDKEYS]" & vbCrLf
404     Put n, , "ChatCombate=" & CByte(UserList(UserIndex).ChatCombate) & vbCrLf
406     Put n, , "ChatGlobal=" & CByte(UserList(UserIndex).ChatGlobal) & vbCrLf
    
408     Put n, , vbCrLf
    
        'HECHIZOS
410     Put n, , "[HECHIZOS]" & vbCrLf

        Dim cad As String

412     For LoopC = 1 To MAXUSERHECHIZOS
414         cad = UserList(UserIndex).Stats.UserHechizos(LoopC)
416         Put n, , "H" & LoopC & "=" & cad & vbCrLf
        Next
      
418     Put n, , vbCrLf
    
420     Put n, , "[CORREO]" & vbCrLf & "NoLeidos=0" & vbCrLf
422     Put n, , "CANTCORREO=0" & vbCrLf
    
        'Correo Ladder
    
424     For LoopC = 1 To UserList(UserIndex).Correo.CantCorreo
    
426         Put n, , "REMITENTE" & LoopC & "=" & UserList(UserIndex).Correo.Mensaje(LoopC).Remitente & vbCrLf
428         Put n, , "MENSAJE" & LoopC & "=" & UserList(UserIndex).Correo.Mensaje(LoopC).Mensaje & vbCrLf
430         Put n, , "Item" & LoopC & "=" & UserList(UserIndex).Correo.Mensaje(LoopC).Item & vbCrLf
432         Put n, , "ItemCount" & LoopC & "=" & UserList(UserIndex).Correo.Mensaje(LoopC).ItemCount & vbCrLf
434         Put n, , "DATE" & LoopC & "=" & UserList(UserIndex).Correo.Mensaje(LoopC).Fecha & vbCrLf
436         Put n, , "LEIDO" & LoopC & "=" & UserList(UserIndex).Correo.Mensaje(LoopC).Leido & vbCrLf
        
438     Next LoopC
    
440     Close #n
    
        'Devuelve el head de muerto
442     If UserList(UserIndex).flags.Muerto = 1 Then
444         UserList(UserIndex).Char.Head = 0

        End If
    
        Exit Sub
    
ErrHandler:
446     Call LogError("Error en SaveNewUserCharfile")
448     Close #n

End Sub

Sub SetUserLogged(ByVal UserIndex As Integer)
        
        On Error GoTo SetUserLogged_Err
        

100     If Database_Enabled Then
102         Call SetUserLoggedDatabase(UserList(UserIndex).Id, UserList(UserIndex).AccountId)
        Else
104         Call WriteVar(CharPath & UCase$(UserList(UserIndex).name) & ".chr", "INIT", "Logged", 1)
106         Call WriteVar(CuentasPath & UCase$(UserList(UserIndex).Cuenta) & ".act", "INIT", "LOGEADA", 1)

        End If

        
        Exit Sub

SetUserLogged_Err:
108     Call RegistrarError(Err.Number, Err.Description, "ES.SetUserLogged", Erl)
110     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.Description, "ES.SaveBattlePoints", Erl)
108     Resume Next
        
End Sub

Function Status(ByVal UserIndex As Integer) As Byte
        
        On Error GoTo Status_Err
        

100     Status = UserList(UserIndex).Faccion.Status

        
        Exit Function

Status_Err:
102     Call RegistrarError(Err.Number, Err.Description, "ES.Status", Erl)
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
162     Call RegistrarError(Err.Number, Err.Description, "ES.BackUPnPc", Erl)
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
180     Call RegistrarError(Err.Number, Err.Description, "ES.CargarNpcBackUp", Erl)
182     Resume Next
        
End Sub

Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal Motivo As String)
        
        On Error GoTo LogBan_Err
        

100     Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).name, "BannedBy", UserList(UserIndex).name)
102     Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).name, "Reason", Motivo)

        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim mifile As Integer

104     mifile = FreeFile
106     Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
108     Print #mifile, UserList(BannedIndex).name
110     Close #mifile

        
        Exit Sub

LogBan_Err:
112     Call RegistrarError(Err.Number, Err.Description, "ES.LogBan", Erl)
114     Resume Next
        
End Sub

Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal Motivo As String)
        
        On Error GoTo LogBanFromName_Err
        

100     Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).name)
102     Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", Motivo)

        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim mifile As Integer

104     mifile = FreeFile
106     Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
108     Print #mifile, BannedName
110     Close #mifile

        
        Exit Sub

LogBanFromName_Err:
112     Call RegistrarError(Err.Number, Err.Description, "ES.LogBanFromName", Erl)
114     Resume Next
        
End Sub

Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal Motivo As String)
        
        On Error GoTo Ban_Err
        

100     Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
102     Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", Motivo)

        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim mifile As Integer

104     mifile = FreeFile
106     Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
108     Print #mifile, BannedName
110     Close #mifile

        
        Exit Sub

Ban_Err:
112     Call RegistrarError(Err.Number, Err.Description, "ES.Ban", Erl)
114     Resume Next
        
End Sub

Public Sub CargaApuestas()
        
        On Error GoTo CargaApuestas_Err
        

100     Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
102     Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
104     Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

        
        Exit Sub

CargaApuestas_Err:
106     Call RegistrarError(Err.Number, Err.Description, "ES.CargaApuestas", Erl)
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
148     Call RegistrarError(Err.Number, Err.Description, "ES.LoadRecursosEspeciales", Erl)
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
154     Call RegistrarError(Err.Number, Err.Description, "ES.LoadPesca", Erl)
156     Resume Next
        
End Sub

' Adaptado de https://www.vbforums.com/showthread.php?231925-VB-Quick-Sort-algorithm-(very-fast-sorting-algorithm)
Private Sub QuickSortPeces(ByVal First As Long, ByVal Last As Long)
        
        On Error GoTo QuickSortPeces_Err
        

        Dim low      As Long, high As Long

        Dim MidValue As String

        Dim Aux      As obj
    
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
116             Aux = Peces(low)
118             Peces(low) = Peces(high)
120             Peces(high) = Aux
122             low = low + 1
124             high = high - 1

            End If

126     Loop While low <= high
    
128     If First < high Then QuickSortPeces First, high
130     If low < Last Then QuickSortPeces low, Last

        
        Exit Sub

QuickSortPeces_Err:
132     Call RegistrarError(Err.Number, Err.Description, "ES.QuickSortPeces", Erl)
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
124     Call RegistrarError(Err.Number, Err.Description, "ES.BinarySearchPeces", Erl)
126     Resume Next
        
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
114         .GolpeUsar = IntervaloGolpeUsar
116         .TrabajarExtraer = IntervaloTrabajarExtraer
118         .TrabajarConstruir = IntervaloTrabajarConstruir
120         .UsarU = IntervaloUserPuedeUsarU
122         .UsarClic = IntervaloUserPuedeUsarClic

        End With

        
        Exit Sub

LoadUserIntervals_Err:
124     Call RegistrarError(Err.Number, Err.Description, "ES.LoadUserIntervals", Erl)
126     Resume Next
        
End Sub

Public Sub RegistrarError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
    '**********************************************************
    'Author: Jopi
    'Guarda una descripcion detallada del error en Errores.log
    '**********************************************************
        
        On Error GoTo RegistrarError_Err
    
        
        
        'Si lo del parametro Componente es ES IGUAL, al Componente del anterior error...
100     If Componente = HistorialError.Componente And _
           Numero = HistorialError.ErrorCode Then
       
           'Si ya recibimos error en el mismo componente 10 veces, es bastante probable que estemos en un bucle
            'x lo que no hace falta registrar el error.
102         If HistorialError.Contador = 10 Then
104             Debug.Print "Mismo error"
                Exit Sub
            End If
        
            'Agregamos el error al historial.
106         HistorialError.Contador = HistorialError.Contador + 1
        
        Else 'Si NO es igual, reestablecemos el contador.

108         HistorialError.Contador = 0
110         HistorialError.ErrorCode = Numero
112         HistorialError.Componente = Componente
            
        End If
    
        'Registramos el error en Errores.log
114     Dim File As Integer: File = FreeFile
        
116     Open App.Path & "\logs\Errores.log" For Append As #File
    
118         Print #File, "Error: " & Numero
120         Print #File, "Descripcion: " & Descripcion
        
122         Print #File, "Componente: " & Componente

124         If LenB(Linea) <> 0 Then
126             Print #File, "Linea: " & Linea
            End If

128         Print #File, "Fecha y Hora: " & Date$ & "-" & Time$
        
130         Print #File, vbNullString
        
132     Close #File
    
134     Debug.Print "Error: " & Numero & vbNewLine & _
                    "Descripcion: " & Descripcion & vbNewLine & _
                    "Componente: " & Componente & vbNewLine & _
                    "Linea: " & Linea & vbNewLine & _
                    "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine
        
        Exit Sub

RegistrarError_Err:
136     Call RegistrarError(Err.Number, Err.Description, "ES.RegistrarError", Erl)

        
End Sub

Function CountFiles(strFolder As String, strPattern As String) As Integer
        
        On Error GoTo CountFiles_Err
    
        
   
        Dim strFile As String
100         strFile = dir$(strFolder & "\" & strPattern)
    
102     Do Until Len(strFile) = 0
104         CountFiles = CountFiles + 1
106         strFile = dir$()
        Loop
    
108     If CountFiles <> 0 Then CountFiles = CountFiles + 1
    
        
        Exit Function

CountFiles_Err:
110     Call RegistrarError(Err.Number, Err.Description, "ES.CountFiles", Erl)

        
End Function

Public Function GetElapsedTime() As Single

    '***********************************************************************
    'Author: Wyrox
    'Obenemos el tiempo (en milisegundos) que pasó desde la ultima llamada.
    '***********************************************************************
    
    Dim end_time As Currency
    Static start_time As Currency
    Static timer_freq As Single

    'Get the timer frequency
    If timer_freq = 0 Then
        Dim temp_time As Currency
        Call QueryPerformanceFrequency(temp_time)
        timer_freq = 1000 / temp_time
    End If

    Call QueryPerformanceCounter(end_time)

    GetElapsedTime = (end_time - start_time) * timer_freq
    
    start_time = end_time

End Function
