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
    amount As Integer

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

Function EsAdmin(ByRef Name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        
        On Error GoTo EsAdmin_Err
        
100     EsAdmin = (val(Administradores.GetValue("Admin", Name)) = 1)

        
        Exit Function

EsAdmin_Err:
102     Call RegistrarError(Err.Number, Err.Description, "ES.EsAdmin", Erl)
104     Resume Next
        
End Function

Function EsDios(ByRef Name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        
        On Error GoTo EsDios_Err
        
100     EsDios = (val(Administradores.GetValue("Dios", Name)) = 1)

        
        Exit Function

EsDios_Err:
102     Call RegistrarError(Err.Number, Err.Description, "ES.EsDios", Erl)
104     Resume Next
        
End Function

Function EsSemiDios(ByRef Name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        
        On Error GoTo EsSemiDios_Err
        
100     EsSemiDios = (val(Administradores.GetValue("SemiDios", Name)) = 1)

        
        Exit Function

EsSemiDios_Err:
102     Call RegistrarError(Err.Number, Err.Description, "ES.EsSemiDios", Erl)
104     Resume Next
        
End Function

Function EsConsejero(ByRef Name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        
        On Error GoTo EsConsejero_Err
        
100     EsConsejero = (val(Administradores.GetValue("Consejero", Name)) = 1)

        
        Exit Function

EsConsejero_Err:
102     Call RegistrarError(Err.Number, Err.Description, "ES.EsConsejero", Erl)
104     Resume Next
        
End Function

Function EsRolesMaster(ByRef Name As String) As Boolean
        '***************************************************
        'Author: Unknown
        'Last Modification: 27/03/2011
        '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
        '***************************************************
        
        On Error GoTo EsRolesMaster_Err
        
100     EsRolesMaster = (val(Administradores.GetValue("RM", Name)) = 1)

        
        Exit Function

EsRolesMaster_Err:
102     Call RegistrarError(Err.Number, Err.Description, "ES.EsRolesMaster", Erl)
104     Resume Next
        
End Function

Public Function EsGmChar(ByRef Name As String) As Boolean
        '***************************************************
        'Author: ZaMa
        'Last Modification: 27/03/2011
        'Returns true if char is administrative user.
        '***************************************************
        
        On Error GoTo EsGmChar_Err
        
    
        Dim EsGM As Boolean
    
        ' Admin?
100     EsGM = EsAdmin(Name)

        ' Dios?
102     If Not EsGM Then EsGM = EsDios(Name)

        ' Semidios?
104     If Not EsGM Then EsGM = EsSemiDios(Name)

        ' Consejero?
106     If Not EsGM Then EsGM = EsConsejero(Name)

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
        Dim Name As String
        
        ' Anti-choreo de GM's
100     Set AdministratorAccounts = New Dictionary
        Dim TempName() As String
       
        ' Public container
102     Set Administradores = New clsIniManager
    
        ' Server ini info file
        Dim ServerIni As clsIniManager
104     Set ServerIni = New clsIniManager
106     Call ServerIni.Initialize(IniPath & "Server.ini")
       
        ' Admines
108     buf = val(ServerIni.GetValue("INIT", "Admines"))
    
110     For i = 1 To buf
112         Name = UCase$(ServerIni.GetValue("Admines", "Admin" & i))
114         TempName = Split(Name, "|", , vbTextCompare)
        
            ' Si NO declara el mail de la cuenta en el Server.ini, NO le doy privilegios.
116         If UBound(TempName()) > 0 Then
                ' AdministratorAccounts("Nick") = "Email"
118             AdministratorAccounts(TempName(0)) = TempName(1)
            
                ' Add key
120             Call Administradores.ChangeValue("Admin", TempName(0), "1")
            End If
        
122     Next i
    
        ' Dioses
124     buf = val(ServerIni.GetValue("INIT", "Dioses"))
    
126     For i = 1 To buf
128         Name = UCase$(ServerIni.GetValue("Dioses", "Dios" & i))
130         TempName = Split(Name, "|", , vbTextCompare)
        
            ' Si NO declara el mail de la cuenta en el Server.ini, NO le doy privilegios.
132         If UBound(TempName()) > 0 Then
                ' AdministratorAccounts("Nick") = "Email"
134             AdministratorAccounts(TempName(0)) = TempName(1)
            
                ' Add key
136             Call Administradores.ChangeValue("Dios", TempName(0), "1")
            End If
        
138     Next i
        
        ' SemiDioses
140     buf = val(ServerIni.GetValue("INIT", "SemiDioses"))
    
142     For i = 1 To buf
144         Name = UCase$(ServerIni.GetValue("SemiDioses", "SemiDios" & i))
146         TempName = Split(Name, "|", , vbTextCompare)
        
            ' Si NO declara el mail de la cuenta en el Server.ini, NO le doy privilegios.
148         If UBound(TempName()) > 0 Then
                ' AdministratorAccounts("Nick") = "Email"
150             AdministratorAccounts(TempName(0)) = TempName(1)
            
                ' Add key
152             Call Administradores.ChangeValue("SemiDios", TempName(0), "1")
            End If
        
154     Next i
    
        ' Consejeros
156     buf = val(ServerIni.GetValue("INIT", "Consejeros"))
        
158     For i = 1 To buf
160         Name = UCase$(ServerIni.GetValue("Consejeros", "Consejero" & i))
162         TempName = Split(Name, "|", , vbTextCompare)
        
            ' Si NO declara el mail de la cuenta en el Server.ini, NO le doy privilegios.
164         If UBound(TempName()) > 0 Then
                ' AdministratorAccounts("Nick") = "Email"
166             AdministratorAccounts(TempName(0)) = TempName(1)
            
                ' Add key
168             Call Administradores.ChangeValue("Consejero", TempName(0), "1")
            End If
        
170     Next i
    
        ' RolesMasters
172     buf = val(ServerIni.GetValue("INIT", "RolesMasters"))
        
174     For i = 1 To buf
176         Name = UCase$(ServerIni.GetValue("RolesMasters", "RM" & i))
178         TempName = Split(Name, "|", , vbTextCompare)
        
            ' Si NO declara el mail de la cuenta en el Server.ini, NO le doy privilegios.
180         If UBound(TempName()) > 0 Then
                ' AdministratorAccounts("Nick") = "Email"
182             AdministratorAccounts(TempName(0)) = TempName(1)
            
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

Public Function TxtDimension(ByVal Name As String) As Long
        
        On Error GoTo TxtDimension_Err
        

        Dim n As Integer, cad As String, Tam As Long

100     n = FreeFile(1)
102     Open Name For Input As #n
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

        Dim Leer    As New clsIniManager

102     Call Leer.Initialize(DatPath & "Hechizos.dat")

        'obtiene el numero de hechizos
104     NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))

106     ReDim Hechizos(1 To NumeroHechizos) As tHechizo

108     frmCargando.cargar.min = 0
110     frmCargando.cargar.max = NumeroHechizos
112     frmCargando.cargar.Value = 0

        'Llena la lista
114     For Hechizo = 1 To NumeroHechizos

116         Hechizos(Hechizo).velocidad = val(Leer.GetValue("Hechizo" & Hechizo, "Velocidad"))
    
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
188         Hechizos(Hechizo).MinMana = val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
190         Hechizos(Hechizo).MaxMana = val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
    
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
202                     Objetos(MH.NumeroOBJs).ObjAmmount = .ObjInfo.amount
               
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
        

        Dim BalanceIni As clsIniManager

100     Set BalanceIni = New clsIniManager
    
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
138             .ResistenciaMagica = val(BalanceIni.GetValue("MODRESISTENCIAMAGICA", SearchVar))
            End With

140     Next i
    
        'Modificadores de Raza
142     For i = 1 To NUMRAZAS
144         SearchVar = Replace$(Tilde(ListaRazas(i)), " ", vbNullString)

146         With ModRaza(i)
148             .Fuerza = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Fuerza"))
150             .Agilidad = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Agilidad"))
152             .Inteligencia = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Inteligencia"))
154             .Carisma = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Carisma"))
156             .Constitucion = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Constitucion"))
            End With

158     Next i

        'Extra
160     PorcentajeRecuperoMana = val(BalanceIni.GetValue("EXTRA", "PorcentajeRecuperoMana"))
162     DificultadSubirSkill = val(BalanceIni.GetValue("EXTRA", "DificultadSubirSkill"))
164     InfluenciaPromedioVidas = val(BalanceIni.GetValue("EXTRA", "InfluenciaPromedioVidas"))
166     DesbalancePromedioVidas = val(BalanceIni.GetValue("EXTRA", "DesbalancePromedioVidas"))
168     RangoVidas = val(BalanceIni.GetValue("EXTRA", "RangoVidas"))
170     ModDañoGolpeCritico = val(BalanceIni.GetValue("EXTRA", "ModDañoGolpeCritico"))

        ' Exp
172     For i = 1 To STAT_MAXELV
174         ExpLevelUp(i) = val(BalanceIni.GetValue("EXP", i))
        Next
    
176     Set BalanceIni = Nothing
    
178     AgregarAConsola "Se cargó el balance (Balance.dat)"

        
        Exit Sub

LoadBalance_Err:
180     Call RegistrarError(Err.Number, Err.Description, "ES.LoadBalance", Erl)
182     Resume Next
        
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

    Dim Leer   As clsIniManager
    Set Leer = New clsIniManager
    Call Leer.Initialize(DatPath & "Obj.dat")

        'obtiene el numero de obj
106     NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))
    
108     With frmCargando.cargar
110         .min = 0
112         .max = NumObjDatas
114         .Value = 0
        End With
    
116     ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
    
        Dim ObjKey As String
        Dim str As String, Field() As String
  
        'Llena la lista
118     For Object = 1 To NumObjDatas
        
120         With ObjData(Object)

122             ObjKey = "OBJ" & Object
        
124             .Name = Leer.GetValue(ObjKey, "Name")
    
                ' If .Name = "" Then
                '   Call LogError("Objeto libre:" & Object)
                ' End If
    
                ' If .name = "" Then
                ' Debug.Print Object
                ' End If
    
                'Pablo (ToxicWaste) Log de Objetos.
126             .Log = val(Leer.GetValue(ObjKey, "Log"))
128             .NoLog = val(Leer.GetValue(ObjKey, "NoLog"))
                '07/09/07
    
130             .GrhIndex = val(Leer.GetValue(ObjKey, "GrhIndex"))
132             .OBJType = val(Leer.GetValue(ObjKey, "ObjType"))
    
134             .Newbie = val(Leer.GetValue(ObjKey, "Newbie"))

                'Propiedades by Lader 05-05-08
136             .Instransferible = val(Leer.GetValue(ObjKey, "Instransferible"))
138             .Destruye = val(Leer.GetValue(ObjKey, "Destruye"))
140             .Intirable = val(Leer.GetValue(ObjKey, "Intirable"))
    
142             .CantidadSkill = val(Leer.GetValue(ObjKey, "CantidadSkill"))
144             .QueSkill = val(Leer.GetValue(ObjKey, "QueSkill"))
146             .QueAtributo = val(Leer.GetValue(ObjKey, "queatributo"))
148             .CuantoAumento = val(Leer.GetValue(ObjKey, "cuantoaumento"))
150             .MinELV = val(Leer.GetValue(ObjKey, "MinELV"))
152             .Subtipo = val(Leer.GetValue(ObjKey, "Subtipo"))
154             .Dorada = val(Leer.GetValue(ObjKey, "Dorada"))
156             .VidaUtil = val(Leer.GetValue(ObjKey, "VidaUtil"))
158             .TiempoRegenerar = val(Leer.GetValue(ObjKey, "TiempoRegenerar"))
    
160             .donador = val(Leer.GetValue(ObjKey, "donador"))
    
                Dim i As Integer

162             Select Case .OBJType

                    Case eOBJType.otHerramientas
164                     .WeaponAnim = val(Leer.GetValue(ObjKey, "Anim"))
166                     .Power = val(Leer.GetValue(ObjKey, "Poder"))
            
168                 Case eOBJType.otArmadura
170                     .Real = val(Leer.GetValue(ObjKey, "Real"))
172                     .Caos = val(Leer.GetValue(ObjKey, "Caos"))
174                     .LingH = val(Leer.GetValue(ObjKey, "LingH"))
176                     .LingP = val(Leer.GetValue(ObjKey, "LingP"))
178                     .LingO = val(Leer.GetValue(ObjKey, "LingO"))
180                     .SkHerreria = val(Leer.GetValue(ObjKey, "SkHerreria"))
182                     .ResistenciaMagica = val(Leer.GetValue(ObjKey, "ResistenciaMagica"))
184                     .Invernal = val(Leer.GetValue(ObjKey, "Invernal")) > 0
        
186                 Case eOBJType.otEscudo
188                     .ShieldAnim = val(Leer.GetValue(ObjKey, "Anim"))
190                     .LingH = val(Leer.GetValue(ObjKey, "LingH"))
192                     .LingP = val(Leer.GetValue(ObjKey, "LingP"))
194                     .LingO = val(Leer.GetValue(ObjKey, "LingO"))
196                     .SkHerreria = val(Leer.GetValue(ObjKey, "SkHerreria"))
198                     .Real = val(Leer.GetValue(ObjKey, "Real"))
200                     .Caos = val(Leer.GetValue(ObjKey, "Caos"))
202                     .ResistenciaMagica = val(Leer.GetValue(ObjKey, "ResistenciaMagica"))
        
204                 Case eOBJType.otCasco
206                     .CascoAnim = val(Leer.GetValue(ObjKey, "Anim"))
208                     .LingH = val(Leer.GetValue(ObjKey, "LingH"))
210                     .LingP = val(Leer.GetValue(ObjKey, "LingP"))
212                     .LingO = val(Leer.GetValue(ObjKey, "LingO"))
214                     .SkHerreria = val(Leer.GetValue(ObjKey, "SkHerreria"))
216                     .Real = val(Leer.GetValue(ObjKey, "Real"))
218                     .Caos = val(Leer.GetValue(ObjKey, "Caos"))
220                     .ResistenciaMagica = val(Leer.GetValue(ObjKey, "ResistenciaMagica"))
        
222                 Case eOBJType.otWeapon
224                     .WeaponAnim = val(Leer.GetValue(ObjKey, "Anim"))
226                     .Apuñala = val(Leer.GetValue(ObjKey, "Apuñala"))
228                     .Envenena = val(Leer.GetValue(ObjKey, "Envenena"))
230                     .Paraliza = val(Leer.GetValue(ObjKey, "Paraliza"))
232                     .Estupidiza = val(Leer.GetValue(ObjKey, "Estupidiza"))
234                     .incinera = val(Leer.GetValue(ObjKey, "Incinera"))
        
236                     .MaxHit = val(Leer.GetValue(ObjKey, "MaxHIT"))
238                     .MinHIT = val(Leer.GetValue(ObjKey, "MinHIT"))
240                     .Proyectil = val(Leer.GetValue(ObjKey, "Proyectil"))
242                     .Municion = val(Leer.GetValue(ObjKey, "Municiones"))
244                     .Power = val(Leer.GetValue(ObjKey, "StaffPower"))
246                     .MagicDamageBonus = val(Leer.GetValue(ObjKey, "MagicDamageBonus"))
248                     .Refuerzo = val(Leer.GetValue(ObjKey, "Refuerzo"))
            
250                     .LingH = val(Leer.GetValue(ObjKey, "LingH"))
252                     .LingP = val(Leer.GetValue(ObjKey, "LingP"))
254                     .LingO = val(Leer.GetValue(ObjKey, "LingO"))
256                     .SkHerreria = val(Leer.GetValue(ObjKey, "SkHerreria"))
258                     .Real = val(Leer.GetValue(ObjKey, "Real"))
260                     .Caos = val(Leer.GetValue(ObjKey, "Caos"))
262                     .EfectoMagico = val(Leer.GetValue(ObjKey, "efectomagico"))
264                     .Revive = val(Leer.GetValue(ObjKey, "Revive")) <> 0
                    
266                     .DosManos = val(Leer.GetValue(ObjKey, "DosManos"))
        
268                 Case eOBJType.otInstrumentos
        
                        'Pablo (ToxicWaste)
270                     .Real = val(Leer.GetValue(ObjKey, "Real"))
272                     .Caos = val(Leer.GetValue(ObjKey, "Caos"))
        
274                 Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
276                     .IndexAbierta = val(Leer.GetValue(ObjKey, "IndexAbierta"))
278                     .IndexCerrada = val(Leer.GetValue(ObjKey, "IndexCerrada"))
280                     .IndexCerradaLlave = val(Leer.GetValue(ObjKey, "IndexCerradaLlave"))
        
282                 Case otPociones
284                     .TipoPocion = val(Leer.GetValue(ObjKey, "TipoPocion"))
286                     .MaxModificador = val(Leer.GetValue(ObjKey, "MaxModificador"))
288                     .MinModificador = val(Leer.GetValue(ObjKey, "MinModificador"))
            
290                     .DuracionEfecto = val(Leer.GetValue(ObjKey, "DuracionEfecto"))
292                     .Raices = val(Leer.GetValue(ObjKey, "Raices"))
294                     .SkPociones = val(Leer.GetValue(ObjKey, "SkPociones"))
296                     .Porcentaje = val(Leer.GetValue(ObjKey, "Porcentaje"))
        
298                 Case eOBJType.otBarcos
300                     .MaxHit = val(Leer.GetValue(ObjKey, "MaxHIT"))
302                     .MinHIT = val(Leer.GetValue(ObjKey, "MinHIT"))
304                     .velocidad = val(Leer.GetValue(ObjKey, "Velocidad"))

306                 Case eOBJType.otMonturas
308                     .MaxHit = val(Leer.GetValue(ObjKey, "MaxHIT"))
310                     .MinHIT = val(Leer.GetValue(ObjKey, "MinHIT"))
312                     .MinDef = val(Leer.GetValue(ObjKey, "MINDEF"))
314                     .MaxDef = val(Leer.GetValue(ObjKey, "MAXDEF"))
316                     .Real = val(Leer.GetValue(ObjKey, "Real"))
318                     .Caos = val(Leer.GetValue(ObjKey, "Caos"))
320                     .velocidad = val(Leer.GetValue(ObjKey, "Velocidad"))
        
322                 Case eOBJType.otFlechas
324                     .MaxHit = val(Leer.GetValue(ObjKey, "MaxHIT"))
326                     .MinHIT = val(Leer.GetValue(ObjKey, "MinHIT"))
328                     .Envenena = val(Leer.GetValue(ObjKey, "Envenena"))
330                     .Paraliza = val(Leer.GetValue(ObjKey, "Paraliza"))
332                     .Estupidiza = val(Leer.GetValue(ObjKey, "Estupidiza"))
334                     .incinera = val(Leer.GetValue(ObjKey, "Incinera"))
        
336                     .Refuerzo = val(Leer.GetValue(ObjKey, "Refuerzo"))
338                     .SkCarpinteria = val(Leer.GetValue(ObjKey, "SkCarpinteria"))
340                     .EfectoMagico = val(Leer.GetValue(ObjKey, "efectomagico"))
            
342                     .Snd1 = val(Leer.GetValue(ObjKey, "SND1"))
344                     .Snd2 = val(Leer.GetValue(ObjKey, "SND2"))
            
                        'Case eOBJType.otAnillos 'Pablo (ToxicWaste)
                        '  .LingH = val(Leer.GetValue(ObjKey, "LingH"))
                        '  .LingP = val(Leer.GetValue(ObjKey, "LingP"))
                        '  .LingO = val(Leer.GetValue(ObjKey, "LingO"))
                        '  .SkHerreria = val(Leer.GetValue(ObjKey, "SkHerreria"))
            
                        'Pasajes Ladder 05-05-08
346                 Case eOBJType.otpasajes
348                     .DesdeMap = val(Leer.GetValue(ObjKey, "DesdeMap"))
350                     .HastaMap = val(Leer.GetValue(ObjKey, "Map"))
352                     .HastaX = val(Leer.GetValue(ObjKey, "X"))
354                     .HastaY = val(Leer.GetValue(ObjKey, "Y"))
356                     .NecesitaNave = val(Leer.GetValue(ObjKey, "NecesitaNave"))
            
358                 Case eOBJType.OtDonador
360                     .HastaMap = val(Leer.GetValue(ObjKey, "Map"))
362                     .HastaX = val(Leer.GetValue(ObjKey, "X"))
364                     .HastaY = val(Leer.GetValue(ObjKey, "Y"))
        
366                 Case eOBJType.otMagicos
368                     .EfectoMagico = val(Leer.GetValue(ObjKey, "efectomagico"))

370                     If .EfectoMagico = 15 Then
372                         PENDIENTE = Object

                        End If
            
374                 Case eOBJType.otRunas
376                     .TipoRuna = val(Leer.GetValue(ObjKey, "TipoRuna"))
378                     .DesdeMap = val(Leer.GetValue(ObjKey, "DesdeMap"))
380                     .HastaMap = val(Leer.GetValue(ObjKey, "Map"))
382                     .HastaX = val(Leer.GetValue(ObjKey, "X"))
384                     .HastaY = val(Leer.GetValue(ObjKey, "Y"))
                    
386                 Case eOBJType.otNudillos
388                     .MinHIT = val(Leer.GetValue(ObjKey, "MinHIT"))
390                     .MaxHit = val(Leer.GetValue(ObjKey, "MaxHit"))
392                     .Envenena = val(Leer.GetValue(ObjKey, "Envenena"))
394                     .Paraliza = val(Leer.GetValue(ObjKey, "Paraliza"))
396                     .Estupidiza = val(Leer.GetValue(ObjKey, "Estupidiza"))
398                     .incinera = val(Leer.GetValue(ObjKey, "Incinera"))
400                     .WeaponAnim = val(Leer.GetValue(ObjKey, "Anim"))
402                     .SkHerreria = val(Leer.GetValue(ObjKey, "SkHerreria"))
            
404                 Case eOBJType.otPergaminos
        
                        ' .ClasePermitida = Leer.GetValue(ObjKey, "CP")
        
406                 Case eOBJType.OtCofre
408                     .CantItem = val(Leer.GetValue(ObjKey, "CantItem"))

410                     If .Subtipo = 1 Then
412                         ReDim .Item(1 To .CantItem)
                
414                         For i = 1 To .CantItem
416                             .Item(i).ObjIndex = val(Leer.GetValue(ObjKey, "Item" & i))
418                             .Item(i).amount = val(Leer.GetValue(ObjKey, "Cantidad" & i))
420                         Next i

                        Else
422                         ReDim .Item(1 To .CantItem)
                
424                         .CantEntrega = val(Leer.GetValue(ObjKey, "CantEntrega"))

426                         For i = 1 To .CantItem
428                             .Item(i).ObjIndex = val(Leer.GetValue(ObjKey, "Item" & i))
430                             .Item(i).amount = val(Leer.GetValue(ObjKey, "Cantidad" & i))
432                         Next i

                        End If
            
434                 Case eOBJType.otYacimiento
                        ' Drop gemas yacimientos
436                     .CantItem = val(Leer.GetValue(ObjKey, "Gemas"))
            
438                     If .CantItem > 0 Then
440                         ReDim .Item(1 To .CantItem)

442                         For i = 1 To .CantItem
444                             str = Leer.GetValue(ObjKey, "Gema" & i)
446                             Field = Split(str, "-")
448                             .Item(i).ObjIndex = val(Field(0))    ' ObjIndex
450                             .Item(i).amount = val(Field(1))      ' Probabilidad de drop (1 en X)
452                         Next i

                        End If
                
454                 Case eOBJType.otDañoMagico
456                     .MagicDamageBonus = val(Leer.GetValue(ObjKey, "MagicDamageBonus"))
458                     .Revive = val(Leer.GetValue(ObjKey, "Revive")) <> 0

460                 Case eOBJType.otResistencia
462                     .ResistenciaMagica = val(Leer.GetValue(ObjKey, "ResistenciaMagica"))
            
                End Select
    
464             .MinSkill = val(Leer.GetValue(ObjKey, "MinSkill"))

466             .Elfico = val(Leer.GetValue(ObjKey, "Elfico"))

468             .Snd1 = val(Leer.GetValue(ObjKey, "SND1"))
470             .Snd2 = val(Leer.GetValue(ObjKey, "SND2"))
472             .Snd3 = val(Leer.GetValue(ObjKey, "SND3"))
                'DELETE
474             .SndAura = val(Leer.GetValue(ObjKey, "SndAura"))
                '
    
476             .NoSeLimpia = val(Leer.GetValue(ObjKey, "NoSeLimpia"))
478             .Subastable = val(Leer.GetValue(ObjKey, "Subastable"))
    
480             .ParticulaGolpe = val(Leer.GetValue(ObjKey, "ParticulaGolpe"))
482             .ParticulaViaje = val(Leer.GetValue(ObjKey, "ParticulaViaje"))
484             .ParticulaGolpeTime = val(Leer.GetValue(ObjKey, "ParticulaGolpeTime"))
    
486             .Ropaje = val(Leer.GetValue(ObjKey, "NumRopaje"))
488             .HechizoIndex = val(Leer.GetValue(ObjKey, "HechizoIndex"))
    
490             .LingoteIndex = val(Leer.GetValue(ObjKey, "LingoteIndex"))
    
492             .MineralIndex = val(Leer.GetValue(ObjKey, "MineralIndex"))
    
494             .MaxHp = val(Leer.GetValue(ObjKey, "MaxHP"))
496             .MinHp = val(Leer.GetValue(ObjKey, "MinHP"))
    
498             .Mujer = val(Leer.GetValue(ObjKey, "Mujer"))
500             .Hombre = val(Leer.GetValue(ObjKey, "Hombre"))
    
502             .PielLobo = val(Leer.GetValue(ObjKey, "PielLobo"))
504             .PielOsoPardo = val(Leer.GetValue(ObjKey, "PielOsoPardo"))
506             .PielOsoPolaR = val(Leer.GetValue(ObjKey, "PielOsoPolaR"))
508             .SkMAGOria = val(Leer.GetValue(ObjKey, "SKSastreria"))
    
510             .CreaParticula = Leer.GetValue(ObjKey, "CreaParticula")
512             .CreaFX = val(Leer.GetValue(ObjKey, "CreaFX"))
514             .CreaGRH = Leer.GetValue(ObjKey, "CreaGRH")
516             .CreaLuz = Leer.GetValue(ObjKey, "CreaLuz")
    
518             .MinHam = val(Leer.GetValue(ObjKey, "MinHam"))
520             .MinSed = val(Leer.GetValue(ObjKey, "MinAgu"))
    
522             .MinDef = val(Leer.GetValue(ObjKey, "MINDEF"))
524             .MaxDef = val(Leer.GetValue(ObjKey, "MAXDEF"))
526             .def = (.MinDef + .MaxDef) / 2
    
528             .ClaseTipo = val(Leer.GetValue(ObjKey, "ClaseTipo"))

530             .RazaEnana = val(Leer.GetValue(ObjKey, "RazaEnana"))
532             .RazaDrow = val(Leer.GetValue(ObjKey, "RazaDrow"))
534             .RazaElfa = val(Leer.GetValue(ObjKey, "RazaElfa"))
536             .RazaGnoma = val(Leer.GetValue(ObjKey, "RazaGnoma"))
538             .RazaOrca = val(Leer.GetValue(ObjKey, "RazaOrca"))
540             .RazaHumana = val(Leer.GetValue(ObjKey, "RazaHumana"))
    
542             .Valor = val(Leer.GetValue(ObjKey, "Valor"))
    
544             .Crucial = val(Leer.GetValue(ObjKey, "Crucial"))
    
                '.Cerrada = val(Leer.GetValue(ObjKey, "abierta")) cerrada = abierta??? WTF???????
546             .Cerrada = val(Leer.GetValue(ObjKey, "Cerrada"))

548             If .Cerrada = 1 Then
550                 .Llave = val(Leer.GetValue(ObjKey, "Llave"))
552                 .clave = val(Leer.GetValue(ObjKey, "Clave"))

                End If
    
                'Puertas y llaves
554             .clave = val(Leer.GetValue(ObjKey, "Clave"))
    
556             .texto = Leer.GetValue(ObjKey, "Texto")
558             .GrhSecundario = val(Leer.GetValue(ObjKey, "VGrande"))
    
560             .Agarrable = val(Leer.GetValue(ObjKey, "Agarrable"))
562             .ForoID = Leer.GetValue(ObjKey, "ID")
    
                'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico  -  Nunca más papu
                Dim n As Integer
                Dim S As String

564             For i = 1 To NUMCLASES
566                 S = UCase$(Leer.GetValue(ObjKey, "CP" & i))
568                 n = 1

570                 Do While LenB(S) > 0 And Tilde(ListaClases(n)) <> Trim$(S)
572                     n = n + 1
                    Loop
            
574                 .ClaseProhibida(i) = IIf(LenB(S) > 0, n, 0)
576             Next i
        
578             For i = 1 To NUMRAZAS
580                 S = UCase$(Leer.GetValue(ObjKey, "RP" & i))
582                 n = 1

584                 Do While LenB(S) > 0 And Tilde(ListaRazas(n)) <> Trim$(S)
586                     n = n + 1
                    Loop
            
588                 .RazaProhibida(i) = IIf(LenB(S) > 0, n, 0)
590             Next i
        
                ' Skill requerido
592             str = Leer.GetValue(ObjKey, "SkillRequerido")

594             If Len(str) > 0 Then
596                 Field = Split(str, "-")
            
598                 n = 1

600                 Do While LenB(Field(0)) > 0 And Tilde(SkillsNames(n)) <> Tilde(Field(0))
602                     n = n + 1
                    Loop
    
604                 .SkillIndex = IIf(LenB(Field(0)) > 0, n, 0)
606                 .SkillRequerido = val(Field(1))

                End If

                ' -----------------
    
608             .SkCarpinteria = val(Leer.GetValue(ObjKey, "SkCarpinteria"))
    
                'If .SkCarpinteria > 0 Then
610             .Madera = val(Leer.GetValue(ObjKey, "Madera"))

612             .MaderaElfica = val(Leer.GetValue(ObjKey, "MaderaElfica"))
    
                'Bebidas
614             .MinSta = val(Leer.GetValue(ObjKey, "MinST"))
    
616             .NoSeCae = val(Leer.GetValue(ObjKey, "NoSeCae"))
    
618             frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        
            End With
            
            ' WyroX: Cada 10 objetos revivo la interfaz
620         If Object Mod 10 = 0 Then DoEvents
        
622     Next Object

624     Set Leer = Nothing
    
626     Call InitTesoro
628     Call InitRegalo

        Exit Sub

ErrHandler:
630     MsgBox "error cargando objetos " & Err.Number & ": " & Err.Description & ". Error producido al cargar el objeto: " & Object

End Sub

Sub LoadUserStats(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)
        
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
152     UserList(UserIndex).Stats.ELV = CByte(UserFile.GetValue("STATS", "ELV"))

154     UserList(UserIndex).flags.Envenena = CByte(UserFile.GetValue("MAGIA", "ENVENENA"))
156     UserList(UserIndex).flags.Paraliza = CByte(UserFile.GetValue("MAGIA", "PARALIZA"))
158     UserList(UserIndex).flags.incinera = CByte(UserFile.GetValue("MAGIA", "INCINERA")) 'Estupidiza
160     UserList(UserIndex).flags.Estupidiza = CByte(UserFile.GetValue("MAGIA", "Estupidiza"))

162     UserList(UserIndex).flags.PendienteDelSacrificio = CByte(UserFile.GetValue("MAGIA", "PENDIENTE"))
164     UserList(UserIndex).flags.CarroMineria = CByte(UserFile.GetValue("MAGIA", "CarroMineria"))
166     UserList(UserIndex).flags.NoPalabrasMagicas = CByte(UserFile.GetValue("MAGIA", "NOPALABRASMAGICAS"))

168     If UserList(UserIndex).flags.Muerto = 0 Then
170         UserList(UserIndex).Char.Otra_Aura = CStr(UserFile.GetValue("MAGIA", "OTRA_AURA"))

        End If

        'UserList(UserIndex).flags.DañoMagico = CByte(UserFile.GetValue("MAGIA", "DañoMagico"))
        'UserList(UserIndex).flags.ResistenciaMagica = CByte(UserFile.GetValue("MAGIA", "ResistenciaMagica"))

        'Nuevos
172     UserList(UserIndex).flags.RegeneracionMana = CByte(UserFile.GetValue("MAGIA", "RegeneracionMana"))
174     UserList(UserIndex).flags.AnilloOcultismo = CByte(UserFile.GetValue("MAGIA", "AnilloOcultismo"))
176     UserList(UserIndex).flags.NoDetectable = CByte(UserFile.GetValue("MAGIA", "NoDetectable"))
178     UserList(UserIndex).flags.NoMagiaEfecto = CByte(UserFile.GetValue("MAGIA", "NoMagiaEfeceto"))
180     UserList(UserIndex).flags.RegeneracionHP = CByte(UserFile.GetValue("MAGIA", "RegeneracionHP"))
182     UserList(UserIndex).flags.RegeneracionSta = CByte(UserFile.GetValue("MAGIA", "RegeneracionSta"))

184     UserList(UserIndex).Stats.UsuariosMatados = CLng(UserFile.GetValue("MUERTES", "UserMuertes"))
186     UserList(UserIndex).Stats.NPCsMuertos = CInt(UserFile.GetValue("MUERTES", "NpcsMuertes"))

188     UserList(UserIndex).Stats.InventLevel = CInt(UserFile.GetValue("STATS", "InventLevel"))

190     If CByte(UserFile.GetValue("CONSEJO", "PERTENECE")) Then UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.RoyalCouncil

192     If CByte(UserFile.GetValue("CONSEJO", "PERTENECECAOS")) Then UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.ChaosCouncil

        
        Exit Sub

LoadUserStats_Err:
194     Call RegistrarError(Err.Number, Err.Description, "ES.LoadUserStats", Erl)
196     Resume Next
        
End Sub

Sub LoadUserInit(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)
        
        On Error GoTo LoadUserInit_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 19/11/2006
        'Loads the Users records
        '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, MatadosIngreso y NextRecompensa.
        '23/01/2007 Pablo (ToxicWaste) - Quito CriminalesMatados de Stats porque era redundante.
        '*************************************************
        Dim LoopC As Long

        Dim ln    As String

100     UserList(UserIndex).Faccion.Status = CByte(UserFile.GetValue("FACCIONES", "Status"))
102     UserList(UserIndex).Faccion.ArmadaReal = CByte(UserFile.GetValue("FACCIONES", "EjercitoReal"))
104     UserList(UserIndex).Faccion.FuerzasCaos = CByte(UserFile.GetValue("FACCIONES", "EjercitoCaos"))
106     UserList(UserIndex).Faccion.ciudadanosMatados = CLng(UserFile.GetValue("FACCIONES", "CiudMatados"))
108     UserList(UserIndex).Faccion.CriminalesMatados = CLng(UserFile.GetValue("FACCIONES", "CrimMatados"))
110     UserList(UserIndex).Faccion.RecibioArmaduraCaos = CByte(UserFile.GetValue("FACCIONES", "rArCaos"))
112     UserList(UserIndex).Faccion.RecibioArmaduraReal = CByte(UserFile.GetValue("FACCIONES", "rArReal"))
114     UserList(UserIndex).Faccion.RecibioExpInicialCaos = CByte(UserFile.GetValue("FACCIONES", "rExCaos"))
116     UserList(UserIndex).Faccion.RecibioExpInicialReal = CByte(UserFile.GetValue("FACCIONES", "rExReal"))
118     UserList(UserIndex).Faccion.RecompensasCaos = CLng(UserFile.GetValue("FACCIONES", "recCaos"))
120     UserList(UserIndex).Faccion.RecompensasReal = CLng(UserFile.GetValue("FACCIONES", "recReal"))
122     UserList(UserIndex).Faccion.Reenlistadas = CByte(UserFile.GetValue("FACCIONES", "Reenlistadas"))
124     UserList(UserIndex).Faccion.NivelIngreso = CInt(UserFile.GetValue("FACCIONES", "NivelIngreso"))
126     UserList(UserIndex).Faccion.MatadosIngreso = CInt(UserFile.GetValue("FACCIONES", "MatadosIngreso"))
128     UserList(UserIndex).Faccion.NextRecompensa = CInt(UserFile.GetValue("FACCIONES", "NextRecompensa"))

130     UserList(UserIndex).flags.Muerto = CByte(UserFile.GetValue("FLAGS", "Muerto"))
132     UserList(UserIndex).flags.Escondido = CByte(UserFile.GetValue("FLAGS", "Escondido"))

134     UserList(UserIndex).flags.Hambre = CByte(UserFile.GetValue("FLAGS", "Hambre"))
136     UserList(UserIndex).flags.Sed = CByte(UserFile.GetValue("FLAGS", "Sed"))
138     UserList(UserIndex).flags.Desnudo = CByte(UserFile.GetValue("FLAGS", "Desnudo"))
140     UserList(UserIndex).flags.Navegando = CByte(UserFile.GetValue("FLAGS", "Navegando"))
142     UserList(UserIndex).flags.Envenenado = CByte(UserFile.GetValue("FLAGS", "Envenenado"))
144     UserList(UserIndex).flags.Paralizado = CByte(UserFile.GetValue("FLAGS", "Paralizado"))
146     UserList(UserIndex).flags.Incinerado = CByte(UserFile.GetValue("FLAGS", "Incinerado"))
148     UserList(UserIndex).flags.Inmovilizado = CByte(UserFile.GetValue("FLAGS", "Inmovilizado"))

150     UserList(UserIndex).flags.ScrollExp = CSng(UserFile.GetValue("FLAGS", "ScrollExp"))
152     UserList(UserIndex).flags.ScrollOro = CSng(UserFile.GetValue("FLAGS", "ScrollOro"))

154     If UserList(UserIndex).flags.Paralizado = 1 Then
156         UserList(UserIndex).Counters.Paralisis = IntervaloParalizado

        End If

158     If UserList(UserIndex).flags.Inmovilizado = 1 Then
160         UserList(UserIndex).Counters.Inmovilizado = 20

        End If

162     UserList(UserIndex).Counters.Pena = CLng(UserFile.GetValue("COUNTERS", "Pena"))

164     UserList(UserIndex).Counters.ScrollExperiencia = CLng(UserFile.GetValue("COUNTERS", "ScrollExperiencia"))
166     UserList(UserIndex).Counters.ScrollOro = CLng(UserFile.GetValue("COUNTERS", "ScrollOro"))

168     UserList(UserIndex).Counters.Oxigeno = CLng(UserFile.GetValue("COUNTERS", "Oxigeno"))

170     UserList(UserIndex).MENSAJEINFORMACION = UserFile.GetValue("INIT", "MENSAJEINFORMACION")

172     UserList(UserIndex).genero = UserFile.GetValue("INIT", "Genero")
174     UserList(UserIndex).clase = UserFile.GetValue("INIT", "Clase")
176     UserList(UserIndex).raza = UserFile.GetValue("INIT", "Raza")
178     UserList(UserIndex).Hogar = UserFile.GetValue("INIT", "Hogar")
180     UserList(UserIndex).Char.Heading = CInt(UserFile.GetValue("INIT", "Heading"))

182     UserList(UserIndex).OrigChar.Head = CInt(UserFile.GetValue("INIT", "Head"))
184     UserList(UserIndex).OrigChar.Body = CInt(UserFile.GetValue("INIT", "Body"))
186     UserList(UserIndex).OrigChar.WeaponAnim = CInt(UserFile.GetValue("INIT", "Arma"))
188     UserList(UserIndex).OrigChar.ShieldAnim = CInt(UserFile.GetValue("INIT", "Escudo"))
190     UserList(UserIndex).OrigChar.CascoAnim = CInt(UserFile.GetValue("INIT", "Casco"))

        #If ConUpTime Then
192         UserList(UserIndex).UpTime = CLng(UserFile.GetValue("INIT", "UpTime"))
        #End If

194     UserList(UserIndex).OrigChar.Heading = UserList(UserIndex).Char.Heading

196     If UserList(UserIndex).flags.Muerto = 0 Then
198         UserList(UserIndex).Char = UserList(UserIndex).OrigChar
        Else
200         UserList(UserIndex).Char.Body = iCuerpoMuerto
202         UserList(UserIndex).Char.Head = 0
204         UserList(UserIndex).Char.WeaponAnim = NingunArma
206         UserList(UserIndex).Char.ShieldAnim = NingunEscudo
208         UserList(UserIndex).Char.CascoAnim = NingunCasco

        End If

210     UserList(UserIndex).Desc = UserFile.GetValue("INIT", "Desc")

212     UserList(UserIndex).flags.BanMotivo = UserFile.GetValue("BAN", "BanMotivo")
214     UserList(UserIndex).flags.Montado = CByte(UserFile.GetValue("FLAGS", "Montado"))
216     UserList(UserIndex).flags.VecesQueMoriste = CLng(UserFile.GetValue("FLAGS", "VecesQueMoriste"))

218     UserList(UserIndex).flags.MinutosRestantes = CLng(UserFile.GetValue("FLAGS", "MinutosRestantes"))
220     UserList(UserIndex).flags.Silenciado = CLng(UserFile.GetValue("FLAGS", "Silenciado"))
222     UserList(UserIndex).flags.SegundosPasados = CLng(UserFile.GetValue("FLAGS", "SegundosPasados"))

        'CASAMIENTO LADDER
224     UserList(UserIndex).flags.Casado = CInt(UserFile.GetValue("FLAGS", "CASADO"))
226     UserList(UserIndex).flags.Pareja = UserFile.GetValue("FLAGS", "PAREJA")

228     UserList(UserIndex).Pos.Map = CInt(ReadField(1, UserFile.GetValue("INIT", "Position"), 45))
230     UserList(UserIndex).Pos.X = CInt(ReadField(2, UserFile.GetValue("INIT", "Position"), 45))
232     UserList(UserIndex).Pos.Y = CInt(ReadField(3, UserFile.GetValue("INIT", "Position"), 45))

234     UserList(UserIndex).Invent.NroItems = CInt(UserFile.GetValue("Inventory", "CantidadItems"))

        '[KEVIN]--------------------------------------------------------------------
        '***********************************************************************************
236     UserList(UserIndex).BancoInvent.NroItems = CInt(UserFile.GetValue("BancoInventory", "CantidadItems"))

        'Lista de objetos del banco
238     For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
240         ln = UserFile.GetValue("BancoInventory", "Obj" & LoopC)
242         UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
244         UserList(UserIndex).BancoInvent.Object(LoopC).amount = CInt(ReadField(2, ln, 45))
246     Next LoopC

        '------------------------------------------------------------------------------------
        '[/KEVIN]*****************************************************************************

        'Lista de objetos
248     For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots
250         ln = UserFile.GetValue("Inventory", "Obj" & LoopC)
252         UserList(UserIndex).Invent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
254         UserList(UserIndex).Invent.Object(LoopC).amount = CInt(ReadField(2, ln, 45))
256         UserList(UserIndex).Invent.Object(LoopC).Equipped = CByte(ReadField(3, ln, 45))
258     Next LoopC

260     UserList(UserIndex).Invent.WeaponEqpSlot = CByte(UserFile.GetValue("Inventory", "WeaponEqpSlot"))
262     UserList(UserIndex).Invent.HerramientaEqpSlot = CByte(UserFile.GetValue("Inventory", "HerramientaEqpSlot"))
264     UserList(UserIndex).Invent.ArmourEqpSlot = CByte(UserFile.GetValue("Inventory", "ArmourEqpSlot"))
266     UserList(UserIndex).Invent.EscudoEqpSlot = CByte(UserFile.GetValue("Inventory", "EscudoEqpSlot"))
268     UserList(UserIndex).Invent.CascoEqpSlot = CByte(UserFile.GetValue("Inventory", "CascoEqpSlot"))
270     UserList(UserIndex).Invent.BarcoSlot = CByte(UserFile.GetValue("Inventory", "BarcoSlot"))
272     UserList(UserIndex).Invent.MonturaSlot = CByte(UserFile.GetValue("Inventory", "MonturaSlot"))
274     UserList(UserIndex).Invent.MunicionEqpSlot = CByte(UserFile.GetValue("Inventory", "MunicionSlot"))
276     UserList(UserIndex).Invent.DañoMagicoEqpSlot = CByte(UserFile.GetValue("Inventory", "DMSlot"))
278     UserList(UserIndex).Invent.ResistenciaEqpSlot = CByte(UserFile.GetValue("Inventory", "RMSlot"))
280     UserList(UserIndex).Invent.MagicoSlot = CByte(UserFile.GetValue("Inventory", "MagicoSlot"))
282     UserList(UserIndex).Invent.NudilloSlot = CByte(UserFile.GetValue("Inventory", "NudilloEqpSlot"))

284     UserList(UserIndex).ChatCombate = CByte(UserFile.GetValue("BINDKEYS", "ChatCombate"))
286     UserList(UserIndex).ChatGlobal = CByte(UserFile.GetValue("BINDKEYS", "ChatGlobal"))

288     UserList(UserIndex).Correo.CantCorreo = CByte(UserFile.GetValue("CORREO", "CantCorreo"))
290     UserList(UserIndex).Correo.NoLeidos = CByte(UserFile.GetValue("CORREO", "NoLeidos"))

292     For LoopC = 1 To UserList(UserIndex).Correo.CantCorreo
294         UserList(UserIndex).Correo.Mensaje(LoopC).Remitente = UserFile.GetValue("CORREO", "REMITENTE" & LoopC)
296         UserList(UserIndex).Correo.Mensaje(LoopC).Mensaje = UserFile.GetValue("CORREO", "MENSAJE" & LoopC)
298         UserList(UserIndex).Correo.Mensaje(LoopC).Item = UserFile.GetValue("CORREO", "Item" & LoopC)
300         UserList(UserIndex).Correo.Mensaje(LoopC).ItemCount = CByte(UserFile.GetValue("CORREO", "ItemCount" & LoopC))
302         UserList(UserIndex).Correo.Mensaje(LoopC).Fecha = UserFile.GetValue("CORREO", "DATE" & LoopC)
304         UserList(UserIndex).Correo.Mensaje(LoopC).Leido = CByte(UserFile.GetValue("CORREO", "LEIDO" & LoopC))
306     Next LoopC

        'Logros Ladder
308     UserList(UserIndex).UserLogros = UserFile.GetValue("LOGROS", "UserLogros")
310     UserList(UserIndex).NPcLogros = UserFile.GetValue("LOGROS", "NPcLogros")
312     UserList(UserIndex).LevelLogros = UserFile.GetValue("LOGROS", "LevelLogros")
        'Logros Ladder

314     ln = UserFile.GetValue("Guild", "GUILDINDEX")

316     If IsNumeric(ln) Then
318         UserList(UserIndex).GuildIndex = CInt(ln)
        Else
320         UserList(UserIndex).GuildIndex = 0

        End If

        
        Exit Sub

LoadUserInit_Err:
322     Call RegistrarError(Err.Number, Err.Description, "ES.LoadUserInit", Erl)
324     Resume Next
        
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
    
102     If RunningInVB() Then
104         NumMaps = 869
        Else
106         NumMaps = CountFiles(MapPath, "*.csm")
108         NumMaps = NumMaps - 1
        End If
110     Call InitAreas
    
112     frmCargando.cargar.min = 0
114     frmCargando.cargar.max = NumMaps
116     frmCargando.cargar.Value = 0
118     frmCargando.ToMapLbl.Visible = True
    
120     ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

122     ReDim MapInfo(1 To NumMaps) As MapInfo
      
124     For Map = 1 To NumMaps
126         frmCargando.ToMapLbl = Map & "/" & NumMaps

128         Call CargarMapaFormatoCSM(Map, App.Path & "\WorldBackUp\Mapa" & Map & ".csm")

130         frmCargando.cargar.Value = frmCargando.cargar.Value + 1

132         DoEvents
134     Next Map

136     Call generateMatrix(MATRIX_INITIAL_MAP)

138     frmCargando.ToMapLbl.Visible = False

        Exit Sub

CargarBackUp_Err:
140     Call RegistrarError(Err.Number, Err.Description, "ES.CargarBackUp", Erl)
142     Resume Next
        
End Sub

Sub LoadMapData()

100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas..."

        Dim Map       As Integer
        Dim TempInt   As Integer
        Dim npcfile   As String

        On Error GoTo man
    
102     If RunningInVB() Then
104         NumMaps = 700
        Else
106         NumMaps = CountFiles(MapPath, "*.csm") - 1
        End If

108     Call InitAreas
    
110     frmCargando.cargar.min = 0
112     frmCargando.cargar.max = NumMaps
114     frmCargando.cargar.Value = 0
116     frmCargando.ToMapLbl.Visible = True

118     ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

120     ReDim MapInfo(1 To NumMaps) As MapInfo

122     For Map = 1 To NumMaps
    
124         frmCargando.ToMapLbl = Map & "/" & NumMaps

126         Call CargarMapaFormatoCSM(Map, MapPath & "Mapa" & Map & ".csm")

128         frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        
130         DoEvents
        
132     Next Map
    
134     Call generateMatrix(MATRIX_INITIAL_MAP)

136     frmCargando.ToMapLbl.Visible = False
    
        Exit Sub

man:
138     Call MsgBox("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
140     Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)

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
266                         MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.amount = ObjData(Objetos(i).ObjIndex).VidaUtil
268                         MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.data = &H7FFFFFFF ' Ultimo uso = Max Long

270                     Case Else
272                         MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.amount = Objetos(i).ObjAmmount

                    End Select

274             Next i

            End If

276         If .NumeroNPCs > 0 Then
278             ReDim NPCs(1 To .NumeroNPCs)
280             Get #fh, , NPCs

                Dim NumNpc As Integer, NpcIndex As Integer
                 
282             For i = 1 To .NumeroNPCs
284                 NumNpc = NPCs(i).NpcIndex
                    
286                 If NumNpc > 0 Then
288                     npcfile = DatPath & "NPCs.dat"
290                     NpcIndex = OpenNPC(NumNpc)
                        
                        ' Jopi: Evitamos meter NPCs en el mapa que no existen o estan mal dateados.
                        If NpcIndex > 0 Then
                        
292                         MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex = NpcIndex
    
294                         NpcList(NpcIndex).Pos.Map = Map
296                         NpcList(NpcIndex).Pos.X = NPCs(i).X
298                         NpcList(NpcIndex).Pos.Y = NPCs(i).Y
    
                            ' WyroX: guardo siempre la pos original... puede sernos útil ;)
300                         NpcList(NpcIndex).Orig = NpcList(NpcIndex).Pos
    
302                         If LenB(NpcList(NpcIndex).Name) = 0 Then

304                             MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex = 0

                            Else
    
306                             Call MakeNPCChar(True, 0, NpcIndex, Map, NPCs(i).X, NPCs(i).Y)
                            
                            End If
                           
                        Else
                            
                            ' Lo guardo en los logs + aparece en el Debug.Print
                            Call RegistrarError(404, "NPC no existe en los .DAT's o está mal dateado. Posicion: " & Map & "-" & NPCs(i).X & "-" & NPCs(i).Y, "ES.CargarMapaFormatoCSM")
                            
                        End If
                    End If

308             Next i
                
            End If
            
310         If .NumeroTE > 0 Then
312             ReDim TEs(1 To .NumeroTE)
314             Get #fh, , TEs

316             For i = 1 To .NumeroTE
318                 MapData(Map, TEs(i).X, TEs(i).Y).TileExit.Map = TEs(i).DestM
320                 MapData(Map, TEs(i).X, TEs(i).Y).TileExit.X = TEs(i).DestX
322                 MapData(Map, TEs(i).X, TEs(i).Y).TileExit.Y = TEs(i).DestY
324             Next i

            End If
        
        End With

326     Close fh

        ' WyroX: Nuevo sistema de restricciones
328     If Not IsNumeric(MapDat.restrict_mode) Then
            ' Solo se usaba el "NEWBIE"
330         If UCase$(MapDat.restrict_mode) = "NEWBIE" Then
332             MapDat.restrict_mode = "1"
            Else
334             MapDat.restrict_mode = "0"
            End If
        End If
    
336     MapInfo(Map).map_name = MapDat.map_name
338     MapInfo(Map).ambient = MapDat.ambient
340     MapInfo(Map).backup_mode = MapDat.backup_mode
342     MapInfo(Map).base_light = MapDat.base_light
344     MapInfo(Map).Newbie = (val(MapDat.restrict_mode) And 1) <> 0
346     MapInfo(Map).SinMagia = (val(MapDat.restrict_mode) And 2) <> 0
348     MapInfo(Map).NoPKs = (val(MapDat.restrict_mode) And 4) <> 0
350     MapInfo(Map).NoCiudadanos = (val(MapDat.restrict_mode) And 8) <> 0
352     MapInfo(Map).SinInviOcul = (val(MapDat.restrict_mode) And 16) <> 0
354     MapInfo(Map).SoloClanes = (val(MapDat.restrict_mode) And 32) <> 0
356     MapInfo(Map).ResuCiudad = val(GetVar(DatPath & "Map.dat", "RESUCIUDAD", Map)) <> 0
358     MapInfo(Map).letter_grh = MapDat.letter_grh
360     MapInfo(Map).lluvia = MapDat.lluvia
362     MapInfo(Map).music_numberHi = MapDat.music_numberHi
364     MapInfo(Map).music_numberLow = MapDat.music_numberLow
366     MapInfo(Map).niebla = MapDat.niebla
368     MapInfo(Map).Nieve = MapDat.Nieve
370     MapInfo(Map).MinLevel = MapDat.level And &HFF
372     MapInfo(Map).MaxLevel = (MapDat.level And &HFF00) / &H100
    
374     MapInfo(Map).Seguro = MapDat.Seguro

376     MapInfo(Map).terrain = MapDat.terrain
378     MapInfo(Map).zone = MapDat.zone

380     If LenB(MapDat.Salida) <> 0 Then
            Dim Fields() As String
382         Fields = Split(MapDat.Salida, "-")
384         MapInfo(Map).Salida.Map = val(Fields(0))
386         MapInfo(Map).Salida.X = val(Fields(1))
388         MapInfo(Map).Salida.Y = val(Fields(2))
        End If
 
        Exit Sub

ErrorHandler:
390     Close fh
392     Call RegistrarError(Err.Number, Err.Description, "ES.CargarMapaFormatoCSM", Erl)
    
End Sub

Sub LoadSini()
        On Error GoTo LoadSini_Err

        Dim Lector   As clsIniManager

        Dim Temporal As Long
    
100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."
    
102     Set Lector = New clsIniManager
104     Call Lector.Initialize(IniPath & "Server.ini")
    
        'Misc
106     BootDelBackUp = val(Lector.GetValue("INIT", "IniciarDesdeBackUp"))
108     Md5Cliente = Lector.GetValue("CHECKSUM", "Cliente")
    
        'Directorios
110     DatPath = Lector.GetValue("DIRECTORIOS", "DatPath")
112     MapPath = Lector.GetValue("DIRECTORIOS", "MapPath")
114     CharPath = Lector.GetValue("DIRECTORIOS", "CharPath")
116     DeletePath = Lector.GetValue("DIRECTORIOS", "DeletePath")
118     CuentasPath = Lector.GetValue("DIRECTORIOS", "CuentasPath")
120     DeleteCuentasPath = Lector.GetValue("DIRECTORIOS", "DeleteCuentasPath")
        'Directorios
    
122     Puerto = val(Lector.GetValue("INIT", "StartPort"))
124     LastSockListen = val(Lector.GetValue("INIT", "LastSockListen"))
126     HideMe = val(Lector.GetValue("INIT", "Hide"))
128     MaxConexionesIP = val(Lector.GetValue("INIT", "MaxConexionesIP"))
130     MaxUsersPorCuenta = val(Lector.GetValue("INIT", "MaxUsersPorCuenta"))
132     IdleLimit = val(Lector.GetValue("INIT", "IdleLimit"))
        'Lee la version correcta del cliente
134     ULTIMAVERSION = Lector.GetValue("INIT", "Version")
    
136     PuedeCrearPersonajes = val(Lector.GetValue("INIT", "PuedeCrearPersonajes"))
138     ServerSoloGMs = val(Lector.GetValue("init", "ServerSoloGMs"))
    
140     DiceMinimum = val(Lector.GetValue("INIT", "MinDados"))
142     DiceMaximum = val(Lector.GetValue("INIT", "MaxDados"))
    
144     EnTesting = val(Lector.GetValue("INIT", "Testing"))
        
        'Ressurect pos
146     ResPos.Map = val(ReadField(1, Lector.GetValue("INIT", "ResPos"), 45))
148     ResPos.X = val(ReadField(2, Lector.GetValue("INIT", "ResPos"), 45))
150     ResPos.Y = val(ReadField(3, Lector.GetValue("INIT", "ResPos"), 45))
      
152     If Not Database_Enabled Then
154         RecordUsuarios = val(Lector.GetValue("INIT", "Record"))
        End If
      
        'Max users
156     Temporal = val(Lector.GetValue("INIT", "MaxUsers"))

158     If MaxUsers = 0 Then
            #If DEBUGGING Then
160              MaxUsers = 3000
            #Else
                MaxUsers = Temporal
            #End If
162         ReDim UserList(1 To MaxUsers) As user

        End If

164     Call CargarCiudades
166     Call ConsultaPopular.LoadData

168     Set Lector = Nothing

        Exit Sub

LoadSini_Err:
170     Set Lector = Nothing
172     Call RegistrarError(Err.Number, Err.Description, "ES.LoadSini", Erl)
174     Resume Next
        
End Sub

Public Sub LoadDatabaseIniFile()
    On Error GoTo LoadDatabaseIniFile_Err

        Dim Lector As clsIniManager
    
100     If frmMain.Visible Then frmMain.txStatus.Caption = "Leyendo credenciales de la DB."
    
102     Set Lector = New clsIniManager
104     Call Lector.Initialize(IniPath & "Database.ini")

106     Database_Enabled = True
108     Database_DataSource = Lector.GetValue("DATABASE", "DSN")
110     Database_Host = Lector.GetValue("DATABASE", "Host")
112     Database_Name = Lector.GetValue("DATABASE", "Name")
114     Database_Username = Lector.GetValue("DATABASE", "Username")
116     Database_Password = Lector.GetValue("DATABASE", "Password")

        Exit Sub

LoadDatabaseIniFile_Err:
118     Set Lector = Nothing
120     Call RegistrarError(Err.Number, Err.Description, "ES.LoadDatabaseIniFile", Erl)
122     Resume Next
End Sub

Sub CargarCiudades()
        
        On Error GoTo CargarCiudades_Err
    
        

        Dim Lector As clsIniManager
100     Set Lector = New clsIniManager
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
    
214     With CityArkhein
216         .Map = val(Lector.GetValue("Arkhein", "Mapa"))
218         .X = val(Lector.GetValue("Arkhein", "X"))
220         .Y = val(Lector.GetValue("Arkhein", "Y"))
222         .MapaViaje = val(Lector.GetValue("Arkhein", "MapaViaje"))
224         .ViajeX = val(Lector.GetValue("Arkhein", "ViajeX"))
226         .ViajeY = val(Lector.GetValue("Arkhein", "ViajeY"))
228         .MapaResu = val(Lector.GetValue("Arkhein", "MapaResu"))
230         .ResuX = val(Lector.GetValue("Arkhein", "ResuX"))
232         .ResuY = val(Lector.GetValue("Arkhein", "ResuY"))
234         .NecesitaNave = val(Lector.GetValue("Arkhein", "NecesitaNave"))
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
    
284     Arkhein.Map = CityArkhein.Map
286     Arkhein.X = CityArkhein.X
288     Arkhein.Y = CityArkhein.Y
    
        'Esto es para el /HOGAR
290     Ciudades(eCiudad.cNix) = Nix
292     Ciudades(eCiudad.cUllathorpe) = Ullathorpe
294     Ciudades(eCiudad.cBanderbill) = Banderbill
296     Ciudades(eCiudad.cLindos) = Lindos
298     Ciudades(eCiudad.cArghal) = Arghal
300     Ciudades(eCiudad.cArkhein) = Arkhein
    
        
        Exit Sub

CargarCiudades_Err:
302     Call RegistrarError(Err.Number, Err.Description, "ES.CargarCiudades", Erl)

        
End Sub

Sub LoadIntervalos()
        
        On Error GoTo LoadIntervalos_Err
        

        Dim Lector As clsIniManager
100     Set Lector = New clsIniManager
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
    
192     IntervaloTrabajarExtraer = val(Lector.GetValue("INTERVALOS", "IntervaloTrabajarExtraer"))
194     FrmInterv.txtTrabajoExtraer.Text = IntervaloTrabajarExtraer

196     IntervaloTrabajarConstruir = val(Lector.GetValue("INTERVALOS", "IntervaloTrabajarConstruir"))
198     FrmInterv.txtTrabajoConstruir.Text = IntervaloTrabajarConstruir
    
200     IntervaloUserPuedeAtacar = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeAtacar"))
202     FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar
    
        'TODO : Agregar estos intervalos al form!!!
204     IntervaloMagiaGolpe = val(Lector.GetValue("INTERVALOS", "IntervaloMagiaGolpe"))
206     IntervaloGolpeMagia = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeMagia"))
    
        'frmMain.tLluvia.Interval = val(Lector.GetValue("INTERVALOS", "IntervaloPerdidaStaminaLluvia"))
        'FrmInterv.txtIntervaloPerdidaStaminaLluvia.Text = frmMain.tLluvia.Interval
    
208     MinutosWs = val(Lector.GetValue("INTERVALOS", "IntervaloWS"))

210     If MinutosWs < 1 Then MinutosWs = 10
    
212     IntervaloCerrarConexion = val(Lector.GetValue("INTERVALOS", "IntervaloCerrarConexion"))
214     IntervaloUserPuedeUsarU = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeUsarU"))
216     IntervaloUserPuedeUsarClic = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeUsarClic"))
218     IntervaloFlechasCazadores = val(Lector.GetValue("INTERVALOS", "IntervaloFlechasCazadores"))
220     IntervaloGolpeUsar = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeUsar"))

222     IntervaloOculto = val(Lector.GetValue("INTERVALOS", "IntervaloOculto"))

224     MargenDeIntervaloPorPing = val(Lector.GetValue("INTERVALOS", "MargenDeIntervaloPorPing"))
    
226     IntervaloPuedeSerAtacado = val(Lector.GetValue("INTERVALOS", "IntervaloPuedeSerAtacado"))

228     IntervaloGuardarUsuarios = val(Lector.GetValue("INTERVALOS", "IntervaloGuardarUsuarios"))
        
230     LimiteGuardarUsuarios = val(Lector.GetValue("INTERVALOS", "LimiteGuardarUsuarios"))

232     IntervaloTimerGuardarUsuarios = val(Lector.GetValue("INTERVALOS", "IntervaloTimerGuardarUsuarios"))

234     IntervaloMensajeGlobal = val(Lector.GetValue("INTERVALOS", "IntervaloMensajeGlobal"))
    
        '&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    
236     Set Lector = Nothing

        
        Exit Sub

LoadIntervalos_Err:
238     Call RegistrarError(Err.Number, Err.Description, "ES.LoadIntervalos", Erl)
240     Resume Next
        
End Sub

Sub LoadConfiguraciones()
        
        On Error GoTo LoadConfiguraciones_Err
        
        Dim Leer As clsIniManager
        Set Leer = New clsIniManager

102     Call Leer.Initialize(IniPath & "Configuracion.ini")

104     ExpMult = val(Leer.GetValue("CONFIGURACIONES", "ExpMult"))
106     OroMult = val(Leer.GetValue("CONFIGURACIONES", "OroMult"))
108     DropMult = val(Leer.GetValue("DROPEO", "DropMult"))
110     DropActive = val(Leer.GetValue("DROPEO", "DropActive"))
112     RecoleccionMult = val(Leer.GetValue("CONFIGURACIONES", "RecoleccionMult"))

114     TimerLimpiarObjetos = val(Leer.GetValue("CONFIGURACIONES", "TimerLimpiarObjetos"))
116     OroPorNivel = val(Leer.GetValue("CONFIGURACIONES", "OroPorNivel"))

118     DuracionDia = val(Leer.GetValue("CONFIGURACIONES", "DuracionDia")) * 60 * 1000 ' De minutos a milisegundos

120     CostoPerdonPorCiudadano = val(Leer.GetValue("CONFIGURACIONES", "CostoPerdonPorCiudadano"))

122     MaximoSpeedHack = val(Leer.GetValue("ANTICHEAT", "MaximoSpeedHack"))

124     frmMain.lblLimpieza.Caption = "Limpieza de objetos cada: " & TimerLimpiarObjetos & " minutos."

126     Set Leer = Nothing

128     Call CargarEventos
130     Call CargarInfoRetos
132     Call CargarInfoEventos
        Call CargarMapasEspeciales

        Exit Sub

LoadConfiguraciones_Err:
134     Call RegistrarError(Err.Number, Err.Description, "ES.LoadConfiguraciones", Erl)
136     Resume Next
        
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
126             .Char.Head = iCabezaMuerto
128             .Char.WeaponAnim = NingunArma
130             .Char.ShieldAnim = NingunEscudo
132             .Char.CascoAnim = NingunCasco
134             .Char.Heading = eHeading.SOUTH
            End If
        
            'Obtiene el indice-objeto del arma
136         If .Invent.WeaponEqpSlot > 0 Then
138             If .Invent.Object(.Invent.WeaponEqpSlot).ObjIndex > 0 Then
140                 .Invent.WeaponEqpObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).ObjIndex
142                 If .flags.Muerto = 0 Then
144                     .Char.Arma_Aura = ObjData(.Invent.WeaponEqpObjIndex).CreaGRH
                    End If
                Else
146                 .Invent.WeaponEqpSlot = 0
                End If
            End If

            'Obtiene el indice-objeto del armadura
148         If .Invent.ArmourEqpSlot > 0 Then
150             If .Invent.Object(.Invent.ArmourEqpSlot).ObjIndex > 0 Then
152                 .Invent.ArmourEqpObjIndex = .Invent.Object(.Invent.ArmourEqpSlot).ObjIndex
154                 If .flags.Muerto = 0 Then
156                     .Char.Body_Aura = ObjData(.Invent.ArmourEqpObjIndex).CreaGRH
                    End If
                Else
158                 .Invent.ArmourEqpSlot = 0
                End If
160             .flags.Desnudo = 0
            Else
162             .flags.Desnudo = 1
            End If

            'Obtiene el indice-objeto del escudo
164         If .Invent.EscudoEqpSlot > 0 Then
166             If .Invent.Object(.Invent.EscudoEqpSlot).ObjIndex > 0 Then
168                 .Invent.EscudoEqpObjIndex = .Invent.Object(.Invent.EscudoEqpSlot).ObjIndex
170                 If .flags.Muerto = 0 Then
172                     .Char.Escudo_Aura = ObjData(.Invent.EscudoEqpObjIndex).CreaGRH
                    End If
                Else
174                 .Invent.EscudoEqpSlot = 0
                End If
            End If
        
            'Obtiene el indice-objeto del casco
176         If .Invent.CascoEqpSlot > 0 Then
178             If .Invent.Object(.Invent.CascoEqpSlot).ObjIndex > 0 Then
180                 .Invent.CascoEqpObjIndex = .Invent.Object(.Invent.CascoEqpSlot).ObjIndex
182                 If .flags.Muerto = 0 Then
184                     .Char.Head_Aura = ObjData(.Invent.CascoEqpObjIndex).CreaGRH
                    End If
                Else
186                 .Invent.CascoEqpSlot = 0
                End If
            End If

            'Obtiene el indice-objeto barco
188         If .Invent.BarcoSlot > 0 Then
190             If .Invent.Object(.Invent.BarcoSlot).ObjIndex > 0 Then
192                  .Invent.BarcoObjIndex = .Invent.Object(.Invent.BarcoSlot).ObjIndex
                Else
194                 .Invent.BarcoSlot = 0
                End If
            End If

            'Obtiene el indice-objeto municion
196         If .Invent.MunicionEqpSlot > 0 Then
198             If .Invent.Object(.Invent.MunicionEqpSlot).ObjIndex > 0 Then
200                 .Invent.MunicionEqpObjIndex = .Invent.Object(.Invent.MunicionEqpSlot).ObjIndex
                Else
202                 .Invent.MunicionEqpSlot = 0
                End If
            End If

            ' DM
204         If .Invent.DañoMagicoEqpSlot > 0 Then
206             If .Invent.Object(.Invent.DañoMagicoEqpSlot).ObjIndex > 0 Then
208                 .Invent.DañoMagicoEqpObjIndex = .Invent.Object(.Invent.DañoMagicoEqpSlot).ObjIndex
210                 If .flags.Muerto = 0 Then
212                     .Char.DM_Aura = ObjData(.Invent.DañoMagicoEqpObjIndex).CreaGRH
                    End If
                Else
214                  .Invent.DañoMagicoEqpSlot = 0
                End If
            End If
            
            ' RM
216         If .Invent.ResistenciaEqpSlot > 0 Then
218             If .Invent.Object(.Invent.ResistenciaEqpSlot).ObjIndex > 0 Then
220                .Invent.ResistenciaEqpObjIndex = .Invent.Object(.Invent.ResistenciaEqpSlot).ObjIndex
222                 If .flags.Muerto = 0 Then
224                     .Char.RM_Aura = ObjData(.Invent.ResistenciaEqpObjIndex).CreaGRH
                    End If
                Else
226                 .Invent.ResistenciaEqpSlot = 0
                End If
            End If

228         If .Invent.MonturaSlot > 0 Then
230             If .Invent.Object(.Invent.MonturaSlot).ObjIndex > 0 Then
232             .Invent.MonturaObjIndex = .Invent.Object(.Invent.MonturaSlot).ObjIndex
                Else
234                 .Invent.MonturaSlot = 0
                End If
            End If
        
236         If .Invent.HerramientaEqpSlot > 0 Then
238             If .Invent.Object(.Invent.HerramientaEqpSlot).ObjIndex Then
240                 .Invent.HerramientaEqpObjIndex = .Invent.Object(.Invent.HerramientaEqpSlot).ObjIndex
                Else
242                 .Invent.HerramientaEqpSlot = 0
                End If
            End If
        
244         If .Invent.NudilloSlot > 0 Then
246             If .Invent.Object(.Invent.NudilloSlot).ObjIndex > 0 Then
248                 .Invent.NudilloObjIndex = .Invent.Object(.Invent.NudilloSlot).ObjIndex
250                 If .flags.Muerto = 0 Then
252                     .Char.Arma_Aura = ObjData(.Invent.NudilloObjIndex).CreaGRH
                    End If
                Else
254                 .Invent.NudilloSlot = 0
                End If
            End If
        
256         If .Invent.MagicoSlot > 0 Then
258             If .Invent.Object(.Invent.MagicoSlot).ObjIndex Then
260                 .Invent.MagicoObjIndex = .Invent.Object(.Invent.MagicoSlot).ObjIndex

262                 If .flags.Muerto = 0 Then
264                     .Char.Otra_Aura = ObjData(.Invent.MagicoObjIndex).CreaGRH
                    End If
                Else
266                 .Invent.MagicoSlot = 0
                End If
            End If

        End With

        Exit Sub

ErrorHandler:
268     Call RegistrarError(Err.Number, Err.Description & " UserName: " & UserList(UserIndex).Name, "ES.LoadUser", Erl)
270     Resume Next
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer, Optional ByVal Logout As Boolean = False)

        On Error GoTo SaveUser_Err
    
        #If DEBUGGING = 1 Then
100         Call GetElapsedTime
        #End If
    
102     Call SaveUserDatabase(UserIndex, Logout)
    
104     UserList(UserIndex).Counters.LastSave = GetTickCount
    
        #If DEBUGGING = 1 Then
106         Call LogPerformance("Guardado de Personaje - Tiempo transcurrido: " & Round(GetElapsedTime(), 1) & " ms")
        #End If
    
        Exit Sub

SaveUser_Err:
108     Call RegistrarError(Err.Number, Err.Description, "ES.SaveUser", Erl)

110     Resume Next

End Sub

Sub LoadUserBinary(ByVal UserIndex As Integer)
        
        On Error GoTo LoadUserBinary_Err
        

        'Cargamos el personaje
        Dim Leer As New clsIniManager
100     Call Leer.Initialize(CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
    
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


Sub SaveNewUser(ByVal UserIndex As Integer)
    On Error GoTo SaveNewUser_Err
            
100 Call SaveNewUserDatabase(UserIndex)
    
    Exit Sub

SaveNewUser_Err:
102     Call RegistrarError(Err.Number, Err.Description, "ES.SaveNewUser", Erl)
104     Resume Next
        
End Sub

Sub SetUserLogged(ByVal UserIndex As Integer)
        
        On Error GoTo SetUserLogged_Err
        

100     If Database_Enabled Then
102         Call SetUserLoggedDatabase(UserList(UserIndex).ID, UserList(UserIndex).AccountId)
        Else
104         Call WriteVar(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", "INIT", "Logged", 1)
106         Call WriteVar(CuentasPath & UCase$(UserList(UserIndex).Cuenta) & ".act", "INIT", "LOGEADA", 1)

        End If

        
        Exit Sub

SetUserLogged_Err:
108     Call RegistrarError(Err.Number, Err.Description, "ES.SetUserLogged", Erl)
110     Resume Next
        
End Sub

Function Status(ByVal UserIndex As Integer) As e_Facciones
        
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

100     NpcNumero = NpcList(NpcIndex).Numero

        'If NpcNumero > 499 Then
        '    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
        'Else
102     npcfile = DatPath & "bkNPCs.dat"
        'End If

        'General
104     Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", NpcList(NpcIndex).Name)
106     Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", NpcList(NpcIndex).Desc)
108     Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(NpcList(NpcIndex).Char.Head))
110     Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(NpcList(NpcIndex).Char.Body))
112     Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(NpcList(NpcIndex).Char.Heading))
114     Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(NpcList(NpcIndex).Movement))
116     Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(NpcList(NpcIndex).Attackable))
118     Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(NpcList(NpcIndex).Comercia))
120     Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(NpcList(NpcIndex).TipoItems))
122     Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(NpcList(NpcIndex).Hostile))
124     Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(NpcList(NpcIndex).GiveEXP))
126     Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", val(NpcList(NpcIndex).GiveGLD))
128     Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(NpcList(NpcIndex).Hostile))
130     Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(NpcList(NpcIndex).InvReSpawn))
132     Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(NpcList(NpcIndex).NPCtype))

        'Stats
134     Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(NpcList(NpcIndex).flags.AIAlineacion))
136     Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(NpcList(NpcIndex).Stats.def))
138     Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(NpcList(NpcIndex).Stats.MaxHit))
140     Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(NpcList(NpcIndex).Stats.MaxHp))
142     Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(NpcList(NpcIndex).Stats.MinHIT))
144     Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(NpcList(NpcIndex).Stats.MinHp))
146     Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(NpcList(NpcIndex).Stats.UsuariosMatados)) 'Que es ESTO?!!

        'Flags
148     Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(NpcList(NpcIndex).flags.Respawn))
150     Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(NpcList(NpcIndex).flags.backup))
152     Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(NpcList(NpcIndex).flags.Domable))

        'Inventario
154     Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(NpcList(NpcIndex).Invent.NroItems))

156     If NpcList(NpcIndex).Invent.NroItems > 0 Then

158         For LoopC = 1 To MAX_INVENTORY_SLOTS
160             Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & LoopC, NpcList(NpcIndex).Invent.Object(LoopC).ObjIndex & "-" & NpcList(NpcIndex).Invent.Object(LoopC).amount)
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

104     NpcList(NpcIndex).Numero = NpcNumber
106     NpcList(NpcIndex).Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
108     NpcList(NpcIndex).Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
110     NpcList(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
112     NpcList(NpcIndex).NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))

114     NpcList(NpcIndex).Char.Body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
116     NpcList(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
118     NpcList(NpcIndex).Char.Heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))

120     NpcList(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
122     NpcList(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
124     NpcList(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
126     NpcList(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))

128     NpcList(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))

130     NpcList(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))

132     NpcList(NpcIndex).Stats.MaxHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
134     NpcList(NpcIndex).Stats.MinHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
136     NpcList(NpcIndex).Stats.MaxHit = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
138     NpcList(NpcIndex).Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
140     NpcList(NpcIndex).Stats.def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
142     NpcList(NpcIndex).flags.AIAlineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))

        Dim LoopC As Integer

        Dim ln    As String

144     NpcList(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))

146     If NpcList(NpcIndex).Invent.NroItems > 0 Then

148         For LoopC = 1 To MAX_INVENTORY_SLOTS
150             ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
152             NpcList(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
154             NpcList(NpcIndex).Invent.Object(LoopC).amount = val(ReadField(2, ln, 45))
       
156         Next LoopC

        Else

158         For LoopC = 1 To MAX_INVENTORY_SLOTS
160             NpcList(NpcIndex).Invent.Object(LoopC).ObjIndex = 0
162             NpcList(NpcIndex).Invent.Object(LoopC).amount = 0
164         Next LoopC

        End If

166     NpcList(NpcIndex).flags.NPCActive = True
168     NpcList(NpcIndex).flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
170     NpcList(NpcIndex).flags.backup = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
172     NpcList(NpcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
174     NpcList(NpcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))

        'Tipo de items con los que comercia
176     NpcList(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

        
        Exit Sub

CargarNpcBackUp_Err:
178     Call RegistrarError(Err.Number, Err.Description, "ES.CargarNpcBackUp", Erl)
180     Resume Next
        
End Sub

Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal Motivo As String)
        
        On Error GoTo LogBan_Err
        

100     Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
102     Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "Reason", Motivo)

        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim mifile As Integer

104     mifile = FreeFile
106     Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
108     Print #mifile, UserList(BannedIndex).Name
110     Close #mifile

        
        Exit Sub

LogBan_Err:
112     Call RegistrarError(Err.Number, Err.Description, "ES.LogBan", Erl)
114     Resume Next
        
End Sub

Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal Motivo As String)
        
        On Error GoTo LogBanFromName_Err
        

100     Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).Name)
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

        Dim IniFile As clsIniManager

106     Set IniFile = New clsIniManager
    
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

        Dim IniFile As clsIniManager

106     Set IniFile = New clsIniManager
    
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
134             Peces(i).amount = nivel
            Next

            ' Los ordeno segun nivel de caña (quick sort)
136         Call QuickSortPeces(1, Count)

            ' Sumo los pesos
138         For i = 1 To Count
140             For j = Peces(i).amount To MaxLvlCania
142                 PesoPeces(j) = PesoPeces(j) + Peces(i).data
144             Next j

146             Peces(i).data = PesoPeces(Peces(i).amount)
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
Private Sub QuickSortPeces(ByVal first As Long, ByVal last As Long)
        
        On Error GoTo QuickSortPeces_Err
        

        Dim low      As Long, high As Long

        Dim MidValue As String

        Dim aux      As obj
    
100     low = first
102     high = last
104     MidValue = Peces((first + last) \ 2).amount
    
        Do

106         While Peces(low).amount < MidValue

108             low = low + 1
            Wend

110         While Peces(high).amount > MidValue

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
    
128     If first < high Then QuickSortPeces first, high
130     If low < last Then QuickSortPeces low, last

        
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

Public Sub LoadRangosFaccion()
            On Error GoTo LoadRangosFaccion_Err

100         If Not FileExist(DatPath & "rangos_faccion.dat", vbArchive) Then
102             ReDim RangosFaccion(0) As tRangoFaccion
                Exit Sub

            End If

        Dim IniFile As clsIniManager
        Set IniFile = New clsIniManager

106         Call IniFile.Initialize(DatPath & "rangos_faccion.dat")

            Dim i As Byte, rankData() As String

108         MaxRangoFaccion = val(IniFile.GetValue("INIT", "NumRangos"))

110         If MaxRangoFaccion > 0 Then
                ' Los rangos de la Armada se guardan en los indices impar, y los del caos en indices pares.
                ' Luego, para acceder es tan facil como usar el Rango directamente para la Armada, y multiplicar por 2 para el Caos.
112             ReDim RangosFaccion(1 To MaxRangoFaccion * 2) As tRangoFaccion

114             For i = 1 To MaxRangoFaccion
                    '<N>Rango=<NivelRequerido>-<AsesinatosRequeridos>-<Título>
116                 rankData = Split(IniFile.GetValue("ArmadaReal", i & "Rango"), "-", , vbTextCompare)
118                 RangosFaccion(2 * i - 1).rank = i
120                 RangosFaccion(2 * i - 1).Titulo = rankData(2)
122                 RangosFaccion(2 * i - 1).NivelRequerido = val(rankData(0))
124                 RangosFaccion(2 * i - 1).AsesinatosRequeridos = val(rankData(1))

126                 rankData = Split(IniFile.GetValue("LegionCaos", i & "Rango"), "-", , vbTextCompare)
128                 RangosFaccion(2 * i).rank = i
130                 RangosFaccion(2 * i).Titulo = rankData(2)
132                 RangosFaccion(2 * i).NivelRequerido = val(rankData(0))
134                 RangosFaccion(2 * i).AsesinatosRequeridos = val(rankData(1))
136             Next i

            End If

138         Set IniFile = Nothing

            Exit Sub

LoadRangosFaccion_Err:
140         Call RegistrarError(Err.Number, Err.Description, "ES.LoadRangosFaccion", Erl)
142         Resume Next

End Sub


Public Sub LoadRecompensasFaccion()
            On Error GoTo LoadRecompensasFaccion_Err

100         If Not FileExist(DatPath & "recompensas_faccion.dat", vbArchive) Then
102             ReDim RecompensasFaccion(0) As tRecompensaFaccion
                Exit Sub

            End If

        Dim IniFile As clsIniManager
        Set IniFile = New clsIniManager

106         Call IniFile.Initialize(DatPath & "recompensas_faccion.dat")

            Dim cantidadRecompensas As Byte, i As Integer, rank_and_objindex() As String

108         cantidadRecompensas = val(IniFile.GetValue("INIT", "NumRecompensas"))

110         If cantidadRecompensas > 0 Then
112             ReDim RecompensasFaccion(1 To cantidadRecompensas) As tRecompensaFaccion

114             For i = 1 To cantidadRecompensas
116                 rank_and_objindex = Split(IniFile.GetValue("Recompensas", "Recompensa" & i), "-", , vbTextCompare)

118                 RecompensasFaccion(i).rank = val(rank_and_objindex(0))
120                 RecompensasFaccion(i).ObjIndex = val(rank_and_objindex(1))
122             Next i

            End If

124         Set IniFile = Nothing

            Exit Sub

LoadRecompensasFaccion_Err:
126         Call RegistrarError(Err.Number, Err.Description, "ES.LoadRecompensasFaccion", Erl)
128         Resume Next

End Sub


Public Sub LoadUserIntervals(ByVal UserIndex As Integer)
        
        On Error GoTo LoadUserIntervals_Err
        

100     With UserList(UserIndex).Intervals
102         .Arco = IntervaloFlechasCazadores
104         .Caminar = IntervaloCaminar
106         .Golpe = IntervaloUserPuedeAtacar
108         .Magia = IntervaloUserPuedeCastear
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
100     If timer_freq = 0 Then
            Dim temp_time As Currency
102         Call QueryPerformanceFrequency(temp_time)
104         timer_freq = 1000 / temp_time
        End If

106     Call QueryPerformanceCounter(end_time)

108     GetElapsedTime = (end_time - start_time) * timer_freq
    
110     start_time = end_time

End Function
