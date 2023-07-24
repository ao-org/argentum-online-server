Attribute VB_Name = "ES"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
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
'MinHIT
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Const MAX_RANDOM_TELEPORT_IN_MAP = 20
Const FISHING_REQUIRED_PERCENT = 95
Const FISHING_TILES_ON_MAP = 10
Const FISHING_POOL_ID = 3740
Private Type t_Position

    x As Integer
    y As Integer

End Type

'Item type
Private Type t_Item

    objIndex As Integer
    amount As Integer

End Type

Private Type t_WorldPos

    map As Integer
    x As Byte
    y As Byte

End Type

Private Type t_Grh

    GrhIndex As Long
    FrameCounter As Single
    Speed As Single
    Started As Byte
    alpha_blend As Boolean
    angle As Single

End Type

Private Type t_GrhData

    sX As Integer
    sY As Integer
    filenum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Integer
    Frames() As Integer
    Speed As Integer
    mini_map_color As Long

End Type

Private Type t_MapHeader

    NumeroBloqueados As Long
    NumeroLayers(1 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long

End Type

Private Type t_DatosBloqueados

    x As Integer
    y As Integer
    Lados As Byte

End Type

Private Type t_DatosGrh

    x As Integer
    y As Integer
    GrhIndex As Long

End Type

Private Type t_DatosTrigger

    x As Integer
    y As Integer
    trigger As Integer

End Type

Private Type t_DatosLuces

    x As Integer
    y As Integer
    Color As Long
    Rango As Byte

End Type

Private Type t_DatosParticulas

    x As Integer
    y As Integer
    Particula As Long

End Type

Private Type t_DatosNPC

    x As Integer
    y As Integer
    npcIndex As Integer

End Type

Private Type t_DatosObjs

    x As Integer
    y As Integer
    objIndex As Integer
    ObjAmmount As Integer

End Type

Private Type t_DatosTE

    x As Integer
    y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer

End Type

Private Type t_MapSize

    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer

End Type

Private Type t_MapDat

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

Private MapSize As t_MapSize
Private MapDat  As t_MapDat
Private FeatureToggles As Dictionary

Public Sub load_stats()
On Error GoTo error_load_stats
    Dim n As Integer
    Dim strFile As String
    strFile = App.Path & "\logs\recordusers.log"
    Dim str As String
    
    If Not FileExist(strFile) Then
        n = FreeFile()
        Open strFile For Append As #n
        Print #n, "1"
        Close #n
    End If
    
    Debug.Assert FileExist(strFile)
    n = FreeFile()
    Open strFile For Input Shared As n
    Line Input #n, str
    RecordUsuarios = val(str)
    Close #n
    
    Exit Sub
error_load_stats:
     Call TraceError(Err.Number, Err.Description, "ES.load_stats", Erl)

End Sub

Public Sub dump_stats()
On Error GoTo error_dump_stats
    Dim n As Integer
    n = FreeFile()
    Open App.Path & "\logs\numusers.log" For Output Shared As n
    Print #n, NumUsers
    Close #n
    
    n = FreeFile()
    Open App.Path & "\logs\recordusers.log" For Output Shared As n
    Print #n, str(RecordUsuarios)
    Close #n
    
    Exit Sub
error_dump_stats:
     Call TraceError(Err.Number, Err.Description, "ES.error_dump_stats", Erl)
End Sub

Public Sub CargarSpawnList()
        
        On Error GoTo CargarSpawnList_Err
        

        Dim n As Integer, LoopC As Integer

100     n = val(GetVar(DatPath & "npcs.dat", "INIT", "NumNPCs"))
102     ReDim SpawnList(n) As t_CriaturasEntrenador

104     For LoopC = 1 To n

106         SpawnList(LoopC).npcIndex = LoopC
108         SpawnList(LoopC).NpcName = GetVar(DatPath & "npcs.dat", "NPC" & LoopC, "Name")
            SpawnList(LoopC).PuedeInvocar = val(GetVar(DatPath & "npcs.dat", "NPC" & LoopC, "PuedeInvocar")) = 1

110         If Len(SpawnList(LoopC).NpcName) = 0 Then
112             SpawnList(LoopC).NpcName = "Nada"
            End If
            
114     Next LoopC

        
        Exit Sub

CargarSpawnList_Err:
116     Call TraceError(Err.Number, Err.Description, "ES.CargarSpawnList", Erl)

        
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
102     Call TraceError(Err.Number, Err.Description, "ES.EsAdmin", Erl)

        
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
102     Call TraceError(Err.Number, Err.Description, "ES.EsDios", Erl)

        
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
102     Call TraceError(Err.Number, Err.Description, "ES.EsSemiDios", Erl)

        
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
102     Call TraceError(Err.Number, Err.Description, "ES.EsConsejero", Erl)

        
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
102     Call TraceError(Err.Number, Err.Description, "ES.EsRolesMaster", Erl)

        
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
110     Call TraceError(Err.Number, Err.Description, "ES.EsGmChar", Erl)

        
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
        Debug.Assert FileExist(IniPath & "Server.ini")
        
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
190     Call TraceError(Err.Number, Err.Description, "ES.loadAdministrativeUsers", Erl)

        
End Sub


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
116     Call TraceError(Err.Number, Err.Description, "ES.TxtDimension", Erl)

        
End Function

Public Sub CargarForbidenWords()
        
        On Error GoTo CargarForbidenWords_Err

        Dim size As Integer

100     size = TxtDimension(DatPath & "NombresInvalidos.txt")
    
102     If size = 0 Then
104         ReDim ForbidenNames(0)
            Exit Sub

        End If
    
106     ReDim ForbidenNames(1 To size)

        Dim n As Integer, i As Integer

108     n = FreeFile(1)
110     Open DatPath & "NombresInvalidos.txt" For Input As #n
    
112     For i = 1 To UBound(ForbidenNames)
114         Line Input #n, ForbidenNames(i)
            ForbidenNames(i) = LCase$(ForbidenNames(i))
116     Next i
    
118     Close n

        
        Exit Sub

CargarForbidenWords_Err:
120     Call TraceError(Err.Number, Err.Description, "ES.CargarForbidenWords", Erl)

        
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

106     ReDim Hechizos(1 To NumeroHechizos) As t_Hechizo

108     frmCargando.cargar.Min = 0
110     frmCargando.cargar.max = NumeroHechizos
112     frmCargando.cargar.value = 0

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
151         Hechizos(Hechizo).SkillType = val(Leer.GetValue("Hechizo" & Hechizo, "SkillType"))
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
    
172         If val(Leer.GetValue("Hechizo" & Hechizo, "Incinera")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.Incinerate)
            If val(Leer.GetValue("Hechizo" & Hechizo, "RemoveDebuff")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.RemoveDebuff)
            If val(Leer.GetValue("Hechizo" & Hechizo, "StealBuff")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.StealBuff)
174         Hechizos(Hechizo).AutoLanzar = val(Leer.GetValue("Hechizo" & Hechizo, "AutoLanzar"))
175         Hechizos(Hechizo).TargetEffectType = val(Leer.GetValue("Hechizo" & Hechizo, "TargetEffectType"))
176         Hechizos(Hechizo).Cooldown = val(Leer.GetValue("Hechizo" & Hechizo, "CoolDown"))
177         Hechizos(Hechizo).CdEffectId = val(Leer.GetValue("Hechizo" & Hechizo, "CdEffectId"))
178         Hechizos(Hechizo).loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
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
            
            
228         If val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.Invisibility)
230         If val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.Paralize)
232         If val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.Immobilize)
234         If val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.RemoveParalysis)
236         If val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.RemoveDumb)
238         If val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.RemoveInvisibility)
240         If val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.CurePoison)
242         Hechizos(Hechizo).Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
244         If val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.Curse)
246         If val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.RemoveCurse)
250         If val(Leer.GetValue("Hechizo" & Hechizo, "Revivir")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.Resurrect)
252         If val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.Blindness)
254         If val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.Dumb)
255         If val(Leer.GetValue("Hechizo" & Hechizo, "ToggleCleave")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.ToggleCleave)
            If val(Leer.GetValue("Hechizo" & Hechizo, "AdjustStatsWithCaster")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.AdjustStatsWithCaster)
            If val(Leer.GetValue("Hechizo" & Hechizo, "CancelActiveEffect")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.CancelActiveEffect)

256         Hechizos(Hechizo).Invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
258         Hechizos(Hechizo).NumNpc = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
260         Hechizos(Hechizo).cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
262         Hechizos(Hechizo).Mimetiza = val(Leer.GetValue("Hechizo" & Hechizo, "Mimetiza"))
    
264         If val(Leer.GetValue("Hechizo" & Hechizo, "GolpeCertero")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.PreciseHit)
266         Hechizos(Hechizo).MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
268         Hechizos(Hechizo).ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
270         Hechizos(Hechizo).RequiredHP = val(Leer.GetValue("Hechizo" & Hechizo, "RequiredHP"))
    
272         Hechizos(Hechizo).duration = val(Leer.GetValue("Hechizo" & Hechizo, "Duration"))
    
            'Barrin 30/9/03
274         Hechizos(Hechizo).StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
    
276         Hechizos(Hechizo).Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
277         Hechizos(Hechizo).RequireTransform = val(Leer.GetValue("Hechizo" & Hechizo, "RequireTransform"))
278         frmCargando.cargar.value = frmCargando.cargar.value + 1
    
280         Hechizos(Hechizo).NeedStaff = val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
281         Hechizos(Hechizo).StaffAffected = CBool(val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))
282         Hechizos(Hechizo).EotId = val(Leer.GetValue("Hechizo" & Hechizo, "EOTID"))

290         If val(Leer.GetValue("Hechizo" & Hechizo, "RequireArmor")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eArmor)
292         If val(Leer.GetValue("Hechizo" & Hechizo, "RequireShip")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eShip)
294         If val(Leer.GetValue("Hechizo" & Hechizo, "RequireHelm")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eHelm)
296         If val(Leer.GetValue("Hechizo" & Hechizo, "RequireKnucle")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eKnucle)
298         If val(Leer.GetValue("Hechizo" & Hechizo, "RequireMagicItem")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eMagicItem)
300         If val(Leer.GetValue("Hechizo" & Hechizo, "RequireProjectile")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eProjectile)
302         If val(Leer.GetValue("Hechizo" & Hechizo, "RequireShield")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eShield)
304         If val(Leer.GetValue("Hechizo" & Hechizo, "RequireWeapon")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eWeapon)
            If val(Leer.GetValue("Hechizo" & Hechizo, "RequireTargetOnLand")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eRequireTargetOnLand)
            If val(Leer.GetValue("Hechizo" & Hechizo, "RequireTargetOnWater")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eRequireTargetOnWater)
305         Hechizos(Hechizo).RequireWeaponType = val(Leer.GetValue("Hechizo" & Hechizo, "RequireWeaponType"))
            Dim SubeHP As Byte
            SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
            If SubeHP = 1 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.eDoHeal)
            If SubeHP = 2 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.eDoDamage)
306     Next Hechizo

400     Set Leer = Nothing
        Exit Sub

ErrHandler:
402     MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.Description
End Sub

Public Sub LoadEffectOverTime()
On Error GoTo ErrHandler

100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."
        Dim i As Integer
        Dim Leer As New clsIniManager
        Dim EffectCount As Integer

102     Call Leer.Initialize(DatPath & "EffectsOverTime.dat")

        'obtiene el numero de hechizos
104     EffectCount = val(Leer.GetValue("INIT", "EffectCount"))

106     ReDim EffectOverTime(1 To EffectCount) As t_EffectOverTime

108     frmCargando.cargar.Min = 0
110     frmCargando.cargar.max = EffectCount
112     frmCargando.cargar.value = 0

114     For i = 1 To EffectCount
            EffectOverTime(i).Type = val(Leer.GetValue("EOT" & i, "Type"))
            EffectOverTime(i).SubType = val(Leer.GetValue("EOT" & i, "SubType"))
            EffectOverTime(i).TickPowerMin = val(Leer.GetValue("EOT" & i, "TickPowerMin"))
            EffectOverTime(i).TickPowerMax = val(Leer.GetValue("EOT" & i, "TickPowerMax"))
            EffectOverTime(i).Ticks = val(Leer.GetValue("EOT" & i, "Ticks"))
            EffectOverTime(i).TickTime = val(Leer.GetValue("EOT" & i, "TickTime"))
            EffectOverTime(i).TickFX = val(Leer.GetValue("EOT" & i, "TickFX"))
            EffectOverTime(i).TickManaConsumption = val(Leer.GetValue("EOT" & i, "TickManaConsumption"))
            EffectOverTime(i).TickStaminaConsumption = val(Leer.GetValue("EOT" & i, "TickStaminaConsumption"))
            EffectOverTime(i).OnHitFx = val(Leer.GetValue("EOT" & i, "OnHitFx"))
            EffectOverTime(i).OnHitWav = val(Leer.GetValue("EOT" & i, "OnHitWav"))
            EffectOverTime(i).Override = val(Leer.GetValue("EOT" & i, "Override"))
            EffectOverTime(i).Limit = val(Leer.GetValue("EOT" & i, "Limit"))
            EffectOverTime(i).PhysicalDamageReduction = val(Leer.GetValue("EOT" & i, "PhysicalDamageReduction"))
            EffectOverTime(i).MagicDamageReduction = val(Leer.GetValue("EOT" & i, "MagicDamageReduction"))
            EffectOverTime(i).PhysicalDamageDone = val(Leer.GetValue("EOT" & i, "PhysicalDamageDone"))
            EffectOverTime(i).SpeedModifier = val(Leer.GetValue("EOT" & i, "SpeedModifier"))
            EffectOverTime(i).HitModifier = val(Leer.GetValue("EOT" & i, "HitModifier"))
            EffectOverTime(i).EvasionModifier = val(Leer.GetValue("EOT" & i, "EvasionModifier"))
            EffectOverTime(i).MagicDamageDone = val(Leer.GetValue("EOT" & i, "MagicDamageDone"))
            EffectOverTime(i).SelfHealingBonus = val(Leer.GetValue("EOT" & i, "SelfHealingBonus"))
            EffectOverTime(i).MagicHealingBonus = val(Leer.GetValue("EOT" & i, "MagicHealingBonus"))
            EffectOverTime(i).ClientEffectTypeId = val(Leer.GetValue("EOT" & i, "ClientEffectTypeId"))
            EffectOverTime(i).PhysicalLinearBonus = val(Leer.GetValue("EOT" & i, "PhysicalLinearBonus"))
            EffectOverTime(i).DefenseBonus = val(Leer.GetValue("EOT" & i, "DefenseBonus"))
            EffectOverTime(i).BuffType = val(Leer.GetValue("EOT" & i, "BuffType"))
            EffectOverTime(i).Area = val(Leer.GetValue("EOT" & i, "Area"))
            EffectOverTime(i).Aura = Leer.GetValue("EOT" & i, "Aura")
            EffectOverTime(i).ApplyeffectID = val(Leer.GetValue("EOT" & i, "ApplyeffectID"))
            If val(Leer.GetValue("EOT" & i, "AffectedByMagicBonus")) > 0 Then
                Call SetMask(EffectOverTime(i).EffectModifiers, e_ModifierTypes.MagicBonus)
            End If
            If val(Leer.GetValue("EOT" & i, "AffectedByMagicReduction")) > 0 Then
                Call SetMask(EffectOverTime(i).EffectModifiers, e_ModifierTypes.MagicReduction)
            End If
            If val(Leer.GetValue("EOT" & i, "AffectedByPhysicalBonus")) > 0 Then
                Call SetMask(EffectOverTime(i).EffectModifiers, e_ModifierTypes.PhysiccalBonus)
            End If
            If val(Leer.GetValue("EOT" & i, "AffectedByPhysicalReduction")) > 0 Then
                Call SetMask(EffectOverTime(i).EffectModifiers, e_ModifierTypes.PhysicalReduction)
            End If
            If val(Leer.GetValue("EOT" & i, "AffectedBySpeedModifier")) > 0 Then
                Call SetMask(EffectOverTime(i).EffectModifiers, e_ModifierTypes.MovementSpeed)
            End If
            If val(Leer.GetValue("EOT" & i, "AffectedByMagicHealing")) > 0 Then
                Call SetMask(EffectOverTime(i).EffectModifiers, e_ModifierTypes.MagicHealingBonus)
            End If
            If val(Leer.GetValue("EOT" & i, "AffectedBySelfHealing")) > 0 Then
                Call SetMask(EffectOverTime(i).EffectModifiers, e_ModifierTypes.SelfHealingBonus)
            End If
290         If val(Leer.GetValue("EOT" & i, "RequireArmor")) > 0 Then Call SetMask(EffectOverTime(i).SpellRequirementMask, e_SpellRequirementMask.eArmor)
292         If val(Leer.GetValue("EOT" & i, "RequireShip")) > 0 Then Call SetMask(EffectOverTime(i).SpellRequirementMask, e_SpellRequirementMask.eShip)
294         If val(Leer.GetValue("EOT" & i, "RequireHelm")) > 0 Then Call SetMask(EffectOverTime(i).SpellRequirementMask, e_SpellRequirementMask.eHelm)
296         If val(Leer.GetValue("EOT" & i, "RequireKnucle")) > 0 Then Call SetMask(EffectOverTime(i).SpellRequirementMask, e_SpellRequirementMask.eKnucle)
298         If val(Leer.GetValue("EOT" & i, "RequireMagicItem")) > 0 Then Call SetMask(EffectOverTime(i).SpellRequirementMask, e_SpellRequirementMask.eMagicItem)
300         If val(Leer.GetValue("EOT" & i, "RequireProjectile")) > 0 Then Call SetMask(EffectOverTime(i).SpellRequirementMask, e_SpellRequirementMask.eProjectile)
302         If val(Leer.GetValue("EOT" & i, "RequireShield")) > 0 Then Call SetMask(EffectOverTime(i).SpellRequirementMask, e_SpellRequirementMask.eShield)
304         If val(Leer.GetValue("EOT" & i, "RequireWeapon")) > 0 Then Call SetMask(EffectOverTime(i).SpellRequirementMask, e_SpellRequirementMask.eWeapon)
316         EffectOverTime(i).npcId = val(Leer.GetValue("EOT" & i, "NpcId"))
            If val(Leer.GetValue("EOT" & i, "LimitCastOnlyOnSelf")) > 0 Then
                Call SetMask(EffectOverTime(i).ApplyStatusMask, e_StatusMask.eCastOnlyOnSelf)
            End If
            If val(Leer.GetValue("EOT" & i, "Transform")) > 0 Then
                Call SetMask(EffectOverTime(i).ApplyStatusMask, e_StatusMask.eTransformed)
            End If
            If val(Leer.GetValue("EOT" & i, "CCInmunity")) > 0 Then
                Call SetMask(EffectOverTime(i).ApplyStatusMask, e_StatusMask.eCCInmunity)
            End If
            
            If val(Leer.GetValue("EOT" & i, "RequireSword")) > 0 Then Call SetIntMask(EffectOverTime(i).RequireWeaponType, ShiftLeft(1, eSword))
            If val(Leer.GetValue("EOT" & i, "RequireDagger")) > 0 Then Call SetIntMask(EffectOverTime(i).RequireWeaponType, ShiftLeft(1, eDagger))
            If val(Leer.GetValue("EOT" & i, "RequireBow")) > 0 Then Call SetIntMask(EffectOverTime(i).RequireWeaponType, ShiftLeft(1, eBow))
            If val(Leer.GetValue("EOT" & i, "RequireStaff")) > 0 Then Call SetIntMask(EffectOverTime(i).RequireWeaponType, ShiftLeft(1, eStaff))
            If val(Leer.GetValue("EOT" & i, "RequireMace")) > 0 Then Call SetIntMask(EffectOverTime(i).RequireWeaponType, ShiftLeft(1, eMace))
            If val(Leer.GetValue("EOT" & i, "RequireThrowableAxe")) > 0 Then Call SetIntMask(EffectOverTime(i).RequireWeaponType, ShiftLeft(1, eThrowableAxe))
            If val(Leer.GetValue("EOT" & i, "RequireAxe")) > 0 Then Call SetIntMask(EffectOverTime(i).RequireWeaponType, ShiftLeft(1, eAxe))
            If val(Leer.GetValue("EOT" & i, "RequireKnucle")) > 0 Then Call SetIntMask(EffectOverTime(i).RequireWeaponType, ShiftLeft(1, eKnuckle))
            If val(Leer.GetValue("EOT" & i, "RequireFist")) > 0 Then Call SetIntMask(EffectOverTime(i).RequireWeaponType, ShiftLeft(1, eFist))
            If val(Leer.GetValue("EOT" & i, "RequireSpear")) > 0 Then Call SetIntMask(EffectOverTime(i).RequireWeaponType, ShiftLeft(1, eSpear))
            If val(Leer.GetValue("EOT" & i, "RequireGunpowder")) > 0 Then Call SetIntMask(EffectOverTime(i).RequireWeaponType, ShiftLeft(1, eGunPowder))
            
        Next i
        
        Call InitializePools
        Exit Sub
ErrHandler:
288     MsgBox "Error cargando EffectsOverTime.dat " & Err.Number & ": " & Err.Description
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
112     Call TraceError(Err.Number, Err.Description, "ES.LoadMotd", Erl)

        
End Sub

Public Sub DoBackUp()
        On Error GoTo DoBackUp_Err
    
        
100     haciendoBK = True

        Dim i As Integer
102     Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
108     Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
110     haciendoBK = False
        Call LogThis(0, "[BackUps.log] DoBackUp", vbLogEventTypeInformation)
        Exit Sub
DoBackUp_Err:
120     Call TraceError(Err.Number, Err.Description, "ES.DoBackUp", Erl)
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
114     Call TraceError(Err.Number, Err.Description, "ES.LoadArmasHerreria", Erl)

        
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
114     Call TraceError(Err.Number, Err.Description, "ES.LoadArmadurasHerreria", Erl)

        
End Sub

Sub LoadBalance()
        
        On Error GoTo LoadBalance_Err
        

        Dim BalanceIni As clsIniManager

100     Set BalanceIni = New clsIniManager
        If IsFeatureEnabled("balance-2") Then
101         BalanceIni.Initialize DatPath & "Balance2.dat"
        Else
102         BalanceIni.Initialize DatPath & "Balance.dat"
        End If
        Dim i, j As Long

        Dim SearchVar As String

        'Modificadores de Clase
104     For i = 1 To NUMCLASES
106         SearchVar = Replace$(Tilde(ListaClases(i)), " ", vbNullString)

108         With ModClase(i)
110             .Evasion = val(BalanceIni.GetValue("MODEVASION", SearchVar))
112             .AtaqueArmas = val(BalanceIni.GetValue("MODATAQUEARMAS", SearchVar))
114             .AtaqueProyectiles = val(BalanceIni.GetValue("MODATAQUEPROYECTILES", SearchVar))
116             .DañoArmas = val(BalanceIni.GetValue("MODDANOARMAS", SearchVar))
118             .DañoProyectiles = val(BalanceIni.GetValue("MODDANOPROYECTILES", SearchVar))
120             .DañoWrestling = val(BalanceIni.GetValue("MODDANOWRESTLING", SearchVar))
122             .Escudo = val(BalanceIni.GetValue("MODESCUDO", SearchVar))
124             .ModApuñalar = val(BalanceIni.GetValue("MODAPUNALAR", SearchVar))
126             .Vida = val(BalanceIni.GetValue("MODVIDA", SearchVar))
128             .ManaInicial = val(BalanceIni.GetValue("MANA_INICIAL", SearchVar))
130             .MultMana = val(BalanceIni.GetValue("MULT_MANA", SearchVar))
132             .AumentoSta = val(BalanceIni.GetValue("AUMENTO_STA", SearchVar))
134             .HitPre36 = val(BalanceIni.GetValue("GOLPE_PRE_36", SearchVar))
136             .HitPost36 = val(BalanceIni.GetValue("GOLPE_POST_36", SearchVar))
138             .ResistenciaMagica = val(BalanceIni.GetValue("MODRESISTENCIAMAGICA", SearchVar))
                For j = 1 To eWeaponTypeCount - 1
                    .WeaponHitBonus(j) = val(BalanceIni.GetValue(SearchVar, WeaponTypeNames(j)))
                Next j
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
171     RequiredSpellDisplayTime = val(BalanceIni.GetValue("EXTRA", "RequiredSpellDisplayTime"))
172     MaxInvisibleSpellDisplayTime = val(BalanceIni.GetValue("EXTRA", "MaxInvisibleSpellDisplayTime"))
        MultiShotReduction = val(BalanceIni.GetValue("EXTRA", "MultiShotReduction"))
        HomeTimer = val(BalanceIni.GetValue("EXTRA", "HomeTimer"))
        MagicSkillBonusDamageModifier = val(BalanceIni.GetValue("EXTRA", "MagicSkillBonusDamageModifier"))
        MRSkillProtectionModifier = val(BalanceIni.GetValue("EXTRA", "MagicResistanceSkillProtectionModifier"))
        'stun
        PlayerStunTime = val(BalanceIni.GetValue("STUN", "PlayerStunTime"))
        NpcStunTime = val(BalanceIni.GetValue("STUN", "NpcStunTime"))
        PlayerInmuneTime = val(BalanceIni.GetValue("STUN", "PlayerInmuneTime"))

        ' Exp
173     For i = 1 To STAT_MAXELV
174         ExpLevelUp(i) = val(BalanceIni.GetValue("EXP", i))
        Next
    
176     Set BalanceIni = Nothing
    
178     AgregarAConsola "Se cargó el balance (Balance.dat)"

        
        Exit Sub

LoadBalance_Err:
180     Call TraceError(Err.Number, Err.Description, "ES.LoadBalance", Erl)

        
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
114     Call TraceError(Err.Number, Err.Description, "ES.LoadObjCarpintero", Erl)

        
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
114     Call TraceError(Err.Number, Err.Description, "ES.LoadObjAlquimista", Erl)

        
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
114     Call TraceError(Err.Number, Err.Description, "ES.LoadObjSastre", Erl)

        
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

        On Error GoTo ErrHandler

100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."

        '*****************************************************************
        'Carga la lista de objetos
        '*****************************************************************
        Dim Object As Integer

        Dim Leer   As clsIniManager
102     Set Leer = New clsIniManager
104     Call Leer.Initialize(DatPath & "Obj.dat")

        'obtiene el numero de obj
106     NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))
    
108     With frmCargando.cargar
110         .Min = 0
112         .max = NumObjDatas
114         .value = 0
        End With
    
116     ReDim Preserve ObjData(1 To NumObjDatas) As t_ObjData

        ReDim ObjShop(1 To 1) As t_ObjData
        
        Dim ObjKey As String
        Dim str As String, Field() As String
        Dim Crafteo As clsCrafteo
        Dim NFT As Boolean
  
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
144             .Que_Skill = val(Leer.GetValue(ObjKey, "QueSkill"))
146             .QueAtributo = val(Leer.GetValue(ObjKey, "queatributo"))
148             .CuantoAumento = val(Leer.GetValue(ObjKey, "cuantoaumento"))
150             .MinELV = val(Leer.GetValue(ObjKey, "MinELV"))
152             .Subtipo = val(Leer.GetValue(ObjKey, "Subtipo"))
154             .Dorada = val(Leer.GetValue(ObjKey, "Dorada"))
155             .Blodium = val(Leer.GetValue(ObjKey, "Blodium"))
156             .VidaUtil = val(Leer.GetValue(ObjKey, "VidaUtil"))
158             .TiempoRegenerar = val(Leer.GetValue(ObjKey, "TiempoRegenerar"))
                .Jerarquia = val(Leer.GetValue(ObjKey, "Jerarquia"))
160             .Cooldown = val(Leer.GetValue(ObjKey, "CD"))
                .CdType = val(Leer.GetValue(ObjKey, "CDType"))
161             .ImprovedRangedHitChance = val(Leer.GetValue(ObjKey, "ImprovedRHit"))
                .ImprovedMeleeHitChance = val(Leer.GetValue(ObjKey, "ImprovedMHit"))
                .ApplyEffectId = val(Leer.GetValue(ObjKey, "ApplyEffectId"))
                Dim i As Integer

162             Select Case .OBJType

                    Case e_OBJType.otHerramientas
164                     .WeaponAnim = val(Leer.GetValue(ObjKey, "Anim"))
166                     .Power = val(Leer.GetValue(ObjKey, "Power"))
            
168                 Case e_OBJType.otArmadura
170                     .Real = val(Leer.GetValue(ObjKey, "Real"))
172                     .Caos = val(Leer.GetValue(ObjKey, "Caos"))
174                     .ResistenciaMagica = val(Leer.GetValue(ObjKey, "ResistenciaMagica"))
176                     .Invernal = val(Leer.GetValue(ObjKey, "Invernal")) > 0
        
178                 Case e_OBJType.otEscudo
180                     .ShieldAnim = val(Leer.GetValue(ObjKey, "Anim"))
182                     .Real = val(Leer.GetValue(ObjKey, "Real"))
184                     .Caos = val(Leer.GetValue(ObjKey, "Caos"))
186                     .ResistenciaMagica = val(Leer.GetValue(ObjKey, "ResistenciaMagica"))
                        .Porcentaje = val(Leer.GetValue(ObjKey, "Porcentaje"))
        
188                 Case e_OBJType.otCasco
190                     .CascoAnim = val(Leer.GetValue(ObjKey, "Anim"))
192                     .Real = val(Leer.GetValue(ObjKey, "Real"))
194                     .Caos = val(Leer.GetValue(ObjKey, "Caos"))
196                     .ResistenciaMagica = val(Leer.GetValue(ObjKey, "ResistenciaMagica"))
        
198                 Case e_OBJType.otWeapon
200                     .WeaponAnim = val(Leer.GetValue(ObjKey, "Anim"))
202                     .Apuñala = val(Leer.GetValue(ObjKey, "Apuñala"))
204                     .Envenena = val(Leer.GetValue(ObjKey, "Envenena"))
206                     .Paraliza = val(Leer.GetValue(ObjKey, "Paraliza"))
208                     .Estupidiza = val(Leer.GetValue(ObjKey, "Estupidiza"))
210                     .incinera = val(Leer.GetValue(ObjKey, "Incinera"))
        
212                     .MaxHit = val(Leer.GetValue(ObjKey, "MaxHIT"))
213                     .MinHIT = val(Leer.GetValue(ObjKey, "MinHIT"))

214                     .IgnoreArmorAmmount = val(Leer.GetValue(ObjKey, "IgnoreArmorAmmount"))
215                     .IgnoreArmorPercent = val(Leer.GetValue(ObjKey, "IgnoreArmorPercent"))
216                     .Proyectil = val(Leer.GetValue(ObjKey, "Proyectil"))
218                     .Municion = val(Leer.GetValue(ObjKey, "Municiones"))
220                     .Power = val(Leer.GetValue(ObjKey, "StaffPower"))
222                     .MagicDamageBonus = val(Leer.GetValue(ObjKey, "MagicDamageBonus"))
226                     .Real = val(Leer.GetValue(ObjKey, "Real"))
228                     .Caos = val(Leer.GetValue(ObjKey, "Caos"))
230                     .EfectoMagico = val(Leer.GetValue(ObjKey, "efectomagico"))
232                     .Revive = val(Leer.GetValue(ObjKey, "Revive")) <> 0
234                     .DosManos = val(Leer.GetValue(ObjKey, "DosManos"))
                        .Porcentaje = val(Leer.GetValue(ObjKey, "Porcentaje"))
                        .WeaponType = val(Leer.GetValue(ObjKey, "WeaponType"))
                        
236                 Case e_OBJType.otInstrumentos
        
                        'Pablo (ToxicWaste)
238                     .Real = val(Leer.GetValue(ObjKey, "Real"))
240                     .Caos = val(Leer.GetValue(ObjKey, "Caos"))
        
242                 Case e_OBJType.otPuertas, e_OBJType.otBotellaVacia, e_OBJType.otBotellaLlena
244                     .IndexAbierta = val(Leer.GetValue(ObjKey, "IndexAbierta"))
246                     .IndexCerrada = val(Leer.GetValue(ObjKey, "IndexCerrada"))
248                     .IndexCerradaLlave = val(Leer.GetValue(ObjKey, "IndexCerradaLlave"))
        
250                 Case otPociones
252                     .TipoPocion = val(Leer.GetValue(ObjKey, "TipoPocion"))
254                     .MaxModificador = val(Leer.GetValue(ObjKey, "MaxModificador"))
256                     .MinModificador = val(Leer.GetValue(ObjKey, "MinModificador"))
            
258                     .DuracionEfecto = val(Leer.GetValue(ObjKey, "DuracionEfecto"))
260                     .Raices = val(Leer.GetValue(ObjKey, "Raices"))
262                     .SkPociones = val(Leer.GetValue(ObjKey, "SkPociones"))
264                     .Porcentaje = val(Leer.GetValue(ObjKey, "Porcentaje"))
        
266                 Case e_OBJType.otBarcos
268                     .MaxHit = val(Leer.GetValue(ObjKey, "MaxHIT"))
270                     .MinHIT = val(Leer.GetValue(ObjKey, "MinHIT"))
272                     .velocidad = val(Leer.GetValue(ObjKey, "Velocidad"))

274                 Case e_OBJType.otMonturas
276                     .MaxHit = val(Leer.GetValue(ObjKey, "MaxHIT"))
278                     .MinHIT = val(Leer.GetValue(ObjKey, "MinHIT"))
280                     .MinDef = val(Leer.GetValue(ObjKey, "MINDEF"))
282                     .MaxDef = val(Leer.GetValue(ObjKey, "MAXDEF"))
284                     .Real = val(Leer.GetValue(ObjKey, "Real"))
286                     .Caos = val(Leer.GetValue(ObjKey, "Caos"))
288                     .velocidad = val(Leer.GetValue(ObjKey, "Velocidad"))
        
290                 Case e_OBJType.otFlechas
292                     .MaxHit = val(Leer.GetValue(ObjKey, "MaxHIT"))
294                     .MinHIT = val(Leer.GetValue(ObjKey, "MinHIT"))
296                     .Envenena = val(Leer.GetValue(ObjKey, "Envenena"))
298                     .Paraliza = val(Leer.GetValue(ObjKey, "Paraliza"))
300                     .Estupidiza = val(Leer.GetValue(ObjKey, "Estupidiza"))
302                     .incinera = val(Leer.GetValue(ObjKey, "Incinera"))
    
306                     .EfectoMagico = val(Leer.GetValue(ObjKey, "efectomagico"))
            
308                     .Snd1 = val(Leer.GetValue(ObjKey, "SND1"))
310                     .Snd2 = val(Leer.GetValue(ObjKey, "SND2"))
            
                        'Pasajes Ladder 05-05-08
312                 Case e_OBJType.otpasajes
314                     .DesdeMap = val(Leer.GetValue(ObjKey, "DesdeMap"))
316                     .HastaMap = val(Leer.GetValue(ObjKey, "Map"))
318                     .HastaX = val(Leer.GetValue(ObjKey, "X"))
320                     .HastaY = val(Leer.GetValue(ObjKey, "Y"))
322                     .NecesitaNave = val(Leer.GetValue(ObjKey, "NecesitaNave"))
324                 Case e_OBJType.OtDonador
326                     .HastaMap = val(Leer.GetValue(ObjKey, "Map"))
328                     .HastaX = val(Leer.GetValue(ObjKey, "X"))
330                     .HastaY = val(Leer.GetValue(ObjKey, "Y"))
                    Case e_OBJType.OtQuest
                        .QuestId = val(Leer.GetValue(ObjKey, "QuestID"))
332                 Case e_OBJType.otMagicos
334                     .EfectoMagico = val(Leer.GetValue(ObjKey, "efectomagico"))
                        .Revive = val(Leer.GetValue(ObjKey, "Revive")) <> 0
336                     If .EfectoMagico = 15 Then
338                         PENDIENTE = Object
                        End If
                        If .EfectoMagico = 12 Then
                            .MaxItems = val(Leer.GetValue(ObjKey, "Peces"))
                        End If
            
340                 Case e_OBJType.otRunas
342                     .TipoRuna = val(Leer.GetValue(ObjKey, "TipoRuna"))
344                     .DesdeMap = val(Leer.GetValue(ObjKey, "DesdeMap"))
346                     .HastaMap = val(Leer.GetValue(ObjKey, "Map"))
348                     .HastaX = val(Leer.GetValue(ObjKey, "X"))
350                     .HastaY = val(Leer.GetValue(ObjKey, "Y"))
            
370                 Case e_OBJType.otTeleport
                        .Radio = val(Leer.GetValue(ObjKey, "Radio"))
                        
372                 Case e_OBJType.OtCofre
374                     .CantItem = val(Leer.GetValue(ObjKey, "CantItem"))

376                     Select Case .Subtipo
                            Case 1
378                             ReDim .Item(1 To .CantItem)
                    
380                             For i = 1 To .CantItem
382                                 .Item(i).objIndex = val(Leer.GetValue(ObjKey, "Item" & i))
384                                 .Item(i).amount = val(Leer.GetValue(ObjKey, "Cantidad" & i))
386                             Next i

                            Case 2
388                             ReDim .Item(1 To .CantItem)
                    
390                             .CantEntrega = val(Leer.GetValue(ObjKey, "CantEntrega"))
    
392                             For i = 1 To .CantItem
394                                 .Item(i).objIndex = val(Leer.GetValue(ObjKey, "Item" & i))
396                                 .Item(i).amount = val(Leer.GetValue(ObjKey, "Cantidad" & i))
398                             Next i

                            Case 3
                                ReDim .Item(1 To .CantItem)
    
                                For i = 1 To .CantItem
                                    .Item(i).objIndex = val(Leer.GetValue(ObjKey, "Item" & i))
                                    .Item(i).amount = val(Leer.GetValue(ObjKey, "Cantidad" & i))
                                    .Item(i).Data = 101 - val(Leer.GetValue(ObjKey, "Drop" & i))
                                Next i

                        End Select
            
400                 Case e_OBJType.otYacimiento
402                     .MineralIndex = val(Leer.GetValue(ObjKey, "MineralIndex"))
                        ' Drop gemas yacimientos
404                     .CantItem = val(Leer.GetValue(ObjKey, "Gemas"))
            
406                     If .CantItem > 0 Then
408                         ReDim .Item(1 To .CantItem)

410                         For i = 1 To .CantItem
412                             str = Leer.GetValue(ObjKey, "Gema" & i)
414                             Field = Split(str, "-")
416                             .Item(i).objIndex = val(Field(0))    ' ObjIndex
418                             .Item(i).amount = val(Field(1))      ' Probabilidad de drop (1 en X)
420                         Next i

                        End If
                    Case e_OBJType.otUsableOntarget
                        .MaxHit = val(Leer.GetValue(ObjKey, "MaxHIT"))
                        .MinHIT = val(Leer.GetValue(ObjKey, "MinHIT"))
                        .Proyectil = val(Leer.GetValue(ObjKey, "Proyectil"))
                
422                 Case e_OBJType.otDañoMagico
424                     .MagicDamageBonus = val(Leer.GetValue(ObjKey, "MagicDamageBonus"))
426                     .Revive = val(Leer.GetValue(ObjKey, "Revive")) <> 0

428                 Case e_OBJType.otResistencia
430                     .ResistenciaMagica = val(Leer.GetValue(ObjKey, "ResistenciaMagica"))


432                 Case e_OBJType.otMinerales
434                     .LingoteIndex = val(Leer.GetValue(ObjKey, "LingoteIndex"))
                    Case e_OBJType.otUsableOntarget
                        .EfectoMagico = val(Leer.GetValue(ObjKey, "efectomagico"))
                End Select
                .EfectoMagico = val(Leer.GetValue(ObjKey, "EfectoMagico"))
                .ProjectileType = val(Leer.GetValue(ObjKey, "ProjectileType"))
436             .MinSkill = val(Leer.GetValue(ObjKey, "MinSkill"))

438             .Elfico = val(Leer.GetValue(ObjKey, "Elfico"))
439             .Pino = val(Leer.GetValue(ObjKey, "Pino"))

440             .Snd1 = val(Leer.GetValue(ObjKey, "SND1"))
442             .Snd2 = val(Leer.GetValue(ObjKey, "SND2"))
444             .Snd3 = val(Leer.GetValue(ObjKey, "SND3"))
                'DELETE
446             .SndAura = val(Leer.GetValue(ObjKey, "SndAura"))
                '
    
448             .NoSeLimpia = val(Leer.GetValue(ObjKey, "NoSeLimpia"))
450             .Subastable = val(Leer.GetValue(ObjKey, "Subastable"))
    
452             .ParticulaGolpe = val(Leer.GetValue(ObjKey, "ParticulaGolpe"))
454             .ParticulaViaje = val(Leer.GetValue(ObjKey, "ParticulaViaje"))
456             .ParticulaGolpeTime = val(Leer.GetValue(ObjKey, "ParticulaGolpeTime"))
    
458             .Ropaje = val(Leer.GetValue(ObjKey, "NumRopaje"))
460             .HechizoIndex = val(Leer.GetValue(ObjKey, "HechizoIndex"))
    
462             .MaxHp = val(Leer.GetValue(ObjKey, "MaxHP"))
464             .MinHp = val(Leer.GetValue(ObjKey, "MinHP"))
    
466             .Mujer = val(Leer.GetValue(ObjKey, "Mujer"))
468             .Hombre = val(Leer.GetValue(ObjKey, "Hombre"))
    
470             .PielLobo = val(Leer.GetValue(ObjKey, "PielLobo"))
472             .PielOsoPardo = val(Leer.GetValue(ObjKey, "PielOsoPardo"))
474             .PielOsoPolaR = val(Leer.GetValue(ObjKey, "PielOsoPolaR"))
476             .SkMAGOria = val(Leer.GetValue(ObjKey, "SKSastreria"))

478             .LingH = val(Leer.GetValue(ObjKey, "LingH"))
480             .LingP = val(Leer.GetValue(ObjKey, "LingP"))
482             .LingO = val(Leer.GetValue(ObjKey, "LingO"))
484             .SkHerreria = val(Leer.GetValue(ObjKey, "SkHerreria"))
    
486             .CreaParticula = Leer.GetValue(ObjKey, "CreaParticula")
488             .CreaFX = val(Leer.GetValue(ObjKey, "CreaFX"))
490             .CreaGRH = Leer.GetValue(ObjKey, "CreaGRH")
492             .CreaLuz = Leer.GetValue(ObjKey, "CreaLuz")
    
494             .MinHam = val(Leer.GetValue(ObjKey, "MinHam"))
496             .MinSed = val(Leer.GetValue(ObjKey, "MinAgu"))
497             .PuntosPesca = val(Leer.GetValue(ObjKey, "PuntosPesca"))
    
498             .MinDef = val(Leer.GetValue(ObjKey, "MINDEF"))
500             .MaxDef = val(Leer.GetValue(ObjKey, "MAXDEF"))
502             .def = (.MinDef + .MaxDef) / 2
    
504             .ClaseTipo = val(Leer.GetValue(ObjKey, "ClaseTipo"))

506             .RazaEnana = val(Leer.GetValue(ObjKey, "RazaEnana"))
508             .RazaDrow = val(Leer.GetValue(ObjKey, "RazaDrow"))
510             .RazaElfa = val(Leer.GetValue(ObjKey, "RazaElfa"))
512             .RazaGnoma = val(Leer.GetValue(ObjKey, "RazaGnoma"))
514             .RazaOrca = val(Leer.GetValue(ObjKey, "RazaOrca"))
516             .RazaHumana = val(Leer.GetValue(ObjKey, "RazaHumana"))
    
518             .Valor = val(Leer.GetValue(ObjKey, "Valor"))
    
520             .Crucial = val(Leer.GetValue(ObjKey, "Crucial"))
    
                '.Cerrada = val(Leer.GetValue(ObjKey, "abierta")) cerrada = abierta??? WTF???????
522             .Cerrada = val(Leer.GetValue(ObjKey, "Cerrada"))

524             If .Cerrada = 1 Then
526                 .Llave = val(Leer.GetValue(ObjKey, "Llave"))
528                 .clave = val(Leer.GetValue(ObjKey, "Clave"))

                End If
    
                'Puertas y llaves
530             .clave = val(Leer.GetValue(ObjKey, "Clave"))
    
532             .texto = Leer.GetValue(ObjKey, "Texto")
534             .GrhSecundario = val(Leer.GetValue(ObjKey, "VGrande"))
    
536             .Agarrable = val(Leer.GetValue(ObjKey, "Agarrable"))
538             .ForoID = Leer.GetValue(ObjKey, "ID")
    
                'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico  -  Nunca más papu
                Dim n As Integer
                Dim s As String

540             For i = 1 To NUMCLASES
542                 s = UCase$(Leer.GetValue(ObjKey, "CP" & i))
544                 n = 1

546                 Do While LenB(s) > 0 And Tilde(ListaClases(n)) <> Trim$(s)
548                     n = n + 1
                    Loop
            
550                 .ClaseProhibida(i) = IIf(LenB(s) > 0, n, 0)
552             Next i
        
554             For i = 1 To NUMRAZAS
556                 s = UCase$(Leer.GetValue(ObjKey, "RP" & i))
558                 n = 1

560                 Do While LenB(s) > 0 And Tilde(ListaRazas(n)) <> Trim$(s)
562                     n = n + 1
                    Loop
            
564                 .RazaProhibida(i) = IIf(LenB(s) > 0, n, 0)
566             Next i
        
                ' Skill requerido
568             str = Leer.GetValue(ObjKey, "SkillRequerido")

570             If Len(str) > 0 Then
572                 Field = Split(str, "-")
            
574                 n = 1

576                 Do While LenB(Field(0)) > 0 And Tilde(SkillsNames(n)) <> Tilde(Field(0))
578                     n = n + 1
                    Loop
    
580                 .SkillIndex = IIf(LenB(Field(0)) > 0, n, 0)
582                 .SkillRequerido = val(Field(1))

                End If

                ' -----------------
    
584             .SkCarpinteria = val(Leer.GetValue(ObjKey, "SkCarpinteria"))
    
                'If .SkCarpinteria > 0 Then
586             .Madera = val(Leer.GetValue(ObjKey, "Madera"))

588             .MaderaElfica = val(Leer.GetValue(ObjKey, "MaderaElfica"))
                
                .MaderaPino = val(Leer.GetValue(ObjKey, "Maderapino"))
    
                'Bebidas
590             .MinSta = val(Leer.GetValue(ObjKey, "MinST"))
    
592             .NoSeCae = val(Leer.GetValue(ObjKey, "NoSeCae"))

                ' Crafteos
594             If val(Leer.GetValue(ObjKey, "Crafteable")) = 1 Then
596                 str = Leer.GetValue(ObjKey, "Materiales")

598                 If LenB(str) Then
600                     Field = Split(str, "-", MAX_SLOTS_CRAFTEO)
                    
                        Dim Items() As Integer
602                     ReDim Items(1 To UBound(Field) + 1)

604                     For i = 0 To UBound(Field)
606                         Items(i + 1) = val(Field(i))
608                         If Items(i + 1) > UBound(ObjData) Then Items(i + 1) = 0
                        Next

610                     Call SortIntegerArray(Items, 1, UBound(Items))
                        
612                     Set Crafteo = New clsCrafteo
614                     Call Crafteo.SetItems(Items)
616                     Crafteo.Tipo = val(Leer.GetValue(ObjKey, "TipoCrafteo"))
618                     Crafteo.Probabilidad = Clamp(val(Leer.GetValue(ObjKey, "ProbCrafteo")), 0, 100)
620                     Crafteo.precio = val(Leer.GetValue(ObjKey, "CostoCrafteo"))
622                     Crafteo.Resultado = Object

624                     If Not Crafteos.Exists(Crafteo.Tipo) Then
626                         Call Crafteos.Add(Crafteo.Tipo, New Dictionary)
                        End If

                        Dim ItemKey As String
628                     ItemKey = GetRecipeKey(Items)
630                     If Not Crafteos.Item(Crafteo.Tipo).Exists(ItemKey) Then
632                         Call Crafteos.Item(Crafteo.Tipo).Add(ItemKey, Crafteo)
                        End If
                    End If
                End If

                ' Catalizadores
634             .CatalizadorTipo = val(Leer.GetValue(ObjKey, "CatalizadorTipo"))
636             If .CatalizadorTipo Then
638                 .CatalizadorAumento = val(Leer.GetValue(ObjKey, "CatalizadorAumento"))
                End If
                
                NFT = val(Leer.GetValue(ObjKey, "NFT"))
                
                .ObjDonador = NFT
                
                If NFT Then
                    ObjShop(UBound(ObjShop)).Name = Leer.GetValue(ObjKey, "Name")
                    ObjShop(UBound(ObjShop)).Valor = val(Leer.GetValue(ObjKey, "Valor"))
                    ObjShop(UBound(ObjShop)).ObjNum = Object
                    ObjShop(UBound(ObjShop)).ObjDonador = 1
                    ReDim Preserve ObjShop(1 To (UBound(ObjShop) + 1)) As t_ObjData
                End If
                
                
640             frmCargando.cargar.value = frmCargando.cargar.value + 1

        
            End With
            
            ' WyroX: Cada 10 objetos revivo la interfaz
642         If Object Mod 10 = 0 Then DoEvents
        
644     Next Object
        ReDim Preserve ObjShop(1 To (UBound(ObjShop) - 1)) As t_ObjData

646     Set Leer = Nothing
    
648     Call InitTesoro
650     Call InitRegalo

        Exit Sub

ErrHandler:
652     MsgBox "error cargando objetos " & Err.Number & ": " & Err.Description & ". Error producido al cargar el objeto: " & Object

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
110     Call TraceError(Err.Number, Err.Description, "ES.GetVar", Erl)

        
End Function

Sub CargarBackUp()
        
        On Error GoTo CargarBackUp_Err
        

100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."

        Dim map       As Integer
        Dim TempInt   As Integer
        Dim npcfile   As String
    
102     If RunningInVB() Then
104         NumMaps = 869
        Else
106         NumMaps = CountFiles(MapPath, "*.csm")
108         NumMaps = NumMaps - 1
        End If
110     Call InitAreas
    
112     frmCargando.cargar.Min = 0
114     frmCargando.cargar.max = NumMaps
116     frmCargando.cargar.value = 0
118     frmCargando.ToMapLbl.Visible = True
    
120     ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As t_MapBlock

122     ReDim MapInfo(1 To NumMaps) As t_MapInfo
      
124     For map = 1 To NumMaps
126         frmCargando.ToMapLbl = map & "/" & NumMaps

128         Call CargarMapaFormatoCSM(map, App.Path & "\WorldBackUp\Mapa" & map & ".csm")

130         frmCargando.cargar.value = frmCargando.cargar.value + 1

132         DoEvents
134     Next map

        'Call generateMatrix(MATRIX_INITIAL_MAP)

136     frmCargando.ToMapLbl.Visible = False

        Exit Sub

CargarBackUp_Err:
138     Call TraceError(Err.Number, Err.Description, "ES.CargarBackUp", Erl)

        
End Sub

Sub LoadMapData()
        On Error GoTo man

100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas..."

        Dim map       As Integer
        Dim TempInt   As Integer
        Dim npcfile   As String

#If UNIT_TEST = 1 Then
        'We only need 10 maps for unit testing
        NumMaps = 10
        Debug.Print "UNIT_TEST Enabled Loading just " & NumMaps & " maps"
#Else

        If RunningInVB() Then
                'VB runs out of memory when debugging
                NumMaps = 512
        Else
                NumMaps = CountFiles(MapPath, "*.csm") - 1
        End If

#End If

104     Call InitAreas
    
106     frmCargando.cargar.Min = 0
108     frmCargando.cargar.max = NumMaps
110     frmCargando.cargar.value = 0
112     frmCargando.ToMapLbl.Visible = True

114     ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As t_MapBlock

116     ReDim MapInfo(1 To NumMaps) As t_MapInfo

118     For map = 1 To NumMaps
    
120         frmCargando.ToMapLbl = map & "/" & NumMaps

122         Call CargarMapaFormatoCSM(map, MapPath & "Mapa" & map & ".csm")

124         frmCargando.cargar.value = frmCargando.cargar.value + 1
        
126         DoEvents
        
128     Next map
    
        'Call generateMatrix(MATRIX_INITIAL_MAP)

130     frmCargando.ToMapLbl.Visible = False
    
        Exit Sub

man:
132     Call MsgBox("Error durante la carga de mapas, el mapa " & map & " contiene errores")
134     Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)

End Sub

Public Sub CargarMapaFormatoCSM(ByVal map As Long, ByVal MAPFl As String)

        On Error GoTo ErrorHandler:

        Dim npcfile      As String
        Dim fh           As Integer
        Dim MH           As t_MapHeader
        Dim Blqs()       As t_DatosBloqueados
        Dim L1()         As t_DatosGrh
        Dim L2()         As t_DatosGrh
        Dim L3()         As t_DatosGrh
        Dim L4()         As t_DatosGrh

        Dim Triggers()   As t_DatosTrigger
        Dim Luces()      As t_DatosLuces
        Dim Particulas() As t_DatosParticulas
        Dim Objetos()    As t_DatosObjs
        Dim NPCs()       As t_DatosNPC
        Dim TEs()        As t_DatosTE
        Dim RandomTeleports(MAX_RANDOM_TELEPORT_IN_MAP) As Integer
        Dim randomTeleportCount As Integer
        Dim body         As Integer
        Dim head         As Integer
        Dim Heading      As Byte
        Dim SailingTiles As Long
        Dim TotalTiles   As Long

        Dim i            As Long
        Dim j            As Long
    
        Dim x As Integer, y As Integer
        randomTeleportCount = 0
100     If Not FileExist(MAPFl, vbNormal) Then
102         Call TraceError(404, "Estas tratando de cargar un MAPA que NO EXISTE" & vbNewLine & "Mapa: " & MAPFl, "ES.CargarMapaFormatoCSM")
            Exit Sub
        End If
        
104     If FileLen(MAPFl) = 0 Then
106         Call TraceError(500, "Se trato de cargar un mapa corrupto o mal generado" & vbNewLine & "Mapa: " & MAPFl, "ES.CargarMapaFormatoCSM")
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
128                 MapData(map, Blqs(i).x, Blqs(i).y).Blocked = Blqs(i).Lados
130             Next i
            End If
        
            'Cargamos Layer 1
        
132         If .NumeroLayers(1) > 0 Then
134             ReDim L1(1 To .NumeroLayers(1))
136             Get #fh, , L1
138             For i = 1 To .NumeroLayers(1)
140                 x = L1(i).x
142                 y = L1(i).y
144                 MapData(map, x, y).Graphic(1) = L1(i).GrhIndex
            
                    TotalTiles = TotalTiles + 1
146                 If HayAgua(map, x, y) Then
148                     MapData(map, x, y).Blocked = MapData(map, x, y).Blocked Or FLAG_AGUA
                        SailingTiles = SailingTiles + 1
                    End If
150             Next i
            End If
        
            'Cargamos Layer 2
152         If .NumeroLayers(2) > 0 Then
154             ReDim L2(1 To .NumeroLayers(2))
156             Get #fh, , L2
158             For i = 1 To .NumeroLayers(2)
160                 x = L2(i).x
162                 y = L2(i).y
164                 MapData(map, x, y).Graphic(2) = L2(i).GrhIndex
166                 MapData(map, x, y).Blocked = MapData(map, x, y).Blocked And Not FLAG_AGUA
168             Next i

            End If
                
170         If .NumeroLayers(3) > 0 Then
172             ReDim L3(1 To .NumeroLayers(3))
174             Get #fh, , L3

176             For i = 1 To .NumeroLayers(3)
178                 x = L3(i).x
180                 y = L3(i).y
182                 MapData(map, x, y).Graphic(3) = L3(i).GrhIndex
                
184                 If EsArbol(L3(i).GrhIndex) Then
186                     MapData(map, x, y).Blocked = MapData(map, x, y).Blocked Or FLAG_ARBOL
                    End If
188             Next i

            End If
        
190         If .NumeroLayers(4) > 0 Then
192             ReDim L4(1 To .NumeroLayers(4))
194             Get #fh, , L4
196             For i = 1 To .NumeroLayers(4)
198                 MapData(map, L4(i).x, L4(i).y).Graphic(4) = L4(i).GrhIndex
200             Next i

            End If

202         If .NumeroTriggers > 0 Then
204             ReDim Triggers(1 To .NumeroTriggers)
206             Get #fh, , Triggers

208             For i = 1 To .NumeroTriggers
210                 x = Triggers(i).x
212                 y = Triggers(i).y
214                 MapData(map, x, y).trigger = Triggers(i).trigger

                    ' Trigger detalles en agua
216                 If Triggers(i).trigger = e_Trigger.DETALLEAGUA Then
                        ' Vuelvo a poner flag agua
218                     MapData(map, x, y).Blocked = MapData(map, x, y).Blocked Or FLAG_AGUA
                    End If
                    
217                 If Triggers(i).trigger = e_Trigger.VALIDONADO Or Triggers(i).trigger = e_Trigger.NADOCOMBINADO Or Triggers(i).trigger = e_Trigger.NADOBAJOTECHO Then
                        ' Vuelvo a poner flag agua
219                     MapData(map, x, y).Blocked = MapData(map, x, y).Blocked Or FLAG_AGUA
                    End If
220             Next i
            End If

222         If .NumeroParticulas > 0 Then
224             ReDim Particulas(1 To .NumeroParticulas)
226             Get #fh, , Particulas

228             For i = 1 To .NumeroParticulas
230                 MapData(map, Particulas(i).x, Particulas(i).y).ParticulaIndex = Particulas(i).Particula
232                 MapData(map, Particulas(i).x, Particulas(i).y).ParticulaIndex = 0
234             Next i
            End If

236         If .NumeroLuces > 0 Then
238             ReDim Luces(1 To .NumeroLuces)
240             Get #fh, , Luces

242             For i = 1 To .NumeroLuces
244                 MapData(map, Luces(i).x, Luces(i).y).Luz.Color = Luces(i).Color
246                 MapData(map, Luces(i).x, Luces(i).y).Luz.Rango = Luces(i).Rango
248                 MapData(map, Luces(i).x, Luces(i).y).Luz.Color = 0
250                 MapData(map, Luces(i).x, Luces(i).y).Luz.Rango = 0
252             Next i
            End If
            
254         If .NumeroOBJs > 0 Then
256             ReDim Objetos(1 To .NumeroOBJs)
258             Get #fh, , Objetos
260             For i = 1 To .NumeroOBJs
262                 MapData(map, Objetos(i).x, Objetos(i).y).ObjInfo.objIndex = Objetos(i).objIndex
                    With ObjData(Objetos(i).objIndex)
264                 Select Case .OBJType
                        Case e_OBJType.otYacimiento, e_OBJType.otArboles
266                         MapData(map, Objetos(i).x, Objetos(i).y).ObjInfo.amount = ObjData(Objetos(i).objIndex).VidaUtil
268                         MapData(map, Objetos(i).x, Objetos(i).y).ObjInfo.Data = &H7FFFFFFF ' Ultimo uso = Max Long
270                     Case Else
272                         MapData(map, Objetos(i).x, Objetos(i).y).ObjInfo.amount = Objetos(i).ObjAmmount
                    End Select
                    If .OBJType = otTeleport And .Subtipo = e_TeleportSubType.eTransportNetwork Then
                        RandomTeleports(randomTeleportCount) = i
                        randomTeleportCount = randomTeleportCount + 1
                    End If
                    End With
274             Next i
            End If

276         If .NumeroNPCs > 0 Then
278             ReDim NPCs(1 To .NumeroNPCs)
280             Get #fh, , NPCs

                Dim NumNpc As Integer, npcIndex As Integer
                 
282             For i = 1 To .NumeroNPCs
284                 NumNpc = NPCs(i).npcIndex
286                 If NumNpc > 0 Then
288                     npcfile = DatPath & "NPCs.dat"
290                     npcIndex = OpenNPC(NumNpc)
                        
292                     If npcIndex > 0 Then
294                         MapData(map, NPCs(i).x, NPCs(i).y).npcIndex = npcIndex
296                         NpcList(npcIndex).pos.map = map
298                         NpcList(npcIndex).pos.x = NPCs(i).x
300                         NpcList(npcIndex).pos.y = NPCs(i).y
                            ' WyroX: guardo siempre la pos original... puede sernos útil ;)
302                         NpcList(npcIndex).Orig = NpcList(npcIndex).pos
    
304                         If LenB(NpcList(npcIndex).Name) = 0 Then
306                             MapData(map, NPCs(i).x, NPCs(i).y).npcIndex = 0
                            Else
308                             Call MakeNPCChar(True, 0, npcIndex, map, NPCs(i).x, NPCs(i).y)
                            End If
                        Else
                            ' Lo guardo en los logs + aparece en el Debug.Print
310                         Call TraceError(404, "NPC no existe en los .DAT's o está mal dateado. Posicion: " & Map & "-" & NPCs(i).x & "-" & NPCs(i).y, "ES.CargarMapaFormatoCSM")
                        End If
                    End If
312             Next i
            End If
            
314         If .NumeroTE > 0 Then
316             ReDim TEs(1 To .NumeroTE)
318             Get #fh, , TEs

320             For i = 1 To .NumeroTE
322                 MapData(map, TEs(i).x, TEs(i).y).TileExit.map = TEs(i).DestM
324                 MapData(map, TEs(i).x, TEs(i).y).TileExit.x = TEs(i).DestX
326                 MapData(map, TEs(i).x, TEs(i).y).TileExit.y = TEs(i).DestY
328             Next i
            End If
        End With
330     Close fh

        ' WyroX: Nuevo sistema de restricciones
332     If Not IsNumeric(MapDat.restrict_mode) Then
            ' Solo se usaba el "NEWBIE"
334         If UCase$(MapDat.restrict_mode) = "NEWBIE" Then
336             MapDat.restrict_mode = "1"
            Else
338             MapDat.restrict_mode = "0"
            End If
        End If
        
        If SailingTiles * 100 / TotalTiles > FISHING_REQUIRED_PERCENT Then
            Call AddFishingPoolsToMap(Map)
        End If
    
340     MapInfo(map).map_name = MapDat.map_name
342     MapInfo(map).ambient = MapDat.ambient
344     MapInfo(map).backup_mode = MapDat.backup_mode
346     MapInfo(map).base_light = MapDat.base_light
348     MapInfo(map).Newbie = (val(MapDat.restrict_mode) And 1) <> 0
350     MapInfo(map).SinMagia = (val(MapDat.restrict_mode) And 2) <> 0
352     MapInfo(map).NoPKs = (val(MapDat.restrict_mode) And 4) <> 0
354     MapInfo(map).NoCiudadanos = (val(MapDat.restrict_mode) And 8) <> 0
356     MapInfo(map).SinInviOcul = (val(MapDat.restrict_mode) And 16) <> 0
358     MapInfo(map).SoloClanes = (val(MapDat.restrict_mode) And 32) <> 0
359     MapInfo(map).NoMascotas = (val(MapDat.restrict_mode) And 64) <> 0
360     MapInfo(map).ResuCiudad = val(GetVar(DatPath & "Map.dat", "RESUCIUDAD", map)) <> 0
362     MapInfo(map).letter_grh = MapDat.letter_grh
364     MapInfo(map).lluvia = MapDat.lluvia
366     MapInfo(map).music_numberHi = MapDat.music_numberHi
368     MapInfo(map).music_numberLow = MapDat.music_numberLow
370     MapInfo(map).niebla = MapDat.niebla
372     MapInfo(map).Nieve = MapDat.Nieve
374     MapInfo(map).MinLevel = MapDat.level And &HFF
376     MapInfo(map).MaxLevel = (MapDat.level And &HFF00) / &H100
    
378     MapInfo(map).Seguro = MapDat.Seguro

380     MapInfo(map).terrain = MapDat.terrain
382     MapInfo(map).zone = MapDat.zone
383     MapInfo(map).DropItems = True
        MapInfo(map).FriendlyFire = True

384     If LenB(MapDat.Salida) <> 0 Then
            Dim Fields() As String
386         Fields = Split(MapDat.Salida, "-")
388         MapInfo(map).Salida.map = val(Fields(0))
390         MapInfo(map).Salida.x = val(Fields(1))
392         MapInfo(map).Salida.y = val(Fields(2))
        End If
        If randomTeleportCount > 0 Then
            ReDim MapInfo(map).TransportNetwork(randomTeleportCount - 1) As t_TransportNetworkExit
            For i = 0 To randomTeleportCount - 1
                MapInfo(map).TransportNetwork(i).TileX = Objetos(RandomTeleports(i)).x
                MapInfo(map).TransportNetwork(i).TileY = Objetos(RandomTeleports(i)).y
            Next i
        End If
        Exit Sub

ErrorHandler:
394     Close fh
396     Call TraceError(Err.Number, Err.Description, "ES.CargarMapaFormatoCSM", Erl)
    
End Sub

Sub AddFishingPoolsToMap(ByVal Map As Integer)
    Dim i As Integer
    For i = 1 To FISHING_TILES_ON_MAP
        Call CreateFishingPool(Map)
    Next i
End Sub

Public Sub CreateFishingPool(ByVal Map As Integer)
    Dim x, y As Integer
    Do
        x = RandomNumber(12, 88)
        y = RandomNumber(12, 88)
    Loop While MapData(Map, x, y).ObjInfo.objIndex <> 0 Or Not HayAgua(Map, x, y)
    MapData(Map, x, y).ObjInfo.objIndex = FISHING_POOL_ID
    MapData(Map, x, y).ObjInfo.amount = ObjData(FISHING_POOL_ID).VidaUtil
    MapData(Map, x, y).ObjInfo.Data = &H7FFFFFFF ' Ultimo uso = Max Long
End Sub

Sub LoadPrivateKey()
    Dim MyLine As String
    Open App.Path & "\..\ao20-ComputePK\crypto-hex.txt" For Input As #1
    Line Input #1, PrivateKey
Close #1
End Sub
Sub LoadMD5()
    Open IniPath & "ClienteMD5.txt" For Input As #1
        Line Input #1, Md5Cliente
    Close #1
    Md5Cliente = Replace(Md5Cliente, " ", "")
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
    
        'Directorios
110     DatPath = Lector.GetValue("DIRECTORIOS", "DatPath")
112     MapPath = Lector.GetValue("DIRECTORIOS", "MapPath")
114     CharPath = Lector.GetValue("DIRECTORIOS", "CharPath")
116     DeletePath = Lector.GetValue("DIRECTORIOS", "DeletePath")
118     CuentasPath = Lector.GetValue("DIRECTORIOS", "CuentasPath")
120     DeleteCuentasPath = Lector.GetValue("DIRECTORIOS", "DeleteCuentasPath")
        'Directorios
    
122     Puerto = val(Lector.GetValue("INIT", "StartPort"))
126     HideMe = val(Lector.GetValue("INIT", "Hide"))
128     MaxConexionesIP = val(Lector.GetValue("INIT", "MaxConexionesIP"))
130     MaxUsersPorCuenta = val(Lector.GetValue("INIT", "MaxUsersPorCuenta"))
132     IdleLimit = val(Lector.GetValue("INIT", "IdleLimit"))
        'Lee la version correcta del cliente
134     ULTIMAVERSION = Lector.GetValue("INIT", "Version")
    
136     PuedeCrearPersonajes = val(Lector.GetValue("INIT", "PuedeCrearPersonajes"))
138     ServerSoloGMs = val(Lector.GetValue("init", "ServerSoloGMs"))
140     DisconnectTimeout = val(Lector.GetValue("INIT", "DisconnectTimeout"))
    
144     EnTesting = val(Lector.GetValue("INIT", "Testing"))
145     EnableTelemetry = val(Lector.GetValue("INIT", "EnableTelemetry"))
        
        'Ressurect pos
146     ResPos.map = val(ReadField(1, Lector.GetValue("INIT", "ResPos"), 45))
148     ResPos.x = val(ReadField(2, Lector.GetValue("INIT", "ResPos"), 45))
150     ResPos.y = val(ReadField(3, Lector.GetValue("INIT", "ResPos"), 45))
      
        'Max users
156     Temporal = val(Lector.GetValue("INIT", "MaxUsers"))

158     If MaxUsers = 0 Then

160         MaxUsers = Temporal
162         ReDim UserList(1 To MaxUsers) As t_User

        End If

164     Call CargarCiudades
166     Call LoadFeatureToggles
167     Call LoadGlobalDropTable

168     Set Lector = Nothing

        Exit Sub

LoadSini_Err:
170     Set Lector = Nothing
172     Call TraceError(Err.Number, Err.Description, "ES.LoadSini", Erl)
End Sub

Sub LoadGlobalDropTable()
    
    Dim Lector   As clsIniManager

    Dim Temporal As Long
    If Not FileExist(DatPath & "GlobalDropTable.dat") Then
        Exit Sub
    End If
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando tabla de drop globales."
    
    Set Lector = New clsIniManager
    Call Lector.Initialize(DatPath & "GlobalDropTable.dat")
    If Lector.NodesCount = 0 Then
        Set Lector = Nothing
        Exit Sub
    End If
    Dim DropCount, i As Integer
    DropCount = val(Lector.GetValue("INIT", "DROPCOUNT"))
    If DropCount = 0 Then
        ReDim GlobalDropTable(0) As t_GlobalDrop
        Set Lector = Nothing
        Exit Sub
    End If
    ReDim GlobalDropTable(1 To DropCount) As t_GlobalDrop
    For i = 1 To DropCount
        GlobalDropTable(i).MaxPercent = val(Lector.GetValue("DROP" & i, "MAXPERCENT"))
        GlobalDropTable(i).MinPercent = val(Lector.GetValue("DROP" & i, "MINPERCENT"))
        GlobalDropTable(i).ObjectNumber = val(Lector.GetValue("DROP" & i, "OBJECTNUMBER"))
        GlobalDropTable(i).RequiredHPForMaxChance = val(Lector.GetValue("DROP" & i, "HPFORMAXCHANCE"))
        GlobalDropTable(i).Amount = val(Lector.GetValue("DROP" & i, "AMOUNT"))
    Next i
    Set Lector = Nothing
End Sub

Sub LoadFeatureToggles()
On Error GoTo LoadFeatureToggles_Err

        Dim Lector   As clsIniManager

        Dim Temporal As Long
        Set FeatureToggles = New Dictionary
        If Not FileExist("feature_toggle.ini") Then
            Exit Sub
        End If
100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de feature toggles."
    
102     Set Lector = New clsIniManager
104     Call Lector.Initialize(IniPath & "feature_toggle.ini")
        If Lector.NodesCount = 0 Then
            Exit Sub
        End If
        Dim TOGGLECOUNT As Integer
        TOGGLECOUNT = val(Lector.GetValue("INIT", "TOGGLECOUNT"))
        
        Dim i As Integer
        Dim key As String
        Dim value As Boolean
        For i = 1 To TOGGLECOUNT
            key = Lector.GetValue("TOGGLE" & i, "name")
            value = val(Lector.GetValue("TOGGLE" & i, "value")) > 0
            Call SetFeatureToggle(key, value)
        Next i
        
168     Set Lector = Nothing
            
        Exit Sub
LoadFeatureToggles_Err:
170     Set Lector = Nothing
172     Call TraceError(Err.Number, Err.Description, "ES.LoadFeatureToggles", Erl)
End Sub

Sub LoadPacketRatePolicy()
        On Error GoTo LoadPacketRatePolicy_Err

        Dim Lector   As clsIniManager
        Dim i As Long
100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando PacketRatePolicy."
    
102     Set Lector = New clsIniManager
104     Call Lector.Initialize(IniPath & "PacketRatePolicy.ini")
            For i = 1 To MAX_PACKET_COUNTERS
                Dim PacketName As String
                PacketName = PacketIdToString(i)
                MacroIterations(i) = val(Lector.GetValue(PacketName, "Iterations"))
                PacketTimerThreshold(i) = val(Lector.GetValue(PacketName, "Limit"))
            Next i

168     Set Lector = Nothing

        Exit Sub

LoadPacketRatePolicy_Err:
170     Set Lector = Nothing
172     Call TraceError(Err.Number, Err.Description, "ES.LoadPacketRatePolicy", Erl)

        
End Sub

Sub CargarCiudades()
        
        On Error GoTo CargarCiudades_Err
    
        Dim i As Long

        Dim Lector As clsIniManager
100     Set Lector = New clsIniManager
102     Call Lector.Initialize(DatPath & "Ciudades.dat")
        Dim MapasCiudades As String
    
104     With CityNix
106         .map = val(Lector.GetValue("NIX", "Mapa"))
108         .x = val(Lector.GetValue("NIX", "X"))
110         .y = val(Lector.GetValue("NIX", "Y"))
112         .MapaViaje = val(Lector.GetValue("NIX", "MapaViaje"))
114         .ViajeX = val(Lector.GetValue("NIX", "ViajeX"))
116         .ViajeY = val(Lector.GetValue("NIX", "ViajeY"))
118         .MapaResu = val(Lector.GetValue("NIX", "MapaResu"))
120         .ResuX = val(Lector.GetValue("NIX", "ResuX"))
122         .ResuY = val(Lector.GetValue("NIX", "ResuY"))
124         .NecesitaNave = val(Lector.GetValue("NIX", "NecesitaNave"))
            MapasCiudades = Lector.GetValue("NIX", "Mapas") & ","
            
        End With
    
126     With CityUllathorpe
128         .map = val(Lector.GetValue("Ullathorpe", "Mapa"))
130         .x = val(Lector.GetValue("Ullathorpe", "X"))
132         .y = val(Lector.GetValue("Ullathorpe", "Y"))
134         .MapaViaje = val(Lector.GetValue("Ullathorpe", "MapaViaje"))
136         .ViajeX = val(Lector.GetValue("Ullathorpe", "ViajeX"))
138         .ViajeY = val(Lector.GetValue("Ullathorpe", "ViajeY"))
140         .MapaResu = val(Lector.GetValue("Ullathorpe", "MapaResu"))
142         .ResuX = val(Lector.GetValue("Ullathorpe", "ResuX"))
144         .ResuY = val(Lector.GetValue("Ullathorpe", "ResuY"))
146         .NecesitaNave = val(Lector.GetValue("Ullathorpe", "NecesitaNave"))
            MapasCiudades = MapasCiudades & Lector.GetValue("Ullathorpe", "Mapas") & ","
        End With
    
148     With CityBanderbill
150         .map = val(Lector.GetValue("Banderbill", "Mapa"))
152         .x = val(Lector.GetValue("Banderbill", "X"))
154         .y = val(Lector.GetValue("Banderbill", "Y"))
156         .MapaViaje = val(Lector.GetValue("Banderbill", "MapaViaje"))
158         .ViajeX = val(Lector.GetValue("Banderbill", "ViajeX"))
160         .ViajeY = val(Lector.GetValue("Banderbill", "ViajeY"))
162         .MapaResu = val(Lector.GetValue("Banderbill", "MapaResu"))
164         .ResuX = val(Lector.GetValue("Banderbill", "ResuX"))
166         .ResuY = val(Lector.GetValue("Banderbill", "ResuY"))
168         .NecesitaNave = val(Lector.GetValue("Banderbill", "NecesitaNave"))
            MapasCiudades = MapasCiudades & Lector.GetValue("Banderbill", "Mapas") & ","
        End With
    
170     With CityLindos
172         .map = val(Lector.GetValue("Lindos", "Mapa"))
174         .x = val(Lector.GetValue("Lindos", "X"))
176         .y = val(Lector.GetValue("Lindos", "Y"))
178         .MapaViaje = val(Lector.GetValue("Lindos", "MapaViaje"))
180         .ViajeX = val(Lector.GetValue("Lindos", "ViajeX"))
182         .ViajeY = val(Lector.GetValue("Lindos", "ViajeY"))
184         .MapaResu = val(Lector.GetValue("Lindos", "MapaResu"))
186         .ResuX = val(Lector.GetValue("Lindos", "ResuX"))
188         .ResuY = val(Lector.GetValue("Lindos", "ResuY"))
190         .NecesitaNave = val(Lector.GetValue("Lindos", "NecesitaNave"))
            MapasCiudades = MapasCiudades & Lector.GetValue("Lindos", "Mapas") & ","
        End With
    
192     With CityArghal
194         .map = val(Lector.GetValue("Arghal", "Mapa"))
196         .x = val(Lector.GetValue("Arghal", "X"))
198         .y = val(Lector.GetValue("Arghal", "Y"))
200         .MapaViaje = val(Lector.GetValue("Arghal", "MapaViaje"))
202         .ViajeX = val(Lector.GetValue("Arghal", "ViajeX"))
204         .ViajeY = val(Lector.GetValue("Arghal", "ViajeY"))
206         .MapaResu = val(Lector.GetValue("Arghal", "MapaResu"))
208         .ResuX = val(Lector.GetValue("Arghal", "ResuX"))
210         .ResuY = val(Lector.GetValue("Arghal", "ResuY"))
212         .NecesitaNave = val(Lector.GetValue("Arghal", "NecesitaNave"))
            MapasCiudades = MapasCiudades & Lector.GetValue("Arghal", "Mapas") & ","
        End With
    
214     With CityArkhein
216         .map = val(Lector.GetValue("Arkhein", "Mapa"))
218         .x = val(Lector.GetValue("Arkhein", "X"))
220         .y = val(Lector.GetValue("Arkhein", "Y"))
222         .MapaViaje = val(Lector.GetValue("Arkhein", "MapaViaje"))
224         .ViajeX = val(Lector.GetValue("Arkhein", "ViajeX"))
226         .ViajeY = val(Lector.GetValue("Arkhein", "ViajeY"))
228         .MapaResu = val(Lector.GetValue("Arkhein", "MapaResu"))
230         .ResuX = val(Lector.GetValue("Arkhein", "ResuX"))
232         .ResuY = val(Lector.GetValue("Arkhein", "ResuY"))
234         .NecesitaNave = val(Lector.GetValue("Arkhein", "NecesitaNave"))
            MapasCiudades = MapasCiudades & Lector.GetValue("Arkhein", "Mapas") & ","
        End With
    
        With CityEleusis
            .map = val(Lector.GetValue("Eleusis", "Mapa"))
            .x = val(Lector.GetValue("Eleusis", "X"))
            .y = val(Lector.GetValue("Eleusis", "Y"))
            .MapaViaje = val(Lector.GetValue("Eleusis", "MapaViaje"))
            .ViajeX = val(Lector.GetValue("Eleusis", "ViajeX"))
            .ViajeY = val(Lector.GetValue("Eleusis", "ViajeY"))
            .MapaResu = val(Lector.GetValue("Eleusis", "MapaResu"))
            .ResuX = val(Lector.GetValue("Eleusis", "ResuX"))
            .ResuY = val(Lector.GetValue("Eleusis", "ResuY"))
            .NecesitaNave = val(Lector.GetValue("Eleusis", "NecesitaNave"))
            MapasCiudades = MapasCiudades & Lector.GetValue("Eleusis", "Mapas") & ","
        End With
        
        With CityPenthar
            .map = val(Lector.GetValue("Penthar", "Mapa"))
            .x = val(Lector.GetValue("Penthar", "X"))
            .y = val(Lector.GetValue("Penthar", "Y"))
            .MapaViaje = val(Lector.GetValue("Penthar", "MapaViaje"))
            .ViajeX = val(Lector.GetValue("Penthar", "ViajeX"))
            .ViajeY = val(Lector.GetValue("Penthar", "ViajeY"))
            .MapaResu = val(Lector.GetValue("Penthar", "MapaResu"))
            .ResuX = val(Lector.GetValue("Penthar", "ResuX"))
            .ResuY = val(Lector.GetValue("Penthar", "ResuY"))
            .NecesitaNave = val(Lector.GetValue("Penthar", "NecesitaNave"))
            MapasCiudades = MapasCiudades & Lector.GetValue("Penthar", "Mapas")
        End With
    
236     With Prision
238         .map = val(Lector.GetValue("Prision", "Mapa"))
240         .x = val(Lector.GetValue("Prision", "X"))
242         .y = val(Lector.GetValue("Prision", "Y"))
        End With
    
244     With Libertad
246         .map = val(Lector.GetValue("Libertad", "Mapa"))
248         .x = val(Lector.GetValue("Libertad", "X"))
250         .y = val(Lector.GetValue("Libertad", "Y"))
        End With
        
         With Renacimiento
            .map = val(Lector.GetValue("Renacimiento", "Mapa"))
            .X = val(Lector.GetValue("Renacimiento", "X"))
            .y = val(Lector.GetValue("Renacimiento", "Y"))
        End With
        
        With BarcoNavegando
            .Map = val(Lector.GetValue("BarcoNavegando", "Mapa"))
            .StartX = val(Lector.GetValue("BarcoNavegando", "StartX"))
            .StartY = val(Lector.GetValue("BarcoNavegando", "StartY"))
            .EndX = val(Lector.GetValue("BarcoNavegando", "EndX"))
            .EndY = val(Lector.GetValue("BarcoNavegando", "EndY"))
            .DestX = val(Lector.GetValue("BarcoNavegando", "DestX"))
            .DestY = val(Lector.GetValue("BarcoNavegando", "DestY"))
            .DockX = val(Lector.GetValue("BarcoNavegando", "DockX"))
            .DockY = val(Lector.GetValue("BarcoNavegando", "DockY"))
            .RequiredPassID = val(Lector.GetValue("BarcoNavegando", "RequiredPassID"))
        End With
        
        With ForgatDock
            .Map = val(Lector.GetValue("ForgatDock", "Mapa"))
            .StartX = val(Lector.GetValue("ForgatDock", "StartX"))
            .StartY = val(Lector.GetValue("ForgatDock", "StartY"))
            .EndX = val(Lector.GetValue("ForgatDock", "EndX"))
            .EndY = val(Lector.GetValue("ForgatDock", "EndY"))
            .DestX = val(Lector.GetValue("ForgatDock", "DestX"))
            .DestY = val(Lector.GetValue("ForgatDock", "DestY"))
            .DockX = val(Lector.GetValue("ForgatDock", "DockX"))
            .DockY = val(Lector.GetValue("ForgatDock", "DockY"))
            .RequiredPassID = val(Lector.GetValue("BarcoNavegando", "RequiredPassID"))
        End With
        
        With NixDock
            .Map = val(Lector.GetValue("NixDock", "Mapa"))
            .StartX = val(Lector.GetValue("NixDock", "StartX"))
            .StartY = val(Lector.GetValue("NixDock", "StartY"))
            .EndX = val(Lector.GetValue("NixDock", "EndX"))
            .EndY = val(Lector.GetValue("NixDock", "EndY"))
            .DestX = val(Lector.GetValue("NixDock", "DestX"))
            .DestY = val(Lector.GetValue("NixDock", "DestY"))
            .DockX = val(Lector.GetValue("NixDock", "DockX"))
            .DockY = val(Lector.GetValue("NixDock", "DockY"))
            .RequiredPassID = val(Lector.GetValue("BarcoNavegando", "RequiredPassID"))
        End With
        
        
        TotalMapasCiudades = Split(MapasCiudades, ",")
    
252     Set Lector = Nothing
    
254     Nix.map = CityNix.map
256     Nix.x = CityNix.x
258     Nix.y = CityNix.y
    
260     Ullathorpe.map = CityUllathorpe.map
262     Ullathorpe.x = CityUllathorpe.x
264     Ullathorpe.y = CityUllathorpe.y
    
266     Banderbill.map = CityBanderbill.map
268     Banderbill.x = CityBanderbill.x
270     Banderbill.y = CityBanderbill.y
    
272     Lindos.map = CityLindos.map
274     Lindos.x = CityLindos.x
276     Lindos.y = CityLindos.y
    
278     Arghal.map = CityArghal.map
280     Arghal.x = CityArghal.x
282     Arghal.y = CityArghal.y
    
284     Arkhein.map = CityArkhein.map
286     Arkhein.x = CityArkhein.x
288     Arkhein.y = CityArkhein.y
    
        'Esto es para el /HOGAR
290     Ciudades(e_Ciudad.cNix) = Nix
292     Ciudades(e_Ciudad.cUllathorpe) = Ullathorpe
294     Ciudades(e_Ciudad.cBanderbill) = Banderbill
296     Ciudades(e_Ciudad.cLindos) = Lindos
298     Ciudades(e_Ciudad.cArghal) = Arghal
300     Ciudades(e_Ciudad.cArkhein) = Arkhein
        
        Exit Sub

CargarCiudades_Err:
302     Call TraceError(Err.Number, Err.Description, "ES.CargarCiudades", Erl)
        
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
    
124     IntervaloSed = val(Lector.GetValue("INTERVALOS", "IntervaloSed")) / 25
126     FrmInterv.txtIntervaloSed.Text = IntervaloSed
    
128     IntervaloHambre = val(Lector.GetValue("INTERVALOS", "IntervaloHambre")) / 25
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
162     FrmInterv.txtTimeoutPrimerPaquete.Text = TimeoutPrimerPaquete / 25
    
164     TimeoutEsperandoLoggear = val(Lector.GetValue("INTERVALOS", "TimeoutEsperandoLoggear"))
166     FrmInterv.txtTimeoutEsperandoLoggear.Text = TimeoutEsperandoLoggear / 25
    
168     IntervaloIncineracion = val(Lector.GetValue("INTERVALOS", "IntervaloFuego"))
170     FrmInterv.txtintervalofuego.Text = IntervaloIncineracion
    
172     IntervaloTirar = val(Lector.GetValue("INTERVALOS", "IntervaloTirar"))
174     FrmInterv.txtintervalotirar.Text = IntervaloTirar

176     IntervaloMeditar = val(Lector.GetValue("INTERVALOS", "IntervaloMeditar"))
178     FrmInterv.txtIntervaloMeditar.Text = IntervaloMeditar
    
180     IntervaloCaminar = val(Lector.GetValue("INTERVALOS", "IntervaloCaminar"))
182     FrmInterv.txtintervalocaminar.Text = IntervaloCaminar
        
184     IntervaloEnCombate = val(Lector.GetValue("INTERVALOS", "IntervaloEnCombate"))
    
        '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    
186     IntervaloUserPuedeCastear = val(Lector.GetValue("INTERVALOS", "IntervaloLanzaHechizo"))
188     FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear
    
190     frmMain.TIMER_AI.Interval = val(Lector.GetValue("INTERVALOS", "IntervaloNpcAI"))
192     FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval
    
194     IntervaloTrabajarExtraer = val(Lector.GetValue("INTERVALOS", "IntervaloTrabajarExtraer"))
196     FrmInterv.txtTrabajoExtraer.Text = IntervaloTrabajarExtraer

198     IntervaloTrabajarConstruir = val(Lector.GetValue("INTERVALOS", "IntervaloTrabajarConstruir"))
200     FrmInterv.txtTrabajoConstruir.Text = IntervaloTrabajarConstruir
    
202     IntervaloUserPuedeAtacar = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeAtacar"))
204     FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar
    
        'TODO : Agregar estos intervalos al form!!!
206     IntervaloMagiaGolpe = val(Lector.GetValue("INTERVALOS", "IntervaloMagiaGolpe"))
208     IntervaloGolpeMagia = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeMagia"))
    
        'frmMain.tLluvia.Interval = val(Lector.GetValue("INTERVALOS", "IntervaloPerdidaStaminaLluvia"))
        'FrmInterv.txtIntervaloPerdidaStaminaLluvia.Text = frmMain.tLluvia.Interval
        
        
210     MinutosWs = val(Lector.GetValue("INTERVALOS", "IntervaloWS"))

212     If MinutosWs < 1 Then MinutosWs = 10
    
214     IntervaloCerrarConexion = val(Lector.GetValue("INTERVALOS", "IntervaloCerrarConexion"))
216     IntervaloUserPuedeUsarU = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeUsarU"))
218     IntervaloUserPuedeUsarClic = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeUsarClic"))
220     IntervaloFlechasCazadores = val(Lector.GetValue("INTERVALOS", "IntervaloFlechasCazadores"))
222     IntervaloGolpeUsar = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeUsar"))

224     IntervaloOculto = val(Lector.GetValue("INTERVALOS", "IntervaloOculto"))

    
228     IntervaloPuedeSerAtacado = val(Lector.GetValue("INTERVALOS", "IntervaloPuedeSerAtacado"))

230     IntervaloGuardarUsuarios = val(Lector.GetValue("INTERVALOS", "IntervaloGuardarUsuarios"))

234     IntervaloTimerGuardarUsuarios = val(Lector.GetValue("INTERVALOS", "IntervaloTimerGuardarUsuarios"))

236     IntervaloMensajeGlobal = val(Lector.GetValue("INTERVALOS", "IntervaloMensajeGlobal"))
    
        '&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    
238     Set Lector = Nothing

        
        Exit Sub

LoadIntervalos_Err:
240     Call TraceError(Err.Number, Err.Description, "ES.LoadIntervalos", Erl)

        
End Sub

Sub LoadConfiguraciones()
        
        On Error GoTo LoadConfiguraciones_Err
        
        Dim Leer As clsIniManager
100     Set Leer = New clsIniManager

102     Call Leer.Initialize(IniPath & "Configuracion.ini")

104     ExpMult = val(Leer.GetValue("CONFIGURACIONES", "ExpMult"))
106     OroMult = val(Leer.GetValue("CONFIGURACIONES", "OroMult"))
108     DropMult = val(Leer.GetValue("DROPEO", "DropMult"))
110     DropActive = val(Leer.GetValue("DROPEO", "DropActive"))
112     RecoleccionMult = val(Leer.GetValue("CONFIGURACIONES", "RecoleccionMult"))
113     OroPorNivelBilletera = val(Leer.GetValue("CONFIGURACIONES", "OroPorNivelBilletera"))

114     TimerLimpiarObjetos = val(Leer.GetValue("CONFIGURACIONES", "TimerLimpiarObjetos"))
116     OroPorNivel = val(Leer.GetValue("CONFIGURACIONES", "OroPorNivel"))

118     DuracionDia = val(Leer.GetValue("CONFIGURACIONES", "DuracionDia")) * 60 * 1000 ' De minutos a milisegundos

120     CostoPerdonPorCiudadano = val(Leer.GetValue("CONFIGURACIONES", "CostoPerdonPorCiudadano"))

122     MaximoSpeedHack = val(Leer.GetValue("ANTICHEAT", "MaximoSpeedHack"))


126     Set Leer = Nothing

128     Call CargarEventos
130     Call CargarInfoRetos
132     Call CargarInfoEventos
134     Call CargarMapasEspeciales

        Exit Sub

LoadConfiguraciones_Err:
136     Call TraceError(Err.Number, Err.Description, "ES.LoadConfiguraciones", Erl)

        
End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
        '*****************************************************************
        'Escribe VAR en un archivo
        '*****************************************************************
        
        On Error GoTo WriteVar_Err
        

100     writeprivateprofilestring Main, Var, value, File
    
        
        Exit Sub

WriteVar_Err:
102     Call TraceError(Err.Number, Err.Description, "ES.WriteVar", Erl)

        
End Sub

Function LoadUser(ByVal userIndex As Integer) As Boolean


        On Error GoTo ErrorHandler
    
        LoadUser = LoadCharacterFromDB(userIndex)
            
        Exit Function

ErrorHandler:
        Call TraceError(Err.Number, Err.Description & " UserName: " & UserList(userIndex).Name, "ES.LoadUser", Erl)
        LoadUser = False
        
End Function

Sub SaveUser(ByVal userIndex As Integer, Optional ByVal Logout As Boolean = False)
On Error GoTo SaveUser_Err
        If Logout Then
            Call UserDisconnected(UserList(UserIndex).pos.map, UserIndex)
        End If
        Call SaveCharacterDB(userIndex)
        If Logout Then
            Call RemoveTokenDatabase(userIndex)
        End If
        UserList(userIndex).Counters.LastSave = GetTickCount
        Exit Sub

SaveUser_Err:
        Call TraceError(Err.Number, Err.Description, "ES.SaveUser", Erl)
End Sub

Public Sub RemoveTokenDatabase(ByVal userIndex As Integer)
    Call Execute("delete from tokens where id =  ?;", UserList(userIndex).encrypted_session_token_db_id)
End Sub

Public Sub AddTokenDatabase(ByVal encrypted_token As String, ByVal decrypted_token As String, ByVal username As String)
#If UNIT_TEST = 1 Then
    'Only used in automated unit testing to create a valid session so that we can then try LoginNewChar and
    'LoginExistingChar
    Call Execute("insert into tokens (encrypted_token, decrypted_token, username, remote_host, timestamp) values(?,?,?,""127.0.0.1"",""123456"") ;", encrypted_token, decrypted_token, username)
#End If
End Sub

Sub SaveNewUser(ByVal userIndex As Integer)
    On Error GoTo SaveNewUser_Err
            
    Call SaveNewCharacterDB(userIndex)
    
    Exit Sub

SaveNewUser_Err:
102     Call TraceError(Err.Number, Err.Description, "ES.SaveNewUser", Erl)

        
End Sub

Function Status(ByVal userIndex As Integer) As e_Facciones
        
        On Error GoTo Status_Err
        

100     Status = UserList(userIndex).Faccion.Status

        
        Exit Function

Status_Err:
102     Call TraceError(Err.Number, Err.Description, "ES.Status", Erl)

        
End Function

Sub BackUPnPc(npcIndex As Integer)
        
        On Error GoTo BackUPnPc_Err
        

        Dim NpcNumero As Integer

        Dim npcfile   As String

        Dim LoopC     As Integer

100     NpcNumero = NpcList(npcIndex).Numero

        'If NpcNumero > 499 Then
        '    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
        'Else
102     npcfile = DatPath & "bkNPCs.dat"
        'End If

        'General
104     Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", NpcList(npcIndex).Name)
106     Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", NpcList(npcIndex).Desc)
108     Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(NpcList(npcIndex).Char.head))
110     Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(NpcList(npcIndex).Char.body))
112     Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(NpcList(npcIndex).Char.Heading))
114     Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(NpcList(npcIndex).Movement))
116     Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(NpcList(npcIndex).Attackable))
118     Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(NpcList(npcIndex).Comercia))
120     Call WriteVar(npcfile, "NPC" & NpcNumero, "Craftea", val(NpcList(npcIndex).Craftea))
122     Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(NpcList(npcIndex).TipoItems))
124     Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(NpcList(npcIndex).Hostile))
126     Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(NpcList(npcIndex).GiveEXP))
128     Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", val(NpcList(npcIndex).GiveGLD))
130     Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(NpcList(npcIndex).Hostile))
132     Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(NpcList(npcIndex).InvReSpawn))
134     Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(NpcList(npcIndex).npcType))

        'Stats
136     Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(NpcList(npcIndex).flags.AIAlineacion))
138     Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(NpcList(npcIndex).Stats.def))
140     Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(NpcList(npcIndex).Stats.MaxHit))
142     Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(NpcList(npcIndex).Stats.MaxHp))
144     Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(NpcList(npcIndex).Stats.MinHIT))
146     Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(NpcList(npcIndex).Stats.MinHp))
148     Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(NpcList(npcIndex).Stats.UsuariosMatados)) 'Que es ESTO?!!

        'Flags
150     Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(NpcList(npcIndex).flags.Respawn))
152     Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(NpcList(npcIndex).flags.backup))
154     Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(NpcList(npcIndex).flags.Domable))

        'Inventario
156     Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(NpcList(npcIndex).Invent.NroItems))

158     If NpcList(npcIndex).Invent.NroItems > 0 Then

160         For LoopC = 1 To MAX_INVENTORY_SLOTS
162             Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & LoopC, NpcList(npcIndex).Invent.Object(LoopC).objIndex & "-" & NpcList(npcIndex).Invent.Object(LoopC).amount)
            Next

        End If

        
        Exit Sub

BackUPnPc_Err:
164     Call TraceError(Err.Number, Err.Description, "ES.BackUPnPc", Erl)

        
End Sub

Sub CargarNpcBackUp(npcIndex As Integer, ByVal NpcNumber As Integer)
        
        On Error GoTo CargarNpcBackUp_Err
        

        'Status
100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup Npc"

        Dim npcfile As String

        'If NpcNumber > 499 Then
        '    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
        'Else
102     npcfile = DatPath & "bkNPCs.dat"
        'End If

104     NpcList(npcIndex).Numero = NpcNumber
106     NpcList(npcIndex).Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
108     NpcList(npcIndex).Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
110     Call SetMovement(NpcIndex, val(GetVar(npcfile, "NPC" & NpcNumber, "Movement")))
112     NpcList(npcIndex).npcType = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))

114     NpcList(npcIndex).Char.body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
116     NpcList(npcIndex).Char.head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
118     NpcList(npcIndex).Char.Heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))

120     NpcList(npcIndex).Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
122     NpcList(npcIndex).Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
124     NpcList(npcIndex).Craftea = val(GetVar(npcfile, "NPC" & NpcNumber, "Craftea"))
126     NpcList(npcIndex).Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
128     NpcList(npcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))

130     NpcList(npcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))

132     NpcList(npcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))

134     NpcList(npcIndex).Stats.MaxHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
136     NpcList(npcIndex).Stats.MinHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
138     NpcList(npcIndex).Stats.MaxHit = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
140     NpcList(npcIndex).Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
142     NpcList(npcIndex).Stats.def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
144     NpcList(npcIndex).flags.AIAlineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))

        Dim LoopC As Integer

        Dim ln    As String

146     NpcList(npcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))

148     If NpcList(npcIndex).Invent.NroItems > 0 Then

150         For LoopC = 1 To MAX_INVENTORY_SLOTS
152             ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
154             NpcList(npcIndex).Invent.Object(LoopC).objIndex = val(ReadField(1, ln, 45))
156             NpcList(npcIndex).Invent.Object(LoopC).amount = val(ReadField(2, ln, 45))
       
158         Next LoopC

        Else

160         For LoopC = 1 To MAX_INVENTORY_SLOTS
162             NpcList(npcIndex).Invent.Object(LoopC).objIndex = 0
164             NpcList(npcIndex).Invent.Object(LoopC).amount = 0
166         Next LoopC

        End If

168     NpcList(npcIndex).flags.NPCActive = True
170     NpcList(npcIndex).flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
172     NpcList(npcIndex).flags.backup = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
174     NpcList(npcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
176     NpcList(npcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))

        'Tipo de items con los que comercia
178     NpcList(npcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

        
        Exit Sub

CargarNpcBackUp_Err:
180     Call TraceError(Err.Number, Err.Description, "ES.CargarNpcBackUp", Erl)

        
End Sub



Sub LogBanFromName(ByVal BannedName As String, ByVal userIndex As Integer, ByVal Motivo As String)
        
        On Error GoTo LogBanFromName_Err
        

100     Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(userIndex).Name)
102     Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", Motivo)

        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim mifile As Integer

104     mifile = FreeFile
106     Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
108     Print #mifile, BannedName
110     Close #mifile

        
        Exit Sub

LogBanFromName_Err:
112     Call TraceError(Err.Number, Err.Description, "ES.LogBanFromName", Erl)

        
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
112     Call TraceError(Err.Number, Err.Description, "ES.Ban", Erl)

        
End Sub

Public Sub CargaApuestas()
        
        On Error GoTo CargaApuestas_Err
        

100     Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
102     Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
104     Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

        
        Exit Sub

CargaApuestas_Err:
106     Call TraceError(Err.Number, Err.Description, "ES.CargaApuestas", Erl)

        
End Sub

Public Sub LoadRecursosEspeciales()
        
        On Error GoTo LoadRecursosEspeciales_Err
        

100     If Not FileExist(DatPath & "RecursosEspeciales.dat", vbArchive) Then
102         ReDim EspecialesTala(0) As t_Obj
104         ReDim EspecialesPesca(0) As t_Obj
            Exit Sub

        End If

        Dim IniFile As clsIniManager

106     Set IniFile = New clsIniManager
    
108     Call IniFile.Initialize(DatPath & "RecursosEspeciales.dat")
    
        Dim Count As Long, i As Long, str As String, Field() As String
    
        ' Tala
110     Count = val(IniFile.GetValue("Tala", "Items"))

112     If Count > 0 Then
114         ReDim EspecialesTala(1 To Count) As t_Obj

116         For i = 1 To Count
118             str = IniFile.GetValue("Tala", "Item" & i)
120             Field = Split(str, "-")
            
122             EspecialesTala(i).objIndex = val(Field(0))
124             EspecialesTala(i).Data = val(Field(1))      ' Probabilidad
            Next
        Else
126         ReDim EspecialesTala(0) As t_Obj

        End If
    
        ' Pesca
128     Count = val(IniFile.GetValue("Pesca", "Items"))

130     If Count > 0 Then
132         ReDim EspecialesPesca(1 To Count) As t_Obj

134         For i = 1 To Count
136             str = IniFile.GetValue("Pesca", "Item" & i)
138             Field = Split(str, "-")
            
140             EspecialesPesca(i).objIndex = val(Field(0))
142             EspecialesPesca(i).Data = val(Field(1))     ' Probabilidad
            Next
        Else
144         ReDim EspecialesPesca(0) As t_Obj

        End If
    
146     Set IniFile = Nothing

        
        Exit Sub

LoadRecursosEspeciales_Err:
148     Call TraceError(Err.Number, Err.Description, "ES.LoadRecursosEspeciales", Erl)

        
End Sub

Public Sub LoadPesca()
        
        On Error GoTo LoadPesca_Err
        

100     If Not FileExist(DatPath & "pesca.dat", vbArchive) Then
102         ReDim Peces(0) As t_Obj
            ReDim PecesEspeciales(0) As t_Obj
104         ReDim PesoPeces(0) As Long
            Exit Sub

        End If

        Dim IniFile As clsIniManager

106     Set IniFile = New clsIniManager
    
108     Call IniFile.Initialize(DatPath & "pesca.dat")
    
        Dim Count As Long, CountEspecial As Long, i As Long, j As Long, str As String, Field() As String, nivel As Integer, MaxLvlCania As Long

110     Count = val(IniFile.GetValue("PECES", "NumPeces"))
112     MaxLvlCania = val(IniFile.GetValue("PECES", "Maxlvlcaña"))
        CountEspecial = 1
114     ReDim PesoPeces(0 To MaxLvlCania) As Long
    
116     If Count > 0 Then
118         ReDim Peces(1 To Count) As t_Obj

            ' Cargo todos los peces
120         For i = 1 To Count
122             str = IniFile.GetValue("PECES", "Pez" & i)
124             Field = Split(str, "-")
                
                'HarThaoS: Si es un pez especial lo guardo en otro array
                If val(Field(3)) = 1 Then
                    ReDim Preserve PecesEspeciales(1 To CountEspecial) As t_Obj
                    PecesEspeciales(CountEspecial).objIndex = val(Field(0))
                    PecesEspeciales(CountEspecial).Data = val(Field(1))
                    nivel = val(Field(2))               ' Nivel de caña
                    
                    If (nivel > MaxLvlCania) Then nivel = MaxLvlCania
                    
                    PecesEspeciales(CountEspecial).amount = nivel
                    CountEspecial = CountEspecial + 1
                End If
126             Peces(i).objIndex = val(Field(0))
128             Peces(i).Data = val(Field(1))       ' Peso
130             nivel = val(Field(2))               ' Nivel de caña

132             If (nivel > MaxLvlCania) Then nivel = MaxLvlCania
134             Peces(i).amount = nivel
                
            Next

            ' Los ordeno segun nivel de caña (quick sort)
136         Call QuickSortPeces(1, Count)

            ' Sumo los pesos
138         For i = 1 To Count
140             For j = Peces(i).amount To MaxLvlCania
142                 PesoPeces(j) = PesoPeces(j) + Peces(i).Data
144             Next j

146             Peces(i).Data = PesoPeces(Peces(i).amount)
148         Next i
        Else
150         ReDim Peces(0) As t_Obj

        End If
    
152     Set IniFile = Nothing

        
        Exit Sub

LoadPesca_Err:
154     Call TraceError(Err.Number, Err.Description, "ES.LoadPesca", Erl)

        
End Sub

' Adaptado de https://www.vbforums.com/showthread.php?231925-VB-Quick-Sort-algorithm-(very-fast-sorting-algorithm)
Private Sub QuickSortPeces(ByVal First As Long, ByVal Last As Long)
        
        On Error GoTo QuickSortPeces_Err
        

        Dim Low      As Long, High As Long

        Dim MidValue As Long

        Dim aux      As t_Obj
    
100     Low = First
102     High = Last
104     MidValue = Peces((First + Last) \ 2).amount
    
        Do

106         While Peces(Low).amount < MidValue

108             Low = Low + 1
            Wend

110         While Peces(High).amount > MidValue

112             High = High - 1
            Wend

114         If Low <= High Then
116             aux = Peces(Low)
118             Peces(Low) = Peces(High)
120             Peces(High) = aux
122             Low = Low + 1
124             High = High - 1

            End If

126     Loop While Low <= High
    
128     If First < High Then QuickSortPeces First, High
130     If Low < Last Then QuickSortPeces Low, Last

        
        Exit Sub

QuickSortPeces_Err:
132     Call TraceError(Err.Number, Err.Description, "ES.QuickSortPeces", Erl)

        
End Sub

' Adaptado de https://www.freevbcode.com/ShowCode.asp?ID=9416
Public Function BinarySearchPeces(ByVal value As Long) As Long
        
        On Error GoTo BinarySearchPeces_Err
        

        Dim Low  As Long

        Dim High As Long

100     Low = 1
102     High = UBound(Peces)

        Dim i              As Long

        Dim valor_anterior As Long
    
104     Do While Low <= High
106         i = (Low + High) \ 2

108         If i > 1 Then
110             valor_anterior = Peces(i - 1).Data
            Else
112             valor_anterior = 0
            End If

114         If value >= valor_anterior And value < Peces(i).Data Then
116             BinarySearchPeces = i
                Exit Do
            
118         ElseIf value < valor_anterior Then
120             High = (i - 1)
            
            Else
122             Low = (i + 1)

            End If

        Loop

        
        Exit Function

BinarySearchPeces_Err:
124     Call TraceError(Err.Number, Err.Description, "ES.BinarySearchPeces", Erl)

        
End Function

Public Sub LoadRangosFaccion()
            On Error GoTo LoadRangosFaccion_Err

100         If Not FileExist(DatPath & "rangos_faccion.dat", vbArchive) Then
102             ReDim RangosFaccion(0) As t_RangoFaccion
                Exit Sub

            End If

        Dim IniFile As clsIniManager
104     Set IniFile = New clsIniManager

106         Call IniFile.Initialize(DatPath & "rangos_faccion.dat")

            Dim i As Byte, rankData() As String

108         MaxRangoFaccion = val(IniFile.GetValue("INIT", "NumRangos"))

110         If MaxRangoFaccion > 0 Then
                ' Los rangos de la Armada se guardan en los indices impar, y los del caos en indices pares.
                ' Luego, para acceder es tan facil como usar el Rango directamente para la Armada, y multiplicar por 2 para el Caos.
112             ReDim RangosFaccion(1 To MaxRangoFaccion * 2) As t_RangoFaccion

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
140         Call TraceError(Err.Number, Err.Description, "ES.LoadRangosFaccion", Erl)


End Sub

Public Sub LoadRecompensasFaccion()
            On Error GoTo LoadRecompensasFaccion_Err

100         If Not FileExist(DatPath & "recompensas_faccion.dat", vbArchive) Then
102             ReDim RecompensasFaccion(0) As t_RecompensaFaccion
                Exit Sub

            End If

        Dim IniFile As clsIniManager
104     Set IniFile = New clsIniManager

106         Call IniFile.Initialize(DatPath & "recompensas_faccion.dat")

            Dim cantidadRecompensas As Byte, i As Integer, rank_and_objindex() As String

108         cantidadRecompensas = val(IniFile.GetValue("INIT", "NumRecompensas"))

110         If cantidadRecompensas > 0 Then
112             ReDim RecompensasFaccion(1 To cantidadRecompensas) As t_RecompensaFaccion

114             For i = 1 To cantidadRecompensas
116                 rank_and_objindex = Split(IniFile.GetValue("Recompensas", "Recompensa" & i), "-", , vbTextCompare)

118                 RecompensasFaccion(i).rank = val(rank_and_objindex(0))
120                 RecompensasFaccion(i).objIndex = val(rank_and_objindex(1))
122             Next i

            End If

124         Set IniFile = Nothing

            Exit Sub

LoadRecompensasFaccion_Err:
126         Call TraceError(Err.Number, Err.Description, "ES.LoadRecompensasFaccion", Erl)


End Sub


Public Sub LoadUserIntervals(ByVal userIndex As Integer)
        
        On Error GoTo LoadUserIntervals_Err
        

100     With UserList(userIndex)
            If False Then '.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios) Then
                .Intervals.Arco = 50
                .Intervals.Caminar = IntervaloCaminar
                .Intervals.Golpe = 50
                .Intervals.Magia = 50
                .Intervals.GolpeMagia = 50
                .Intervals.MagiaGolpe = 50
                .Intervals.GolpeUsar = 0
                .Intervals.TrabajarExtraer = IntervaloTrabajarExtraer
                .Intervals.TrabajarConstruir = IntervaloTrabajarConstruir
                .Intervals.UsarU = 50
                .Intervals.UsarClic = 50
            Else
102             .Intervals.Arco = IntervaloFlechasCazadores
104             .Intervals.Caminar = IntervaloCaminar
106             .Intervals.Golpe = IntervaloUserPuedeAtacar
108             .Intervals.Magia = IntervaloUserPuedeCastear
110             .Intervals.GolpeMagia = IntervaloGolpeMagia
112             .Intervals.MagiaGolpe = IntervaloMagiaGolpe
114             .Intervals.GolpeUsar = IntervaloGolpeUsar
116             .Intervals.TrabajarExtraer = IntervaloTrabajarExtraer
118             .Intervals.TrabajarConstruir = IntervaloTrabajarConstruir
120             .Intervals.UsarU = IntervaloUserPuedeUsarU
122             .Intervals.UsarClic = IntervaloUserPuedeUsarClic
            
            End If

        End With

        
        Exit Sub

LoadUserIntervals_Err:
124     Call TraceError(Err.Number, Err.Description, "ES.LoadUserIntervals", Erl)

        
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
110     Call TraceError(Err.Number, Err.Description, "ES.CountFiles", Erl)

        
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


Public Sub CargarDonadores()
100     If Not FileExist(DatPath & "donadores.dat", vbArchive) Then
            Exit Sub
        End If
        Dim IniFile As clsIniManager
106     Set IniFile = New clsIniManager
108     Call IniFile.Initialize(DatPath & "donadores.dat")
        Dim cantidadDonadores As Integer
        cantidadDonadores = val(IniFile.GetValue("INIT", "Cantidad"))
        ReDim lstUsuariosDonadores(0 To cantidadDonadores)
        If cantidadDonadores > 0 Then
            Dim i As Integer
            For i = 1 To cantidadDonadores
                lstUsuariosDonadores(i) = IniFile.GetValue("DONADOR", "Donador" & i)
            Next i
        End If
End Sub

Public Function IsFeatureEnabled(ByVal featureName As String)
    If FeatureToggles.Exists(featureName) Then
        IsFeatureEnabled = FeatureToggles.Item(featureName)
    Else
        IsFeatureEnabled = False
    End If
End Function

Public Sub SetFeatureToggle(ByVal Name As String, ByVal State As Boolean)
    If FeatureToggles.Exists(Name) Then
        FeatureToggles.Remove Name
    End If
    Call FeatureToggles.Add(Name, State)
End Sub
