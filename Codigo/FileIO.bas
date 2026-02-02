Attribute VB_Name = "ES"
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
Const MAX_RANDOM_TELEPORT_IN_MAP = 20

Private Type t_Position
    x As Integer
    y As Integer
End Type

'Item type
Private Type t_Item
    ObjIndex As Integer
    amount As Integer
End Type

Private Type t_WorldPos
    Map As Integer
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
    NpcIndex As Integer
End Type

Private Type t_DatosObjs
    x As Integer
    y As Integer
    ObjIndex As Integer
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

Private MapSize        As t_MapSize
Private MapDat         As t_MapDat
Private FeatureToggles As Dictionary

Public Sub load_stats()
    On Error GoTo error_load_stats
    Dim n       As Integer
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
    If EOF(n) Then
        RecordUsuarios = 1
    Else
        Line Input #n, str
        RecordUsuarios = val(str)
    End If
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
    n = val(GetVar(DatPath & "npcs.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(n) As t_CriaturasEntrenador
    For LoopC = 1 To n
        SpawnList(LoopC).NpcIndex = LoopC
        SpawnList(LoopC).NpcName = GetVar(DatPath & "npcs.dat", "NPC" & LoopC, "Name")
        SpawnList(LoopC).PuedeInvocar = val(GetVar(DatPath & "npcs.dat", "NPC" & LoopC, "PuedeInvocar")) = 1
        If Len(SpawnList(LoopC).NpcName) = 0 Then
            SpawnList(LoopC).NpcName = "Nada"
        End If
    Next LoopC
    Exit Sub
CargarSpawnList_Err:
    Call TraceError(Err.Number, Err.Description, "ES.CargarSpawnList", Erl)
End Sub

Function EsAdmin(ByRef name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    On Error GoTo EsAdmin_Err
    EsAdmin = (val(Administradores.GetValue("Admin", name)) = 1)
    Exit Function
EsAdmin_Err:
    Call TraceError(Err.Number, Err.Description, "ES.EsAdmin", Erl)
End Function

Function EsDios(ByRef name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    On Error GoTo EsDios_Err
    EsDios = (val(Administradores.GetValue("Dios", name)) = 1)
    Exit Function
EsDios_Err:
    Call TraceError(Err.Number, Err.Description, "ES.EsDios", Erl)
End Function

Function EsSemiDios(ByRef name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    On Error GoTo EsSemiDios_Err
    EsSemiDios = (val(Administradores.GetValue("SemiDios", name)) = 1)
    Exit Function
EsSemiDios_Err:
    Call TraceError(Err.Number, Err.Description, "ES.EsSemiDios", Erl)
End Function

Function EsConsejero(ByRef name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    On Error GoTo EsConsejero_Err
    EsConsejero = (val(Administradores.GetValue("Consejero", name)) = 1)
    Exit Function
EsConsejero_Err:
    Call TraceError(Err.Number, Err.Description, "ES.EsConsejero", Erl)
End Function

Function EsRolesMaster(ByRef name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    On Error GoTo EsRolesMaster_Err
    EsRolesMaster = (val(Administradores.GetValue("RM", name)) = 1)
    Exit Function
EsRolesMaster_Err:
    Call TraceError(Err.Number, Err.Description, "ES.EsRolesMaster", Erl)
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
    EsGM = EsAdmin(name)
    ' Dios?
    If Not EsGM Then EsGM = EsDios(name)
    ' Semidios?
    If Not EsGM Then EsGM = EsSemiDios(name)
    ' Consejero?
    If Not EsGM Then EsGM = EsConsejero(name)
    EsGmChar = EsGM
    Exit Function
EsGmChar_Err:
    Call TraceError(Err.Number, Err.Description, "ES.EsGmChar", Erl)
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
    Set AdministratorAccounts = New Dictionary
    Dim TempName() As String
    ' Public container
    Set Administradores = New clsIniManager
    ' Server ini info file
    Dim ServerIni As clsIniManager
    Set ServerIni = New clsIniManager
    Debug.Assert FileExist(IniPath & "Server.ini")
    Call ServerIni.Initialize(IniPath & "Server.ini")
    ' Admines
    buf = val(ServerIni.GetValue("INIT", "Admines"))
    For i = 1 To buf
        name = UCase$(ServerIni.GetValue("Admines", "Admin" & i))
        TempName = Split(name, "|", , vbTextCompare)
        ' Si NO declara el mail de la cuenta en el Server.ini, NO le doy privilegios.
        If UBound(TempName()) > 0 Then
            ' AdministratorAccounts("Nick") = "Email"
            AdministratorAccounts(TempName(0)) = TempName(1)
            ' Add key
            Call Administradores.ChangeValue("Admin", TempName(0), "1")
        End If
    Next i
    ' Dioses
    buf = val(ServerIni.GetValue("INIT", "Dioses"))
    For i = 1 To buf
        name = UCase$(ServerIni.GetValue("Dioses", "Dios" & i))
        TempName = Split(name, "|", , vbTextCompare)
        ' Si NO declara el mail de la cuenta en el Server.ini, NO le doy privilegios.
        If UBound(TempName()) > 0 Then
            ' AdministratorAccounts("Nick") = "Email"
            AdministratorAccounts(TempName(0)) = TempName(1)
            ' Add key
            Call Administradores.ChangeValue("Dios", TempName(0), "1")
        End If
    Next i
    ' SemiDioses
    buf = val(ServerIni.GetValue("INIT", "SemiDioses"))
    For i = 1 To buf
        name = UCase$(ServerIni.GetValue("SemiDioses", "SemiDios" & i))
        TempName = Split(name, "|", , vbTextCompare)
        ' Si NO declara el mail de la cuenta en el Server.ini, NO le doy privilegios.
        If UBound(TempName()) > 0 Then
            ' AdministratorAccounts("Nick") = "Email"
            AdministratorAccounts(TempName(0)) = TempName(1)
            ' Add key
            Call Administradores.ChangeValue("SemiDios", TempName(0), "1")
        End If
    Next i
    ' Consejeros
    buf = val(ServerIni.GetValue("INIT", "Consejeros"))
    For i = 1 To buf
        name = UCase$(ServerIni.GetValue("Consejeros", "Consejero" & i))
        TempName = Split(name, "|", , vbTextCompare)
        ' Si NO declara el mail de la cuenta en el Server.ini, NO le doy privilegios.
        If UBound(TempName()) > 0 Then
            ' AdministratorAccounts("Nick") = "Email"
            AdministratorAccounts(TempName(0)) = TempName(1)
            ' Add key
            Call Administradores.ChangeValue("Consejero", TempName(0), "1")
        End If
    Next i
    ' RolesMasters
    buf = val(ServerIni.GetValue("INIT", "RolesMasters"))
    For i = 1 To buf
        name = UCase$(ServerIni.GetValue("RolesMasters", "RM" & i))
        TempName = Split(name, "|", , vbTextCompare)
        ' Si NO declara el mail de la cuenta en el Server.ini, NO le doy privilegios.
        If UBound(TempName()) > 0 Then
            ' AdministratorAccounts("Nick") = "Email"
            AdministratorAccounts(TempName(0)) = TempName(1)
            ' Add key
            Call Administradores.ChangeValue("RM", TempName(0), "1")
        End If
    Next i
    Set ServerIni = Nothing
    'If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & Time & " - Los Administradores/Dioses/Gms se han cargado correctamente."
    Exit Sub
loadAdministrativeUsers_Err:
    Call TraceError(Err.Number, Err.Description, "ES.loadAdministrativeUsers", Erl)
End Sub

Public Function TxtDimension(ByVal name As String) As Long
    On Error GoTo TxtDimension_Err
    Dim n As Integer, cad As String, Tam As Long
    n = FreeFile(1)
    If FileExist(name, vbArchive) Then
    Open name For Input As #n
    Tam = 0
    Do While Not EOF(n)
        Tam = Tam + 1
        Line Input #n, cad
    Loop
    Close n
    Else
        Debug.print "No existe el archivo " & name
    End If

    TxtDimension = Tam
    Exit Function
TxtDimension_Err:
    Call TraceError(Err.Number, Err.Description, "ES.TxtDimension", Erl)
End Function

Public Sub CargarForbidenWords()
    On Error GoTo CargarForbidenWords_Err
    Dim Size As Integer
    Size = TxtDimension(DatPath & "NombresInvalidos.txt")
    If Size = 0 Then
        ReDim ForbidenNames(0)
        Exit Sub
    End If
    ReDim ForbidenNames(1 To Size)
    Dim n As Integer, i As Integer
    n = FreeFile(1)
    Open DatPath & "NombresInvalidos.txt" For Input As #n
    For i = 1 To UBound(ForbidenNames)
        Line Input #n, ForbidenNames(i)
        ForbidenNames(i) = LCase$(ForbidenNames(i))
    Next i
    Close n
    Exit Sub
CargarForbidenWords_Err:
    Call TraceError(Err.Number, Err.Description, "ES.CargarForbidenWords", Erl)
End Sub

Public Sub LoadBlockedWordsDescription()
    On Error GoTo LoadBlockedWordsDescription_Err
    Dim Size As Integer
    Size = TxtDimension(DatPath & "BlockedWordsDescription.txt")
    If Size = 0 Then
        ReDim BlockedWordsDescription(0)
        Exit Sub
    End If
    ReDim BlockedWordsDescription(1 To Size)
    Dim n As Integer, i As Integer
    n = FreeFile(1)
    Open DatPath & "BlockedWordsDescription.txt" For Input As #n
    For i = LBound(BlockedWordsDescription) To UBound(BlockedWordsDescription)
        Line Input #n, BlockedWordsDescription(i)
        BlockedWordsDescription(i) = LCase$(BlockedWordsDescription(i))
    Next i
    Close n
    Exit Sub
LoadBlockedWordsDescription_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadBlockedWordsDescription", Erl)
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
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."
    Dim Hechizo As Integer
    Dim Leer    As New clsIniManager
    Call Leer.Initialize(DatPath & "Hechizos.dat")
    'obtiene el numero de hechizos
    NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))
    ReDim Hechizos(1 To NumeroHechizos) As t_Hechizo
    frmCargando.cargar.Min = 0
    frmCargando.cargar.max = NumeroHechizos
    frmCargando.cargar.value = 0
    'Llena la lista
    For Hechizo = 1 To NumeroHechizos
        Hechizos(Hechizo).velocidad = val(Leer.GetValue("Hechizo" & Hechizo, "Velocidad"))
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
        Hechizos(Hechizo).NecesitaObj = val(Leer.GetValue("Hechizo" & Hechizo, "NecesitaObj"))
        Hechizos(Hechizo).NecesitaObj2 = val(Leer.GetValue("Hechizo" & Hechizo, "NecesitaObj2"))
        Hechizos(Hechizo).Tipo = val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
        Hechizos(Hechizo).SkillType = val(Leer.GetValue("Hechizo" & Hechizo, "SkillType"))
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
        If val(Leer.GetValue("Hechizo" & Hechizo, "Incinera")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.Incinerate)
        If val(Leer.GetValue("Hechizo" & Hechizo, "RemoveDebuff")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.RemoveDebuff)
        If val(Leer.GetValue("Hechizo" & Hechizo, "StealBuff")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.StealBuff)
        Hechizos(Hechizo).AutoLanzar = val(Leer.GetValue("Hechizo" & Hechizo, "AutoLanzar"))
        Hechizos(Hechizo).TargetEffectType = val(Leer.GetValue("Hechizo" & Hechizo, "TargetEffectType"))
        Hechizos(Hechizo).Cooldown = val(Leer.GetValue("Hechizo" & Hechizo, "CoolDown"))
        Hechizos(Hechizo).CdEffectId = val(Leer.GetValue("Hechizo" & Hechizo, "CdEffectId"))
        Hechizos(Hechizo).loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
        Hechizos(Hechizo).MinHp = val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
        Hechizos(Hechizo).MaxHp = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
        Hechizos(Hechizo).SubeMana = val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
        Hechizos(Hechizo).MinMana = val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
        Hechizos(Hechizo).MaxMana = val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
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
        If val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.Invisibility)
        If val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.Paralize)
        If val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.Immobilize)
        If val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.RemoveParalysis)
        If val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.RemoveDumb)
        If val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.RemoveInvisibility)
        If val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.CurePoison)
        Hechizos(Hechizo).Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
        If val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.Curse)
        If val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.RemoveCurse)
        If val(Leer.GetValue("Hechizo" & Hechizo, "Revivir")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.Resurrect)
        If val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.Blindness)
        If val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.Dumb)
        If val(Leer.GetValue("Hechizo" & Hechizo, "ToggleCleave")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.ToggleCleave)
        If val(Leer.GetValue("Hechizo" & Hechizo, "ToggleDivineBlood")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.ToggleDivineBlood)
        If val(Leer.GetValue("Hechizo" & Hechizo, "AdjustStatsWithCaster")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.AdjustStatsWithCaster)
        If val(Leer.GetValue("Hechizo" & Hechizo, "CancelActiveEffect")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.CancelActiveEffect)
        Hechizos(Hechizo).Invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
        Hechizos(Hechizo).NumNpc = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
        Hechizos(Hechizo).cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
        Hechizos(Hechizo).Mimetiza = val(Leer.GetValue("Hechizo" & Hechizo, "Mimetiza"))
        If val(Leer.GetValue("Hechizo" & Hechizo, "GolpeCertero")) > 0 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.PreciseHit)
        Hechizos(Hechizo).MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
        Hechizos(Hechizo).ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
        Hechizos(Hechizo).RequiredHP = val(Leer.GetValue("Hechizo" & Hechizo, "RequiredHP"))
        Hechizos(Hechizo).Duration = val(Leer.GetValue("Hechizo" & Hechizo, "Duration"))
        'Barrin 30/9/03
        Hechizos(Hechizo).StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
        Hechizos(Hechizo).Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
        Hechizos(Hechizo).RequireTransform = val(Leer.GetValue("Hechizo" & Hechizo, "RequireTransform"))
        frmCargando.cargar.value = frmCargando.cargar.value + 1
        Hechizos(Hechizo).NeedStaff = val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
        Hechizos(Hechizo).RequiereInstrumento = val(Leer.GetValue("Hechizo" & Hechizo, "RequiereInstrumento"))
        Hechizos(Hechizo).StaffAffected = CBool(val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))
        Hechizos(Hechizo).EotId = val(Leer.GetValue("Hechizo" & Hechizo, "EOTID"))
        Hechizos(Hechizo).MaxLevelCasteable = val(Leer.GetValue("Hechizo" & Hechizo, "MaxLevelCasteable"))
        If val(Leer.GetValue("Hechizo" & Hechizo, "RequireArmor")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eArmor)
        If val(Leer.GetValue("Hechizo" & Hechizo, "RequireShip")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eShip)
        If val(Leer.GetValue("Hechizo" & Hechizo, "RequireHelm")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eHelm)
        If val(Leer.GetValue("Hechizo" & Hechizo, "RequireKnucle")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eKnucle)
        If val(Leer.GetValue("Hechizo" & Hechizo, "RequireMagicItem")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eMagicItem)
        If val(Leer.GetValue("Hechizo" & Hechizo, "RequireProjectile")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eProjectile)
        If val(Leer.GetValue("Hechizo" & Hechizo, "RequireShield")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eShield)
        If val(Leer.GetValue("Hechizo" & Hechizo, "RequireWeapon")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eWeapon)
        If val(Leer.GetValue("Hechizo" & Hechizo, "RequireTargetOnLand")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eRequireTargetOnLand)
        If val(Leer.GetValue("Hechizo" & Hechizo, "RequireTargetOnWater")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, _
                e_SpellRequirementMask.eRequireTargetOnWater)
        If val(Leer.GetValue("Hechizo" & Hechizo, "WorkOnDead")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eWorkOnDead)
        If val(Leer.GetValue("Hechizo" & Hechizo, "IsSkill")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eIsSkill)
        If val(Leer.GetValue("Hechizo" & Hechizo, "IsBindable")) > 0 Then Call SetMask(Hechizos(Hechizo).SpellRequirementMask, e_SpellRequirementMask.eIsBindable)
        Hechizos(Hechizo).RequireWeaponType = val(Leer.GetValue("Hechizo" & Hechizo, "RequireWeaponType"))
        Hechizos(Hechizo).IsElementalTagsOnly = val(Leer.GetValue("Hechizo" & Hechizo, "IsElementalTagsOnly")) > 0
        Dim SubeHP As Byte
        SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
        If SubeHP = 1 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.eDoHeal)
        If SubeHP = 2 Then Call SetMask(Hechizos(Hechizo).Effects, e_SpellEffects.eDoDamage)
    Next Hechizo
    Set Leer = Nothing
    Exit Sub
ErrHandler:
    MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.Description
End Sub

Public Sub LoadEffectOverTime()
    On Error GoTo ErrHandler
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."
    Dim i           As Integer
    Dim Leer        As New clsIniManager
    Dim EffectCount As Integer
    Call Leer.Initialize(DatPath & "EffectsOverTime.dat")
    'obtiene el numero de hechizos
    EffectCount = val(Leer.GetValue("INIT", "EffectCount"))
    ReDim EffectOverTime(1 To EffectCount) As t_EffectOverTime
    frmCargando.cargar.Min = 0
    frmCargando.cargar.max = EffectCount
    frmCargando.cargar.value = 0
    For i = 1 To EffectCount
        EffectOverTime(i).Type = val(Leer.GetValue("EOT" & i, "Type"))
        EffectOverTime(i).SubType = val(Leer.GetValue("EOT" & i, "SubType"))
        EffectOverTime(i).SharedTypeId = val(Leer.GetValue("EOT" & i, "SharedTypeId"))
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
        EffectOverTime(i).buffType = val(Leer.GetValue("EOT" & i, "BuffType"))
        EffectOverTime(i).Area = val(Leer.GetValue("EOT" & i, "Area"))
        EffectOverTime(i).Aura = Leer.GetValue("EOT" & i, "Aura")
        EffectOverTime(i).ApplyEffectId = val(Leer.GetValue("EOT" & i, "ApplyeffectID"))
        EffectOverTime(i).SecondaryEffectId = val(Leer.GetValue("EOT" & i, "SecondaryEffectId"))
        EffectOverTime(i).RequireTransform = val(Leer.GetValue("EOT" & i, "RequireTransform"))
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
        If val(Leer.GetValue("EOT" & i, "RequireArmor")) > 0 Then Call SetMask(EffectOverTime(i).SpellRequirementMask, e_SpellRequirementMask.eArmor)
        If val(Leer.GetValue("EOT" & i, "RequireShip")) > 0 Then Call SetMask(EffectOverTime(i).SpellRequirementMask, e_SpellRequirementMask.eShip)
        If val(Leer.GetValue("EOT" & i, "RequireHelm")) > 0 Then Call SetMask(EffectOverTime(i).SpellRequirementMask, e_SpellRequirementMask.eHelm)
        If val(Leer.GetValue("EOT" & i, "RequireKnucle")) > 0 Then Call SetMask(EffectOverTime(i).SpellRequirementMask, e_SpellRequirementMask.eKnucle)
        If val(Leer.GetValue("EOT" & i, "RequireMagicItem")) > 0 Then Call SetMask(EffectOverTime(i).SpellRequirementMask, e_SpellRequirementMask.eMagicItem)
        If val(Leer.GetValue("EOT" & i, "RequireProjectile")) > 0 Then Call SetMask(EffectOverTime(i).SpellRequirementMask, e_SpellRequirementMask.eProjectile)
        If val(Leer.GetValue("EOT" & i, "RequireShield")) > 0 Then Call SetMask(EffectOverTime(i).SpellRequirementMask, e_SpellRequirementMask.eShield)
        If val(Leer.GetValue("EOT" & i, "RequireWeapon")) > 0 Then Call SetMask(EffectOverTime(i).SpellRequirementMask, e_SpellRequirementMask.eWeapon)
        EffectOverTime(i).NpcId = val(Leer.GetValue("EOT" & i, "NpcId"))
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
        EffectOverTime(i).SecondaryTargetModifier = val(Leer.GetValue("EOT" & i, "SecondaryTargetModifier"))
    Next i
    Call InitializePools
    Exit Sub
ErrHandler:
    MsgBox "Error cargando EffectsOverTime.dat " & Err.Number & ": " & Err.Description
End Sub

Sub LoadMotd()
    On Error GoTo LoadMotd_Err
    Dim i As Integer
    MaxLines = val(GetVar(DatPath & "Motd.ini", "INIT", "NumLines"))
    ReDim MOTD(1 To MaxLines)
    For i = 1 To MaxLines
        MOTD(i).texto = GetVar(DatPath & "Motd.ini", "Motd", "Line" & i)
        MOTD(i).Formato = vbNullString
    Next i
    Exit Sub
LoadMotd_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadMotd", Erl)
End Sub

Public Sub DoBackUp()
    On Error GoTo DoBackUp_Err
    haciendoBK = True
    Dim i As Integer
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    haciendoBK = False
    Call LogThis(0, "[BackUps.log] DoBackUp", vbLogEventTypeInformation)
    Exit Sub
DoBackUp_Err:
    Call TraceError(Err.Number, Err.Description, "ES.DoBackUp", Erl)
End Sub

Sub LoadArmasHerreria()
    On Error GoTo LoadArmasHerreria_Err
    Dim n As Integer, lc As Integer
    n = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))
    If n = 0 Then
        ReDim ArmasHerrero(0) As Integer
        Exit Sub
    End If
    ReDim Preserve ArmasHerrero(1 To n) As Integer
    For lc = 1 To n
        ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
    Next lc
    Exit Sub
LoadArmasHerreria_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadArmasHerreria", Erl)
End Sub

Sub LoadArmadurasHerreria()
    On Error GoTo LoadArmadurasHerreria_Err
    Dim n As Integer, lc As Integer
    n = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))
    If n = 0 Then
        ReDim ArmadurasHerrero(0) As Integer
        Exit Sub
    End If
    ReDim Preserve ArmadurasHerrero(1 To n) As Integer
    For lc = 1 To n
        ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
    Next lc
    Exit Sub
LoadArmadurasHerreria_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadArmadurasHerreria", Erl)
End Sub

Sub LoadBlackSmithElementalRunes()
    On Error GoTo LoadRunasHerreria_Err
    Dim n As Integer, lc As Integer
    n = val(GetVar(DatPath & "BlackSmithElementalRunes.dat", "INIT", "NumRunas"))
    If n = 0 Then
        ReDim BlackSmithElementalRunes(0) As Integer
        Exit Sub
    End If
    ReDim Preserve BlackSmithElementalRunes(1 To n) As Integer
    For lc = 1 To n
        BlackSmithElementalRunes(lc) = val(GetVar(DatPath & "BlackSmithElementalRunes.dat", "Runa" & lc, "Index"))
    Next lc
    Exit Sub
LoadRunasHerreria_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadBlackSmithElementalRunes", Erl)
End Sub

Sub LoadBalance()
    On Error GoTo LoadBalance_Err
    Dim BalanceIni As clsIniManager
    Set BalanceIni = New clsIniManager
    BalanceIni.Initialize DatPath & "Balance.dat"
    Dim i, j As Long
    Dim SearchVar As String
    'Modificadores de Clase
    For i = 1 To NUMCLASES
        SearchVar = Replace$(Tilde(ListaClases(i)), " ", vbNullString)
        With ModClase(i)
            .Evasion = val(BalanceIni.GetValue("MODEVASION", SearchVar))
            .AtaqueArmas = val(BalanceIni.GetValue("MODATAQUEARMAS", SearchVar))
            .AtaqueProyectiles = val(BalanceIni.GetValue("MODATAQUEPROYECTILES", SearchVar))
            .DañoArmas = val(BalanceIni.GetValue("MODDANOARMAS", SearchVar))
            .DañoProyectiles = val(BalanceIni.GetValue("MODDANOPROYECTILES", SearchVar))
            .DañoWrestling = val(BalanceIni.GetValue("MODDANOWRESTLING", SearchVar))
            .Escudo = val(BalanceIni.GetValue("MODESCUDO", SearchVar))
            .ModApunalar = val(BalanceIni.GetValue("MODAPUNALAR", SearchVar, 1))
            .ModStabbingNPCMin = val(BalanceIni.GetValue("MODAPUNALARNPCMIN", SearchVar, 1))
            .ModStabbingNPCMax = val(BalanceIni.GetValue("MODAPUNALARNPCMAX", SearchVar, 1))
            .Vida = val(BalanceIni.GetValue("MODVIDA", SearchVar))
            .ManaInicial = val(BalanceIni.GetValue("MANA_INICIAL", SearchVar))
            .MultMana = val(BalanceIni.GetValue("MULT_MANA", SearchVar))
            .AumentoSta = val(BalanceIni.GetValue("AUMENTO_STA", SearchVar))
            .HitPre36 = val(BalanceIni.GetValue("GOLPE_PRE_36", SearchVar))
            .HitPost36 = val(BalanceIni.GetValue("GOLPE_POST_36", SearchVar))
            .ResistenciaMagica = val(BalanceIni.GetValue("MODRESISTENCIAMAGICA", SearchVar))
            .LevelSkillPoints = val(BalanceIni.GetValue("MODSKILLPOINTS", SearchVar))
            For j = 1 To eWeaponTypeCount - 1
                .WeaponHitBonus(j) = val(BalanceIni.GetValue(SearchVar, WeaponTypeNames(j)))
            Next j
        End With
    Next i
    'Modificadores de Raza
    For i = 1 To NUMRAZAS
        SearchVar = Replace$(Tilde(ListaRazas(i)), " ", vbNullString)
        With ModRaza(i)
            .Fuerza = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Fuerza"))
            .Agilidad = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Agilidad"))
            .Inteligencia = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Inteligencia"))
            .Carisma = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Carisma"))
            .Constitucion = val(BalanceIni.GetValue("MODRAZA", SearchVar + "Constitucion"))
        End With
    Next i
    CriticalHitDmgModifier = val(BalanceIni.GetValue("BACKSTAB", "CriticalHitDmgModifier"))
    IgnoreArmorChance = val(BalanceIni.GetValue("BACKSTAB", "IgnoreArmorChance"))
    ExtraBackstabChance = val(BalanceIni.GetValue("BACKSTAB", "ExtraBackstabChance"))
    AssasinStabbingChance = val(BalanceIni.GetValue("BACKSTAB", "AssasinStabbingChance"))
    HunterStabbingChance = val(BalanceIni.GetValue("BACKSTAB", "HunterStabbingChance"))
    BardStabbingChance = val(BalanceIni.GetValue("BACKSTAB", "BardStabbingChance"))
    GenericStabbingChance = val(BalanceIni.GetValue("BACKSTAB", "GenericStabbingChance"))
    BanditCriticalHitChance = val(BalanceIni.GetValue("BACKSTAB", "BanditCriticalHitChance"))
    'Extra
    PorcentajeRecuperoMana = val(BalanceIni.GetValue("EXTRA", "PorcentajeRecuperoMana"))
    RecoveryMana = val(BalanceIni.GetValue("EXTRA", "RecoveryMana"))
    MultiplierManaxSkills = val(BalanceIni.GetValue("EXTRA", "MultiplierManaxSkills"))
    ManaCommonLute = val(BalanceIni.GetValue("EXTRA", "ManaCommonLute"))
    ManaMagicLute = val(BalanceIni.GetValue("EXTRA", "ManaMagicLute"))
    ManaElvenLute = val(BalanceIni.GetValue("EXTRA", "ManaElvenLute"))
    DificultadSubirSkill = val(BalanceIni.GetValue("EXTRA", "DificultadSubirSkill"))
    InfluenciaPromedioVidas = val(BalanceIni.GetValue("EXTRA", "InfluenciaPromedioVidas"))
    DesbalancePromedioVidas = val(BalanceIni.GetValue("EXTRA", "DesbalancePromedioVidas"))
    RangoVidas = val(BalanceIni.GetValue("EXTRA", "RangoVidas"))
    CapVidaMax = val(BalanceIni.GetValue("EXTRA", "CapVidaMax"))
    CapVidaMin = val(BalanceIni.GetValue("EXTRA", "CapVidaMin"))
    RequiredSpellDisplayTime = val(BalanceIni.GetValue("EXTRA", "RequiredSpellDisplayTime"))
    MaxInvisibleSpellDisplayTime = val(BalanceIni.GetValue("EXTRA", "MaxInvisibleSpellDisplayTime"))
    MultiShotReduction = val(BalanceIni.GetValue("EXTRA", "MultiShotReduction"))
    HomeTimer = val(BalanceIni.GetValue("EXTRA", "HomeTimer"))
    HomeTimerAdventurer = val(BalanceIni.GetValue("EXTRA", "HomeTimerAdventurer"))
    HomeTimerHero = val(BalanceIni.GetValue("EXTRA", "HomeTimerHero"))
    HomeTimerLegend = val(BalanceIni.GetValue("EXTRA", "HomeTimerLegend"))
    MagicSkillBonusDamageModifier = val(BalanceIni.GetValue("EXTRA", "MagicSkillBonusDamageModifier"))
    MRSkillProtectionModifier = val(BalanceIni.GetValue("EXTRA", "MagicResistanceSkillProtectionModifier"))
    MRSkillNpcProtectionModifier = val(BalanceIni.GetValue("EXTRA", "MagicResistanceSkillProtectionModifierNpc"))
    AssistDamageValidTime = val(BalanceIni.GetValue("EXTRA", "AssistDamageValidTime"))
    AssistHelpValidTime = val(BalanceIni.GetValue("EXTRA", "AssistHelpValidTime"))
    HideAfterHitTime = val(BalanceIni.GetValue("EXTRA", "HideAfterHitTime"))
    FactionReKillTime = val(BalanceIni.GetValue("EXTRA", "FactionReKillTime"))
    AirHitReductParalisisTime = val(BalanceIni.GetValue("EXTRA", "AirHitReductParalisisTime"))
    PorcentajePescaSegura = val(BalanceIni.GetValue("EXTRA", "PorcentajePescaSegura"))
    DivineBloodHealingMultiplierBonus = val(BalanceIni.GetValue("EXTRA", "DivineBloodHealingMultiplierBonus"))
    DivineBloodManaCostMultiplier = val(BalanceIni.GetValue("EXTRA", "DivineBloodManaCostMultiplier"))
    WarriorLifeStealOnHitMultiplier = val(BalanceIni.GetValue("EXTRA", "WarriorLifeStealOnHitMultiplier"))
    'stun
    PlayerStunTime = val(BalanceIni.GetValue("STUN", "PlayerStunTime"))
    NpcStunTime = val(BalanceIni.GetValue("STUN", "NpcStunTime"))
    PlayerInmuneTime = val(BalanceIni.GetValue("STUN", "PlayerInmuneTime"))
    ' Exp
    For i = 1 To STAT_MAXELV
        ExpLevelUp(i) = val(BalanceIni.GetValue("EXP", i))
    Next
    'ElementalMatrixForNpcs
    Dim vals() As String
    Dim row    As String
    For i = 0 To MAX_ELEMENT_TAGS - 1
        row = (CStr(BalanceIni.GetValue("ElementalMatrixForNpcs", "Row" & i + 1, "1")))
        vals = Split(row, " ")
        For j = 0 To MAX_ELEMENT_TAGS - 1
            ElementalMatrixForNpcs(i + 1, j + 1) = val(vals(j))
        Next j
    Next i
    Set BalanceIni = Nothing
    AgregarAConsola "Se cargó el balance (Balance.dat)"
    Exit Sub
LoadBalance_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadBalance", Erl)
End Sub

Sub LoadObjCarpintero()
    On Error GoTo LoadObjCarpintero_Err
    Dim n As Integer, lc As Integer
    n = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))
    If n = 0 Then
        ReDim ObjCarpintero(0) As Integer
        Exit Sub
    End If
    ReDim Preserve ObjCarpintero(1 To n) As Integer
    For lc = 1 To n
        ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
    Next lc
    Exit Sub
LoadObjCarpintero_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadObjCarpintero", Erl)
End Sub

Sub LoadObjAlquimista()
    On Error GoTo LoadObjAlquimista_Err
    Dim n As Integer, lc As Integer
    n = val(GetVar(DatPath & "ObjAlquimista.dat", "INIT", "NumObjs"))
    If n = 0 Then
        ReDim ObjAlquimista(0) As Integer
        Exit Sub
    End If
    ReDim Preserve ObjAlquimista(1 To n) As Integer
    For lc = 1 To n
        ObjAlquimista(lc) = val(GetVar(DatPath & "ObjAlquimista.dat", "Obj" & lc, "Index"))
    Next lc
    Exit Sub
LoadObjAlquimista_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadObjAlquimista", Erl)
End Sub

Sub LoadObjSastre()
    On Error GoTo LoadObjSastre_Err
    Dim n As Integer, lc As Integer
    n = val(GetVar(DatPath & "ObjSastre.dat", "INIT", "NumObjs"))
    If n = 0 Then
        ReDim ObjSastre(0) As Integer
        Exit Sub
    End If
    ReDim Preserve ObjSastre(1 To n) As Integer
    For lc = 1 To n
        ObjSastre(lc) = val(GetVar(DatPath & "ObjSastre.dat", "Obj" & lc, "Index"))
    Next lc
    Exit Sub
LoadObjSastre_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadObjSastre", Erl)
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
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."
    '*****************************************************************
    'Carga la lista de objetos
    '*****************************************************************
    Dim Object As Integer
    Dim Leer   As clsIniManager
    Set Leer = New clsIniManager
    Call Leer.Initialize(DatPath & "Obj.dat")
    'obtiene el numero de obj
    NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))
    With frmCargando.cargar
        .Min = 0
        .max = NumObjDatas
        .value = 0
    End With
    ReDim Preserve ObjData(1 To NumObjDatas) As t_ObjData
    ReDim ObjShop(1 To 1) As t_ObjData
    Dim ObjKey  As String
    Dim str     As String, Field() As String
    Dim Crafteo As clsCrafteo
    Dim NFT     As Boolean
    'Llena la lista
    For Object = 1 To NumObjDatas
        With ObjData(Object)
            ObjKey = "OBJ" & Object
            .name = Leer.GetValue(ObjKey, "Name")
            .Log = val(Leer.GetValue(ObjKey, "Log"))
            .NoLog = val(Leer.GetValue(ObjKey, "NoLog"))
            '07/09/07
            .GrhIndex = val(Leer.GetValue(ObjKey, "GrhIndex"))
            .OBJType = val(Leer.GetValue(ObjKey, "ObjType"))
            .Newbie = val(Leer.GetValue(ObjKey, "Newbie"))
            'Propiedades by Lader 05-05-08
            .Instransferible = val(Leer.GetValue(ObjKey, "Instransferible"))
            .Destruye = val(Leer.GetValue(ObjKey, "Destruye"))
            .Intirable = val(Leer.GetValue(ObjKey, "Intirable"))
            .CantidadSkill = val(Leer.GetValue(ObjKey, "CantidadSkill"))
            .Que_Skill = val(Leer.GetValue(ObjKey, "QueSkill"))
            .QueAtributo = val(Leer.GetValue(ObjKey, "queatributo"))
            .CuantoAumento = val(Leer.GetValue(ObjKey, "cuantoaumento"))
            .MinELV = val(Leer.GetValue(ObjKey, "MinELV"))
            .MaxLEV = val(Leer.GetValue(ObjKey, "MaxLEV"))
            .InstrumentoRequerido = val(Leer.GetValue(ObjKey, "InstrumentoRequerido"))
            .Subtipo = val(Leer.GetValue(ObjKey, "Subtipo"))
            .Dorada = val(Leer.GetValue(ObjKey, "Dorada"))
            .Blodium = val(Leer.GetValue(ObjKey, "Blodium"))
            .FireEssence = val(Leer.GetValue(ObjKey, "FireEssence"))
            .WaterEssence = val(Leer.GetValue(ObjKey, "WaterEssence"))
            .EarthEssence = val(Leer.GetValue(ObjKey, "EarthEssence"))
            .WindEssence = val(Leer.GetValue(ObjKey, "WindEssence"))
            .VidaUtil = val(Leer.GetValue(ObjKey, "VidaUtil"))
            .TiempoRegenerar = val(Leer.GetValue(ObjKey, "TiempoRegenerar"))
            .Jerarquia = val(Leer.GetValue(ObjKey, "Jerarquia"))
            .Cooldown = val(Leer.GetValue(ObjKey, "CD"))
            .cdType = val(Leer.GetValue(ObjKey, "CDType"))
            .ImprovedRangedHitChance = val(Leer.GetValue(ObjKey, "ImprovedRHit"))
            .ImprovedMeleeHitChance = val(Leer.GetValue(ObjKey, "ImprovedMHit"))
            .ApplyEffectId = val(Leer.GetValue(ObjKey, "ApplyEffectId"))
            .JineteLevel = val(Leer.GetValue(ObjKey, "JineteLevel"))
            .ElementalTags = val(Leer.GetValue(ObjKey, "ElementalTags"))
            .BowCategory = val(Leer.GetValue(ObjKey, "BowCategory"))
            .ArrowCategory = val(Leer.GetValue(ObjKey, "ArrowCategory"))
            If val(Leer.GetValue(ObjKey, "Bindable")) > 0 Then Call SetMask(.ObjFlags, e_ObjFlags.e_Bindable)
            If val(Leer.GetValue(ObjKey, "UseOnSafeAreaOnly")) > 0 Then Call SetMask(.ObjFlags, e_ObjFlags.e_UseOnSafeAreaOnly)
            Dim i As Integer
            Select Case .OBJType
                Case e_OBJType.otWorkingTools
                    .WeaponAnim = val(Leer.GetValue(ObjKey, "Anim"))
                    .Power = val(Leer.GetValue(ObjKey, "Power"))
                Case e_OBJType.otArmor, e_OBJType.otSkinsArmours
                    .RequiereObjeto = val(Leer.GetValue(ObjKey, "RequiereObjeto"))
                    .Real = val(Leer.GetValue(ObjKey, "Real"))
                    .Caos = val(Leer.GetValue(ObjKey, "Caos"))
                    .LeadersOnly = val(Leer.GetValue(ObjKey, "LeadersOnly")) <> 0
                    .ResistenciaMagica = val(Leer.GetValue(ObjKey, "ResistenciaMagica"))
                    .Invernal = val(Leer.GetValue(ObjKey, "Invernal")) > 0
                    .Camouflage = val(Leer.GetValue(ObjKey, "Camouflage")) > 0
                    If val(Leer.GetValue(ObjKey, "velocidad")) = 0 Then
                        .velocidad = 1
                    Else
                        .velocidad = val(Leer.GetValue(ObjKey, "velocidad"))
                    End If
                Case e_OBJType.otShield, e_OBJType.otSkinsShields
                    .RequiereObjeto = val(Leer.GetValue(ObjKey, "RequiereObjeto"))
                    .ShieldAnim = val(Leer.GetValue(ObjKey, "Anim"))
                    .Real = val(Leer.GetValue(ObjKey, "Real"))
                    .Caos = val(Leer.GetValue(ObjKey, "Caos"))
                    .LeadersOnly = val(Leer.GetValue(ObjKey, "LeadersOnly")) <> 0
                    .ResistenciaMagica = val(Leer.GetValue(ObjKey, "ResistenciaMagica"))
                    .Porcentaje = val(Leer.GetValue(ObjKey, "Porcentaje"))
                Case e_OBJType.otHelmet, e_OBJType.otSkinsHelmets
                    .RequiereObjeto = val(Leer.GetValue(ObjKey, "RequiereObjeto"))
                    .CascoAnim = val(Leer.GetValue(ObjKey, "Anim"))
                    .Real = val(Leer.GetValue(ObjKey, "Real"))
                    .Caos = val(Leer.GetValue(ObjKey, "Caos"))
                    .LeadersOnly = val(Leer.GetValue(ObjKey, "LeadersOnly")) <> 0
                    .ResistenciaMagica = val(Leer.GetValue(ObjKey, "ResistenciaMagica"))
                Case e_OBJType.otBackpack, e_OBJType.otSkinsWings
                    .RequiereObjeto = val(Leer.GetValue(ObjKey, "RequiereObjeto"))
                    '.BackpackAnim = val(Leer.GetValue(ObjKey, "Anim"))
                Case e_OBJType.otMagicalInstrument
                    .Revive = val(Leer.GetValue(ObjKey, "Revive")) <> 0
                Case e_OBJType.otWeapon, e_OBJType.otSkinsWeapons
                    .RequiereObjeto = val(Leer.GetValue(ObjKey, "RequiereObjeto"))
                    .WeaponAnim = val(Leer.GetValue(ObjKey, "Anim"))
                    .Apuñala = val(Leer.GetValue(ObjKey, "Apuñala"))
                    .Envenena = val(Leer.GetValue(ObjKey, "Envenena"))
                    .Paraliza = val(Leer.GetValue(ObjKey, "Paraliza"))
                    .Estupidiza = val(Leer.GetValue(ObjKey, "Estupidiza"))
                    .incinera = val(Leer.GetValue(ObjKey, "Incinera"))
                    .MaxHit = val(Leer.GetValue(ObjKey, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue(ObjKey, "MinHIT"))
                    .MinArmorPenetrationFlat = val(Leer.GetValue(ObjKey, "MinArmorPenetrationFlat"))
                    .MaxArmorPenetrationFlat = val(Leer.GetValue(ObjKey, "MaxArmorPenetrationFlat"))
                    .ArmorPenetrationPercent = val(Leer.GetValue(ObjKey, "ArmorPenetrationPercent"))
                    .Proyectil = val(Leer.GetValue(ObjKey, "Proyectil"))
                    .Municion = val(Leer.GetValue(ObjKey, "Municiones"))
                    .Power = val(Leer.GetValue(ObjKey, "StaffPower"))
                    .Real = val(Leer.GetValue(ObjKey, "Real"))
                    .Caos = val(Leer.GetValue(ObjKey, "Caos"))
                    .LeadersOnly = val(Leer.GetValue(ObjKey, "LeadersOnly")) <> 0
                    .EfectoMagico = val(Leer.GetValue(ObjKey, "efectomagico"))
                    .Revive = val(Leer.GetValue(ObjKey, "Revive")) <> 0
                    .DosManos = val(Leer.GetValue(ObjKey, "DosManos"))
                    .Porcentaje = val(Leer.GetValue(ObjKey, "Porcentaje"))
                    .WeaponType = val(Leer.GetValue(ObjKey, "WeaponType"))
                Case e_OBJType.otMusicalInstruments
                    'Pablo (ToxicWaste)
                    .Real = val(Leer.GetValue(ObjKey, "Real"))
                    .Caos = val(Leer.GetValue(ObjKey, "Caos"))
                    .LeadersOnly = val(Leer.GetValue(ObjKey, "LeadersOnly")) <> 0
                Case e_OBJType.otDoors, e_OBJType.otEmptyBottle, e_OBJType.otFullBottle
                    .IndexAbierta = val(Leer.GetValue(ObjKey, "IndexAbierta"))
                    .IndexCerrada = val(Leer.GetValue(ObjKey, "IndexCerrada"))
                    .IndexCerradaLlave = val(Leer.GetValue(ObjKey, "IndexCerradaLlave"))
                Case otPotions
                    .TipoPocion = val(Leer.GetValue(ObjKey, "TipoPocion"))
                    .MaxModificador = val(Leer.GetValue(ObjKey, "MaxModificador"))
                    .MinModificador = val(Leer.GetValue(ObjKey, "MinModificador"))
                    .DuracionEfecto = val(Leer.GetValue(ObjKey, "DuracionEfecto"))
                    .Hechizo = val(Leer.GetValue(ObjKey, "Hechizo"))
                    .Raices = val(Leer.GetValue(ObjKey, "Raices"))
                    .Cuchara = val(Leer.GetValue(ObjKey, "Cuchara"))
                    .Botella = val(Leer.GetValue(ObjKey, "Botella"))
                    .Mortero = val(Leer.GetValue(ObjKey, "Mortero"))
                    .FrascoAlq = val(Leer.GetValue(ObjKey, "FrascoAlq"))
                    .FrascoElixir = val(Leer.GetValue(ObjKey, "FrascoElixir"))
                    .Dosificador = val(Leer.GetValue(ObjKey, "Dosificador"))
                    .Orquidea = val(Leer.GetValue(ObjKey, "Orquidea"))
                    .Carmesi = val(Leer.GetValue(ObjKey, "Carmesi"))
                    .HongoDeLuz = val(Leer.GetValue(ObjKey, "HongoDeLuz"))
                    .Esporas = val(Leer.GetValue(ObjKey, "Esporas"))
                    .Tuna = val(Leer.GetValue(ObjKey, "Tuna"))
                    .Cala = val(Leer.GetValue(ObjKey, "Cala"))
                    .ColaDeZorro = val(Leer.GetValue(ObjKey, "ColaDeZorro"))
                    .FlorOceano = val(Leer.GetValue(ObjKey, "FlorOceano"))
                    .FlorRoja = val(Leer.GetValue(ObjKey, "FlorRoja"))
                    .Hierva = val(Leer.GetValue(ObjKey, "Hierva"))
                    .HojasDeRin = val(Leer.GetValue(ObjKey, "HojasDeRin"))
                    .HojasRojas = val(Leer.GetValue(ObjKey, "HojasRojas"))
                    .SemillasPros = val(Leer.GetValue(ObjKey, "SemillasPros"))
                    .Pimiento = val(Leer.GetValue(ObjKey, "Pimiento"))
                    .SkPociones = val(Leer.GetValue(ObjKey, "SkPociones"))
                    .Porcentaje = val(Leer.GetValue(ObjKey, "Porcentaje"))
                Case e_OBJType.otShips, e_OBJType.otSkinsBoats
                    .RequiereObjeto = val(Leer.GetValue(ObjKey, "RequiereObjeto"))
                    .MaxHit = val(Leer.GetValue(ObjKey, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue(ObjKey, "MinHIT"))
                    .velocidad = val(Leer.GetValue(ObjKey, "Velocidad"))
                Case e_OBJType.otSaddles
                    .MaxHit = val(Leer.GetValue(ObjKey, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue(ObjKey, "MinHIT"))
                    .MinDef = val(Leer.GetValue(ObjKey, "MINDEF"))
                    .MaxDef = val(Leer.GetValue(ObjKey, "MAXDEF"))
                    .Real = val(Leer.GetValue(ObjKey, "Real"))
                    .Caos = val(Leer.GetValue(ObjKey, "Caos"))
                    .LeadersOnly = val(Leer.GetValue(ObjKey, "LeadersOnly")) <> 0
                    .velocidad = val(Leer.GetValue(ObjKey, "Velocidad"))
                Case e_OBJType.otArrows
                    .MaxHit = val(Leer.GetValue(ObjKey, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue(ObjKey, "MinHIT"))
                    .Envenena = val(Leer.GetValue(ObjKey, "Envenena"))
                    .Paraliza = val(Leer.GetValue(ObjKey, "Paraliza"))
                    .Estupidiza = val(Leer.GetValue(ObjKey, "Estupidiza"))
                    .incinera = val(Leer.GetValue(ObjKey, "Incinera"))
                    .EfectoMagico = val(Leer.GetValue(ObjKey, "efectomagico"))
                    .Snd1 = val(Leer.GetValue(ObjKey, "SND1"))
                    .Snd2 = val(Leer.GetValue(ObjKey, "SND2"))
                    'Pasajes Ladder 05-05-08
                Case e_OBJType.otPassageTicket
                    .DesdeMap = val(Leer.GetValue(ObjKey, "DesdeMap"))
                    .HastaMap = val(Leer.GetValue(ObjKey, "Map"))
                    .HastaX = val(Leer.GetValue(ObjKey, "X"))
                    .HastaY = val(Leer.GetValue(ObjKey, "Y"))
                    .NecesitaNave = val(Leer.GetValue(ObjKey, "NecesitaNave"))
                Case e_OBJType.otDonator
                    .HastaMap = val(Leer.GetValue(ObjKey, "Map"))
                    .HastaX = val(Leer.GetValue(ObjKey, "X"))
                    .HastaY = val(Leer.GetValue(ObjKey, "Y"))
                Case e_OBJType.OtQuest
                    .QuestId = val(Leer.GetValue(ObjKey, "QuestID"))
                Case e_OBJType.otAmulets
                    .EfectoMagico = val(Leer.GetValue(ObjKey, "efectomagico"))
                    .Revive = val(Leer.GetValue(ObjKey, "Revive")) <> 0
                    If .EfectoMagico = 15 Then
                        PENDIENTE = Object
                    End If
                    If .EfectoMagico = 12 Then
                        .MaxItems = val(Leer.GetValue(ObjKey, "Peces"))
                    End If
                Case e_OBJType.otRecallStones
                    .TipoRuna = val(Leer.GetValue(ObjKey, "TipoRuna"))
                    .DesdeMap = val(Leer.GetValue(ObjKey, "DesdeMap"))
                    .HastaMap = val(Leer.GetValue(ObjKey, "Map"))
                    .HastaX = val(Leer.GetValue(ObjKey, "X"))
                    .HastaY = val(Leer.GetValue(ObjKey, "Y"))
                Case e_OBJType.otTeleport
                    .Radio = val(Leer.GetValue(ObjKey, "Radio"))
                Case e_OBJType.otChest
                    .CantItem = val(Leer.GetValue(ObjKey, "CantItem"))
                    Select Case .Subtipo
                        Case 1
                            ReDim .Item(1 To .CantItem)
                            For i = 1 To .CantItem
                                .Item(i).ObjIndex = val(Leer.GetValue(ObjKey, "Item" & i))
                                .Item(i).amount = val(Leer.GetValue(ObjKey, "Cantidad" & i))
                            Next i
                        Case 2
                            ReDim .Item(1 To .CantItem)
                            .CantEntrega = val(Leer.GetValue(ObjKey, "CantEntrega"))
                            For i = 1 To .CantItem
                                .Item(i).ObjIndex = val(Leer.GetValue(ObjKey, "Item" & i))
                                .Item(i).amount = val(Leer.GetValue(ObjKey, "Cantidad" & i))
                            Next i
                        Case 3
                            ReDim .Item(1 To .CantItem)
                            For i = 1 To .CantItem
                                .Item(i).ObjIndex = val(Leer.GetValue(ObjKey, "Item" & i))
                                .Item(i).amount = val(Leer.GetValue(ObjKey, "Cantidad" & i))
                                .Item(i).data = 101 - val(Leer.GetValue(ObjKey, "Drop" & i))
                            Next i
                    End Select
                Case e_OBJType.otOreDeposit
                    .MineralIndex = val(Leer.GetValue(ObjKey, "MineralIndex"))
                    ' Drop gemas yacimientos
                    .CantItem = val(Leer.GetValue(ObjKey, "Gemas"))
                    If .CantItem > 0 Then
                        ReDim .Item(1 To .CantItem)
                        For i = 1 To .CantItem
                            str = Leer.GetValue(ObjKey, "Gema" & i)
                            Field = Split(str, "-")
                            .Item(i).ObjIndex = val(Field(0))    ' ObjIndex
                            .Item(i).amount = val(Field(1))      ' Probabilidad de drop (1 en X)
                        Next i
                    End If
                Case e_OBJType.otUsableOntarget
                    .MaxHit = val(Leer.GetValue(ObjKey, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue(ObjKey, "MinHIT"))
                    .Proyectil = val(Leer.GetValue(ObjKey, "Proyectil"))
                Case e_OBJType.otRingAccesory
                    .ResistenciaMagica = val(Leer.GetValue(ObjKey, "ResistenciaMagica"))
                Case e_OBJType.otMinerals
                    .LingoteIndex = val(Leer.GetValue(ObjKey, "LingoteIndex"))
                Case e_OBJType.otUsableOntarget
                    .EfectoMagico = val(Leer.GetValue(ObjKey, "efectomagico"))
                Case e_OBJType.otElementalRune
                    .Hechizo = val(Leer.GetValue(ObjKey, "Hechizo"))
                Case e_OBJType.otParchment, e_OBJType.otSkinsSpells
                    .RequiereObjeto = val(Leer.GetValue(ObjKey, "RequiereObjeto"))
            End Select
            .MagicDamageBonus = val(Leer.GetValue(ObjKey, "MagicDamageBonus"))
            .MagicAbsoluteBonus = val(Leer.GetValue(ObjKey, "MagicAbsoluteBonus"))
            .MagicPenetration = val(Leer.GetValue(ObjKey, "MagicPenetration"))
            .EfectoMagico = val(Leer.GetValue(ObjKey, "EfectoMagico"))
            .ProjectileType = val(Leer.GetValue(ObjKey, "ProjectileType"))
            .MinSkill = val(Leer.GetValue(ObjKey, "MinSkill"))
            .Elfico = val(Leer.GetValue(ObjKey, "Elfico"))
            .Pino = val(Leer.GetValue(ObjKey, "Pino"))
            .Snd1 = val(Leer.GetValue(ObjKey, "SND1"))
            .Snd2 = val(Leer.GetValue(ObjKey, "SND2"))
            .Snd3 = val(Leer.GetValue(ObjKey, "SND3"))
            'DELETE
            .SndAura = val(Leer.GetValue(ObjKey, "SndAura"))
            '
            .NoSeLimpia = val(Leer.GetValue(ObjKey, "NoSeLimpia"))
            .Subastable = val(Leer.GetValue(ObjKey, "Subastable"))
            .ParticulaGolpe = val(Leer.GetValue(ObjKey, "ParticulaGolpe"))
            .ParticulaViaje = val(Leer.GetValue(ObjKey, "ParticulaViaje"))
            .ParticulaGolpeTime = val(Leer.GetValue(ObjKey, "ParticulaGolpeTime"))
            .Ropaje = val(Leer.GetValue(ObjKey, "NumRopaje"))
            .RopajeHumano = val(Leer.GetValue(ObjKey, "RopajeHumano"))
            .RopajeElfo = val(Leer.GetValue(ObjKey, "RopajeElfo"))
            .RopajeElfoOscuro = val(Leer.GetValue(ObjKey, "RopajeElfoOscuro"))
            .RopajeOrco = val(Leer.GetValue(ObjKey, "RopajeOrco"))
            .RopajeEnano = val(Leer.GetValue(ObjKey, "RopajeEnano"))
            .RopajeGnomo = val(Leer.GetValue(ObjKey, "RopajeGnomo"))
            .RopajeHumana = val(Leer.GetValue(ObjKey, "RopajeHumana"))
            .RopajeElfa = val(Leer.GetValue(ObjKey, "RopajeElfa"))
            .RopajeElfaOscura = val(Leer.GetValue(ObjKey, "RopajeElfaOscura"))
            .RopajeOrca = val(Leer.GetValue(ObjKey, "RopajeOrca"))
            .RopajeEnana = val(Leer.GetValue(ObjKey, "RopajeEnana"))
            .RopajeGnoma = val(Leer.GetValue(ObjKey, "RopajeGnoma"))
            .RazaAltos = val(Leer.GetValue(ObjKey, "RazaAltos"))
            .RazaBajos = val(Leer.GetValue(ObjKey, "RazaBajos"))
            .HechizoIndex = val(Leer.GetValue(ObjKey, "HechizoIndex"))
            .MaxHp = val(Leer.GetValue(ObjKey, "MaxHP"))
            .MinHp = val(Leer.GetValue(ObjKey, "MinHP"))
            .Mujer = val(Leer.GetValue(ObjKey, "Mujer"))
            .Hombre = val(Leer.GetValue(ObjKey, "Hombre"))
            .PielLobo = val(Leer.GetValue(ObjKey, "PielLobo"))
            .PielOsoPardo = val(Leer.GetValue(ObjKey, "PielOsoPardo"))
            .PielOsoPolaR = val(Leer.GetValue(ObjKey, "PielOsoPolaR"))
            .PielLoboNegro = val(Leer.GetValue(ObjKey, "PielLoboNegro"))
            .PielTigre = val(Leer.GetValue(ObjKey, "PielTigre"))
            .PielTigreBengala = val(Leer.GetValue(ObjKey, "PielTigreBengala"))
            .SkSastreria = val(Leer.GetValue(ObjKey, "SKSastreria"))
            .LingH = val(Leer.GetValue(ObjKey, "LingH"))
            .LingP = val(Leer.GetValue(ObjKey, "LingP"))
            .LingO = val(Leer.GetValue(ObjKey, "LingO"))
            .Coal = val(Leer.GetValue(ObjKey, "Coal"))
            .SkHerreria = val(Leer.GetValue(ObjKey, "SkHerreria"))
            .CreaParticula = Leer.GetValue(ObjKey, "CreaParticula")
            .CreaFX = val(Leer.GetValue(ObjKey, "CreaFX"))
            .CreaGRH = Leer.GetValue(ObjKey, "CreaGRH")
            .CreaLuz = Leer.GetValue(ObjKey, "CreaLuz")
            .CreaWav = val(Leer.GetValue(ObjKey, "CreaWav"))
            .MinHam = val(Leer.GetValue(ObjKey, "MinHam"))
            .MinSed = val(Leer.GetValue(ObjKey, "MinAgu"))
            .PuntosPesca = val(Leer.GetValue(ObjKey, "PuntosPesca"))
            .MinDef = val(Leer.GetValue(ObjKey, "MINDEF"))
            .MaxDef = val(Leer.GetValue(ObjKey, "MAXDEF"))
            .def = (.MinDef + .MaxDef) / 2
            .ClaseTipo = val(Leer.GetValue(ObjKey, "ClaseTipo"))
            .RazaEnana = val(Leer.GetValue(ObjKey, "RazaEnana"))
            .RazaDrow = val(Leer.GetValue(ObjKey, "RazaDrow"))
            .RazaElfa = val(Leer.GetValue(ObjKey, "RazaElfa"))
            .RazaGnoma = val(Leer.GetValue(ObjKey, "RazaGnoma"))
            .RazaOrca = val(Leer.GetValue(ObjKey, "RazaOrca"))
            .RazaHumana = val(Leer.GetValue(ObjKey, "RazaHumana"))
            .Valor = val(Leer.GetValue(ObjKey, "Valor"))
            .Crucial = val(Leer.GetValue(ObjKey, "Crucial"))
            '.Cerrada = val(Leer.GetValue(ObjKey, "abierta")) cerrada = abierta??? WTF???????
            .Cerrada = val(Leer.GetValue(ObjKey, "Cerrada"))
            If .Cerrada = 1 Then
                .Llave = val(Leer.GetValue(ObjKey, "Llave"))
                .clave = val(Leer.GetValue(ObjKey, "Clave"))
            End If
            'Puertas y llaves
            .clave = val(Leer.GetValue(ObjKey, "Clave"))
            .texto = Leer.GetValue(ObjKey, "Texto")
            .GrhSecundario = val(Leer.GetValue(ObjKey, "VGrande"))
            .Agarrable = val(Leer.GetValue(ObjKey, "Agarrable"))
            .ForoID = Leer.GetValue(ObjKey, "ID")
            'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico  -  Nunca más papu
            Dim n As Integer
            Dim s As String
            For i = 1 To NUMCLASES
                s = UCase$(Leer.GetValue(ObjKey, "CP" & i))
                n = 1
                Do While LenB(s) > 0 And Tilde(ListaClases(n)) <> Trim$(s)
                    n = n + 1
                Loop
                .ClaseProhibida(i) = IIf(LenB(s) > 0, n, 0)
            Next i
            For i = 1 To NUMRAZAS
                s = UCase$(Leer.GetValue(ObjKey, "RP" & i))
                n = 1
                Do While LenB(s) > 0 And Tilde(ListaRazas(n)) <> Trim$(s)
                    n = n + 1
                Loop
                .RazaProhibida(i) = IIf(LenB(s) > 0, n, 0)
            Next i
            ' Skill requerido
            str = Leer.GetValue(ObjKey, "SkillRequerido")
            If Len(str) > 0 Then
                Field = Split(str, "-")
                n = 1
                Do While LenB(Field(0)) > 0 And Tilde(SkillsNames(n)) <> Tilde(Field(0))
                    n = n + 1
                Loop
                .SkillIndex = IIf(LenB(Field(0)) > 0, n, 0)
                .SkillRequerido = val(Field(1))
            End If
            ' -----------------
            .SkCarpinteria = val(Leer.GetValue(ObjKey, "SkCarpinteria"))
            'If .SkCarpinteria > 0 Then
            .Madera = val(Leer.GetValue(ObjKey, "Madera"))
            .MaderaElfica = val(Leer.GetValue(ObjKey, "MaderaElfica"))
            .MaderaPino = val(Leer.GetValue(ObjKey, "Maderapino"))
            'Bebidas
            .MinSta = val(Leer.GetValue(ObjKey, "MinST"))
            .NoSeCae = val(Leer.GetValue(ObjKey, "NoSeCae"))
            ' Crafteos
            If val(Leer.GetValue(ObjKey, "Crafteable")) = 1 Then
                str = Leer.GetValue(ObjKey, "Materiales")
                If LenB(str) Then
                    Field = Split(str, "-", MAX_SLOTS_CRAFTEO)
                    Dim Items() As Integer
                    ReDim Items(1 To UBound(Field) + 1)
                    For i = 0 To UBound(Field)
                        Items(i + 1) = val(Field(i))
                        If Items(i + 1) > UBound(ObjData) Then Items(i + 1) = 0
                    Next
                    Call SortIntegerArray(Items, 1, UBound(Items))
                    Set Crafteo = New clsCrafteo
                    Call Crafteo.SetItems(Items)
                    Crafteo.Tipo = val(Leer.GetValue(ObjKey, "TipoCrafteo"))
                    Crafteo.Probabilidad = Clamp(val(Leer.GetValue(ObjKey, "ProbCrafteo")), 0, 100)
                    Crafteo.precio = val(Leer.GetValue(ObjKey, "CostoCrafteo"))
                    Crafteo.Resultado = Object
                    If Not Crafteos.Exists(Crafteo.Tipo) Then
                        Call Crafteos.Add(Crafteo.Tipo, New Dictionary)
                    End If
                    Dim ItemKey As String
                    ItemKey = GetRecipeKey(Items)
                    If Not Crafteos.Item(Crafteo.Tipo).Exists(ItemKey) Then
                        Call Crafteos.Item(Crafteo.Tipo).Add(ItemKey, Crafteo)
                    End If
                End If
            End If
            ' Catalizadores
            .CatalizadorTipo = val(Leer.GetValue(ObjKey, "CatalizadorTipo"))
            If .CatalizadorTipo Then
                .CatalizadorAumento = val(Leer.GetValue(ObjKey, "CatalizadorAumento"))
            End If
            NFT = val(Leer.GetValue(ObjKey, "NFT"))
            .ObjDonador = NFT
            If NFT Then
                ObjShop(UBound(ObjShop)).name = Leer.GetValue(ObjKey, "Name")
                ObjShop(UBound(ObjShop)).Valor = val(Leer.GetValue(ObjKey, "Valor"))
                ObjShop(UBound(ObjShop)).ObjNum = Object
                ObjShop(UBound(ObjShop)).ObjDonador = 1
                ReDim Preserve ObjShop(1 To (UBound(ObjShop) + 1)) As t_ObjData
            End If
            frmCargando.cargar.value = frmCargando.cargar.value + 1
        End With
        '  Cada 10 objetos revivo la interfaz
        If Object Mod 10 = 0 Then DoEvents
    Next Object
    ReDim Preserve ObjShop(1 To (UBound(ObjShop) - 1)) As t_ObjData
    Set Leer = Nothing
    Call InitTesoro
    Call InitRegalo
    Exit Sub
ErrHandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.Description & ". Error producido al cargar el objeto: " & Object
End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String
    On Error GoTo GetVar_Err
    Dim sSpaces  As String ' This will hold the input that the program will retrieve
    Dim szReturn As String ' This will be the defaul value if the string is not found
    szReturn = vbNullString
    sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
    GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, File
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
    Exit Function
GetVar_Err:
    Call TraceError(Err.Number, Err.Description, "ES.GetVar", Erl)
End Function

Sub CargarBackUp()
    On Error GoTo CargarBackUp_Err
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."
    Dim Map     As Integer
    Dim TempInt As Integer
    Dim npcfile As String
    If RunningInVB() Then
        NumMaps = 869
    Else
        NumMaps = CountFiles(MapPath, "*.csm")
        NumMaps = NumMaps - 1
    End If
    Call InitAreas
    frmCargando.cargar.Min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.value = 0
    frmCargando.ToMapLbl.Visible = True
    ReDim MapData(1 To (NumMaps + InstanceMapCount), XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As t_MapBlock
    ReDim MapInfo(1 To (NumMaps + InstanceMapCount)) As t_MapInfo
    For Map = 1 To NumMaps
        frmCargando.ToMapLbl = Map & "/" & NumMaps
        Call CargarMapaFormatoCSM(Map, App.Path & "\WorldBackUp\Mapa" & Map & ".csm")
        frmCargando.cargar.value = frmCargando.cargar.value + 1
        DoEvents
    Next Map
    'Call generateMatrix(MATRIX_INITIAL_MAP)
    frmCargando.ToMapLbl.Visible = False
    Exit Sub
CargarBackUp_Err:
    Call TraceError(Err.Number, Err.Description, "ES.CargarBackUp", Erl)
End Sub

Sub LoadMapData()
    On Error GoTo man
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas..."
    Dim Map     As Integer
    Dim TempInt As Integer
    Dim npcfile As String
    #If UNIT_TEST = 1 Then
        'We only need 50 maps for unit testing
        NumMaps = 50
        Debug.Print "UNIT_TEST Enabled Loading just " & NumMaps & " maps"
    #ElseIf LOGIN_STRESS_TEST = 1 Then
        NumMaps = 100
    #Else
        If RunningInVB() Then
            'VB runs out of memory when debugging
            NumMaps = 300
        Else
            NumMaps = CountFiles(MapPath, "*.csm") - 1
        End If
    #End If
    Dim NormalMapsCount As Integer: NormalMapsCount = NumMaps
    NumMaps = NormalMapsCount + InstanceMapCount
    Call InitAreas
    frmCargando.cargar.Min = 0
    frmCargando.cargar.max = NormalMapsCount
    frmCargando.cargar.value = 0
    frmCargando.ToMapLbl.Visible = True
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As t_MapBlock
    ReDim MapInfo(1 To NumMaps) As t_MapInfo
    For Map = 1 To NormalMapsCount
        frmCargando.ToMapLbl = Map & "/" & NormalMapsCount
        Call CargarMapaFormatoCSM(Map, MapPath & "Mapa" & Map & ".csm")
        frmCargando.cargar.value = frmCargando.cargar.value + 1
        DoEvents
    Next Map
    frmCargando.ToMapLbl.Visible = False
    Call InstanceManager.InitializeInstanceHeap(InstanceMapCount, NormalMapsCount + 1)
    Exit Sub
man:
    Call MsgBox("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)
End Sub

Public Sub CargarMapaFormatoCSM(ByVal Map As Long, ByVal MAPFl As String)
    On Error GoTo ErrorHandler:
    Dim npcfile                                     As String
    Dim fh                                          As Integer
    Dim MH                                          As t_MapHeader
    Dim Blqs()                                      As t_DatosBloqueados
    Dim L1()                                        As t_DatosGrh
    Dim L2()                                        As t_DatosGrh
    Dim L3()                                        As t_DatosGrh
    Dim L4()                                        As t_DatosGrh
    Dim Triggers()                                  As t_DatosTrigger
    Dim Luces()                                     As t_DatosLuces
    Dim Particulas()                                As t_DatosParticulas
    Dim Objetos()                                   As t_DatosObjs
    Dim NPCs()                                      As t_DatosNPC
    Dim TEs()                                       As t_DatosTE
    Dim RandomTeleports(MAX_RANDOM_TELEPORT_IN_MAP) As Integer
    Dim randomTeleportCount                         As Integer
    Dim body                                        As Integer
    Dim head                                        As Integer
    Dim Heading                                     As Byte
    Dim SailingTiles                                As Long
    Dim TotalTiles                                  As Long
    Dim i                                           As Long
    Dim j                                           As Long
    Dim x                                           As Integer, y As Integer
    randomTeleportCount = 0
    If Not FileExist(MAPFl, vbNormal) Then
        Call TraceError(404, "Estas tratando de cargar un MAPA que NO EXISTE" & vbNewLine & "Mapa: " & MAPFl, "ES.CargarMapaFormatoCSM")
        Exit Sub
    End If
    If FileLen(MAPFl) = 0 Then
        Call TraceError(500, "Se trato de cargar un mapa corrupto o mal generado" & vbNewLine & "Mapa: " & MAPFl, "ES.CargarMapaFormatoCSM")
        Exit Sub
    End If
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
                MapData(Map, Blqs(i).x, Blqs(i).y).Blocked = Blqs(i).Lados
            Next i
        End If
        'Cargamos Layer 1
        If .NumeroLayers(1) > 0 Then
            ReDim L1(1 To .NumeroLayers(1))
            Get #fh, , L1
            For i = 1 To .NumeroLayers(1)
                x = L1(i).x
                y = L1(i).y
                MapData(Map, x, y).Graphic(1) = L1(i).GrhIndex
                TotalTiles = TotalTiles + 1
                If HayAgua(Map, x, y) Then
                    MapData(Map, x, y).Blocked = MapData(Map, x, y).Blocked Or FLAG_AGUA
                    SailingTiles = SailingTiles + 1
                End If
            Next i
        End If
        'Cargamos Layer 2
        If .NumeroLayers(2) > 0 Then
            ReDim L2(1 To .NumeroLayers(2))
            Get #fh, , L2
            For i = 1 To .NumeroLayers(2)
                x = L2(i).x
                y = L2(i).y
                MapData(Map, x, y).Graphic(2) = L2(i).GrhIndex
                MapData(Map, x, y).Blocked = MapData(Map, x, y).Blocked And Not FLAG_AGUA
            Next i
        End If
        If .NumeroLayers(3) > 0 Then
            ReDim L3(1 To .NumeroLayers(3))
            Get #fh, , L3
            For i = 1 To .NumeroLayers(3)
                x = L3(i).x
                y = L3(i).y
                MapData(Map, x, y).Graphic(3) = L3(i).GrhIndex
                If EsArbol(L3(i).GrhIndex) Then
                    MapData(Map, x, y).Blocked = MapData(Map, x, y).Blocked Or FLAG_ARBOL
                End If
            Next i
        End If
        If .NumeroLayers(4) > 0 Then
            ReDim L4(1 To .NumeroLayers(4))
            Get #fh, , L4
            For i = 1 To .NumeroLayers(4)
                MapData(Map, L4(i).x, L4(i).y).Graphic(4) = L4(i).GrhIndex
            Next i
        End If
        If .NumeroTriggers > 0 Then
            ReDim Triggers(1 To .NumeroTriggers)
            Get #fh, , Triggers
            For i = 1 To .NumeroTriggers
                x = Triggers(i).x
                y = Triggers(i).y
                MapData(Map, x, y).trigger = Triggers(i).trigger
                ' Trigger detalles en agua
                If Triggers(i).trigger = e_Trigger.DETALLEAGUA Then
                    ' Vuelvo a poner flag agua
                    MapData(Map, x, y).Blocked = MapData(Map, x, y).Blocked Or FLAG_AGUA
                End If
                If Triggers(i).trigger = e_Trigger.VALIDONADO Or Triggers(i).trigger = e_Trigger.NADOCOMBINADO Or Triggers(i).trigger = e_Trigger.NADOBAJOTECHO Then
                    ' Vuelvo a poner flag agua
                    MapData(Map, x, y).Blocked = MapData(Map, x, y).Blocked Or FLAG_AGUA
                End If
            Next i
        End If
        If .NumeroParticulas > 0 Then
            ReDim Particulas(1 To .NumeroParticulas)
            Get #fh, , Particulas
            For i = 1 To .NumeroParticulas
                MapData(Map, Particulas(i).x, Particulas(i).y).ParticulaIndex = Particulas(i).Particula
                MapData(Map, Particulas(i).x, Particulas(i).y).ParticulaIndex = 0
            Next i
        End If
        If .NumeroLuces > 0 Then
            ReDim Luces(1 To .NumeroLuces)
            Get #fh, , Luces
            For i = 1 To .NumeroLuces
                MapData(Map, Luces(i).x, Luces(i).y).Luz.Color = Luces(i).Color
                MapData(Map, Luces(i).x, Luces(i).y).Luz.Rango = Luces(i).Rango
                MapData(Map, Luces(i).x, Luces(i).y).Luz.Color = 0
                MapData(Map, Luces(i).x, Luces(i).y).Luz.Rango = 0
            Next i
        End If
        If .NumeroOBJs > 0 Then
            ReDim Objetos(1 To .NumeroOBJs)
            Get #fh, , Objetos
            For i = 1 To .NumeroOBJs
                MapData(Map, Objetos(i).x, Objetos(i).y).ObjInfo.ObjIndex = Objetos(i).ObjIndex
                With ObjData(Objetos(i).ObjIndex)
                    Select Case .OBJType
                        Case e_OBJType.otOreDeposit, e_OBJType.otTrees
                            MapData(Map, Objetos(i).x, Objetos(i).y).ObjInfo.amount = ObjData(Objetos(i).ObjIndex).VidaUtil
                            MapData(Map, Objetos(i).x, Objetos(i).y).ObjInfo.data = &H7FFFFFFF ' Ultimo uso = Max Long
                        Case Else
                            MapData(Map, Objetos(i).x, Objetos(i).y).ObjInfo.amount = Objetos(i).ObjAmmount
                    End Select
                    If .OBJType = otTeleport And .Subtipo = e_TeleportSubType.eTransportNetwork Then
                        RandomTeleports(randomTeleportCount) = i
                        randomTeleportCount = randomTeleportCount + 1
                    End If
                End With
            Next i
        End If
        If .NumeroNPCs > 0 Then
            ReDim NPCs(1 To .NumeroNPCs)
            Get #fh, , NPCs
            Dim NumNpc As Integer, NpcIndex As Integer
            For i = 1 To .NumeroNPCs
                NumNpc = NPCs(i).NpcIndex
                If NumNpc > 0 Then
                    npcfile = DatPath & "NPCs.dat"
                    NpcIndex = OpenNPC(NumNpc)
                    If NpcIndex > 0 Then
                        MapData(Map, NPCs(i).x, NPCs(i).y).NpcIndex = NpcIndex
                        NpcList(NpcIndex).pos.Map = Map
                        NpcList(NpcIndex).pos.x = NPCs(i).x
                        NpcList(NpcIndex).pos.y = NPCs(i).y
                        '  guardo siempre la pos original... puede sernos útil ;)
                        NpcList(NpcIndex).Orig = NpcList(NpcIndex).pos
                        If LenB(NpcList(NpcIndex).name) = 0 Then
                            MapData(Map, NPCs(i).x, NPCs(i).y).NpcIndex = 0
                        Else
                            Call MakeNPCChar(True, 0, NpcIndex, Map, NPCs(i).x, NPCs(i).y)
                        End If
                    Else
                        ' Lo guardo en los logs + aparece en el Debug.Print
                        Call TraceError(404, "NPC no existe en los .DAT's o está mal dateado. Posicion: " & Map & "-" & NPCs(i).x & "-" & NPCs(i).y, "ES.CargarMapaFormatoCSM")
                    End If
                End If
            Next i
        End If
        If .NumeroTE > 0 Then
            ReDim TEs(1 To .NumeroTE)
            Get #fh, , TEs
            For i = 1 To .NumeroTE
                MapData(Map, TEs(i).x, TEs(i).y).TileExit.Map = TEs(i).DestM
                MapData(Map, TEs(i).x, TEs(i).y).TileExit.x = TEs(i).DestX
                MapData(Map, TEs(i).x, TEs(i).y).TileExit.y = TEs(i).DestY
            Next i
        End If
    End With
    Close fh
    '  Nuevo sistema de restricciones
    If Not IsNumeric(MapDat.restrict_mode) Then
        ' Solo se usaba el "NEWBIE"
        If UCase$(MapDat.restrict_mode) = "NEWBIE" Then
            MapDat.restrict_mode = "1"
        Else
            MapDat.restrict_mode = "0"
        End If
    End If
    If SailingTiles * 100 / TotalTiles > SvrConfig.GetValue("FISHING_REQUIRED_PERCENT") And Not MapDat.Seguro Then
        Call AddFishingPoolsToMap(Map)
    End If
    MapInfo(Map).map_name = MapDat.map_name
    MapInfo(Map).MapResource = Map
    MapInfo(Map).ambient = MapDat.ambient
    MapInfo(Map).backup_mode = MapDat.backup_mode
    MapInfo(Map).base_light = MapDat.base_light
    MapInfo(Map).Newbie = (val(MapDat.restrict_mode) And 1) <> 0
    MapInfo(Map).SinMagia = (val(MapDat.restrict_mode) And 2) <> 0
    MapInfo(Map).NoPKs = (val(MapDat.restrict_mode) And 4) <> 0
    MapInfo(Map).NoCiudadanos = (val(MapDat.restrict_mode) And 8) <> 0
    MapInfo(Map).SinInviOcul = (val(MapDat.restrict_mode) And 16) <> 0
    MapInfo(Map).SoloClanes = (val(MapDat.restrict_mode) And 32) <> 0
    MapInfo(Map).NoMascotas = (val(MapDat.restrict_mode) And 64) <> 0
    MapInfo(Map).OnlyGroups = (val(MapDat.restrict_mode) And 128) <> 0
    MapInfo(Map).OnlyPatreon = (val(MapDat.restrict_mode) And 256) <> 0
    MapInfo(Map).ResuCiudad = val(GetVar(DatPath & "Map.dat", "RESUCIUDAD", Map)) <> 0
    MapInfo(Map).letter_grh = MapDat.letter_grh
    MapInfo(Map).lluvia = MapDat.lluvia
    MapInfo(Map).music_numberHi = MapDat.music_numberHi
    MapInfo(Map).music_numberLow = MapDat.music_numberLow
    MapInfo(Map).niebla = MapDat.niebla
    MapInfo(Map).Nieve = MapDat.Nieve
    MapInfo(Map).MinLevel = MapDat.level And &HFF
    MapInfo(Map).MaxLevel = (MapDat.level And &HFF00) / &H100
    MapInfo(Map).Seguro = MapDat.Seguro
    MapInfo(Map).terrain = MapDat.terrain
    MapInfo(Map).zone = MapDat.zone
    MapInfo(Map).DropItems = True
    If EsMapaNoDrop(Map) Then
        MapInfo(Map).DropItems = False
    End If
    MapInfo(Map).FriendlyFire = True
    MapInfo(Map).KeepInviOnAttack = val(GetVar(DatPath & "Map.dat", "KeepInviOnAttack", Map)) <> 0
    MapInfo(Map).ForceUpdate = val(GetVar(DatPath & "Map.dat", "ForceUpdateAi", Map)) <> 0
    If LenB(MapDat.Salida) <> 0 Then
        Dim Fields() As String
        Fields = Split(MapDat.Salida, "-")
        MapInfo(Map).Salida.Map = val(Fields(0))
        MapInfo(Map).Salida.x = val(Fields(1))
        MapInfo(Map).Salida.y = val(Fields(2))
    End If
    If randomTeleportCount > 0 Then
        ReDim MapInfo(Map).TransportNetwork(randomTeleportCount - 1) As t_TransportNetworkExit
        For i = 0 To randomTeleportCount - 1
            MapInfo(Map).TransportNetwork(i).TileX = Objetos(RandomTeleports(i)).x
            MapInfo(Map).TransportNetwork(i).TileY = Objetos(RandomTeleports(i)).y
        Next i
    End If
    Exit Sub
ErrorHandler:
    Close fh
    Call TraceError(Err.Number, Err.Description, "ES.CargarMapaFormatoCSM", Erl)
End Sub

Sub AddFishingPoolsToMap(ByVal Map As Integer)
    Dim i As Integer
    For i = 1 To SvrConfig.GetValue("FISHING_TILES_ON_MAP")
        Call CreateFishingPool(Map)
    Next i
End Sub

Public Sub CreateFishingPool(ByVal Map As Integer)
    Dim x, y As Integer
    Do
        x = RandomNumber(12, 88)
        y = RandomNumber(12, 88)
    Loop While MapData(Map, x, y).ObjInfo.ObjIndex <> 0 Or Not HayAgua(Map, x, y)
    MapData(Map, x, y).ObjInfo.ObjIndex = SvrConfig.GetValue("FISHING_POOL_ID")
    MapData(Map, x, y).ObjInfo.amount = ObjData(SvrConfig.GetValue("FISHING_POOL_ID")).VidaUtil
    MapData(Map, x, y).ObjInfo.data = &H7FFFFFFF ' Ultimo uso = Max Long
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
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."
    If Not FileExist(IniPath & "Server.ini", vbArchive) Then
        MsgBox "Se requiere de la configuración en Server.ini", vbCritical + vbOKOnly
        End
    End If
    Set Lector = New clsIniManager
    Call Lector.Initialize(IniPath & "Server.ini")
    'Misc
    BootDelBackUp = val(Lector.GetValue("INIT", "IniciarDesdeBackUp"))
    'Directorios
    DatPath = Lector.GetValue("DIRECTORIOS", "DatPath")
    MapPath = Lector.GetValue("DIRECTORIOS", "MapPath")
    CharPath = Lector.GetValue("DIRECTORIOS", "CharPath")
    DeletePath = Lector.GetValue("DIRECTORIOS", "DeletePath")
    CuentasPath = Lector.GetValue("DIRECTORIOS", "CuentasPath")
    DeleteCuentasPath = Lector.GetValue("DIRECTORIOS", "DeleteCuentasPath")
    'Directorios
    Puerto = val(Lector.GetValue("INIT", "StartPort"))
    ListenIp = Lector.GetValue("INIT", "ListenIp")
    If ListenIp = "" Then ListenIp = "0.0.0.0"
    HideMe = val(Lector.GetValue("INIT", "Hide"))
    MaxConexionesIP = val(Lector.GetValue("INIT", "MaxConexionesIP"))
    MaxUsersPorCuenta = val(Lector.GetValue("INIT", "MaxUsersPorCuenta"))
    IdleLimit = val(Lector.GetValue("INIT", "IdleLimit"))
    'Lee la version correcta del cliente
    ULTIMAVERSION = Lector.GetValue("INIT", "Version")
    PuedeCrearPersonajes = val(Lector.GetValue("INIT", "PuedeCrearPersonajes"))
    MinimumPriceMao = val(Lector.GetValue("INIT", "MinimumPriceMao"))
    GoldPriceMao = val(Lector.GetValue("INIT", "GoldPriceMao"))
    MinimumLevelMao = val(Lector.GetValue("INIT", "MinimumLevelMao"))
    ServerSoloGMs = val(Lector.GetValue("init", "ServerSoloGMs"))
    DisconnectTimeout = val(Lector.GetValue("INIT", "DisconnectTimeout"))
    InstanceMapCount = val(Lector.GetValue("INIT", "InstanceMaps"))
    EnTesting = val(Lector.GetValue("INIT", "Testing"))
    PendingConnectionTimeout = val(Lector.GetValue("INIT", "PendingConnectionTimeout"))
    If PendingConnectionTimeout = 0 Then
        PendingConnectionTimeout = 1000
    End If
    'Ressurect pos
    ResPos.Map = val(ReadField(1, Lector.GetValue("INIT", "ResPos"), 45))
    ResPos.x = val(ReadField(2, Lector.GetValue("INIT", "ResPos"), 45))
    ResPos.y = val(ReadField(3, Lector.GetValue("INIT", "ResPos"), 45))
    'Max users
    Temporal = val(Lector.GetValue("INIT", "MaxUsers"))
    If MaxUsers = 0 Then
        MaxUsers = Temporal
        ReDim UserList(1 To MaxUsers) As t_User
        InitializeUserIndexHeap (MaxUsers)
    End If
    Call CargarCiudades
    Call LoadFeatureToggles
    Call LoadGlobalDropTable
    Set Lector = Nothing
    Exit Sub
LoadSini_Err:
    Set Lector = Nothing
    Call TraceError(Err.Number, Err.Description, "ES.LoadSini", Erl)
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
        GlobalDropTable(i).amount = val(Lector.GetValue("DROP" & i, "AMOUNT"))
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
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de feature toggles."
    Set Lector = New clsIniManager
    Call Lector.Initialize(IniPath & "feature_toggle.ini")
    If Lector.NodesCount = 0 Then
        Exit Sub
    End If
    Dim TOGGLECOUNT As Integer
    TOGGLECOUNT = val(Lector.GetValue("INIT", "TOGGLECOUNT"))
    Dim i     As Integer
    Dim key   As String
    Dim value As Boolean
    For i = 1 To TOGGLECOUNT
        key = Lector.GetValue("TOGGLE" & i, "name")
        value = val(Lector.GetValue("TOGGLE" & i, "value")) > 0
        Call SetFeatureToggle(key, value)
    Next i
    Set Lector = Nothing
    Exit Sub
LoadFeatureToggles_Err:
    Set Lector = Nothing
    Call TraceError(Err.Number, Err.Description, "ES.LoadFeatureToggles", Erl)
End Sub

Sub LoadPacketRatePolicy()
    On Error GoTo LoadPacketRatePolicy_Err
    Dim Lector As clsIniManager
    Dim i      As Long
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando PacketRatePolicy."
    Set Lector = New clsIniManager
    Call Lector.Initialize(IniPath & "PacketRatePolicy.ini")
    For i = 1 To MAX_PACKET_COUNTERS
        Dim PacketName As String
        PacketName = PacketIdToString(i)
        MacroIterations(i) = val(Lector.GetValue(PacketName, "Iterations"))
        PacketTimerThreshold(i) = val(Lector.GetValue(PacketName, "Limit"))
    Next i
    Set Lector = Nothing
    Exit Sub
LoadPacketRatePolicy_Err:
    Set Lector = Nothing
    Call TraceError(Err.Number, Err.Description, "ES.LoadPacketRatePolicy", Erl)
End Sub

Sub CargarCiudades()
    On Error GoTo CargarCiudades_Err
    Dim i      As Long
    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    Call Lector.Initialize(DatPath & "Ciudades.dat")
    Dim MapasCiudades As String
    With CityNix
        .Map = val(Lector.GetValue("NIX", "Mapa"))
        .x = val(Lector.GetValue("NIX", "X"))
        .y = val(Lector.GetValue("NIX", "Y"))
        .MapaViaje = val(Lector.GetValue("NIX", "MapaViaje"))
        .ViajeX = val(Lector.GetValue("NIX", "ViajeX"))
        .ViajeY = val(Lector.GetValue("NIX", "ViajeY"))
        .MapaResu = val(Lector.GetValue("NIX", "MapaResu"))
        .ResuX = val(Lector.GetValue("NIX", "ResuX"))
        .ResuY = val(Lector.GetValue("NIX", "ResuY"))
        .NecesitaNave = val(Lector.GetValue("NIX", "NecesitaNave"))
        MapasCiudades = Lector.GetValue("NIX", "Mapas") & ","
    End With
    With CityUllathorpe
        .Map = val(Lector.GetValue("Ullathorpe", "Mapa"))
        .x = val(Lector.GetValue("Ullathorpe", "X"))
        .y = val(Lector.GetValue("Ullathorpe", "Y"))
        .MapaViaje = val(Lector.GetValue("Ullathorpe", "MapaViaje"))
        .ViajeX = val(Lector.GetValue("Ullathorpe", "ViajeX"))
        .ViajeY = val(Lector.GetValue("Ullathorpe", "ViajeY"))
        .MapaResu = val(Lector.GetValue("Ullathorpe", "MapaResu"))
        .ResuX = val(Lector.GetValue("Ullathorpe", "ResuX"))
        .ResuY = val(Lector.GetValue("Ullathorpe", "ResuY"))
        .NecesitaNave = val(Lector.GetValue("Ullathorpe", "NecesitaNave"))
        MapasCiudades = MapasCiudades & Lector.GetValue("Ullathorpe", "Mapas") & ","
    End With
    With CityBanderbill
        .Map = val(Lector.GetValue("Banderbill", "Mapa"))
        .x = val(Lector.GetValue("Banderbill", "X"))
        .y = val(Lector.GetValue("Banderbill", "Y"))
        .MapaViaje = val(Lector.GetValue("Banderbill", "MapaViaje"))
        .ViajeX = val(Lector.GetValue("Banderbill", "ViajeX"))
        .ViajeY = val(Lector.GetValue("Banderbill", "ViajeY"))
        .MapaResu = val(Lector.GetValue("Banderbill", "MapaResu"))
        .ResuX = val(Lector.GetValue("Banderbill", "ResuX"))
        .ResuY = val(Lector.GetValue("Banderbill", "ResuY"))
        .NecesitaNave = val(Lector.GetValue("Banderbill", "NecesitaNave"))
        MapasCiudades = MapasCiudades & Lector.GetValue("Banderbill", "Mapas") & ","
    End With
    With CityLindos
        .Map = val(Lector.GetValue("Lindos", "Mapa"))
        .x = val(Lector.GetValue("Lindos", "X"))
        .y = val(Lector.GetValue("Lindos", "Y"))
        .MapaViaje = val(Lector.GetValue("Lindos", "MapaViaje"))
        .ViajeX = val(Lector.GetValue("Lindos", "ViajeX"))
        .ViajeY = val(Lector.GetValue("Lindos", "ViajeY"))
        .MapaResu = val(Lector.GetValue("Lindos", "MapaResu"))
        .ResuX = val(Lector.GetValue("Lindos", "ResuX"))
        .ResuY = val(Lector.GetValue("Lindos", "ResuY"))
        .NecesitaNave = val(Lector.GetValue("Lindos", "NecesitaNave"))
        MapasCiudades = MapasCiudades & Lector.GetValue("Lindos", "Mapas") & ","
    End With
    With CityArghal
        .Map = val(Lector.GetValue("Arghal", "Mapa"))
        .x = val(Lector.GetValue("Arghal", "X"))
        .y = val(Lector.GetValue("Arghal", "Y"))
        .MapaViaje = val(Lector.GetValue("Arghal", "MapaViaje"))
        .ViajeX = val(Lector.GetValue("Arghal", "ViajeX"))
        .ViajeY = val(Lector.GetValue("Arghal", "ViajeY"))
        .MapaResu = val(Lector.GetValue("Arghal", "MapaResu"))
        .ResuX = val(Lector.GetValue("Arghal", "ResuX"))
        .ResuY = val(Lector.GetValue("Arghal", "ResuY"))
        .NecesitaNave = val(Lector.GetValue("Arghal", "NecesitaNave"))
        MapasCiudades = MapasCiudades & Lector.GetValue("Arghal", "Mapas") & ","
    End With
    With CityForgat
        .Map = val(Lector.GetValue("Forgat", "Mapa"))
        .x = val(Lector.GetValue("Forgat", "X"))
        .y = val(Lector.GetValue("Forgat", "Y"))
        .MapaViaje = val(Lector.GetValue("Forgat", "MapaViaje"))
        .ViajeX = val(Lector.GetValue("Forgat", "ViajeX"))
        .ViajeY = val(Lector.GetValue("Forgat", "ViajeY"))
        .MapaResu = val(Lector.GetValue("Forgat", "MapaResu"))
        .ResuX = val(Lector.GetValue("Forgat", "ResuX"))
        .ResuY = val(Lector.GetValue("Forgat", "ResuY"))
        .NecesitaNave = val(Lector.GetValue("Forgat", "NecesitaNave"))
        MapasCiudades = MapasCiudades & Lector.GetValue("Forgat", "Mapas") & ","
    End With
    With CityEldoria
        .Map = val(Lector.GetValue("Eldoria", "Mapa"))
        .x = val(Lector.GetValue("Eldoria", "X"))
        .y = val(Lector.GetValue("Eldoria", "Y"))
        .MapaViaje = val(Lector.GetValue("Eldoria", "MapaViaje"))
        .ViajeX = val(Lector.GetValue("Eldoria", "ViajeX"))
        .ViajeY = val(Lector.GetValue("Eldoria", "ViajeY"))
        .MapaResu = val(Lector.GetValue("Eldoria", "MapaResu"))
        .ResuX = val(Lector.GetValue("Eldoria", "ResuX"))
        .ResuY = val(Lector.GetValue("Eldoria", "ResuY"))
        .NecesitaNave = val(Lector.GetValue("Eldoria", "NecesitaNave"))
        MapasCiudades = MapasCiudades & Lector.GetValue("Eldoria", "Mapas") & ","
    End With
    With CityArkhein
        .Map = val(Lector.GetValue("Arkhein", "Mapa"))
        .x = val(Lector.GetValue("Arkhein", "X"))
        .y = val(Lector.GetValue("Arkhein", "Y"))
        .MapaViaje = val(Lector.GetValue("Arkhein", "MapaViaje"))
        .ViajeX = val(Lector.GetValue("Arkhein", "ViajeX"))
        .ViajeY = val(Lector.GetValue("Arkhein", "ViajeY"))
        .MapaResu = val(Lector.GetValue("Arkhein", "MapaResu"))
        .ResuX = val(Lector.GetValue("Arkhein", "ResuX"))
        .ResuY = val(Lector.GetValue("Arkhein", "ResuY"))
        .NecesitaNave = val(Lector.GetValue("Arkhein", "NecesitaNave"))
        MapasCiudades = MapasCiudades & Lector.GetValue("Arkhein", "Mapas") & ","
    End With
    With CityEleusis
        .Map = val(Lector.GetValue("Eleusis", "Mapa"))
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
        .Map = val(Lector.GetValue("Penthar", "Mapa"))
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
    With Prision
        .Map = val(Lector.GetValue("Prision", "Mapa"))
        .x = val(Lector.GetValue("Prision", "X"))
        .y = val(Lector.GetValue("Prision", "Y"))
    End With
    With Libertad
        .Map = val(Lector.GetValue("Libertad", "Mapa"))
        .x = val(Lector.GetValue("Libertad", "X"))
        .y = val(Lector.GetValue("Libertad", "Y"))
    End With
    With Renacimiento
        .Map = val(Lector.GetValue("Renacimiento", "Mapa"))
        .x = val(Lector.GetValue("Renacimiento", "X"))
        .y = val(Lector.GetValue("Renacimiento", "Y"))
    End With
    With BarcoNavegandoForgatNix
        .Map = val(Lector.GetValue("BarcoNavegandoForgatNix", "Mapa"))
        .startX = val(Lector.GetValue("BarcoNavegandoForgatNix", "StartX"))
        .startY = val(Lector.GetValue("BarcoNavegandoForgatNix", "StartY"))
        .EndX = val(Lector.GetValue("BarcoNavegandoForgatNix", "EndX"))
        .EndY = val(Lector.GetValue("BarcoNavegandoForgatNix", "EndY"))
        .DestX = val(Lector.GetValue("BarcoNavegandoForgatNix", "DestX"))
        .DestY = val(Lector.GetValue("BarcoNavegandoForgatNix", "DestY"))
        .DockX = val(Lector.GetValue("BarcoNavegandoForgatNix", "DockX"))
        .DockY = val(Lector.GetValue("BarcoNavegandoForgatNix", "DockY"))
        .RequiredPassID = val(Lector.GetValue("BarcoNavegandoForgatNix", "RequiredPassID"))
    End With
    With BarcoNavegandoNixArghal
        .Map = val(Lector.GetValue("BarcoNavegandoNixArghal", "Mapa"))
        .startX = val(Lector.GetValue("BarcoNavegandoNixArghal", "StartX"))
        .startY = val(Lector.GetValue("BarcoNavegandoNixArghal", "StartY"))
        .EndX = val(Lector.GetValue("BarcoNavegandoNixArghal", "EndX"))
        .EndY = val(Lector.GetValue("BarcoNavegandoNixArghal", "EndY"))
        .DestX = val(Lector.GetValue("BarcoNavegandoNixArghal", "DestX"))
        .DestY = val(Lector.GetValue("BarcoNavegandoNixArghal", "DestY"))
        .DockX = val(Lector.GetValue("BarcoNavegandoNixArghal", "DockX"))
        .DockY = val(Lector.GetValue("BarcoNavegandoNixArghal", "DockY"))
        .RequiredPassID = val(Lector.GetValue("BarcoNavegandoNixArghal", "RequiredPassID"))
    End With
    With BarcoNavegandoArghalForgat
        .Map = val(Lector.GetValue("BarcoNavegandoArghalForgat", "Mapa"))
        .startX = val(Lector.GetValue("BarcoNavegandoArghalForgat", "StartX"))
        .startY = val(Lector.GetValue("BarcoNavegandoArghalForgat", "StartY"))
        .EndX = val(Lector.GetValue("BarcoNavegandoArghalForgat", "EndX"))
        .EndY = val(Lector.GetValue("BarcoNavegandoArghalForgat", "EndY"))
        .DestX = val(Lector.GetValue("BarcoNavegandoArghalForgat", "DestX"))
        .DestY = val(Lector.GetValue("BarcoNavegandoArghalForgat", "DestY"))
        .DockX = val(Lector.GetValue("BarcoNavegandoArghalForgat", "DockX"))
        .DockY = val(Lector.GetValue("BarcoNavegandoArghalForgat", "DockY"))
        .RequiredPassID = val(Lector.GetValue("BarcoNavegandoArghalForgat", "RequiredPassID"))
    End With
    With ForgatDock
        .Map = val(Lector.GetValue("ForgatDock", "Mapa"))
        .startX = val(Lector.GetValue("ForgatDock", "StartX"))
        .startY = val(Lector.GetValue("ForgatDock", "StartY"))
        .EndX = val(Lector.GetValue("ForgatDock", "EndX"))
        .EndY = val(Lector.GetValue("ForgatDock", "EndY"))
        .DestX = val(Lector.GetValue("ForgatDock", "DestX"))
        .DestY = val(Lector.GetValue("ForgatDock", "DestY"))
        .DockX = val(Lector.GetValue("ForgatDock", "DockX"))
        .DockY = val(Lector.GetValue("ForgatDock", "DockY"))
        .RequiredPassID = val(Lector.GetValue("BarcoNavegandoForgatNix", "RequiredPassID"))
    End With
    With NixDock
        .Map = val(Lector.GetValue("NixDock", "Mapa"))
        .startX = val(Lector.GetValue("NixDock", "StartX"))
        .startY = val(Lector.GetValue("NixDock", "StartY"))
        .EndX = val(Lector.GetValue("NixDock", "EndX"))
        .EndY = val(Lector.GetValue("NixDock", "EndY"))
        .DestX = val(Lector.GetValue("NixDock", "DestX"))
        .DestY = val(Lector.GetValue("NixDock", "DestY"))
        .DockX = val(Lector.GetValue("NixDock", "DockX"))
        .DockY = val(Lector.GetValue("NixDock", "DockY"))
        .RequiredPassID = val(Lector.GetValue("BarcoNavegandoNixArghal", "RequiredPassID"))
    End With
    With ArghalDock
        .Map = val(Lector.GetValue("ArghalDock", "Mapa"))
        .startX = val(Lector.GetValue("ArghalDock", "StartX"))
        .startY = val(Lector.GetValue("ArghalDock", "StartY"))
        .EndX = val(Lector.GetValue("ArghalDock", "EndX"))
        .EndY = val(Lector.GetValue("ArghalDock", "EndY"))
        .DestX = val(Lector.GetValue("ArghalDock", "DestX"))
        .DestY = val(Lector.GetValue("ArghalDock", "DestY"))
        .DockX = val(Lector.GetValue("ArghalDock", "DockX"))
        .DockY = val(Lector.GetValue("ArghalDock", "DockY"))
        .RequiredPassID = val(Lector.GetValue("BarcoNavegandoArghalForgat", "RequiredPassID"))
    End With
    TotalMapasCiudades = Split(MapasCiudades, ",")
    Set Lector = Nothing
    Nix.Map = CityNix.Map
    Nix.x = CityNix.x
    Nix.y = CityNix.y
    Ullathorpe.Map = CityUllathorpe.Map
    Ullathorpe.x = CityUllathorpe.x
    Ullathorpe.y = CityUllathorpe.y
    Banderbill.Map = CityBanderbill.Map
    Banderbill.x = CityBanderbill.x
    Banderbill.y = CityBanderbill.y
    Lindos.Map = CityLindos.Map
    Lindos.x = CityLindos.x
    Lindos.y = CityLindos.y
    Arghal.Map = CityArghal.Map
    Arghal.x = CityArghal.x
    Arghal.y = CityArghal.y
    Forgat.Map = CityForgat.Map
    Forgat.x = CityForgat.x
    Forgat.y = CityForgat.y
    Eldoria.Map = CityEldoria.Map
    Eldoria.x = CityEldoria.x
    Eldoria.y = CityEldoria.y
    Arkhein.Map = CityArkhein.Map
    Arkhein.x = CityArkhein.x
    Arkhein.y = CityArkhein.y
    Penthar.Map = CityPenthar.Map
    Penthar.x = CityPenthar.x
    Penthar.y = CityPenthar.y
    'Esto es para el /HOGAR
    Ciudades(e_Ciudad.cNix) = Nix
    Ciudades(e_Ciudad.cUllathorpe) = Ullathorpe
    Ciudades(e_Ciudad.cBanderbill) = Banderbill
    Ciudades(e_Ciudad.cLindos) = Lindos
    Ciudades(e_Ciudad.cArghal) = Arghal
    Ciudades(e_Ciudad.cForgat) = Forgat
    Ciudades(e_Ciudad.cArkhein) = Arkhein
    Ciudades(e_Ciudad.cEldoria) = Eldoria
    Ciudades(e_Ciudad.cPenthar) = Penthar
    Exit Sub
CargarCiudades_Err:
    Call TraceError(Err.Number, Err.Description, "ES.CargarCiudades", Erl)
End Sub

Sub LoadIntervalos()
    On Error GoTo LoadIntervalos_Err
    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    Call Lector.Initialize(IniPath & "intervalos.ini")
    'Intervalos
    SanaIntervaloSinDescansar = val(Lector.GetValue("INTERVALOS", "SanaIntervaloSinDescansar"))
    FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar
    StaminaIntervaloSinDescansar = val(Lector.GetValue("INTERVALOS", "StaminaIntervaloSinDescansar"))
    FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar
    SanaIntervaloDescansar = val(Lector.GetValue("INTERVALOS", "SanaIntervaloDescansar"))
    FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar
    StaminaIntervaloDescansar = val(Lector.GetValue("INTERVALOS", "StaminaIntervaloDescansar"))
    FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar
    IntervaloPerderStamina = val(Lector.GetValue("INTERVALOS", "IntervaloPerderStamina"))
    FrmInterv.txtIntervaloPerderStamina.Text = IntervaloPerderStamina
    IntervaloSed = val(Lector.GetValue("INTERVALOS", "IntervaloSed")) / 25
    FrmInterv.txtIntervaloSed.Text = IntervaloSed
    IntervaloHambre = val(Lector.GetValue("INTERVALOS", "IntervaloHambre")) / 25
    FrmInterv.txtIntervaloHambre.Text = IntervaloHambre
    IntervaloVeneno = val(Lector.GetValue("INTERVALOS", "IntervaloVeneno"))
    FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno
    IntervaloParalizado = val(Lector.GetValue("INTERVALOS", "IntervaloParalizado"))
    FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado
    IntervaloInmovilizado = val(Lector.GetValue("INTERVALOS", "IntervaloInmovilizado"))
    FrmInterv.txtIntervaloInmovilizado.Text = IntervaloInmovilizado
    IntervaloInvisible = val(Lector.GetValue("INTERVALOS", "IntervaloInvisible"))
    FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible
    IntervaloFrio = val(Lector.GetValue("INTERVALOS", "IntervaloFrio"))
    FrmInterv.txtIntervaloFrio.Text = IntervaloFrio
    IntervaloWavFx = val(Lector.GetValue("INTERVALOS", "IntervaloWAVFX"))
    FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx
    IntervaloInvocacion = val(Lector.GetValue("INTERVALOS", "IntervaloInvocacion"))
    FrmInterv.txtInvocacion.Text = IntervaloInvocacion
    TimeoutPrimerPaquete = val(Lector.GetValue("INTERVALOS", "TimeoutPrimerPaquete"))
    FrmInterv.txtTimeoutPrimerPaquete.Text = TimeoutPrimerPaquete / 25
    TimeoutEsperandoLoggear = val(Lector.GetValue("INTERVALOS", "TimeoutEsperandoLoggear"))
    FrmInterv.txtTimeoutEsperandoLoggear.Text = TimeoutEsperandoLoggear / 25
    IntervaloIncineracion = val(Lector.GetValue("INTERVALOS", "IntervaloFuego"))
    FrmInterv.txtintervalofuego.Text = IntervaloIncineracion
    IntervaloTirar = val(Lector.GetValue("INTERVALOS", "IntervaloTirar"))
    FrmInterv.txtintervalotirar.Text = IntervaloTirar
    IntervaloMeditar = val(Lector.GetValue("INTERVALOS", "IntervaloMeditar"))
    FrmInterv.txtIntervaloMeditar.Text = IntervaloMeditar
    IntervaloCaminar = val(Lector.GetValue("INTERVALOS", "IntervaloCaminar"))
    FrmInterv.txtintervalocaminar.Text = IntervaloCaminar
    IntervaloEnCombate = val(Lector.GetValue("INTERVALOS", "IntervaloEnCombate"))
    '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    IntervaloUserPuedeCastear = val(Lector.GetValue("INTERVALOS", "IntervaloLanzaHechizo"))
    FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear
    Call InitializeNpcAiInterval(val(Lector.GetValue("INTERVALOS", "IntervaloNpcAI", CStr(DEFAULT_NPC_AI_INTERVAL_MS))))
    FrmInterv.txtAI.Text = IntervaloNPCAI
    IntervaloTrabajarExtraer = val(Lector.GetValue("INTERVALOS", "IntervaloTrabajarExtraer"))
    FrmInterv.txtTrabajoExtraer.Text = IntervaloTrabajarExtraer
    IntervaloTrabajarConstruir = val(Lector.GetValue("INTERVALOS", "IntervaloTrabajarConstruir"))
    FrmInterv.txtTrabajoConstruir.Text = IntervaloTrabajarConstruir
    IntervaloUserPuedeAtacar = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeAtacar"))
    FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar
    'TODO : Agregar estos intervalos al form!!!
    IntervaloMagiaGolpe = val(Lector.GetValue("INTERVALOS", "IntervaloMagiaGolpe"))
    IntervaloGolpeMagia = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeMagia"))
    IntervaloNpcOwner = val(Lector.GetValue("INTERVALOS", "IntervaloNpcOwner", "10000"))
    'frmMain.tLluvia.Interval = val(Lector.GetValue("INTERVALOS", "IntervaloPerdidaStaminaLluvia"))
    'FrmInterv.txtIntervaloPerdidaStaminaLluvia.Text = frmMain.tLluvia.Interval
    MinutosWs = val(Lector.GetValue("INTERVALOS", "IntervaloWS"))
    If MinutosWs < 1 Then MinutosWs = 10
    IntervaloCerrarConexion = val(Lector.GetValue("INTERVALOS", "IntervaloCerrarConexion"))
    IntervaloUserPuedeUsarU = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeUsarU"))
    IntervaloUserPuedeUsarClic = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeUsarClic"))
    IntervaloFlechasCazadores = val(Lector.GetValue("INTERVALOS", "IntervaloFlechasCazadores"))
    IntervaloGolpeUsar = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeUsar"))
    IntervaloOculto = val(Lector.GetValue("INTERVALOS", "IntervaloOculto"))
    IntervaloPuedeSerAtacado = val(Lector.GetValue("INTERVALOS", "IntervaloPuedeSerAtacado"))
    IntervaloGuardarUsuarios = val(Lector.GetValue("INTERVALOS", "IntervaloGuardarUsuarios"))
    IntervaloTimerGuardarUsuarios = val(Lector.GetValue("INTERVALOS", "IntervaloTimerGuardarUsuarios"))
    IntervaloMensajeGlobal = val(Lector.GetValue("INTERVALOS", "IntervaloMensajeGlobal"))
    IntervalAutomatedAction = val(Lector.GetValue("INTERVALOS", "IntervalAutomatedAction"))
    IntervalChangeGlobalQuestsState = val(Lector.GetValue("INTERVALOS", "IntervalChangeGlobalQuestsState"))
    IntervalPhoenixSpawn = val(Lector.GetValue("INTERVALOS", "IntervalPhoenixSpawn"))
    '&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    Set Lector = Nothing
    Exit Sub
LoadIntervalos_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadIntervalos", Erl)
End Sub

Sub LoadMainConfigFile()
    On Error GoTo LoadMainConfigFile_Err
    Set SvrConfig = New ServerConfig
    Call SvrConfig.LoadSettings(IniPath & "Configuracion.ini")
    Call CargarEventos
    Call CargarInfoRetos
    Call CargarInfoEventos
    Call CargarMapasEspeciales
    Exit Sub
LoadMainConfigFile_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadMainConfigFile", Erl)
End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
    '*****************************************************************
    'Escribe VAR en un archivo
    '*****************************************************************
    On Error GoTo WriteVar_Err
    writeprivateprofilestring Main, Var, value, File
    Exit Sub
WriteVar_Err:
    Call TraceError(Err.Number, Err.Description, "ES.WriteVar", Erl)
End Sub

Sub SaveUser(ByVal UserIndex As Integer, Optional ByVal Logout As Boolean = False)
    On Error GoTo SaveUser_Err
    If Logout Then
        Call UserDisconnected(UserList(UserIndex).pos.Map, UserIndex)
    End If
    Call SaveCharacterDB(UserIndex)
    If Logout Then
        Call RemoveTokenDatabase(UserIndex)
    End If
    UserList(UserIndex).Counters.LastSave = GetTickCountRaw()
    Exit Sub
SaveUser_Err:
    Call TraceError(Err.Number, Err.Description, "ES.SaveUser", Erl)
End Sub

Public Sub RemoveTokenDatabase(ByVal UserIndex As Integer)
    Call Execute("delete from tokens where id =  ?;", UserList(UserIndex).encrypted_session_token_db_id)
End Sub

Public Sub AddTokenDatabase(ByVal encrypted_token As String, ByVal decrypted_token As String, ByVal username As String)
    #If UNIT_TEST = 1 Then
        'Only used in automated unit testing to create a valid session so that we can then try LoginNewChar and
        'LoginExistingChar
        Call Execute("insert into tokens (encrypted_token, decrypted_token, username, remote_host, timestamp) values(?,?,?,""127.0.0.1"",""123456"") ;", encrypted_token, _
                decrypted_token, username)
    #End If
End Sub

Sub SaveNewUser(ByVal UserIndex As Integer)
    On Error GoTo SaveNewUser_Err
    Call SaveNewCharacterDB(UserIndex)
    Exit Sub
SaveNewUser_Err:
    Call TraceError(Err.Number, Err.Description, "ES.SaveNewUser", Erl)
End Sub

Function Status(ByVal UserIndex As Integer) As e_Facciones
    On Error GoTo Status_Err
    Status = UserList(UserIndex).Faccion.Status
    Exit Function
Status_Err:
    Call TraceError(Err.Number, Err.Description, "ES.Status", Erl)
End Function

Sub BackUPnPc(NpcIndex As Integer)
    On Error GoTo BackUPnPc_Err
    Dim NpcNumero As Integer
    Dim npcfile   As String
    Dim LoopC     As Integer
    NpcNumero = NpcList(NpcIndex).Numero
    'If NpcNumero > 499 Then
    '    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
    'Else
    npcfile = DatPath & "bkNPCs.dat"
    'End If
    'General
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", NpcList(NpcIndex).name)
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", NpcList(NpcIndex).Desc)
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(NpcList(NpcIndex).Char.head))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(NpcList(NpcIndex).Char.body))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(NpcList(NpcIndex).Char.Heading))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(NpcList(NpcIndex).Movement))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(NpcList(NpcIndex).Attackable))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(NpcList(NpcIndex).Comercia))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Craftea", val(NpcList(NpcIndex).Craftea))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(NpcList(NpcIndex).TipoItems))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(NpcList(NpcIndex).Hostile))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(NpcList(NpcIndex).GiveEXP))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", val(NpcList(NpcIndex).GiveGLD))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(NpcList(NpcIndex).Hostile))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(NpcList(NpcIndex).InvReSpawn))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(NpcList(NpcIndex).npcType))
    'Stats
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(NpcList(NpcIndex).flags.AIAlineacion))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(NpcList(NpcIndex).Stats.def))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(NpcList(NpcIndex).Stats.MaxHit))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(NpcList(NpcIndex).Stats.MaxHp))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(NpcList(NpcIndex).Stats.MinHIT))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(NpcList(NpcIndex).Stats.MinHp))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(NpcList(NpcIndex).Stats.UsuariosMatados)) 'Que es ESTO?!!
    'Flags
    Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(NpcList(NpcIndex).flags.Respawn))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(NpcList(NpcIndex).flags.backup))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(NpcList(NpcIndex).flags.Domable))
    'Inventario
    Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(NpcList(NpcIndex).invent.NroItems))
    If NpcList(NpcIndex).invent.NroItems > 0 Then
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & LoopC, NpcList(NpcIndex).invent.Object(LoopC).ObjIndex & "-" & NpcList(NpcIndex).invent.Object(LoopC).amount)
        Next
    End If
    Exit Sub
BackUPnPc_Err:
    Call TraceError(Err.Number, Err.Description, "ES.BackUPnPc", Erl)
End Sub

Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NpcNumber As Integer)
    On Error GoTo CargarNpcBackUp_Err
    'Status
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup Npc"
    Dim npcfile As String
    'If NpcNumber > 499 Then
    '    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
    'Else
    npcfile = DatPath & "bkNPCs.dat"
    'End If
    NpcList(NpcIndex).Numero = NpcNumber
    NpcList(NpcIndex).name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
    NpcList(NpcIndex).Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
    Call SetMovement(NpcIndex, val(GetVar(npcfile, "NPC" & NpcNumber, "Movement")))
    NpcList(NpcIndex).npcType = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))
    NpcList(NpcIndex).Char.body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
    NpcList(NpcIndex).Char.head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
    NpcList(NpcIndex).Char.Heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))
    NpcList(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
    NpcList(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
    NpcList(NpcIndex).Craftea = val(GetVar(npcfile, "NPC" & NpcNumber, "Craftea"))
    NpcList(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
    NpcList(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))
    NpcList(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))
    NpcList(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))
    NpcList(NpcIndex).Stats.MaxHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
    NpcList(NpcIndex).Stats.MinHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
    NpcList(NpcIndex).Stats.MaxHit = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
    NpcList(NpcIndex).Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
    NpcList(NpcIndex).Stats.def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
    NpcList(NpcIndex).flags.AIAlineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))
    Dim LoopC As Integer
    Dim ln    As String
    NpcList(NpcIndex).invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
    If NpcList(NpcIndex).invent.NroItems > 0 Then
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
            NpcList(NpcIndex).invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
            NpcList(NpcIndex).invent.Object(LoopC).amount = val(ReadField(2, ln, 45))
        Next LoopC
    Else
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            NpcList(NpcIndex).invent.Object(LoopC).ObjIndex = 0
            NpcList(NpcIndex).invent.Object(LoopC).amount = 0
        Next LoopC
    End If
    NpcList(NpcIndex).flags.NPCActive = True
    NpcList(NpcIndex).flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
    NpcList(NpcIndex).flags.backup = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
    NpcList(NpcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
    NpcList(NpcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))
    'Tipo de items con los que comercia
    NpcList(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))
    Exit Sub
CargarNpcBackUp_Err:
    Call TraceError(Err.Number, Err.Description, "ES.CargarNpcBackUp", Erl)
End Sub

Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal Motivo As String)
    On Error GoTo LogBanFromName_Err
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).name)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", Motivo)
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile
    Exit Sub
LogBanFromName_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LogBanFromName", Erl)
End Sub

Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal Motivo As String)
    On Error GoTo Ban_Err
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", Motivo)
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile
    Exit Sub
Ban_Err:
    Call TraceError(Err.Number, Err.Description, "ES.Ban", Erl)
End Sub

Public Sub CargaApuestas()
    On Error GoTo CargaApuestas_Err
    Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))
    Exit Sub
CargaApuestas_Err:
    Call TraceError(Err.Number, Err.Description, "ES.CargaApuestas", Erl)
End Sub

Public Sub LoadRecursosEspeciales()
    On Error GoTo LoadRecursosEspeciales_Err
    If Not FileExist(DatPath & "RecursosEspeciales.dat", vbArchive) Then
        ReDim EspecialesTala(0) As t_Obj
        ReDim EspecialesPesca(0) As t_Obj
        Exit Sub
    End If
    Dim IniFile As clsIniManager
    Set IniFile = New clsIniManager
    Call IniFile.Initialize(DatPath & "RecursosEspeciales.dat")
    Dim count As Long, i As Long, str As String, Field() As String
    ' Tala
    count = val(IniFile.GetValue("Tala", "Items"))
    If count > 0 Then
        ReDim EspecialesTala(1 To count) As t_Obj
        For i = 1 To count
            str = IniFile.GetValue("Tala", "Item" & i)
            Field = Split(str, "-")
            EspecialesTala(i).ObjIndex = val(Field(0))
            EspecialesTala(i).data = val(Field(1))      ' Probabilidad
        Next
    Else
        ReDim EspecialesTala(0) As t_Obj
    End If
    ' Pesca
    count = val(IniFile.GetValue("Pesca", "Items"))
    If count > 0 Then
        ReDim EspecialesPesca(1 To count) As t_Obj
        For i = 1 To count
            str = IniFile.GetValue("Pesca", "Item" & i)
            Field = Split(str, "-")
            EspecialesPesca(i).ObjIndex = val(Field(0))
            EspecialesPesca(i).data = val(Field(1))     ' Probabilidad
        Next
    Else
        ReDim EspecialesPesca(0) As t_Obj
    End If
    Set IniFile = Nothing
    Exit Sub
LoadRecursosEspeciales_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadRecursosEspeciales", Erl)
End Sub

Public Sub LoadPesca()
    On Error GoTo LoadPesca_Err
    If Not FileExist(DatPath & "pesca.dat", vbArchive) Then
        ReDim Peces(0) As t_Obj
        ReDim PecesEspeciales(0) As t_Obj
        ReDim PesoPeces(0) As Long
        ReDim PoderCanas(0) As Integer
        Exit Sub
    End If
    Dim IniFile As clsIniManager
    Set IniFile = New clsIniManager
    Call IniFile.Initialize(DatPath & "pesca.dat")
    Dim count As Long, CountEspecial As Long, i As Long, j As Long, str As String, Field() As String, nivel As Integer, MaxLvlCania As Long
    count = val(IniFile.GetValue("PECES", "NumPeces"))
    MaxLvlCania = val(IniFile.GetValue("PECES", "Maxlvlcaña"))
    CountEspecial = 1
    ReDim PesoPeces(0 To MaxLvlCania) As Long
    ReDim PoderCanas(0 To MaxLvlCania) As Integer
    For i = 1 To MaxLvlCania
        PoderCanas(i) = val(IniFile.GetValue("POWERCANAS", "Power" & i))
    Next i
    If count > 0 Then
        ReDim Peces(1 To count) As t_Obj
        ' Cargo todos los peces
        For i = 1 To count
            str = IniFile.GetValue("PECES", "Pez" & i)
            Field = Split(str, "-")
            'HarThaoS: Si es un pez especial lo guardo en otro array
            If val(Field(3)) = 1 Then
                ReDim Preserve PecesEspeciales(1 To CountEspecial) As t_Obj
                PecesEspeciales(CountEspecial).ObjIndex = val(Field(0))
                PecesEspeciales(CountEspecial).data = val(Field(1))
                nivel = val(Field(2))               ' Nivel de caña
                If (nivel > MaxLvlCania) Then nivel = MaxLvlCania
                PecesEspeciales(CountEspecial).amount = nivel
                CountEspecial = CountEspecial + 1
            End If
            Peces(i).ObjIndex = val(Field(0))
            Peces(i).data = val(Field(1))       ' Peso
            nivel = val(Field(2))               ' Nivel de caña
            If (nivel > MaxLvlCania) Then nivel = MaxLvlCania
            Peces(i).amount = nivel
        Next
        ' Los ordeno segun nivel de caña (quick sort)
        Call QuickSortPeces(1, count)
        ' Sumo los pesos
        For i = 1 To count
            For j = Peces(i).amount To MaxLvlCania
                PesoPeces(j) = PesoPeces(j) + Peces(i).data
            Next j
            Peces(i).data = PesoPeces(Peces(i).amount)
        Next i
    Else
        ReDim Peces(0) As t_Obj
    End If
    
    ' Cargar UniqueMapfish
    Dim uniqueCount As Long
    Dim uniqueValue As String
    uniqueCount = 0
    i = 1
    Do
        uniqueValue = IniFile.GetValue("UNIQUEMAPFISH", "UniqueMapfish" & i)
        If Len(Trim$(uniqueValue)) = 0 Then Exit Do
    
        uniqueCount = uniqueCount + 1
        ReDim Preserve UniqueMapFishIDs(1 To uniqueCount) As Long
        UniqueMapFishIDs(uniqueCount) = CLng(val(uniqueValue))
    
        i = i + 1
    Loop
    UniqueMapFishCount = uniqueCount
    
    Set IniFile = Nothing
    Exit Sub
LoadPesca_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadPesca", Erl)
End Sub

' Adaptado de https://www.vbforums.com/showthread.php?231925-VB-Quick-Sort-algorithm-(very-fast-sorting-algorithm)
Private Sub QuickSortPeces(ByVal First As Long, ByVal Last As Long)
    On Error GoTo QuickSortPeces_Err
    Dim Low      As Long, High As Long
    Dim MidValue As Long
    Dim aux      As t_Obj
    Low = First
    High = Last
    MidValue = Peces((First + Last) \ 2).amount
    Do
        While Peces(Low).amount < MidValue
            Low = Low + 1
        Wend
        While Peces(High).amount > MidValue
            High = High - 1
        Wend
        If Low <= High Then
            aux = Peces(Low)
            Peces(Low) = Peces(High)
            Peces(High) = aux
            Low = Low + 1
            High = High - 1
        End If
    Loop While Low <= High
    If First < High Then QuickSortPeces First, High
    If Low < Last Then QuickSortPeces Low, Last
    Exit Sub
QuickSortPeces_Err:
    Call TraceError(Err.Number, Err.Description, "ES.QuickSortPeces", Erl)
End Sub

' Adaptado de https://www.freevbcode.com/ShowCode.asp?ID=9416
Public Function BinarySearchPeces(ByVal value As Long) As Long
    On Error GoTo BinarySearchPeces_Err
    Dim Low  As Long
    Dim High As Long
    Low = 1
    High = UBound(Peces)
    Dim i              As Long
    Dim valor_anterior As Long
    Do While Low <= High
        i = (Low + High) \ 2
        If i > 1 Then
            valor_anterior = Peces(i - 1).data
        Else
            valor_anterior = 0
        End If
        If value >= valor_anterior And value < Peces(i).data Then
            BinarySearchPeces = i
            Exit Do
        ElseIf value < valor_anterior Then
            High = (i - 1)
        Else
            Low = (i + 1)
        End If
    Loop
    Exit Function
BinarySearchPeces_Err:
    Call TraceError(Err.Number, Err.Description, "ES.BinarySearchPeces", Erl)
End Function

Public Sub LoadRangosFaccion()
    On Error GoTo LoadRangosFaccion_Err
    If Not FileExist(DatPath & "rangos_faccion.dat", vbArchive) Then
        ReDim RangosFaccion(0) As t_RangoFaccion
        Exit Sub
    End If
    Dim IniFile As clsIniManager
    Set IniFile = New clsIniManager
    Call IniFile.Initialize(DatPath & "rangos_faccion.dat")
    Dim i As Byte, rankData() As String
    MaxRangoFaccion = val(IniFile.GetValue("INIT", "NumRangos"))
    If MaxRangoFaccion > 0 Then
        ' Los rangos de la Armada se guardan en los indices impar, y los del caos en indices pares.
        ' Luego, para acceder es tan facil como usar el Rango directamente para la Armada, y multiplicar por 2 para el Caos.
        ReDim RangosFaccion(1 To MaxRangoFaccion * 2) As t_RangoFaccion
        For i = 1 To MaxRangoFaccion
            '<N>Rango=<NivelRequerido>-<AsesinatosRequeridos>-<Título>
            rankData = Split(IniFile.GetValue("ArmadaReal", i & "Rango"), "-", , vbTextCompare)
            RangosFaccion(2 * i - 1).rank = i
            RangosFaccion(2 * i - 1).Titulo = rankData(2)
            RangosFaccion(2 * i - 1).NivelRequerido = val(rankData(0))
            RangosFaccion(2 * i - 1).RequiredScore = val(rankData(1))
            rankData = Split(IniFile.GetValue("LegionCaos", i & "Rango"), "-", , vbTextCompare)
            RangosFaccion(2 * i).rank = i
            RangosFaccion(2 * i).Titulo = rankData(2)
            RangosFaccion(2 * i).NivelRequerido = val(rankData(0))
            RangosFaccion(2 * i).RequiredScore = val(rankData(1))
        Next i
    End If
    Set IniFile = Nothing
    Exit Sub
LoadRangosFaccion_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadRangosFaccion", Erl)
End Sub

Public Sub LoadRecompensasFaccion()
    On Error GoTo LoadRecompensasFaccion_Err
    If Not FileExist(DatPath & "recompensas_faccion.dat", vbArchive) Then
        ReDim RecompensasFaccion(0) As t_RecompensaFaccion
        Exit Sub
    End If
    Dim IniFile As clsIniManager
    Set IniFile = New clsIniManager
    Call IniFile.Initialize(DatPath & "recompensas_faccion.dat")
    Dim cantidadRecompensas As Byte, i As Integer, rank_and_objindex() As String
    cantidadRecompensas = val(IniFile.GetValue("INIT", "NumRecompensas"))
    If cantidadRecompensas > 0 Then
        ReDim RecompensasFaccion(1 To cantidadRecompensas) As t_RecompensaFaccion
        For i = 1 To cantidadRecompensas
            rank_and_objindex = Split(IniFile.GetValue("Recompensas", "Recompensa" & i), "-", , vbTextCompare)
            RecompensasFaccion(i).rank = val(rank_and_objindex(0))
            RecompensasFaccion(i).ObjIndex = val(rank_and_objindex(1))
        Next i
    End If
    Set IniFile = Nothing
    Exit Sub
LoadRecompensasFaccion_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadRecompensasFaccion", Erl)
End Sub

Public Sub LoadUserIntervals(ByVal UserIndex As Integer)
    On Error GoTo LoadUserIntervals_Err
    With UserList(UserIndex)
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
            .Intervals.Arco = IntervaloFlechasCazadores
            .Intervals.Caminar = IntervaloCaminar
            .Intervals.Golpe = IntervaloUserPuedeAtacar
            .Intervals.Magia = IntervaloUserPuedeCastear
            .Intervals.GolpeMagia = IntervaloGolpeMagia
            .Intervals.MagiaGolpe = IntervaloMagiaGolpe
            .Intervals.GolpeUsar = IntervaloGolpeUsar
            .Intervals.TrabajarExtraer = IntervaloTrabajarExtraer
            .Intervals.TrabajarConstruir = IntervaloTrabajarConstruir
            .Intervals.UsarU = IntervaloUserPuedeUsarU
            .Intervals.UsarClic = IntervaloUserPuedeUsarClic
        End If
    End With
    Exit Sub
LoadUserIntervals_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadUserIntervals", Erl)
End Sub

Function CountFiles(strFolder As String, strPattern As String) As Integer
    On Error GoTo CountFiles_Err
    Dim strFile As String
    strFile = dir$(strFolder & "\" & strPattern)
    Do Until Len(strFile) = 0
        CountFiles = CountFiles + 1
        strFile = dir$()
    Loop
    If CountFiles <> 0 Then CountFiles = CountFiles + 1
    Exit Function
CountFiles_Err:
    Call TraceError(Err.Number, Err.Description, "ES.CountFiles", Erl)
End Function

Public Function GetElapsedTime() As Single
    '***********************************************************************
    'Author: Wyrox
    'Obenemos el tiempo (en milisegundos) que pasó desde la ultima llamada.
    '***********************************************************************
    Dim end_time      As Currency
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

Public Sub CargarDonadores()
    If Not FileExist(DatPath & "donadores.dat", vbArchive) Then
        Exit Sub
    End If
    Dim IniFile As clsIniManager
    Set IniFile = New clsIniManager
    Call IniFile.Initialize(DatPath & "donadores.dat")
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

Public Sub SetFeatureToggle(ByVal name As String, ByVal State As Boolean)
    If FeatureToggles.Exists(name) Then
        FeatureToggles.Remove name
    End If
    Call FeatureToggles.Add(name, State)
End Sub

Public Function GetActiveToggles(ByRef ActiveCount As Integer) As String()
    Dim key          As Variant
    Dim ActiveKeys() As String
    ReDim ActiveKeys(FeatureToggles.count) As String
    ActiveCount = 0
    For Each key In FeatureToggles.Keys
        If FeatureToggles.Item(key) Then
            ActiveKeys(ActiveCount) = key
            ActiveCount = ActiveCount + 1
        End If
    Next key
    GetActiveToggles = ActiveKeys
End Function

Sub LoadGuildsConfig()
    On Error GoTo LoadGuildsConfig_Err
    
    Dim GuildsIni As clsIniManager
    Set GuildsIni = New clsIniManager
    GuildsIni.Initialize DatPath & "Clanes.dat"
    
    Dim i As Long

    'Experiencia de niveles de clan
    For i = 1 To MAX_LEVEL_GUILD
        ExpLevelUpGuild(i) = CLng(val(GuildsIni.GetValue("GUILDEXP", "GuildExpLevel" & CStr(i), "0")))
    Next i
    
    'Miembros máximos por nivel de clan
    For i = 1 To MAX_LEVEL_GUILD
        MembersByLevel(i) = CByte(val(GuildsIni.GetValue("MEMBERSBYLEVEL", "GuildMembersLevel" & CStr(i), "0")))
    Next i
    
    'Requisito para usar llamada de clan
    RequiredGuildLevelCallSupport = CByte(val(GuildsIni.GetValue("GUILDREWARDS", "CallSupportRequiredLevel", "4")))
    
    'Requisito para ver miembros invisibles/ocultos
    RequiredGuildLevelSeeInvisible = CByte(val(GuildsIni.GetValue("GUILDREWARDS", "SeeInvisibleRequiredLevel", "6")))
    
    'Requisito para seguro de clan
    RequiredGuildLevelSafe = CByte(val(GuildsIni.GetValue("GUILDREWARDS", "SafeGuildRequiredLevel", "5")))
    
    'Requisito para ver barra de vida
    RequiredGuildLevelShowHPBar = CByte(val(GuildsIni.GetValue("GUILDREWARDS", "ShowHPBarRequiredLevel", "6")))
    
    Set GuildsIni = Nothing
    AgregarAConsola "Se cargó la configuración de clanes (Clanes.dat)"
    Exit Sub
    
LoadGuildsConfig_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadGuildsConfig", Erl)
End Sub
Sub LoadMeditations()
    On Error GoTo LoadMeditations_Err
    
    Dim MeditationsIni As clsIniManager
    Set MeditationsIni = New clsIniManager
    MeditationsIni.Initialize DatPath & "Meditaciones.dat"
    
    MeditationLevel1to12 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSBYLEVEL", "MeditationLevel1to12", "153")))
    MeditationLevel13to17 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSBYLEVEL", "MeditationLevel13to17", "155")))
    MeditationLevel18to24 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSBYLEVEL", "MeditationLevel18to24", "157")))
    MeditationLevel25to28 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSBYLEVEL", "MeditationLevel25to28", "159")))
    MeditationLevel29to32 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSBYLEVEL", "MeditationLevel29to32", "161")))
    MeditationLevel33to36 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSBYLEVEL", "MeditationLevel33to36", "163")))
    MeditationLevel37to39 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSBYLEVEL", "MeditationLevel37to39", "165")))
    MeditationLevel40to42 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSBYLEVEL", "MeditationLevel40to42", "167")))
    MeditationLevel43to44 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSBYLEVEL", "MeditationLevel43to44", "169")))
    MeditationLevel45to46 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSBYLEVEL", "MeditationLevel45to46", "171")))
    
    'Meditaciones para criminales
    MeditationCriminalLevel1to12 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSCRIMINALBYLEVEL", "MeditationCriminalLevel1to12", "154")))
    MeditationCriminalLevel13to17 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSCRIMINALBYLEVEL", "MeditationCriminalLevel13to17", "156")))
    MeditationCriminalLevel18to24 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSCRIMINALBYLEVEL", "MeditationCriminalLevel18to24", "158")))
    MeditationCriminalLevel25to28 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSCRIMINALBYLEVEL", "MeditationCriminalLevel25to28", "160")))
    MeditationCriminalLevel29to32 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSCRIMINALBYLEVEL", "MeditationCriminalLevel29to32", "162")))
    MeditationCriminalLevel33to36 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSCRIMINALBYLEVEL", "MeditationCriminalLevel33to36", "164")))
    MeditationCriminalLevel37to39 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSCRIMINALBYLEVEL", "MeditationCriminalLevel37to39", "166")))
    MeditationCriminalLevel40to42 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSCRIMINALBYLEVEL", "MeditationCriminalLevel40to42", "168")))
    MeditationCriminalLevel43to44 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSCRIMINALBYLEVEL", "MeditationCriminalLevel43to44", "170")))
    MeditationCriminalLevel45to46 = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSCRIMINALBYLEVEL", "MeditationCriminalLevel45to46", "172")))
    
    MeditationLevelMax = CInt(val(MeditationsIni.GetValue("FXsMEDITATIONSBYLEVEL", "MeditationLevelMax", "120")))
    
    Set MeditationsIni = Nothing
    AgregarAConsola "Se cargaron las meditaciones (Meditaciones.dat)"
    Exit Sub
    
LoadMeditations_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadMeditations", Erl)
End Sub
