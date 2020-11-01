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
Private MapDat As tMapDat

Public Sub CargarSpawnList()
    Dim n As Integer, LoopC As Integer
    n = val(GetVar(DatPath & "npcs.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(n) As tCriaturasEntrenador
    For LoopC = 1 To n
        SpawnList(LoopC).NpcIndex = LoopC
        SpawnList(LoopC).NpcName = GetVar(DatPath & "npcs.dat", "NPC" & LoopC, "Name")
        If SpawnList(LoopC).NpcName = "" Then SpawnList(LoopC).NpcName = "Nada"
    Next LoopC
    End Sub
Function EsAdmin(ByRef name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsAdmin = (val(Administradores.GetValue("Admin", name)) = 1)

End Function

Function EsDios(ByRef name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsDios = (val(Administradores.GetValue("Dios", name)) = 1)

End Function

Function EsSemiDios(ByRef name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsSemiDios = (val(Administradores.GetValue("SemiDios", name)) = 1)

End Function

Function EsConsejero(ByRef name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsConsejero = (val(Administradores.GetValue("Consejero", name)) = 1)

End Function

Function EsRolesMaster(ByRef name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsRolesMaster = (val(Administradores.GetValue("RM", name)) = 1)

End Function

Public Function EsGmChar(ByRef name As String) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 27/03/2011
    'Returns true if char is administrative user.
    '***************************************************
    
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

End Function
Public Sub loadAdministrativeUsers()
    'Admines     => Admin
    'Dioses      => Dios
    'SemiDioses  => SemiDios
    'Especiales  => Especial
    'Consejeros  => Consejero
    'RoleMasters => RM
   ' If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando Administradores/Dioses/Gms."

    'Si esta mierda tuviese array asociativos el codigo seria tan lindo.
    Dim buf  As Integer

    Dim i    As Long

    Dim name As String
       
    ' Public container
    Set Administradores = New clsIniReader
    
    ' Server ini info file
    Dim ServerIni As clsIniReader

    Set ServerIni = New clsIniReader
    
    Call ServerIni.Initialize(IniPath & "Server.ini")
       
    ' Admines
    buf = val(ServerIni.GetValue("INIT", "Admines"))
    
    For i = 1 To buf
        name = UCase$(ServerIni.GetValue("Admines", "Admin" & i))
        
        If Left$(name, 1) = "*" Or Left$(name, 1) = "+" Then name = Right$(name, Len(name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Admin", name, "1")

    Next i
    
    ' Dioses
    buf = val(ServerIni.GetValue("INIT", "Dioses"))
    
    For i = 1 To buf
        name = UCase$(ServerIni.GetValue("Dioses", "Dios" & i))
        
        If Left$(name, 1) = "*" Or Left$(name, 1) = "+" Then name = Right$(name, Len(name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Dios", name, "1")
        
    Next i
        
    ' SemiDioses
    buf = val(ServerIni.GetValue("INIT", "SemiDioses"))
    
    For i = 1 To buf
        name = UCase$(ServerIni.GetValue("SemiDioses", "SemiDios" & i))
        
        If Left$(name, 1) = "*" Or Left$(name, 1) = "+" Then name = Right$(name, Len(name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("SemiDios", name, "1")
        
    Next i
    
    ' Consejeros
    buf = val(ServerIni.GetValue("INIT", "Consejeros"))
        
    For i = 1 To buf
        name = UCase$(ServerIni.GetValue("Consejeros", "Consejero" & i))
        
        If Left$(name, 1) = "*" Or Left$(name, 1) = "+" Then name = Right$(name, Len(name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Consejero", name, "1")
        
    Next i
    
    ' RolesMasters
    buf = val(ServerIni.GetValue("INIT", "RolesMasters"))
        
    For i = 1 To buf
        name = UCase$(ServerIni.GetValue("RolesMasters", "RM" & i))
        
        If Left$(name, 1) = "*" Or Left$(name, 1) = "+" Then name = Right$(name, Len(name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("RM", name, "1")
    Next i
    
    Set ServerIni = Nothing

    'If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & Time & " - Los Administradores/Dioses/Gms se han cargado correctamente."

End Sub

Public Function GetCharPrivs(ByRef UserName As String) As PlayerType
    '****************************************************
    'Author: ZaMa
    'Last Modification: 18/11/2010
    'Reads the user's charfile and retrieves its privs.
    '***************************************************

    Dim privs As PlayerType

    If EsAdmin(UserName) Then
        privs = PlayerType.Admin
        
    ElseIf EsDios(UserName) Then
        privs = PlayerType.Dios

    ElseIf EsSemiDios(UserName) Then
        privs = PlayerType.SemiDios
        
    ElseIf EsConsejero(UserName) Then
        privs = PlayerType.Consejero
    
    Else
        privs = PlayerType.user

    End If

    GetCharPrivs = privs

End Function


Public Function TxtDimension(ByVal name As String) As Long
Dim n As Integer, cad As String, Tam As Long
n = FreeFile(1)
Open name For Input As #n
Tam = 0
Do While Not EOF(n)
    Tam = Tam + 1
    Line Input #n, cad
Loop
Close n
TxtDimension = Tam
End Function

Public Sub CargarForbidenWords()

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
    Next i
    
    Close n

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
Dim Leer As New clsIniReader

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
Dim i As Integer

MaxLines = val(GetVar(DatPath & "Motd.ini", "INIT", "NumLines"))

ReDim MOTD(1 To MaxLines)
For i = 1 To MaxLines
    MOTD(i).texto = GetVar(DatPath & "Motd.ini", "Motd", "Line" & i)
    MOTD(i).Formato = vbNullString
Next i

End Sub

Public Sub DoBackUp()
'Call LogTarea("Sub DoBackUp")
haciendoBK = True
Dim i As Integer







Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())


Call LimpiarMundo
' Call WorldSave
'Call modGuilds.v_RutinaElecciones
Call ResetCentinelaInfo     'Reseteamos al centinela


Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

'Call EstadisticasWeb.Informar(EVENTO_NUEVO_CLAN, 0)

haciendoBK = False

'Log
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time
Close #nfile
End Sub

Public Sub GrabarMapa(ByVal Map As Long, ByVal MAPFILE As String)
Debug.Print "Empezamos a grabar"
On Error GoTo ErrorHandler
Dim MapRoute As String


    
MapRoute = MAPFILE & ".csm"


Dim fh As Integer
Dim MH As tMapHeader
Dim Blqs() As tDatosBloqueados
Dim L1() As tDatosGrh
Dim L2() As tDatosGrh
Dim L3() As tDatosGrh
Dim L4() As tDatosGrh
Dim Triggers() As tDatosTrigger
Dim Luces() As tDatosLuces
Dim Particulas() As tDatosParticulas
Dim Objetos() As tDatosObjs
Dim NPCs() As tDatosNPC
Dim TEs() As tDatosTE



Dim i As Long
Dim j As Integer
Dim tmpLng As Long


For j = 1 To 100
    For i = 1 To 100
        With MapData(Map, i, j)
        
        
    
            
            If .Blocked Then
                MH.NumeroBloqueados = MH.NumeroBloqueados + 1
                ReDim Preserve Blqs(1 To MH.NumeroBloqueados)
                Blqs(MH.NumeroBloqueados).X = i
                Blqs(MH.NumeroBloqueados).Y = j
            End If
            
          Rem L1(i, j) = .Graphic(1).grhindex
  
            If .Graphic(1) > 0 Then
                MH.NumeroLayers(1) = MH.NumeroLayers(1) + 1
                ReDim Preserve L1(1 To MH.NumeroLayers(1))
                L1(MH.NumeroLayers(1)).X = i
                L1(MH.NumeroLayers(1)).Y = j
                L1(MH.NumeroLayers(1)).GrhIndex = .Graphic(1)
            End If
            
            
            If .Graphic(2) > 0 Then
                MH.NumeroLayers(2) = MH.NumeroLayers(2) + 1
                ReDim Preserve L2(1 To MH.NumeroLayers(2))
                L2(MH.NumeroLayers(2)).X = i
                L2(MH.NumeroLayers(2)).Y = j
                L2(MH.NumeroLayers(2)).GrhIndex = .Graphic(2)
            End If
            
            If .Graphic(3) > 0 Then
                MH.NumeroLayers(3) = MH.NumeroLayers(3) + 1
                ReDim Preserve L3(1 To MH.NumeroLayers(3))
                L3(MH.NumeroLayers(3)).X = i
                L3(MH.NumeroLayers(3)).Y = j
                L3(MH.NumeroLayers(3)).GrhIndex = .Graphic(3)
            End If
            
            If .Graphic(4) > 0 Then
                MH.NumeroLayers(4) = MH.NumeroLayers(4) + 1
                ReDim Preserve L4(1 To MH.NumeroLayers(4))
                L4(MH.NumeroLayers(4)).X = i
                L4(MH.NumeroLayers(4)).Y = j
                L4(MH.NumeroLayers(4)).GrhIndex = .Graphic(4)
            End If
            
            If .trigger > 0 Then
                MH.NumeroTriggers = MH.NumeroTriggers + 1
                ReDim Preserve Triggers(1 To MH.NumeroTriggers)
                Triggers(MH.NumeroTriggers).X = i
                Triggers(MH.NumeroTriggers).Y = j
                Triggers(MH.NumeroTriggers).trigger = .trigger
            End If
            
             If .ParticulaIndex > 0 Then
                 MH.NumeroParticulas = MH.NumeroParticulas + 1
                 ReDim Preserve Particulas(1 To MH.NumeroParticulas)
                 Particulas(MH.NumeroParticulas).X = i
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
               
                Objetos(MH.NumeroOBJs).X = i
                Objetos(MH.NumeroOBJs).Y = j
                
            End If
            
            If .NpcIndex > 0 Then
                MH.NumeroNPCs = MH.NumeroNPCs + 1
                ReDim Preserve NPCs(1 To MH.NumeroNPCs)
                NPCs(MH.NumeroNPCs).NpcIndex = .NpcIndex
                NPCs(MH.NumeroNPCs).X = i
                NPCs(MH.NumeroNPCs).Y = j
            End If
            
            If .TileExit.Map > 0 Then
                MH.NumeroTE = MH.NumeroTE + 1
                ReDim Preserve TEs(1 To MH.NumeroTE)
                TEs(MH.NumeroTE).DestM = .TileExit.Map
                TEs(MH.NumeroTE).DestX = .TileExit.X
                TEs(MH.NumeroTE).DestY = .TileExit.Y
                TEs(MH.NumeroTE).X = i
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
        If .NumeroBloqueados > 0 Then _
            Put #fh, , Blqs
        If .NumeroLayers(1) > 0 Then _
            Put #fh, , L1
        If .NumeroLayers(2) > 0 Then _
            Put #fh, , L2
        If .NumeroLayers(3) > 0 Then _
            Put #fh, , L3
        If .NumeroLayers(4) > 0 Then _
            Put #fh, , L4
        If .NumeroTriggers > 0 Then _
            Put #fh, , Triggers
        If .NumeroParticulas > 0 Then _
            Put #fh, , Particulas
        If .NumeroLuces > 0 Then _
            Put #fh, , Luces
        If .NumeroOBJs > 0 Then _
            Put #fh, , Objetos
        If .NumeroNPCs > 0 Then _
            Put #fh, , NPCs
        If .NumeroTE > 0 Then _
            Put #fh, , TEs
    End With

Close fh

Rem MsgBox "Mapa grabado"

Debug.Print "Mapa grabado"



ErrorHandler:
    If fh <> 0 Then Close fh

End Sub
Sub LoadArmasHerreria()

Dim n As Integer, lc As Integer

n = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))

ReDim Preserve ArmasHerrero(1 To n) As Integer

For lc = 1 To n
    ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
Next lc

End Sub

Sub LoadArmadurasHerreria()

Dim n As Integer, lc As Integer

n = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))

ReDim Preserve ArmadurasHerrero(1 To n) As Integer

For lc = 1 To n
    ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
Next lc

End Sub

Sub LoadBalance()
    Dim i As Long
    Dim X As Byte
    'Modificadores de Clase
    For i = 1 To NUMCLASES
        With ModClase(i)
            .Evasion = val(GetVar(DatPath & "Balance.dat", "MODEVASION", ListaClases(i)))
            .AtaqueArmas = val(GetVar(DatPath & "Balance.dat", "MODATAQUEARMAS", ListaClases(i)))
            .AtaqueProyectiles = val(GetVar(DatPath & "Balance.dat", "MODATAQUEPROYECTILES", ListaClases(i)))
           ' .DañoWrestling = val(GetVar(DatPath & "Balance.dat", "MODATAQUEWRESTLING", ListaClases(i)))
            .DañoArmas = val(GetVar(DatPath & "Balance.dat", "MODDANOARMAS", ListaClases(i)))
            .DañoProyectiles = val(GetVar(DatPath & "Balance.dat", "MODDANOPROYECTILES", ListaClases(i)))
            .DañoWrestling = val(GetVar(DatPath & "Balance.dat", "MODDANOWRESTLING", ListaClases(i)))
            .Escudo = val(GetVar(DatPath & "Balance.dat", "MODESCUDO", ListaClases(i)))
        End With
    Next i
    
    'Modificadores de Vida
    For i = 1 To NUMRAZAS
             ModRaza(i).FuerzaGolpe = val(GetVar(DatPath & "Balance.dat", "MODFUERZAGOLPERAZA", ListaRazas(i)))
        For X = 1 To NUMCLASES
           
    
            ModVida(i).Inicial(X) = val(GetVar(DatPath & "BalanceVida.dat", "VIDAINICIAL", ListaRazas(i)))
            ModVida(i).N1TO15(X) = val(GetVar(DatPath & "BalanceVida.dat", ListaClases(X) & ListaRazas(i), "N1TO15"))
            ModVida(i).N16TO35(X) = val(GetVar(DatPath & "BalanceVida.dat", ListaClases(X) & ListaRazas(i), "N16TO35"))
            ModVida(i).N36TO45(X) = val(GetVar(DatPath & "BalanceVida.dat", ListaClases(X) & ListaRazas(i), "N36TO45"))
            ModVida(i).N46TO50(X) = val(GetVar(DatPath & "BalanceVida.dat", ListaClases(X) & ListaRazas(i), "N46TO50"))
        Next X
    Next i
    
    'Extra
    PorcentajeRecuperoMana = val(GetVar(DatPath & "Balance.dat", "EXTRA", "PorcentajeRecuperoMana"))
End Sub

Sub LoadObjCarpintero()

Dim n As Integer, lc As Integer

n = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))

ReDim Preserve ObjCarpintero(1 To n) As Integer

For lc = 1 To n
    ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
Next lc

End Sub

Sub LoadObjAlquimista()

Dim n As Integer, lc As Integer

n = val(GetVar(DatPath & "ObjAlquimista.dat", "INIT", "NumObjs"))

ReDim Preserve ObjAlquimista(1 To n) As Integer

For lc = 1 To n
    ObjAlquimista(lc) = val(GetVar(DatPath & "ObjAlquimista.dat", "Obj" & lc, "Index"))
Next lc

End Sub
Sub LoadObjSastre()

Dim n As Integer, lc As Integer

n = val(GetVar(DatPath & "ObjSastre.dat", "INIT", "NumObjs"))

ReDim Preserve ObjSastre(1 To n) As Integer

For lc = 1 To n
    ObjSastre(lc) = val(GetVar(DatPath & "ObjSastre.dat", "Obj" & lc, "Index"))
Next lc

End Sub
Sub LoadObjDonador()

Dim n As Integer, lc As Integer

n = val(GetVar(DatPath & "ObjDonador.dat", "INIT", "NumObjs"))

ReDim Preserve ObjDonador(1 To n) As tObjDonador

For lc = 1 To n
    ObjDonador(lc).ObjIndex = val(GetVar(DatPath & "ObjDonador.dat", "Obj" & lc, "Index"))
    ObjDonador(lc).Cantidad = val(GetVar(DatPath & "ObjDonador.dat", "Obj" & lc, "Cant"))
    ObjDonador(lc).Valor = val(GetVar(DatPath & "ObjDonador.dat", "Obj" & lc, "Valor"))
Next lc

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
Dim Leer As New clsIniReader

Call Leer.Initialize(DatPath & "Obj.dat")

'obtiene el numero de obj
NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))

frmCargando.cargar.min = 0
frmCargando.cargar.max = NumObjDatas
frmCargando.cargar.Value = 0


ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
  
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
        Case eOBJType.OtHerramientas
            ObjData(Object).WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
        Case eOBJType.otArmadura
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
        
        Case eOBJType.otESCUDO
            ObjData(Object).ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otCASCO
            ObjData(Object).CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
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
            ObjData(Object).StaffPower = val(Leer.GetValue("OBJ" & Object, "StaffPower"))
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
            ObjData(Object).CantItem = Leer.GetValue("OBJ" & Object, "CantItem")
            
            
            ObjData(Object).Subtipo = Leer.GetValue("OBJ" & Object, "SubTipo")
            
            
            
            If ObjData(Object).Subtipo = 1 Then
                ReDim ObjData(Object).Item(1 To ObjData(Object).CantItem)
                
                For i = 1 To ObjData(Object).CantItem
                    ObjData(Object).Item(i).ObjIndex = Leer.GetValue("OBJ" & Object, "Item" & i)
                    ObjData(Object).Item(i).Amount = Leer.GetValue("OBJ" & Object, "Cantidad" & i)
                Next i
            Else
                ReDim ObjData(Object).Item(1 To ObjData(Object).CantItem)
                
                ObjData(Object).CantEntrega = Leer.GetValue("OBJ" & Object, "CantEntrega")
                For i = 1 To ObjData(Object).CantItem
                    ObjData(Object).Item(i).ObjIndex = Leer.GetValue("OBJ" & Object, "Item" & i)
                    ObjData(Object).Item(i).Amount = Leer.GetValue("OBJ" & Object, "Cantidad" & i)
                Next i
            End If
            
            
        Case eOBJType.otYacimiento
            ' Drop gemas yacimientos
            ObjData(Object).CantItem = val(Leer.GetValue("OBJ" & Object, "Gemas"))
            
            If ObjData(Object).CantItem > 0 Then
                ReDim ObjData(Object).Item(1 To ObjData(Object).CantItem)

                Dim str As String, Field() As String
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
    ObjData(Object).RopajeBajo = val(Leer.GetValue("OBJ" & Object, "NumRopajeBajo"))
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
    
    ObjData(Object).ResistenciaMagica = val(Leer.GetValue("OBJ" & Object, "ResistenciaMagica"))
    
    
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
    For i = 1 To 9
        S = UCase$(Leer.GetValue("OBJ" & Object, "CP" & i))
        n = 1
        Do While LenB(S) > 0 And UCase$(ListaClases(n)) <> S
            n = n + 1
        Loop
        ObjData(Object).ClaseProhibida(i) = IIf(LenB(S) > 0, n, 0)
    Next i
    
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

Dim LoopC As Long

For LoopC = 1 To NUMATRIBUTOS
  UserList(UserIndex).Stats.UserAtributos(LoopC) = CInt(UserFile.GetValue("ATRIBUTOS", "AT" & LoopC))
  UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) = UserList(UserIndex).Stats.UserAtributos(LoopC)
Next LoopC

For LoopC = 1 To NUMSKILLS
  UserList(UserIndex).Stats.UserSkills(LoopC) = CInt(UserFile.GetValue("SKILLS", "SK" & LoopC))
Next LoopC

For LoopC = 1 To MAXUSERHECHIZOS
  UserList(UserIndex).Stats.UserHechizos(LoopC) = CInt(UserFile.GetValue("Hechizos", "H" & LoopC))
Next LoopC

UserList(UserIndex).Stats.GLD = CLng(UserFile.GetValue("STATS", "GLD"))
UserList(UserIndex).Stats.Banco = CLng(UserFile.GetValue("STATS", "BANCO"))

UserList(UserIndex).Stats.MaxHp = CInt(UserFile.GetValue("STATS", "MaxHP"))
UserList(UserIndex).Stats.MinHp = CInt(UserFile.GetValue("STATS", "MinHP"))

UserList(UserIndex).Stats.MinSta = CInt(UserFile.GetValue("STATS", "MinSTA"))
UserList(UserIndex).Stats.MaxSta = CInt(UserFile.GetValue("STATS", "MaxSTA"))

UserList(UserIndex).Stats.MaxMAN = CInt(UserFile.GetValue("STATS", "MaxMAN"))
UserList(UserIndex).Stats.MinMAN = CInt(UserFile.GetValue("STATS", "MinMAN"))

UserList(UserIndex).Stats.MaxHit = CInt(UserFile.GetValue("STATS", "MaxHIT"))
UserList(UserIndex).Stats.MinHIT = CInt(UserFile.GetValue("STATS", "MinHIT"))


UserList(UserIndex).Stats.MaxAGU = CByte(UserFile.GetValue("STATS", "MaxAGU"))
UserList(UserIndex).Stats.MinAGU = CByte(UserFile.GetValue("STATS", "MinAGU"))

UserList(UserIndex).Stats.MaxHam = CByte(UserFile.GetValue("STATS", "MaxHAM"))
UserList(UserIndex).Stats.MinHam = CByte(UserFile.GetValue("STATS", "MinHAM"))

UserList(UserIndex).Stats.SkillPts = CInt(UserFile.GetValue("STATS", "SkillPtsLibres"))

UserList(UserIndex).Stats.Exp = CDbl(UserFile.GetValue("STATS", "EXP"))
UserList(UserIndex).Stats.ELU = CLng(UserFile.GetValue("STATS", "ELU"))
UserList(UserIndex).Stats.ELV = CByte(UserFile.GetValue("STATS", "ELV"))




UserList(UserIndex).flags.Envenena = CByte(UserFile.GetValue("MAGIA", "ENVENENA"))
UserList(UserIndex).flags.Paraliza = CByte(UserFile.GetValue("MAGIA", "PARALIZA"))
UserList(UserIndex).flags.incinera = CByte(UserFile.GetValue("MAGIA", "INCINERA")) 'Estupidiza
UserList(UserIndex).flags.Estupidiza = CByte(UserFile.GetValue("MAGIA", "Estupidiza"))

UserList(UserIndex).flags.PendienteDelSacrificio = CByte(UserFile.GetValue("MAGIA", "PENDIENTE"))
UserList(UserIndex).flags.CarroMineria = CByte(UserFile.GetValue("MAGIA", "CarroMineria"))
UserList(UserIndex).flags.NoPalabrasMagicas = CByte(UserFile.GetValue("MAGIA", "NOPALABRASMAGICAS"))
If UserList(UserIndex).flags.Muerto = 0 Then
UserList(UserIndex).Char.Otra_Aura = CStr(UserFile.GetValue("MAGIA", "OTRA_AURA"))
End If

UserList(UserIndex).flags.DañoMagico = CByte(UserFile.GetValue("MAGIA", "DañoMagico"))
UserList(UserIndex).flags.ResistenciaMagica = CByte(UserFile.GetValue("MAGIA", "ResistenciaMagica"))

'Nuevos
UserList(UserIndex).flags.RegeneracionMana = CByte(UserFile.GetValue("MAGIA", "RegeneracionMana"))
UserList(UserIndex).flags.AnilloOcultismo = CByte(UserFile.GetValue("MAGIA", "AnilloOcultismo"))
UserList(UserIndex).flags.NoDetectable = CByte(UserFile.GetValue("MAGIA", "NoDetectable"))
UserList(UserIndex).flags.NoMagiaEfeceto = CByte(UserFile.GetValue("MAGIA", "NoMagiaEfeceto"))
UserList(UserIndex).flags.RegeneracionHP = CByte(UserFile.GetValue("MAGIA", "RegeneracionHP"))
UserList(UserIndex).flags.RegeneracionSta = CByte(UserFile.GetValue("MAGIA", "RegeneracionSta"))



UserList(UserIndex).Stats.UsuariosMatados = CLng(UserFile.GetValue("MUERTES", "UserMuertes"))
UserList(UserIndex).Stats.NPCsMuertos = CInt(UserFile.GetValue("MUERTES", "NpcsMuertes"))


UserList(UserIndex).Stats.InventLevel = CInt(UserFile.GetValue("STATS", "InventLevel"))



If CByte(UserFile.GetValue("CONSEJO", "PERTENECE")) Then _
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.RoyalCouncil

If CByte(UserFile.GetValue("CONSEJO", "PERTENECECAOS")) Then _
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.ChaosCouncil

End Sub
Sub LoadUserInit(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)
'*************************************************
'Author: Unknown
'Last modified: 19/11/2006
'Loads the Users records
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'23/01/2007 Pablo (ToxicWaste) - Quito CriminalesMatados de Stats porque era redundante.
'*************************************************
Dim LoopC As Long
Dim ln As String

UserList(UserIndex).Faccion.Status = CByte(UserFile.GetValue("FACCIONES", "Status"))
UserList(UserIndex).Faccion.ArmadaReal = CByte(UserFile.GetValue("FACCIONES", "EjercitoReal"))
UserList(UserIndex).Faccion.FuerzasCaos = CByte(UserFile.GetValue("FACCIONES", "EjercitoCaos"))
UserList(UserIndex).Faccion.CiudadanosMatados = CLng(UserFile.GetValue("FACCIONES", "CiudMatados"))
UserList(UserIndex).Faccion.CriminalesMatados = CLng(UserFile.GetValue("FACCIONES", "CrimMatados"))
UserList(UserIndex).Faccion.RecibioArmaduraCaos = CByte(UserFile.GetValue("FACCIONES", "rArCaos"))
UserList(UserIndex).Faccion.RecibioArmaduraReal = CByte(UserFile.GetValue("FACCIONES", "rArReal"))
UserList(UserIndex).Faccion.RecibioExpInicialCaos = CByte(UserFile.GetValue("FACCIONES", "rExCaos"))
UserList(UserIndex).Faccion.RecibioExpInicialReal = CByte(UserFile.GetValue("FACCIONES", "rExReal"))
UserList(UserIndex).Faccion.RecompensasCaos = CLng(UserFile.GetValue("FACCIONES", "recCaos"))
UserList(UserIndex).Faccion.RecompensasReal = CLng(UserFile.GetValue("FACCIONES", "recReal"))
UserList(UserIndex).Faccion.Reenlistadas = CByte(UserFile.GetValue("FACCIONES", "Reenlistadas"))
UserList(UserIndex).Faccion.NivelIngreso = CInt(UserFile.GetValue("FACCIONES", "NivelIngreso"))
UserList(UserIndex).Faccion.FechaIngreso = UserFile.GetValue("FACCIONES", "FechaIngreso")
UserList(UserIndex).Faccion.MatadosIngreso = CInt(UserFile.GetValue("FACCIONES", "MatadosIngreso"))
UserList(UserIndex).Faccion.NextRecompensa = CInt(UserFile.GetValue("FACCIONES", "NextRecompensa"))

UserList(UserIndex).flags.Muerto = CByte(UserFile.GetValue("FLAGS", "Muerto"))
UserList(UserIndex).flags.Escondido = CByte(UserFile.GetValue("FLAGS", "Escondido"))

UserList(UserIndex).flags.Hambre = CByte(UserFile.GetValue("FLAGS", "Hambre"))
UserList(UserIndex).flags.Sed = CByte(UserFile.GetValue("FLAGS", "Sed"))
UserList(UserIndex).flags.Desnudo = CByte(UserFile.GetValue("FLAGS", "Desnudo"))
UserList(UserIndex).flags.Navegando = CByte(UserFile.GetValue("FLAGS", "Navegando"))
UserList(UserIndex).flags.Envenenado = CByte(UserFile.GetValue("FLAGS", "Envenenado"))
UserList(UserIndex).flags.Paralizado = CByte(UserFile.GetValue("FLAGS", "Paralizado"))
UserList(UserIndex).flags.Incinerado = CByte(UserFile.GetValue("FLAGS", "Incinerado"))
UserList(UserIndex).flags.Inmovilizado = CByte(UserFile.GetValue("FLAGS", "Inmovilizado"))

UserList(UserIndex).flags.ScrollExp = CSng(UserFile.GetValue("FLAGS", "ScrollExp"))
UserList(UserIndex).flags.ScrollOro = CSng(UserFile.GetValue("FLAGS", "ScrollOro"))

If UserList(UserIndex).flags.Paralizado = 1 Then
    UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
End If

UserList(UserIndex).flags.BattlePuntos = CLng(UserFile.GetValue("Battle", "Puntos"))


If UserList(UserIndex).flags.Inmovilizado = 1 Then
    UserList(UserIndex).Counters.Inmovilizado = 20
End If

UserList(UserIndex).Counters.Pena = CLng(UserFile.GetValue("COUNTERS", "Pena"))

UserList(UserIndex).Counters.ScrollExperiencia = CLng(UserFile.GetValue("COUNTERS", "ScrollExperiencia"))
UserList(UserIndex).Counters.ScrollOro = CLng(UserFile.GetValue("COUNTERS", "ScrollOro"))


UserList(UserIndex).Counters.Oxigeno = CLng(UserFile.GetValue("COUNTERS", "Oxigeno"))

UserList(UserIndex).MENSAJEINFORMACION = UserFile.GetValue("INIT", "MENSAJEINFORMACION")


UserList(UserIndex).genero = UserFile.GetValue("INIT", "Genero")
UserList(UserIndex).clase = UserFile.GetValue("INIT", "Clase")
UserList(UserIndex).raza = UserFile.GetValue("INIT", "Raza")
UserList(UserIndex).Hogar = UserFile.GetValue("INIT", "Hogar")
UserList(UserIndex).Char.heading = CInt(UserFile.GetValue("INIT", "Heading"))

UserList(UserIndex).OrigChar.Head = CInt(UserFile.GetValue("INIT", "Head"))
UserList(UserIndex).OrigChar.Body = CInt(UserFile.GetValue("INIT", "Body"))
UserList(UserIndex).OrigChar.WeaponAnim = CInt(UserFile.GetValue("INIT", "Arma"))
UserList(UserIndex).OrigChar.ShieldAnim = CInt(UserFile.GetValue("INIT", "Escudo"))
UserList(UserIndex).OrigChar.CascoAnim = CInt(UserFile.GetValue("INIT", "Casco"))

#If ConUpTime Then
    UserList(UserIndex).UpTime = CLng(UserFile.GetValue("INIT", "UpTime"))
#End If

 UserList(UserIndex).OrigChar.heading = UserList(UserIndex).Char.heading

If UserList(UserIndex).flags.Muerto = 0 Then
    UserList(UserIndex).Char = UserList(UserIndex).OrigChar
Else
    UserList(UserIndex).Char.Body = iCuerpoMuerto
    UserList(UserIndex).Char.Head = iCabezaMuerto
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.CascoAnim = NingunCasco
End If


UserList(UserIndex).Desc = UserFile.GetValue("INIT", "Desc")





UserList(UserIndex).flags.BanMotivo = UserFile.GetValue("BAN", "BanMotivo")
UserList(UserIndex).flags.Montado = CByte(UserFile.GetValue("FLAGS", "Montado"))
UserList(UserIndex).flags.VecesQueMoriste = CLng(UserFile.GetValue("FLAGS", "VecesQueMoriste"))

UserList(UserIndex).flags.MinutosRestantes = CLng(UserFile.GetValue("FLAGS", "MinutosRestantes"))
UserList(UserIndex).flags.Silenciado = CLng(UserFile.GetValue("FLAGS", "Silenciado"))
UserList(UserIndex).flags.SegundosPasados = CLng(UserFile.GetValue("FLAGS", "SegundosPasados"))

'CASAMIENTO LADDER
UserList(UserIndex).flags.Casado = CInt(UserFile.GetValue("FLAGS", "CASADO"))
UserList(UserIndex).flags.Pareja = UserFile.GetValue("FLAGS", "PAREJA")





UserList(UserIndex).Pos.Map = CInt(ReadField(1, UserFile.GetValue("INIT", "Position"), 45))
UserList(UserIndex).Pos.X = CInt(ReadField(2, UserFile.GetValue("INIT", "Position"), 45))
UserList(UserIndex).Pos.Y = CInt(ReadField(3, UserFile.GetValue("INIT", "Position"), 45))

UserList(UserIndex).Invent.NroItems = CInt(UserFile.GetValue("Inventory", "CantidadItems"))

'[KEVIN]--------------------------------------------------------------------
'***********************************************************************************
UserList(UserIndex).BancoInvent.NroItems = CInt(UserFile.GetValue("BancoInventory", "CantidadItems"))
'Lista de objetos del banco
For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
    ln = UserFile.GetValue("BancoInventory", "Obj" & LoopC)
    UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
    UserList(UserIndex).BancoInvent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
Next LoopC
'------------------------------------------------------------------------------------
'[/KEVIN]*****************************************************************************


'Lista de objetos
For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots
    ln = UserFile.GetValue("Inventory", "Obj" & LoopC)
    UserList(UserIndex).Invent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
    UserList(UserIndex).Invent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
    UserList(UserIndex).Invent.Object(LoopC).Equipped = CByte(ReadField(3, ln, 45))
Next LoopC

UserList(UserIndex).Invent.WeaponEqpSlot = CByte(UserFile.GetValue("Inventory", "WeaponEqpSlot"))
UserList(UserIndex).Invent.HerramientaEqpSlot = CByte(UserFile.GetValue("Inventory", "HerramientaEqpSlot"))
UserList(UserIndex).Invent.ArmourEqpSlot = CByte(UserFile.GetValue("Inventory", "ArmourEqpSlot"))
UserList(UserIndex).Invent.EscudoEqpSlot = CByte(UserFile.GetValue("Inventory", "EscudoEqpSlot"))
UserList(UserIndex).Invent.CascoEqpSlot = CByte(UserFile.GetValue("Inventory", "CascoEqpSlot"))
UserList(UserIndex).Invent.BarcoSlot = CByte(UserFile.GetValue("Inventory", "BarcoSlot"))
UserList(UserIndex).Invent.MonturaSlot = CByte(UserFile.GetValue("Inventory", "MonturaSlot"))
UserList(UserIndex).Invent.MunicionEqpSlot = CByte(UserFile.GetValue("Inventory", "MunicionSlot"))
UserList(UserIndex).Invent.AnilloEqpSlot = CByte(UserFile.GetValue("Inventory", "AnilloSlot"))
UserList(UserIndex).Invent.MagicoSlot = CByte(UserFile.GetValue("Inventory", "MagicoSlot"))
UserList(UserIndex).Invent.NudilloSlot = CByte(UserFile.GetValue("Inventory", "NudilloEqpSlot"))

UserList(UserIndex).ChatCombate = CByte(UserFile.GetValue("BINDKEYS", "ChatCombate"))
UserList(UserIndex).ChatGlobal = CByte(UserFile.GetValue("BINDKEYS", "ChatGlobal"))

UserList(UserIndex).Correo.CantCorreo = CByte(UserFile.GetValue("CORREO", "CantCorreo"))
UserList(UserIndex).Correo.NoLeidos = CByte(UserFile.GetValue("CORREO", "NoLeidos"))


For LoopC = 1 To UserList(UserIndex).Correo.CantCorreo
    UserList(UserIndex).Correo.Mensaje(LoopC).Remitente = UserFile.GetValue("CORREO", "REMITENTE" & LoopC)
    UserList(UserIndex).Correo.Mensaje(LoopC).Mensaje = UserFile.GetValue("CORREO", "MENSAJE" & LoopC)
    UserList(UserIndex).Correo.Mensaje(LoopC).Item = UserFile.GetValue("CORREO", "Item" & LoopC)
    UserList(UserIndex).Correo.Mensaje(LoopC).ItemCount = CByte(UserFile.GetValue("CORREO", "ItemCount" & LoopC))
    UserList(UserIndex).Correo.Mensaje(LoopC).Fecha = UserFile.GetValue("CORREO", "DATE" & LoopC)
    UserList(UserIndex).Correo.Mensaje(LoopC).Leido = CByte(UserFile.GetValue("CORREO", "LEIDO" & LoopC))
Next LoopC

'Logros Ladder
UserList(UserIndex).UserLogros = UserFile.GetValue("LOGROS", "UserLogros")
UserList(UserIndex).NPcLogros = UserFile.GetValue("LOGROS", "NPcLogros")
UserList(UserIndex).LevelLogros = UserFile.GetValue("LOGROS", "LevelLogros")
'Logros Ladder

ln = UserFile.GetValue("Guild", "GUILDINDEX")
If IsNumeric(ln) Then
    UserList(UserIndex).GuildIndex = CInt(ln)
Else
    UserList(UserIndex).GuildIndex = 0
End If

End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
szReturn = vbNullString
  
sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
  
  
GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, File
  
GetVar = RTrim$(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub CargarBackUp()

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."

Dim Map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String



    
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.Value = 0
    frmCargando.ToMapLbl.Visible = True
   ' MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
        ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

        ReDim MapInfo(1 To NumMaps) As MapInfo
      
    For Map = 1 To NumMaps
        frmCargando.ToMapLbl = Map & "/" & NumMaps
       Rem If val(GetVar(App.Path & MapPath & "Mapa" & map & ".Dat", "Mapa" & map, "BackUp")) <> 0 Then
            tFileName = App.Path & "\WorldBackUp\Mapa" & Map & ".csm"
       Rem Else
       Rem     tFileName = App.Path & MapPath & "Mapa" & map
       Rem End If

       Call CargarMapaFormatoCSM(Map, tFileName)
        frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        DoEvents
    Next Map
frmCargando.ToMapLbl.Visible = False
Exit Sub



End Sub

Sub LoadMapData()

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas..."

Dim Map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

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
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

End Sub
Public Sub CargarMapaFormatoCSM(ByVal Map As Long, ByVal MAPFl As String)

  On Error GoTo errh:
Dim npcfile As String
Dim fh As Integer
Dim MH As tMapHeader
Dim Blqs() As tDatosBloqueados
Dim L1() As tDatosGrh
Dim L2() As tDatosGrh
Dim L3() As tDatosGrh
Dim L4() As tDatosGrh
Dim Triggers() As tDatosTrigger
Dim Luces() As tDatosLuces
Dim Particulas() As tDatosParticulas
Dim Objetos() As tDatosObjs
Dim NPCs() As tDatosNPC
Dim TEs() As tDatosTE
Dim Body As Integer
Dim Head As Integer
Dim heading As Byte
Dim i As Long
Dim j As Long
         


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

                MapData(Map, Blqs(i).X, Blqs(i).Y).Blocked = 1
            Next i

        End If

       
        
        'Cargamos Layer 1
        
        If .NumeroLayers(1) > 0 Then
        
            ReDim L1(1 To .NumeroLayers(1))
            Get #fh, , L1
            For i = 1 To .NumeroLayers(1)
                        
           MapData(Map, L1(i).X, L1(i).Y).Graphic(1) = L1(i).GrhIndex
            
            'InitGrh MapData(L1(i).X, L1(i).Y).Graphic(1), MapData(L1(i).X, L1(i).Y).Graphic(1).GrhIndex
               ' Call Map_Grh_Set(L2(i).X, L2(i).Y, L2(i).GrhIndex, 2)
            Next i
        End If
        
        'Cargamos Layer 2
        If .NumeroLayers(2) > 0 Then
            ReDim L2(1 To .NumeroLayers(2))
            Get #fh, , L2
            For i = 1 To .NumeroLayers(2)
            MapData(Map, L2(i).X, L2(i).Y).Graphic(2) = L2(i).GrhIndex
            Next i
        End If
                
        If .NumeroLayers(3) > 0 Then
            ReDim L3(1 To .NumeroLayers(3))
            Get #fh, , L3
            For i = 1 To .NumeroLayers(3)
                MapData(Map, L3(i).X, L3(i).Y).Graphic(3) = L3(i).GrhIndex
            Next i
        End If
        
        If .NumeroLayers(4) > 0 Then
            ReDim L4(1 To .NumeroLayers(4))
            Get #fh, , L4
            For i = 1 To .NumeroLayers(4)
                MapData(Map, L4(i).X, L4(i).Y).Graphic(4) = L4(i).GrhIndex
            Next i
        End If


        If .NumeroTriggers > 0 Then
            ReDim Triggers(1 To .NumeroTriggers)
            Get #fh, , Triggers
            For i = 1 To .NumeroTriggers
                MapData(Map, Triggers(i).X, Triggers(i).Y).trigger = Triggers(i).trigger
            Next i
        End If



        If .NumeroParticulas > 0 Then
            ReDim Particulas(1 To .NumeroParticulas)
            Get #fh, , Particulas
            For i = 1 To .NumeroParticulas
                MapData(Map, Particulas(i).X, Particulas(i).Y).ParticulaIndex = Particulas(i).Particula
                MapData(Map, Particulas(i).X, Particulas(i).Y).ParticulaIndex = 0
            Next i
        End If


        If .NumeroLuces > 0 Then
            ReDim Luces(1 To .NumeroLuces)
            Get #fh, , Luces
            For i = 1 To .NumeroLuces
            MapData(Map, Luces(i).X, Luces(i).Y).Luz.Color = Luces(i).Color
            MapData(Map, Luces(i).X, Luces(i).Y).Luz.Rango = Luces(i).Rango
            MapData(Map, Luces(i).X, Luces(i).Y).Luz.Color = 0
            MapData(Map, Luces(i).X, Luces(i).Y).Luz.Rango = 0
            Next i
        End If

            
        If .NumeroOBJs > 0 Then
            ReDim Objetos(1 To .NumeroOBJs)
            Get #fh, , Objetos
            For i = 1 To .NumeroOBJs
                MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.ObjIndex = Objetos(i).ObjIndex

                Select Case ObjData(Objetos(i).ObjIndex).OBJType
                    Case eOBJType.otYacimiento, eOBJType.otArboles
                        MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.Amount = ObjData(Objetos(i).ObjIndex).VidaUtil
                        MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.UltimoUso = &H7FFFFFFF ' Max Long
                    Case Else
                        MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.Amount = Objetos(i).ObjAmmount
                End Select
            Next i
        End If

        If .NumeroNPCs > 0 Then
                ReDim NPCs(1 To .NumeroNPCs)
                Get #fh, , NPCs
                 
                For i = 1 To .NumeroNPCs




                    MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex = NPCs(i).NpcIndex
                    

                    
                    
                    If MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex > 0 Then
                           npcfile = DatPath & "NPCs.dat"
                        'Si el npc debe hacer respawn en la pos
                        'original la guardamos
                            If val(GetVar(npcfile, "NPC" & MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex, "PosOrig")) = 1 Then
                                    MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex = OpenNPC(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex)
                                    Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Orig.Map = Map
                                    Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Orig.X = NPCs(i).X
                                    Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Orig.Y = NPCs(i).Y
                            Else
                                    MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex = OpenNPC(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex)
                            End If
                        Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Pos.Map = Map
                        Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Pos.X = NPCs(i).X
                        Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).Pos.Y = NPCs(i).Y
                        
                        
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
                   
                            
                        If Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex).name = "" Then
                       
                        MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex = 0
                        Else
                        
                       Call MakeNPCChar(True, 0, MapData(Map, NPCs(i).X, NPCs(i).Y).NpcIndex, Map, NPCs(i).X, NPCs(i).Y)
                        
                        End If
                    End If



                Next i
                
            End If
            
            
            
            
        If .NumeroTE > 0 Then
                    ReDim TEs(1 To .NumeroTE)
                    Get #fh, , TEs
                    For i = 1 To .NumeroTE
                        MapData(Map, TEs(i).X, TEs(i).Y).TileExit.Map = TEs(i).DestM
                        MapData(Map, TEs(i).X, TEs(i).Y).TileExit.X = TEs(i).DestX
                        MapData(Map, TEs(i).X, TEs(i).Y).TileExit.Y = TEs(i).DestY
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

    Dim Lector As clsIniReader
    Dim Temporal As Long
    
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."
    
    Set Lector = New clsIniReader
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
    LastSockListen = val(Lector.GetValue("INIT", "LastSockListen"))
    HideMe = val(Lector.GetValue("INIT", "Hide"))
    AllowMultiLogins = val(Lector.GetValue("INIT", "AllowMultiLogins"))
    IdleLimit = val(Lector.GetValue("INIT", "IdleLimit"))
    'Lee la version correcta del cliente
    ULTIMAVERSION = Lector.GetValue("INIT", "Version")
    
    PuedeCrearPersonajes = val(Lector.GetValue("INIT", "PuedeCrearPersonajes"))
    ServerSoloGMs = val(Lector.GetValue("init", "ServerSoloGMs"))
    
    DiceMinimum = val(Lector.GetValue("INIT", "MinDados"))
    DiceMaximum = val(Lector.GetValue("INIT", "MaxDados"))
    
    
    EnTesting = val(Lector.GetValue("INIT", "Testing"))
    
    'Start pos
    
    
    'Intervalos
    SanaIntervaloSinDescansar = val(Lector.GetValue("INTERVALOS", "SanaIntervaloSinDescansar"))
    FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar
    
    StaminaIntervaloSinDescansar = val(Lector.GetValue("INTERVALOS", "StaminaIntervaloSinDescansar"))
    FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar
    
    SanaIntervaloDescansar = val(Lector.GetValue("INTERVALOS", "SanaIntervaloDescansar"))
    FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar
    
    StaminaIntervaloDescansar = val(Lector.GetValue("INTERVALOS", "StaminaIntervaloDescansar"))
    FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar
    
    IntervaloSed = val(Lector.GetValue("INTERVALOS", "IntervaloSed"))
    FrmInterv.txtIntervaloSed.Text = IntervaloSed
    
    IntervaloHambre = val(Lector.GetValue("INTERVALOS", "IntervaloHambre"))
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
    FrmInterv.txtTimeoutPrimerPaquete.Text = TimeoutPrimerPaquete
    
    TimeoutEsperandoLoggear = val(Lector.GetValue("INTERVALOS", "TimeoutEsperandoLoggear"))
    FrmInterv.txtTimeoutEsperandoLoggear.Text = TimeoutEsperandoLoggear
    
    IntervaloIncineracion = val(Lector.GetValue("INTERVALOS", "IntervaloFuego"))
    FrmInterv.txtintervalofuego.Text = IntervaloIncineracion
    'Ladder
    
    '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    
    
    IntervaloUserPuedeCastear = val(Lector.GetValue("INTERVALOS", "IntervaloLanzaHechizo"))
    FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear
    
    frmMain.TIMER_AI.Interval = val(Lector.GetValue("INTERVALOS", "IntervaloNpcAI"))
    FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval
    
    frmMain.npcataca.Interval = val(Lector.GetValue("INTERVALOS", "IntervaloNpcPuedeAtacar"))
    FrmInterv.txtNPCPuedeAtacar.Text = frmMain.npcataca.Interval
    
    IntervaloUserPuedeTrabajar = val(Lector.GetValue("INTERVALOS", "IntervaloTrabajo"))
    FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar
    
    IntervaloUserPuedeAtacar = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeAtacar"))
    FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar
    
    'TODO : Agregar estos intervalos al form!!!
    IntervaloMagiaGolpe = val(Lector.GetValue("INTERVALOS", "IntervaloMagiaGolpe"))
    IntervaloGolpeMagia = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeMagia"))
    
    'frmMain.tLluvia.Interval = val(Lector.GetValue("INTERVALOS", "IntervaloPerdidaStaminaLluvia"))
    'FrmInterv.txtIntervaloPerdidaStaminaLluvia.Text = frmMain.tLluvia.Interval
    
    MinutosWs = val(Lector.GetValue("INTERVALOS", "IntervaloWS"))
    If MinutosWs < 1 Then MinutosWs = 10
    
    IntervaloCerrarConexion = val(Lector.GetValue("INTERVALOS", "IntervaloCerrarConexion"))
    IntervaloUserPuedeUsar = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeUsar"))
    IntervaloFlechasCazadores = val(Lector.GetValue("INTERVALOS", "IntervaloFlechasCazadores"))
    
    IntervaloOculto = val(Lector.GetValue("INTERVALOS", "IntervaloOculto"))
    
    IntervaloPuedeSerAtacado = val(Lector.GetValue("INTERVALOS", "IntervaloPuedeSerAtacado"))
    
    '&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    
    ' Database
    Database_Enabled = CBool(val(Lector.GetValue("DATABASE", "Enabled")))
    Database_DataSource = Lector.GetValue("DATABASE", "DSN")
    Database_Host = Lector.GetValue("DATABASE", "Host")
    Database_Name = Lector.GetValue("DATABASE", "Name")
    Database_Username = Lector.GetValue("DATABASE", "Username")
    Database_Password = Lector.GetValue("DATABASE", "Password")
    
    'Ressurect pos
    ResPos.Map = val(ReadField(1, Lector.GetValue("INIT", "ResPos"), 45))
    ResPos.X = val(ReadField(2, Lector.GetValue("INIT", "ResPos"), 45))
    ResPos.Y = val(ReadField(3, Lector.GetValue("INIT", "ResPos"), 45))
      
    recordusuarios = val(Lector.GetValue("INIT", "Record"))
      
    'Max users
    Temporal = val(Lector.GetValue("INIT", "MaxUsers"))
    If MaxUsers = 0 Then
        MaxUsers = Temporal
        ReDim UserList(1 To MaxUsers) As user
    End If
    NumCuentas = val(Lector.GetValue("INIT", "NumCuentas"))
    frmMain.cuentas.Caption = NumCuentas
    #If DEBUGGING Then
    'Shell App.Path & "\estadisticas.exe" & " " & "NUEVACUENTALADDER" & "*" & NumCuentas & "*" & MaxUsers
    #End If
    
    '&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    'Se agregó en LoadBalance y en el Balance.dat
    'PorcentajeRecuperoMana = val(Lector.GetValue("BALANCE", "PorcentajeRecuperoMana"))
    
    ''&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    'Call Statistics.Initialize
    
    Nix.Map = GetVar(DatPath & "Ciudades.dat", "NIX", "Mapa")
    Nix.X = GetVar(DatPath & "Ciudades.dat", "NIX", "X")
    Nix.Y = GetVar(DatPath & "Ciudades.dat", "NIX", "Y")
    
    Ullathorpe.Map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
    Ullathorpe.X = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
    Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")
    
    Banderbill.Map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
    Banderbill.X = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
    Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")
    
    Lindos.Map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
    Lindos.X = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
    Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")
    
    Arghal.Map = GetVar(DatPath & "Ciudades.dat", "Arghal", "Mapa")
    Arghal.X = GetVar(DatPath & "Ciudades.dat", "Arghal", "X")
    Arghal.Y = GetVar(DatPath & "Ciudades.dat", "Arghal", "Y")
    
    Hillidan.Map = GetVar(DatPath & "Ciudades.dat", "Hillidan", "Mapa")
    Hillidan.X = GetVar(DatPath & "Ciudades.dat", "Hillidan", "X")
    Hillidan.Y = GetVar(DatPath & "Ciudades.dat", "Hillidan", "Y")

    CityNix.Map = GetVar(DatPath & "Ciudades.dat", "NIX", "Mapa")
    CityNix.X = GetVar(DatPath & "Ciudades.dat", "NIX", "X")
    CityNix.Y = GetVar(DatPath & "Ciudades.dat", "NIX", "Y")
    CityNix.MapaViaje = GetVar(DatPath & "Ciudades.dat", "NIX", "MapaViaje")
    CityNix.ViajeX = GetVar(DatPath & "Ciudades.dat", "NIX", "ViajeX")
    CityNix.ViajeY = GetVar(DatPath & "Ciudades.dat", "NIX", "ViajeY")
    CityNix.MapaResu = GetVar(DatPath & "Ciudades.dat", "NIX", "MapaResu")
    CityNix.ResuX = GetVar(DatPath & "Ciudades.dat", "NIX", "ResuX")
    CityNix.ResuY = GetVar(DatPath & "Ciudades.dat", "NIX", "ResuY")
    CityNix.NecesitaNave = GetVar(DatPath & "Ciudades.dat", "NIX", "NecesitaNave")

    CityUllathorpe.Map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
    CityUllathorpe.X = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
    CityUllathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")
    CityUllathorpe.MapaViaje = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "MapaViaje")
    CityUllathorpe.ViajeX = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "ViajeX")
    CityUllathorpe.ViajeY = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "ViajeY")
    CityUllathorpe.MapaResu = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "MapaResu")
    CityUllathorpe.ResuX = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "ResuX")
    CityUllathorpe.ResuY = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "ResuY")
    CityUllathorpe.NecesitaNave = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "NecesitaNave")
    
    CityBanderbill.Map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
    CityBanderbill.X = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
    CityBanderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")
    CityBanderbill.MapaViaje = GetVar(DatPath & "Ciudades.dat", "Banderbill", "MapaViaje")
    CityBanderbill.ViajeX = GetVar(DatPath & "Ciudades.dat", "Banderbill", "ViajeX")
    CityBanderbill.ViajeY = GetVar(DatPath & "Ciudades.dat", "Banderbill", "ViajeY")
    CityBanderbill.MapaResu = GetVar(DatPath & "Ciudades.dat", "Banderbill", "MapaResu")
    CityBanderbill.ResuX = GetVar(DatPath & "Ciudades.dat", "Banderbill", "ResuX")
    CityBanderbill.ResuY = GetVar(DatPath & "Ciudades.dat", "Banderbill", "ResuY")
    CityBanderbill.NecesitaNave = GetVar(DatPath & "Ciudades.dat", "Banderbill", "NecesitaNave")
    
    CityLindos.Map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
    CityLindos.X = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
    CityLindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")
    CityLindos.MapaViaje = GetVar(DatPath & "Ciudades.dat", "Lindos", "MapaViaje")
    CityLindos.ViajeX = GetVar(DatPath & "Ciudades.dat", "Lindos", "ViajeX")
    CityLindos.ViajeY = GetVar(DatPath & "Ciudades.dat", "Lindos", "ViajeY")
    CityLindos.MapaResu = GetVar(DatPath & "Ciudades.dat", "Lindos", "MapaResu")
    CityLindos.ResuX = GetVar(DatPath & "Ciudades.dat", "Lindos", "ResuX")
    CityLindos.ResuY = GetVar(DatPath & "Ciudades.dat", "Lindos", "ResuY")
    CityLindos.NecesitaNave = GetVar(DatPath & "Ciudades.dat", "Lindos", "NecesitaNave")
    
    CityArghal.Map = GetVar(DatPath & "Ciudades.dat", "Arghal", "Mapa")
    CityArghal.X = GetVar(DatPath & "Ciudades.dat", "Arghal", "X")
    CityArghal.Y = GetVar(DatPath & "Ciudades.dat", "Arghal", "Y")
    CityArghal.MapaViaje = GetVar(DatPath & "Ciudades.dat", "Arghal", "MapaViaje")
    CityArghal.ViajeX = GetVar(DatPath & "Ciudades.dat", "Arghal", "ViajeX")
    CityArghal.ViajeY = GetVar(DatPath & "Ciudades.dat", "Arghal", "ViajeY")
    CityArghal.MapaResu = GetVar(DatPath & "Ciudades.dat", "Arghal", "MapaResu")
    CityArghal.ResuX = GetVar(DatPath & "Ciudades.dat", "Arghal", "ResuX")
    CityArghal.ResuY = GetVar(DatPath & "Ciudades.dat", "Arghal", "ResuY")
    CityArghal.NecesitaNave = GetVar(DatPath & "Ciudades.dat", "Arghal", "NecesitaNave")
    
    CityHillidan.Map = GetVar(DatPath & "Ciudades.dat", "Hillidan", "Mapa")
    CityHillidan.X = GetVar(DatPath & "Ciudades.dat", "Hillidan", "X")
    CityHillidan.Y = GetVar(DatPath & "Ciudades.dat", "Hillidan", "Y")
    CityHillidan.MapaViaje = GetVar(DatPath & "Ciudades.dat", "Hillidan", "MapaViaje")
    CityHillidan.ViajeX = GetVar(DatPath & "Ciudades.dat", "Hillidan", "ViajeX")
    CityHillidan.ViajeY = GetVar(DatPath & "Ciudades.dat", "Hillidan", "ViajeY")
    CityHillidan.MapaResu = GetVar(DatPath & "Ciudades.dat", "Hillidan", "MapaResu")
    CityHillidan.ResuX = GetVar(DatPath & "Ciudades.dat", "Hillidan", "ResuX")
    CityHillidan.ResuY = GetVar(DatPath & "Ciudades.dat", "Hillidan", "ResuY")
    CityHillidan.NecesitaNave = GetVar(DatPath & "Ciudades.dat", "Hillidan", "NecesitaNave")
    
    Call MD5sCarga
    
    Call ConsultaPopular.LoadData
    
    Set Lector = Nothing

End Sub
Sub LoadConfiguraciones()
ExpMult = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "ExpMult"))
OroMult = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "OroMult"))
OroAutoEquipable = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "OroAutoEquipable"))
DropMult = val(GetVar(IniPath & "Configuracion.ini", "DROPEO", "DropMult"))
DropActive = val(GetVar(IniPath & "Configuracion.ini", "DROPEO", "DropActive"))
RecoleccionMult = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "RecoleccionMult"))

TimerCleanWorld = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "TimerCleanWorld"))
OroPorNivel = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "OroPorNivel"))


TimerHoraFantasia = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "TimerHoraFantasia"))

BattleActivado = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "BattleActivado"))
BattleMinNivel = val(GetVar(IniPath & "Configuracion.ini", "CONFIGURACIONES", "BattleMinNivel"))

frmMain.HoraFantasia.Interval = TimerHoraFantasia

LimpiezaTimerMinutos = TimerCleanWorld
frmMain.lblLimpieza.Caption = "Limpieza del mundo en: " & LimpiezaTimerMinutos & " minutos."

End Sub
Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************

writeprivateprofilestring Main, Var, Value, File
    
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

    If Database_Enabled Then
        Call SaveUserDatabase(UserIndex, Logout)
    Else
        Call SaveUserBinary(UserIndex, Logout)
    End If

End Sub

Sub LoadUserBinary(ByVal UserIndex As Integer)

    'Cargamos el personaje
    Dim Leer As New clsIniReader
    
    Call Leer.Initialize(CharPath & UCase$(UserList(UserIndex).name) & ".chr")
    
    'Cargamos los datos del personaje

    Call LoadUserInit(UserIndex, Leer)
    
    
    Call LoadUserStats(UserIndex, Leer)
    
    
    Call LoadQuestStats(UserIndex, Leer)
    
    Set Leer = Nothing

End Sub

Sub SaveUserBinary(ByVal UserIndex As Integer, Optional ByVal Logout As Boolean)
    '*************************************************
    'Author: Unknown
    'Last modified: 23/01/2007
    'Saves the Users records
    '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
    '*************************************************
    
    On Error GoTo Errhandler
    
    Dim UserFile As String
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
        Put n, , "Position=" & UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.Y & vbCrLf
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
    
    If Database_Enabled Then
        Call SaveNewUserDatabase(UserIndex)
    Else
        Call SaveNewUserCharfile(UserIndex)
    End If
    
End Sub

Sub SaveNewUserCharfile(ByVal UserIndex As Integer)
    '*************************************************
    'Author: Unknown
    'Last modified: 23/01/2007
    'Saves the Users records
    '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
    '*************************************************
    
    On Error GoTo Errhandler
    
    Dim UserFile As String
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
    Put n, , "Position=" & UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.Y & vbCrLf
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

    If Database_Enabled Then
        Call SetUserLoggedDatabase(UserList(UserIndex).Id, UserList(UserIndex).AccountID)
    Else
        Call WriteVar(CharPath & UCase$(UserList(UserIndex).name) & ".chr", "INIT", "Logged", 1)
        Call WriteVar(CuentasPath & UCase$(UserList(UserIndex).Cuenta) & ".act", "INIT", "LOGEADA", 1)
    End If

End Sub

Sub SaveBattlePoints(ByVal UserIndex As Integer)
    
    If Database_Enabled Then
        Call SaveBattlePointsDatabase(UserList(UserIndex).Id, UserList(UserIndex).flags.BattlePuntos)
    Else
        Call WriteVar(CharPath & UserList(UserIndex).name & ".chr", "Battle", "Puntos", UserList(UserIndex).flags.BattlePuntos)
    End If
    
End Sub

Function Status(ByVal UserIndex As Integer) As Byte


Status = UserList(UserIndex).Faccion.Status

End Function

Sub BackUPnPc(NpcIndex As Integer)

Dim NpcNumero As Integer
Dim npcfile As String
Dim LoopC As Integer


NpcNumero = Npclist(NpcIndex).Numero

'If NpcNumero > 499 Then
'    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
'Else
    npcfile = DatPath & "bkNPCs.dat"
'End If

'General
Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", Npclist(NpcIndex).name)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", Npclist(NpcIndex).Desc)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(Npclist(NpcIndex).Char.Head))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(Npclist(NpcIndex).Char.Body))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(Npclist(NpcIndex).Char.heading))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(Npclist(NpcIndex).Movement))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(Npclist(NpcIndex).Attackable))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(Npclist(NpcIndex).Comercia))
Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(Npclist(NpcIndex).TipoItems))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(Npclist(NpcIndex).GiveEXP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", val(Npclist(NpcIndex).GiveGLD))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(Npclist(NpcIndex).InvReSpawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(Npclist(NpcIndex).NPCtype))


'Stats
Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(Npclist(NpcIndex).Stats.Alineacion))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.def))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(Npclist(NpcIndex).Stats.MaxHit))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(Npclist(NpcIndex).Stats.MaxHp))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(Npclist(NpcIndex).Stats.MinHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(Npclist(NpcIndex).Stats.MinHp))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.UsuariosMatados)) 'Que es ESTO?!!




'Flags
Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(Npclist(NpcIndex).flags.Respawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(Npclist(NpcIndex).flags.backup))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(Npclist(NpcIndex).flags.Domable))

'Inventario
Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(Npclist(NpcIndex).Invent.NroItems))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
   For LoopC = 1 To MAX_INVENTORY_SLOTS
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & LoopC, Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex & "-" & Npclist(NpcIndex).Invent.Object(LoopC).Amount)
   Next
End If


End Sub



Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NpcNumber As Integer)

'Status
If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup Npc"

Dim npcfile As String

'If NpcNumber > 499 Then
'    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
'Else
    npcfile = DatPath & "bkNPCs.dat"
'End If

Npclist(NpcIndex).Numero = NpcNumber
Npclist(NpcIndex).name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
Npclist(NpcIndex).Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
Npclist(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
Npclist(NpcIndex).NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))

Npclist(NpcIndex).Char.Body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
Npclist(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
Npclist(NpcIndex).Char.heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))

Npclist(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
Npclist(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
Npclist(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
Npclist(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))


Npclist(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))

Npclist(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))

Npclist(NpcIndex).Stats.MaxHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
Npclist(NpcIndex).Stats.MinHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
Npclist(NpcIndex).Stats.MaxHit = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
Npclist(NpcIndex).Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
Npclist(NpcIndex).Stats.def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
Npclist(NpcIndex).Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))


Dim LoopC As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
        Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
       
    Next LoopC
Else
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = 0
        Npclist(NpcIndex).Invent.Object(LoopC).Amount = 0
    Next LoopC
End If



Npclist(NpcIndex).flags.NPCActive = True
Npclist(NpcIndex).flags.UseAINow = False
Npclist(NpcIndex).flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
Npclist(NpcIndex).flags.backup = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
Npclist(NpcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
Npclist(NpcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))

'Tipo de items con los que comercia
Npclist(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

End Sub


Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).name, "BannedBy", UserList(UserIndex).name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).name, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, UserList(BannedIndex).name
Close #mifile

End Sub


Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, BannedName
Close #mifile

End Sub


Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)



'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, BannedName
Close #mifile

End Sub

Public Sub CargaApuestas()

    Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

End Sub

Public Sub LoadRecursosEspeciales()
    If Not FileExist(DatPath & "RecursosEspeciales.dat", vbArchive) Then
        ReDim EspecialesTala(0) As obj
        ReDim EspecialesPesca(0) As obj
        Exit Sub
    End If

    Dim IniFile As clsIniReader
    Set IniFile = New clsIniReader
    
    Call IniFile.Initialize(DatPath & "RecursosEspeciales.dat")
    
    Dim Count As Long, i As Long, str As String, Field() As String
    
    ' Tala
    Count = val(IniFile.GetValue("Tala", "Items"))
    If Count > 0 Then
        ReDim EspecialesTala(1 To Count) As obj
        For i = 1 To Count
            str = IniFile.GetValue("Tala", "Item" & i)
            Field = Split(str, "-")
            
            EspecialesTala(i).ObjIndex = val(Field(0))
            EspecialesTala(i).Amount = val(Field(1))
        Next
    Else
        ReDim EspecialesTala(0) As obj
    End If
    
    ' Pesca
    Count = val(IniFile.GetValue("Pesca", "Items"))
    If Count > 0 Then
        ReDim EspecialesPesca(1 To Count) As obj
        For i = 1 To Count
            str = IniFile.GetValue("Pesca", "Item" & i)
            Field = Split(str, "-")
            
            EspecialesPesca(i).ObjIndex = val(Field(0))
            EspecialesPesca(i).Amount = val(Field(1))
        Next
    Else
        ReDim EspecialesPesca(0) As obj
    End If
    
    Set IniFile = Nothing
End Sub
