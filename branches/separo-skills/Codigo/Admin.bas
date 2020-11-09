Attribute VB_Name = "Admin"
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

Public Type tMotd

    texto As String
    Formato As String

End Type

Public MaxLines As Integer

Public MOTD()   As tMotd

Public Type tAPuestas

    Ganancias As Long
    Perdidas As Long
    Jugadas As Long

End Type

Public Apuestas                     As tAPuestas

Public NPCs                         As Long

Public DebugSocket                  As Boolean

Public horas                        As Long

Public dias                         As Long

Public MinsRunning                  As Long

Public ReiniciarServer              As Long

Public tInicioServer                As Long

Public EstadisticasWeb              As New clsEstadisticasIPC

'INTERVALOS
Public SanaIntervaloSinDescansar    As Integer

Public StaminaIntervaloSinDescansar As Integer

Public SanaIntervaloDescansar       As Integer

Public StaminaIntervaloDescansar    As Integer

Public IntervaloSed                 As Integer

Public IntervaloHambre              As Integer

Public IntervaloVeneno              As Integer

'Ladder
Public IntervaloIncineracion        As Integer

Public IntervaloInmovilizado        As Integer

Public IntervaloMaldicion           As Integer
'Ladder

Public IntervaloParalizado          As Integer

Public IntervaloInvisible           As Integer

Public IntervaloFrio                As Integer

Public IntervaloWavFx               As Integer

Public IntervaloNPCPuedeAtacar      As Integer

Public IntervaloNPCAI               As Integer

Public IntervaloInvocacion          As Integer

Public IntervaloOculto              As Integer '[Nacho]

Public IntervaloUserPuedeAtacar     As Long

Public IntervaloMagiaGolpe          As Long

Public IntervaloGolpeMagia          As Long

Public IntervaloUserPuedeCastear    As Long

Public IntervaloUserPuedeTrabajar   As Long

Public IntervaloCerrarConexion      As Long '[Gonzalo]

Public IntervaloUserPuedeUsar       As Long

Public IntervaloFlechasCazadores    As Long

Public TimeoutPrimerPaquete         As Long

Public TimeoutEsperandoLoggear      As Long

Public IntervaloTirar               As Long

Public IntervaloCaminar             As Long

Public IntervaloPuedeSerAtacado     As Long

'BALANCE

Public PorcentajeRecuperoMana       As Integer

Public MinutosWs                    As Long

Public Puerto                       As Integer

Public MAXPASOS                     As Long

Public BootDelBackUp                As Byte

Public Lloviendo                    As Boolean

Public Nebando                      As Boolean

Public Nieblando                    As Boolean

Public IpList                       As New Collection

Public Type TCPESStats

    BytesEnviados As Double
    BytesRecibidos As Double
    BytesEnviadosXSEG As Long
    BytesRecibidosXSEG As Long
    BytesEnviadosXSEGMax As Long
    BytesRecibidosXSEGMax As Long
    BytesEnviadosXSEGCuando As Date
    BytesRecibidosXSEGCuando As Date

End Type

Public TCPESStats As TCPESStats

Public Baneos     As New Collection

'Public ResetThread As New clsThreading

Function VersionOK(ByVal Ver As String) As Boolean
    VersionOK = (Ver = ULTIMAVERSION)

End Function

Sub ReSpawnOrigPosNpcs()

    On Error Resume Next

    Dim i     As Integer

    Dim MiNPC As npc
   
    For i = 1 To LastNPC

        'OJO
        If Npclist(i).flags.NPCActive Then
        
            If InMapBounds(Npclist(i).Orig.Map, Npclist(i).Orig.x, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)
                Call ReSpawnNpc(MiNPC)

            End If
        
            'tildada por sugerencia de yind
            'If Npclist(i).Contadores.TiempoExistencia > 0 Then
            '        Call MuereNpc(i, 0)
            'End If
        End If
   
    Next i

End Sub

Sub WorldSave()

    On Error Resume Next

    'Call LogTarea("Sub WorldSave")

    Dim LoopX As Integer

    Dim Porc  As Long

    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Iniciando WorldSave", FontTypeNames.FONTTYPE_SERVER))

    Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales

    Dim j As Integer, K As Integer

    For j = 1 To NumMaps

        If MapInfo(j).backup_mode = 1 Then K = K + 1
    Next j

    FrmStat.ProgressBar1.min = 0
    FrmStat.ProgressBar1.max = K
    FrmStat.ProgressBar1.Value = 0

    For LoopX = 1 To NumMaps
        'DoEvents
    
        If MapInfo(LoopX).backup_mode = 1 Then
    
            '  Call GrabarMapa(LoopX, App.Path & "\WorldBackUp\Mapa" & LoopX)
            FrmStat.ProgressBar1.Value = FrmStat.ProgressBar1.Value + 1

        End If

    Next LoopX

    FrmStat.Visible = False

    'If FileExist(DatPath & "\bkNpc.dat", vbNormal) Then Kill (DatPath & "bkNpc.dat")
    'If FileExist(DatPath & "\bkNPCs-HOSTILES.dat", vbNormal) Then Kill (DatPath & "bkNPCs-HOSTILES.dat")

    'For LoopX = 1 To LastNPC
    '    If Npclist(LoopX).flags.BackUp = 1 Then
    '            Call BackUPnPc(LoopX)
    '    End If
    'Next

    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> WorldSave ha concluído", FontTypeNames.FONTTYPE_SERVER))

End Sub

Public Sub PurgarPenas()

    Dim i As Long
    
    For i = 1 To LastUser

        If UserList(i).flags.UserLogged Then
            If UserList(i).Counters.Pena > 0 Then
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, Libertad.Map, Libertad.x, Libertad.Y, True)
                    Call WriteConsoleMsg(i, "Has sido liberado!", FontTypeNames.FONTTYPE_INFO)
                    
                    Call FlushBuffer(i)

                End If

            End If

        End If

    Next i

End Sub

Public Sub PurgarScroll()

    Dim i As Long
    
    For i = 1 To LastUser

        If UserList(i).flags.UserLogged Then
            If UserList(i).Counters.ScrollExperiencia > 0 Then
                UserList(i).Counters.ScrollExperiencia = UserList(i).Counters.ScrollExperiencia - 1

                If UserList(i).Counters.ScrollExperiencia < 1 Then
                    UserList(i).Counters.ScrollExperiencia = 0
                    UserList(i).flags.ScrollExp = 1
                    Call WriteConsoleMsg(i, "Tu scroll de experiencia a finalizado.", FontTypeNames.FONTTYPE_New_DONADOR)
                    Call WriteContadores(i)
                    Call FlushBuffer(i)

                End If

            End If

            If UserList(i).Counters.ScrollOro > 0 Then
                UserList(i).Counters.ScrollOro = UserList(i).Counters.ScrollOro - 1

                If UserList(i).Counters.ScrollOro < 1 Then
                    UserList(i).Counters.ScrollOro = 0
                    UserList(i).flags.ScrollOro = 1
                    Call WriteConsoleMsg(i, "Tu scroll de oro a finalizado.", FontTypeNames.FONTTYPE_New_DONADOR)
                    Call WriteContadores(i)
                    Call FlushBuffer(i)

                End If

            End If

        End If

    Next i

End Sub

Public Sub PurgarOxigeno()

    Dim i As Long
    
    For i = 1 To LastUser

        If UserList(i).flags.UserLogged Then
            If Not EsGM(i) Then
                If UserList(i).flags.NecesitaOxigeno Then
                    If UserList(i).Counters.Oxigeno > 0 Then
                        UserList(i).Counters.Oxigeno = UserList(i).Counters.Oxigeno - 1

                        If UserList(i).Counters.Oxigeno < 1 Then
                            UserList(i).Counters.Oxigeno = 0
                            Call WriteOxigeno(i)
                            Call WriteConsoleMsg(i, "Te has quedado sin oxigeno.", FontTypeNames.FONTTYPE_EJECUCION)
                            UserList(i).flags.Ahogandose = 1
                            Call WriteContadores(i)
                            Call FlushBuffer(i)

                        End If

                    End If

                End If

            End If

        End If
            
    Next i

End Sub

Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal minutos As Long, Optional ByVal GmName As String = vbNullString)
        
    UserList(UserIndex).Counters.Pena = minutos
        
    Call WarpUserChar(UserIndex, Prision.Map, Prision.x, Prision.Y, True)
        
    If LenB(GmName) = 0 Then
        Call WriteConsoleMsg(UserIndex, "Has sido encarcelado, deberas permanecer en la carcel " & minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, GmName & " te ha encarcelado, deberas permanecer en la carcel " & minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)

    End If
        
End Sub

Public Sub BorrarUsuario(ByVal UserName As String)
    
    If Database_Enabled Then
        Call BorrarUsuarioDatabase(UserName)
    
    Else

        On Error Resume Next
        
        If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
            Kill CharPath & UCase$(UserName) & ".chr"

        End If

    End If
    
End Sub

Public Function BANCheck(ByVal name As String) As Boolean

    If Database_Enabled Then
        BANCheck = BANCheckDatabase(name)
    Else
        BANCheck = (val(GetVar(CharPath & name & ".chr", "BAN", "Baneado")) = 1)

    End If

End Function

Public Function DonadorCheck(ByVal name As String) As Boolean

    If Database_Enabled Then
        DonadorCheck = CheckUserDonatorDatabase(name)
    Else
        DonadorCheck = val(GetVar(CuentasPath & name & ".act", "DONADOR", "DONADOR"))

    End If

End Function

Public Function CreditosDonadorCheck(ByVal name As String) As Long

    If Database_Enabled Then
        CreditosDonadorCheck = GetUserCreditosDatabase(name)
    Else
        CreditosDonadorCheck = val(GetVar(CuentasPath & name & ".act", "DONADOR", "CREDITOS"))

    End If

End Function

Public Function CreditosCanjeadosCheck(ByVal name As String) As Long

    If Database_Enabled Then
        CreditosCanjeadosCheck = GetUserCreditosCanjeadosDatabase(name)
    Else
        CreditosCanjeadosCheck = val(GetVar(CuentasPath & name & ".act", "DONADOR", "CREDITOSCANJEADOS"))

    End If

End Function

Public Function DiasDonadorCheck(ByVal name As String) As Integer

    If Database_Enabled Then
        ' Uso una funcion que hace ambas queries a la vez para optimizar
        DiasDonadorCheck = GetUserDiasDonadorDatabase(name)
    Else

        If DonadorCheck(name) Then

            Dim Diasrestantes As Integer

            Dim fechadonador  As Date

            fechadonador = GetVar(CuentasPath & name & ".act", "DONADOR", "FECHAEXPIRACION")
            DiasDonadorCheck = DateDiff("d", Date, fechadonador)

        End If

    End If

End Function

Public Function ComprasDonadorCheck(ByVal name As String) As Long

    If Database_Enabled Then
        ComprasDonadorCheck = GetUserComprasDonadorDatabase(name)
    Else
        ComprasDonadorCheck = val(GetVar(CuentasPath & name & ".act", "COMPRAS", "CANTIDAD"))

    End If

End Function

Public Function PersonajeExiste(ByVal name As String) As Boolean

    If Database_Enabled Then
        PersonajeExiste = CheckUserExists(name)
    Else
        PersonajeExiste = FileExist(CharPath & name & ".chr", vbNormal)

    End If

End Function

Public Function UnBan(ByVal name As String) As Boolean

    If Database_Enabled Then
        Call UnBanDatabase(name)
    Else
        Call WriteVar(CharPath & name & ".chr", "BAN", "Baneado", "0")
        Call WriteVar(CharPath & name & ".chr", "BAN", "BannedBy", "")
        Call WriteVar(CharPath & name & ".chr", "BAN", "BanMotivo", "")

    End If
    
    'Remove it from the banned people database
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", name, "BannedBy", "")
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", name, "Reason", "")

End Function

Public Function MD5ok(ByVal md5formateado As String) As Boolean

    Dim i As Integer

    If MD5ClientesActivado = 1 Then

        For i = 0 To UBound(MD5s)

            If (md5formateado = MD5s(i)) Then
                MD5ok = True
                Exit Function

            End If

        Next i

        MD5ok = False
    Else
        MD5ok = True

    End If

End Function

Public Sub MD5sCarga()

    Dim LoopC As Integer

    MD5ClientesActivado = val(GetVar(IniPath & "Server.ini", "MD5Hush", "Activado"))

    If MD5ClientesActivado = 1 Then
        ReDim MD5s(val(GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptados")))

        For LoopC = 0 To UBound(MD5s)
            MD5s(LoopC) = GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptado" & (LoopC + 1))
            MD5s(LoopC) = txtOffset(hexMd52Asc(MD5s(LoopC)), 55)
        Next LoopC

    End If

End Sub

Public Sub BanIpAgrega(ByVal ip As String)
    BanIps.Add ip
    
    Call BanIpGuardar

End Sub

Public Function CheckHD(ByVal hd As String) As Boolean

    '***************************************************
    'Author: Nahuel Casas (Zagen)
    'Last Modify Date: 07/12/2009
    ' 07/12/2009: Zagen - Agregè la funcion de agregar los digitos de un Serial Baneado.
    '***************************************************
    Dim handle As Integer

    handle = FreeFile

    Open DatPath & "\BanHds.dat" For Input As #handle

    Dim Linea As String, Total As String

    Do Until EOF(handle)
        Line Input #handle, Linea
        Total = Total + Linea + vbCrLf
    Loop
    Close #handle
    
    Dim Ret As String

    If InStr(1, Total, hd) Then
        CheckHD = True

    End If

End Function

Public Function CheckMAC(ByVal Mac As String) As Boolean

    '***************************************************
    'Author: Nahuel Casas (Zagen)
    'Last Modify Date: 07/12/2009
    ' 07/12/2009: Zagen - Agregè la funcion de agregar los digitos de un Serial Baneado.
    '***************************************************
    Dim handle As Integer

    handle = FreeFile

    Open DatPath & "\BanMacs.dat" For Input As #handle

    Dim Linea As String, Total As String

    Do Until EOF(handle)
        Line Input #handle, Linea
        Total = Total + Linea + vbCrLf
    Loop
    Close #handle

    Dim Ret As String

    If InStr(1, Total, Mac) Then
        CheckMAC = True

    End If

End Function

Public Function BanIpBuscar(ByVal ip As String) As Long

    Dim Dale  As Boolean

    Dim LoopC As Long

    Dale = True
    LoopC = 1

    Do While LoopC <= BanIps.Count And Dale
        Dale = (BanIps.Item(LoopC) <> ip)
        LoopC = LoopC + 1
    Loop

    If Dale Then
        BanIpBuscar = 0
    Else
        BanIpBuscar = LoopC - 1

    End If

End Function

Public Function BanIpQuita(ByVal ip As String) As Boolean

    On Error Resume Next

    Dim n As Long

    n = BanIpBuscar(ip)

    If n > 0 Then
        BanIps.Remove n
        BanIpGuardar
        BanIpQuita = True
    Else
        BanIpQuita = False

    End If

End Function

Public Sub BanIpGuardar()

    Dim ArchivoBanIp As String

    Dim ArchN        As Long

    Dim LoopC        As Long

    ArchivoBanIp = DatPath & "BanIps.dat"

    ArchN = FreeFile()
    Open ArchivoBanIp For Output As #ArchN

    For LoopC = 1 To BanIps.Count
        Print #ArchN, BanIps.Item(LoopC)
    Next LoopC

    Close #ArchN

End Sub

Public Sub BanIpCargar()

    Dim ArchN        As Long

    Dim Tmp          As String

    Dim ArchivoBanIp As String

    ArchivoBanIp = DatPath & "BanIps.dat"

    Do While BanIps.Count > 0
        BanIps.Remove 1
    Loop

    ArchN = FreeFile()
    Open ArchivoBanIp For Input As #ArchN

    Do While Not EOF(ArchN)
        Line Input #ArchN, Tmp
        BanIps.Add Tmp
    Loop

    Close #ArchN

End Sub

Public Sub ActualizaEstadisticasWeb()

    Static Andando  As Boolean

    Static Contador As Long

    Dim Tmp         As Boolean

    Contador = Contador + 1

    If Contador >= 10 Then
        Contador = 0
        Tmp = EstadisticasWeb.EstadisticasAndando()
    
        If Andando = False And Tmp = True Then
            Call InicializaEstadisticas

        End If
    
        Andando = Tmp

    End If

End Sub

Public Sub ActualizaStatsES()

    Static TUlt      As Single

    Dim Transcurrido As Single

    Transcurrido = Timer - TUlt

    If Transcurrido >= 5 Then
        TUlt = Timer

        With TCPESStats
            .BytesEnviadosXSEG = CLng(.BytesEnviados / Transcurrido)
            .BytesRecibidosXSEG = CLng(.BytesRecibidos / Transcurrido)
            .BytesEnviados = 0
            .BytesRecibidos = 0
        
            If .BytesEnviadosXSEG > .BytesEnviadosXSEGMax Then
                .BytesEnviadosXSEGMax = .BytesEnviadosXSEG
                .BytesEnviadosXSEGCuando = CDate(Now)

            End If
        
            If .BytesRecibidosXSEG > .BytesRecibidosXSEGMax Then
                .BytesRecibidosXSEGMax = .BytesRecibidosXSEG
                .BytesRecibidosXSEGCuando = CDate(Now)

            End If
        
            If frmEstadisticas.Visible Then
                Call frmEstadisticas.ActualizaStats

            End If

        End With

    End If

End Sub

Public Function UserDarPrivilegioLevel(ByVal name As String) As PlayerType

    '***************************************************
    'Author: Unknown
    'Last Modification: 03/02/07
    'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
    '***************************************************
    If EsAdmin(name) Then
        UserDarPrivilegioLevel = PlayerType.Admin
    ElseIf EsDios(name) Then
        UserDarPrivilegioLevel = PlayerType.Dios
    ElseIf EsSemiDios(name) Then
        UserDarPrivilegioLevel = PlayerType.SemiDios
    ElseIf EsConsejero(name) Then
        UserDarPrivilegioLevel = PlayerType.Consejero
    Else
        UserDarPrivilegioLevel = PlayerType.user

    End If

End Function

Public Sub BanCharacter(ByVal bannerUserIndex As Integer, ByVal UserName As String, ByVal Reason As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 03/02/07
    '
    '***************************************************
    Dim tUser     As Integer

    Dim userPriv  As Byte

    Dim cantPenas As Byte

    Dim rank      As Integer
    
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")

    End If
    
    tUser = NameIndex(UserName)
    
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
    With UserList(bannerUserIndex)

        If tUser <= 0 Then
            Call WriteConsoleMsg(bannerUserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_TALK)
            
            If PersonajeExiste(UserName) Then
                userPriv = UserDarPrivilegioLevel(UserName)
                
                If (userPriv And rank) > (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(bannerUserIndex, "No podes banear a al alguien de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

                Call LogBanFromName(UserName, bannerUserIndex, Reason)
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha baneado a " & UserName & " debido a: " & LCase$(Reason) & ".", FontTypeNames.FONTTYPE_SERVER))
                    
                If Database_Enabled Then
                    Call SaveBanDatabase(UserName, Reason, .name)
                Else
                    'ponemos el flag de ban a 1
                    Call WriteVar(CharPath & UserName & ".chr", "BAN", "Baneado", "1")
                    Call WriteVar(CharPath & UserName & ".chr", "BAN", "BanMotivo", LCase$(Reason))
                    Call WriteVar(CharPath & UserName & ".chr", "BAN", "BannedBy", LCase$(.name))
            
                    'ponemos la pena
                    cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.name) & ": " & LCase$(Reason) & " " & Date & " " & Time)

                End If
                    
                If (userPriv And rank) = (.flags.Privilegios And rank) Then
                    .flags.Ban = 1
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                    Call CloseSocket(bannerUserIndex)

                End If
                    
                Call LogGM(.name, "BAN a " & UserName)
            Else
                Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

            End If

        Else

            If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                Call WriteConsoleMsg(bannerUserIndex, "No podes banear a al alguien de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            Call LogBan(tUser, bannerUserIndex, Reason)
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha baneado a " & UserList(tUser).name & ".", FontTypeNames.FONTTYPE_SERVER))
            
            'Ponemos el flag de ban a 1
            UserList(tUser).flags.Ban = 1
            
            If (UserList(tUser).flags.Privilegios And rank) = (.flags.Privilegios And rank) Then
                .flags.Ban = 1
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                Call CloseSocket(bannerUserIndex)

            End If
            
            Call LogGM(.name, "BAN a " & UserName)
            
            If Database_Enabled Then
                Call SaveBanDatabase(UserName, Reason, .name)
            Else
                'ponemos el flag de ban a 1
                Call WriteVar(CharPath & UserName & ".chr", "BAN", " Baneado", "1")
                Call WriteVar(CharPath & UserName & ".chr", "BAN", "BanMotivo", LCase$(Reason))
                Call WriteVar(CharPath & UserName & ".chr", "BAN", "BannedBy", LCase$(.name))
                'ponemos la pena
                cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.name) & ": " & LCase$(Reason) & " " & Date & " " & Time)

            End If
            
            Call CloseSocket(tUser)

        End If

    End With

End Sub

Public Sub BanAccount(ByVal bannerUserIndex As Integer, ByVal UserName As String, ByVal Reason As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 03/02/07
    '
    '***************************************************
    Dim tUser     As Integer

    Dim userPriv  As Byte

    Dim cantPenas As Byte

    Dim rank      As Integer

    Dim Cuenta    As String
    
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")

    End If
    
    tUser = NameIndex(UserName)
    
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
    With UserList(bannerUserIndex)

        If tUser <= 0 Then
            Call WriteConsoleMsg(bannerUserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_SERVER)
            
            If PersonajeExiste(UserName) Then
                userPriv = UserDarPrivilegioLevel(UserName)
                
                Cuenta = ObtenerCuenta(UserName)

                'Call LogBanFromName(UserName, bannerUserIndex, reason)
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha baneado la cuenta de " & UserName & "(" & Cuenta & ") debido a: " & LCase$(Reason) & ".", FontTypeNames.FONTTYPE_SERVER))
                
                Call WriteVar(CuentasPath & Cuenta & ".act", "BAN", "Baneada", "1")
                Call WriteVar(CuentasPath & Cuenta & ".act", "BAN", "Motivo", Reason)
                Call WriteVar(CuentasPath & Cuenta & ".act", "BAN", "BANEO", .name)
            
                Call LogGM(.name, "Baneo la cuenta de " & UserName & " por: " & Reason)

            Else
                Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

            End If

        Else
            Call WriteConsoleMsg(bannerUserIndex, "Servidor> Cuenta baneada.", FontTypeNames.FONTTYPE_SERVER)
            Cuenta = ObtenerCuenta(UserName)
            
            Call WriteVar(CuentasPath & Cuenta & ".act", "BAN", "Baneada", "1")
            Call WriteVar(CuentasPath & Cuenta & ".act", "BAN", "Motivo", Reason)
            Call WriteVar(CuentasPath & Cuenta & ".act", "BAN", "BANEO", .name)
            Call CloseSocket(tUser)

        End If

    End With

End Sub

Public Sub UnBanAccount(ByVal bannerUserIndex As Integer, ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 03/02/07
    '
    '***************************************************
    Dim tUser     As Integer

    Dim userPriv  As Byte

    Dim cantPenas As Byte

    Dim rank      As Integer

    Dim Cuenta    As String
    
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")

    End If
    
    tUser = NameIndex(UserName)
    
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
    With UserList(bannerUserIndex)
            
        If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                
            Cuenta = ObtenerCuenta(UserName)

            'Call LogBanFromName(UserName, bannerUserIndex, reason)
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha desbaneado la cuenta de " & UserName & "(" & Cuenta & ").", FontTypeNames.FONTTYPE_SERVER))
                
            Call WriteVar(CuentasPath & Cuenta & ".act", "BAN", "Baneada", "0")
            Call WriteVar(CuentasPath & Cuenta & ".act", "BAN", "Motivo", "")
            Call WriteVar(CuentasPath & Cuenta & ".act", "BAN", "BANEO", "")
            
            Call LogGM(.name, "Desbaneo la cuenta de " & UserName & ".")
                
        Else
            Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

Public Sub BanSerialOK(ByVal bannerUserIndex As Integer, ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 03/02/07
    '
    '***************************************************
    Dim tUser     As Integer

    Dim userPriv  As Byte

    Dim cantPenas As Byte

    Dim rank      As Integer

    Dim Cuenta    As String
    
    Dim Serial    As Long

    Dim MacAdress As String
    
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")

    End If
    
    tUser = NameIndex(UserName)
    
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
    With UserList(bannerUserIndex)
            
        If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                
            Cuenta = ObtenerCuenta(UserName)
            Serial = ObtenerHDserial(Cuenta)
            MacAdress = ObtenerMacAdress(Cuenta)

            Open "" & DatPath & "\BanHds.dat" For Append As #1
            Print #1, Serial
            Close #1
                
            Open "" & DatPath & "\BanMacs.dat" For Append As #1
            Print #1, MacAdress
            Close #1

            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha baneado la computadora de: " & UserName & "(" & Cuenta & ").", FontTypeNames.FONTTYPE_SERVER))
            
            Call LogGM(.name, "Baneo la computadora de " & UserName & ".")

        Else
            Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

        End If
        
        If tUser > 0 Then
            Call WriteConsoleMsg(bannerUserIndex, "Servidor> Usuario expulsado.", FontTypeNames.FONTTYPE_SERVER)
            ' Call CloseSocket(tUser)

        End If

    End With

End Sub

Public Sub UnBanSerialOK(ByVal bannerUserIndex As Integer, ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 03/02/07
    '
    '***************************************************
    Dim tUser     As Integer

    Dim userPriv  As Byte

    Dim cantPenas As Byte

    Dim rank      As Integer

    Dim Cuenta    As String
    
    Dim Serial    As Long

    Dim MacAdress As String
    
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")

    End If
    
    tUser = NameIndex(UserName)
    
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
    With UserList(bannerUserIndex)

        If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                
            Cuenta = ObtenerCuenta(UserName)
            Serial = ObtenerHDserial(Cuenta)
            MacAdress = ObtenerMacAdress(Cuenta)
            
            Call WriteConsoleMsg(bannerUserIndex, "Solamente desbaneo manual: HDSerial:" & Serial & ". MacAdress:" & MacAdress & ".", FontTypeNames.FONTTYPE_INFO)

        Else
            Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

        End If
        
    End With

End Sub

Public Sub BanTemporal(ByVal nombre As String, ByVal dias As Integer, Causa As String, Baneador As String)

    Dim tBan As tBaneo

    Set tBan = New tBaneo
    tBan.name = UCase$(nombre)
    tBan.FechaLiberacion = (Now + dias)
    tBan.Causa = Causa
    tBan.Baneador = Baneador

    Call Baneos.Add(tBan)
    Call SaveBan(Baneos.Count)
    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & nombre & " fue baneado por " & Causa & " durante los próximos " & dias & " días. La medida fue tomada por: " & Baneador, FontTypeNames.FONTTYPE_SERVER))

End Sub

Sub SaveBans()

    Dim num As Integer

    Call WriteVar(DatPath & "baneos.dat", "INIT", "NumeroBans", Baneos.Count)

    For num = 1 To Baneos.Count
        Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "USER", Baneos(num).name)
        Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "FECHA", Baneos(num).FechaLiberacion)
        Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "BANEADOR", Baneos(num).Baneador)
        Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "CAUSA", Baneos(num).Causa)
    Next

End Sub

Sub SaveBan(num As Integer)

    Call WriteVar(DatPath & "baneos.dat", "INIT", "NumeroBans", Baneos.Count)
    Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "USER", Baneos(num).name)
    Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "FECHA", Baneos(num).FechaLiberacion)
    Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "BANEADOR", Baneos(num).Baneador)
    Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "CAUSA", Baneos(num).Causa)
    
    If Database_Enabled Then
        Call SaveBanDatabase(Baneos(num).name, Baneos(num).Causa, Baneos(num).Baneador)
    Else
        Call WriteVar(CharPath & Baneos(num).name & ".chr", "BAN", "Baneado", "1")
        Call WriteVar(CharPath & Baneos(num).name & ".chr", "BAN", "BanMotivo", Baneos(num).Causa)
        Call WriteVar(CharPath & Baneos(num).name & ".chr", "BAN", "BannedBy", Baneos(num).Baneador)

    End If

End Sub

Sub LoadBans()

    Dim BaneosTemporales As Integer

    Dim tBan             As tBaneo, i As Integer

    If Not FileExist(DatPath & "baneos.dat", vbNormal) Then Exit Sub

    BaneosTemporales = val(GetVar(DatPath & "baneos.dat", "INIT", "NumeroBans"))

    For i = 1 To BaneosTemporales
        Set tBan = New tBaneo

        With tBan
            .name = GetVar(DatPath & "baneos.dat", "BANEO" & i, "USER")
            .FechaLiberacion = GetVar(DatPath & "baneos.dat", "BANEO" & i, "FECHA")
            .Causa = GetVar(DatPath & "baneos.dat", "BANEO" & i, "CAUSA")
            .Baneador = GetVar(DatPath & "baneos.dat", "BANEO" & i, "BANEADOR")
        
            Call Baneos.Add(tBan)

        End With

    Next

End Sub

Public Function ChangeBan(ByVal name As String, ByVal Baneado As Byte) As Boolean

    If FileExist(CharPath & name & ".chr", vbNormal) Then
        If (val(GetVar(CharPath & name & ".chr", "BAN", "BANEADO")) = 1) Then
            Call UnBan(name)

        End If

    End If

End Function

Public Function CompararPrivilegios(ByVal Personaje1 As Integer, ByVal Personaje2 As Integer) As Byte
'**************************************************************************************************************************
'Author: Jopi
'Last Modification: 05/07/2020
'   Funcion encargada de comparar los privilegios entre 2 Game Masters.
'   Funciona de la misma forma que el operador spaceship de PHP.
'       - Si los privilegios de el de la izquierda [Personaje1] son MAYORES que el de la derecha [Personaje2], devuelve -1
'       - Si los privilegios de el de la izquierda [Personaje1] son IGUALES que el de la derecha [Personaje2], devuelve 0
'       - Si los privilegios de el de la izquierda [Personaje1] son MENORES que el de la derecha [Personaje2], devuelve 1
'**************************************************************************************************************************

    Dim PrivilegiosGM As PlayerType
    Dim Izquierda As PlayerType
    Dim Derecha As PlayerType

    PrivilegiosGM = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero Or PlayerType.RoleMaster

    ' Obtenemos el rango de los 2 personajes.
    Izquierda = (UserList(Personaje1).flags.Privilegios And PrivilegiosGM)
    Derecha = (UserList(Personaje2).flags.Privilegios And PrivilegiosGM)

    Select Case Izquierda

        Case Is > Derecha
            CompararPrivilegios = -1

        Case Is = Derecha
            CompararPrivilegios = 0

        Case Is < Derecha
            CompararPrivilegios = 1

    End Select

End Function
