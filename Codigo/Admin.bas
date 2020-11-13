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

Public IntervaloUserPuedeUsarU      As Long

Public IntervaloUserPuedeUsarClic   As Long

Public IntervaloGolpeUsar           As Long

Public MargenDeIntervaloPorPing     As Long

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
        
        On Error GoTo VersionOK_Err
        
100     VersionOK = (Ver = ULTIMAVERSION)

        
        Exit Function

VersionOK_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.VersionOK", Erl)
        Resume Next
        
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
        
        On Error GoTo PurgarPenas_Err
        

        Dim i As Long
    
100     For i = 1 To LastUser

102         If UserList(i).flags.UserLogged Then
104             If UserList(i).Counters.Pena > 0 Then
106                 UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
108                 If UserList(i).Counters.Pena < 1 Then
110                     UserList(i).Counters.Pena = 0
112                     Call WarpUserChar(i, Libertad.Map, Libertad.x, Libertad.Y, True)
114                     Call WriteConsoleMsg(i, "Has sido liberado!", FontTypeNames.FONTTYPE_INFO)
                    
                    

                    End If

                End If

            End If

116     Next i

        
        Exit Sub

PurgarPenas_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.PurgarPenas", Erl)
        Resume Next
        
End Sub

Public Sub PurgarScroll()
        
        On Error GoTo PurgarScroll_Err
        

        Dim i As Long
    
100     For i = 1 To LastUser

102         If UserList(i).flags.UserLogged Then
104             If UserList(i).Counters.ScrollExperiencia > 0 Then
106                 UserList(i).Counters.ScrollExperiencia = UserList(i).Counters.ScrollExperiencia - 1

108                 If UserList(i).Counters.ScrollExperiencia < 1 Then
110                     UserList(i).Counters.ScrollExperiencia = 0
112                     UserList(i).flags.ScrollExp = 1
114                     Call WriteConsoleMsg(i, "Tu scroll de experiencia a finalizado.", FontTypeNames.FONTTYPE_New_DONADOR)
116                     Call WriteContadores(i)
                    

                    End If

                End If

118             If UserList(i).Counters.ScrollOro > 0 Then
120                 UserList(i).Counters.ScrollOro = UserList(i).Counters.ScrollOro - 1

122                 If UserList(i).Counters.ScrollOro < 1 Then
124                     UserList(i).Counters.ScrollOro = 0
126                     UserList(i).flags.ScrollOro = 1
128                     Call WriteConsoleMsg(i, "Tu scroll de oro a finalizado.", FontTypeNames.FONTTYPE_New_DONADOR)
130                     Call WriteContadores(i)
                    

                    End If

                End If

            End If

132     Next i

        
        Exit Sub

PurgarScroll_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.PurgarScroll", Erl)
        Resume Next
        
End Sub

Public Sub PurgarOxigeno()
        
        On Error GoTo PurgarOxigeno_Err
        

        Dim i As Long
    
100     For i = 1 To LastUser

102         If UserList(i).flags.UserLogged Then
104             If Not EsGM(i) Then
106                 If UserList(i).flags.NecesitaOxigeno Then
108                     If UserList(i).Counters.Oxigeno > 0 Then
110                         UserList(i).Counters.Oxigeno = UserList(i).Counters.Oxigeno - 1

112                         If UserList(i).Counters.Oxigeno < 1 Then
114                             UserList(i).Counters.Oxigeno = 0
116                             Call WriteOxigeno(i)
118                             Call WriteConsoleMsg(i, "Te has quedado sin oxigeno.", FontTypeNames.FONTTYPE_EJECUCION)
120                             UserList(i).flags.Ahogandose = 1
122                             Call WriteContadores(i)
                            

                            End If

                        End If

                    End If

                End If

            End If
            
124     Next i

        
        Exit Sub

PurgarOxigeno_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.PurgarOxigeno", Erl)
        Resume Next
        
End Sub

Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal minutos As Long, Optional ByVal GmName As String = vbNullString)
        
        On Error GoTo Encarcelar_Err
        
        
100     UserList(UserIndex).Counters.Pena = minutos
        
102     Call WarpUserChar(UserIndex, Prision.Map, Prision.x, Prision.Y, True)
        
104     If LenB(GmName) = 0 Then
106         Call WriteConsoleMsg(UserIndex, "Has sido encarcelado, deberas permanecer en la carcel " & minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
        Else
108         Call WriteConsoleMsg(UserIndex, GmName & " te ha encarcelado, deberas permanecer en la carcel " & minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)

        End If
        
        
        Exit Sub

Encarcelar_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.Encarcelar", Erl)
        Resume Next
        
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
        
        On Error GoTo BANCheck_Err
        

100     If Database_Enabled Then
102         BANCheck = BANCheckDatabase(name)
        Else
104         BANCheck = (val(GetVar(CharPath & name & ".chr", "BAN", "Baneado")) = 1)

        End If

        
        Exit Function

BANCheck_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.BANCheck", Erl)
        Resume Next
        
End Function

Public Function DonadorCheck(ByVal name As String) As Boolean
        
        On Error GoTo DonadorCheck_Err
        

100     If Database_Enabled Then
102         DonadorCheck = CheckUserDonatorDatabase(name)
        Else
104         DonadorCheck = val(GetVar(CuentasPath & name & ".act", "DONADOR", "DONADOR"))

        End If

        
        Exit Function

DonadorCheck_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.DonadorCheck", Erl)
        Resume Next
        
End Function

Public Function CreditosDonadorCheck(ByVal name As String) As Long
        
        On Error GoTo CreditosDonadorCheck_Err
        

100     If Database_Enabled Then
102         CreditosDonadorCheck = GetUserCreditosDatabase(name)
        Else
104         CreditosDonadorCheck = val(GetVar(CuentasPath & name & ".act", "DONADOR", "CREDITOS"))

        End If

        
        Exit Function

CreditosDonadorCheck_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.CreditosDonadorCheck", Erl)
        Resume Next
        
End Function

Public Function CreditosCanjeadosCheck(ByVal name As String) As Long
        
        On Error GoTo CreditosCanjeadosCheck_Err
        

100     If Database_Enabled Then
102         CreditosCanjeadosCheck = GetUserCreditosCanjeadosDatabase(name)
        Else
104         CreditosCanjeadosCheck = val(GetVar(CuentasPath & name & ".act", "DONADOR", "CREDITOSCANJEADOS"))

        End If

        
        Exit Function

CreditosCanjeadosCheck_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.CreditosCanjeadosCheck", Erl)
        Resume Next
        
End Function

Public Function DiasDonadorCheck(ByVal name As String) As Integer
        
        On Error GoTo DiasDonadorCheck_Err
        

100     If Database_Enabled Then
            ' Uso una funcion que hace ambas queries a la vez para optimizar
102         DiasDonadorCheck = GetUserDiasDonadorDatabase(name)
        Else

104         If DonadorCheck(name) Then

                Dim Diasrestantes As Integer

                Dim fechadonador  As Date

106             fechadonador = GetVar(CuentasPath & name & ".act", "DONADOR", "FECHAEXPIRACION")
108             DiasDonadorCheck = DateDiff("d", Date, fechadonador)

            End If

        End If

        
        Exit Function

DiasDonadorCheck_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.DiasDonadorCheck", Erl)
        Resume Next
        
End Function

Public Function ComprasDonadorCheck(ByVal name As String) As Long
        
        On Error GoTo ComprasDonadorCheck_Err
        

100     If Database_Enabled Then
102         ComprasDonadorCheck = GetUserComprasDonadorDatabase(name)
        Else
104         ComprasDonadorCheck = val(GetVar(CuentasPath & name & ".act", "COMPRAS", "CANTIDAD"))

        End If

        
        Exit Function

ComprasDonadorCheck_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.ComprasDonadorCheck", Erl)
        Resume Next
        
End Function

Public Function PersonajeExiste(ByVal name As String) As Boolean
        
        On Error GoTo PersonajeExiste_Err
        

100     If Database_Enabled Then
102         PersonajeExiste = CheckUserExists(name)
        Else
104         PersonajeExiste = FileExist(CharPath & name & ".chr", vbNormal)

        End If

        
        Exit Function

PersonajeExiste_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.PersonajeExiste", Erl)
        Resume Next
        
End Function

Public Function UnBan(ByVal name As String) As Boolean
        
        On Error GoTo UnBan_Err
        

100     If Database_Enabled Then
102         Call UnBanDatabase(name)
        Else
104         Call WriteVar(CharPath & name & ".chr", "BAN", "Baneado", "0")
106         Call WriteVar(CharPath & name & ".chr", "BAN", "BannedBy", "")
108         Call WriteVar(CharPath & name & ".chr", "BAN", "BanMotivo", "")

        End If
    
        'Remove it from the banned people database
110     Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", name, "BannedBy", "")
112     Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", name, "Reason", "")

        
        Exit Function

UnBan_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.UnBan", Erl)
        Resume Next
        
End Function

Public Function MD5ok(ByVal md5formateado As String) As Boolean
        
        On Error GoTo MD5ok_Err
        

        Dim i As Integer

100     If MD5ClientesActivado = 1 Then

102         For i = 0 To UBound(MD5s)

104             If (md5formateado = MD5s(i)) Then
106                 MD5ok = True
                    Exit Function

                End If

108         Next i

110         MD5ok = False
        Else
112         MD5ok = True

        End If

        
        Exit Function

MD5ok_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.MD5ok", Erl)
        Resume Next
        
End Function

Public Sub MD5sCarga()
        
        On Error GoTo MD5sCarga_Err
        

        Dim LoopC As Integer

100     MD5ClientesActivado = val(GetVar(IniPath & "Server.ini", "MD5Hush", "Activado"))

102     If MD5ClientesActivado = 1 Then
104         ReDim MD5s(val(GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptados")))

106         For LoopC = 0 To UBound(MD5s)
108             MD5s(LoopC) = GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptado" & (LoopC + 1))
110             MD5s(LoopC) = txtOffset(hexMd52Asc(MD5s(LoopC)), 55)
112         Next LoopC

        End If

        
        Exit Sub

MD5sCarga_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.MD5sCarga", Erl)
        Resume Next
        
End Sub

Public Sub BanIpAgrega(ByVal ip As String)
        
        On Error GoTo BanIpAgrega_Err
        
100     BanIps.Add ip
    
102     Call BanIpGuardar

        
        Exit Sub

BanIpAgrega_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.BanIpAgrega", Erl)
        Resume Next
        
End Sub

Public Function CheckHD(ByVal hd As String) As Boolean
        
        On Error GoTo CheckHD_Err
        

        '***************************************************
        'Author: Nahuel Casas (Zagen)
        'Last Modify Date: 07/12/2009
        ' 07/12/2009: Zagen - Agregè la funcion de agregar los digitos de un Serial Baneado.
        '***************************************************
        Dim handle As Integer

100     handle = FreeFile

102     Open DatPath & "\BanHds.dat" For Input As #handle

        Dim Linea As String, Total As String

104     Do Until EOF(handle)
106         Line Input #handle, Linea
108         Total = Total + Linea + vbCrLf
        Loop
110     Close #handle
    
        Dim ret As String

112     If InStr(1, Total, hd) Then
114         CheckHD = True

        End If

        
        Exit Function

CheckHD_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.CheckHD", Erl)
        Resume Next
        
End Function

Public Function CheckMAC(ByVal Mac As String) As Boolean
        
        On Error GoTo CheckMAC_Err
        

        '***************************************************
        'Author: Nahuel Casas (Zagen)
        'Last Modify Date: 07/12/2009
        ' 07/12/2009: Zagen - Agregè la funcion de agregar los digitos de un Serial Baneado.
        '***************************************************
        Dim handle As Integer

100     handle = FreeFile

102     Open DatPath & "\BanMacs.dat" For Input As #handle

        Dim Linea As String, Total As String

104     Do Until EOF(handle)
106         Line Input #handle, Linea
108         Total = Total + Linea + vbCrLf
        Loop
110     Close #handle

        Dim ret As String

112     If InStr(1, Total, Mac) Then
114         CheckMAC = True

        End If

        
        Exit Function

CheckMAC_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.CheckMAC", Erl)
        Resume Next
        
End Function

Public Function BanIpBuscar(ByVal ip As String) As Long
        
        On Error GoTo BanIpBuscar_Err
        

        Dim Dale  As Boolean

        Dim LoopC As Long

100     Dale = True
102     LoopC = 1

104     Do While LoopC <= BanIps.Count And Dale
106         Dale = (BanIps.Item(LoopC) <> ip)
108         LoopC = LoopC + 1
        Loop

110     If Dale Then
112         BanIpBuscar = 0
        Else
114         BanIpBuscar = LoopC - 1

        End If

        
        Exit Function

BanIpBuscar_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.BanIpBuscar", Erl)
        Resume Next
        
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
        
        On Error GoTo BanIpGuardar_Err
        

        Dim ArchivoBanIp As String

        Dim ArchN        As Long

        Dim LoopC        As Long

100     ArchivoBanIp = DatPath & "BanIps.dat"

102     ArchN = FreeFile()
104     Open ArchivoBanIp For Output As #ArchN

106     For LoopC = 1 To BanIps.Count
108         Print #ArchN, BanIps.Item(LoopC)
110     Next LoopC

112     Close #ArchN

        
        Exit Sub

BanIpGuardar_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.BanIpGuardar", Erl)
        Resume Next
        
End Sub

Public Sub BanIpCargar()
        
        On Error GoTo BanIpCargar_Err
        

        Dim ArchN        As Long

        Dim Tmp          As String

        Dim ArchivoBanIp As String

100     ArchivoBanIp = DatPath & "BanIps.dat"

102     Do While BanIps.Count > 0
104         BanIps.Remove 1
        Loop

106     ArchN = FreeFile()
108     Open ArchivoBanIp For Input As #ArchN

110     Do While Not EOF(ArchN)
112         Line Input #ArchN, Tmp
114         BanIps.Add Tmp
        Loop

116     Close #ArchN

        
        Exit Sub

BanIpCargar_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.BanIpCargar", Erl)
        Resume Next
        
End Sub

Public Sub ActualizaEstadisticasWeb()
        
        On Error GoTo ActualizaEstadisticasWeb_Err
        

        Static Andando  As Boolean

        Static Contador As Long

        Dim Tmp         As Boolean

100     Contador = Contador + 1

102     If Contador >= 10 Then
104         Contador = 0
106         Tmp = EstadisticasWeb.EstadisticasAndando()
    
108         If Andando = False And Tmp = True Then
110             Call InicializaEstadisticas

            End If
    
112         Andando = Tmp

        End If

        
        Exit Sub

ActualizaEstadisticasWeb_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.ActualizaEstadisticasWeb", Erl)
        Resume Next
        
End Sub

Public Sub ActualizaStatsES()
        
        On Error GoTo ActualizaStatsES_Err
        

        Static TUlt      As Single

        Dim Transcurrido As Single

100     Transcurrido = Timer - TUlt

102     If Transcurrido >= 5 Then
104         TUlt = Timer

106         With TCPESStats
108             .BytesEnviadosXSEG = CLng(.BytesEnviados / Transcurrido)
110             .BytesRecibidosXSEG = CLng(.BytesRecibidos / Transcurrido)
112             .BytesEnviados = 0
114             .BytesRecibidos = 0
        
116             If .BytesEnviadosXSEG > .BytesEnviadosXSEGMax Then
118                 .BytesEnviadosXSEGMax = .BytesEnviadosXSEG
120                 .BytesEnviadosXSEGCuando = CDate(Now)

                End If
        
122             If .BytesRecibidosXSEG > .BytesRecibidosXSEGMax Then
124                 .BytesRecibidosXSEGMax = .BytesRecibidosXSEG
126                 .BytesRecibidosXSEGCuando = CDate(Now)

                End If
        
128             If frmEstadisticas.Visible Then
130                 Call frmEstadisticas.ActualizaStats

                End If

            End With

        End If

        
        Exit Sub

ActualizaStatsES_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.ActualizaStatsES", Erl)
        Resume Next
        
End Sub

Public Function UserDarPrivilegioLevel(ByVal name As String) As PlayerType
        
        On Error GoTo UserDarPrivilegioLevel_Err
        

        '***************************************************
        'Author: Unknown
        'Last Modification: 03/02/07
        'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
        '***************************************************
100     If EsAdmin(name) Then
102         UserDarPrivilegioLevel = PlayerType.Admin
104     ElseIf EsDios(name) Then
106         UserDarPrivilegioLevel = PlayerType.Dios
108     ElseIf EsSemiDios(name) Then
110         UserDarPrivilegioLevel = PlayerType.SemiDios
112     ElseIf EsConsejero(name) Then
114         UserDarPrivilegioLevel = PlayerType.Consejero
        Else
116         UserDarPrivilegioLevel = PlayerType.user

        End If

        
        Exit Function

UserDarPrivilegioLevel_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.UserDarPrivilegioLevel", Erl)
        Resume Next
        
End Function

Public Sub BanCharacter(ByVal bannerUserIndex As Integer, ByVal UserName As String, ByVal Reason As String)
        
        On Error GoTo BanCharacter_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 03/02/07
        '
        '***************************************************
        Dim tUser     As Integer

        Dim userPriv  As Byte

        Dim cantPenas As Byte

        Dim rank      As Integer
    
100     If InStrB(UserName, "+") Then
102         UserName = Replace(UserName, "+", " ")

        End If
    
104     tUser = NameIndex(UserName)
    
106     rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
108     With UserList(bannerUserIndex)

110         If tUser <= 0 Then
112             Call WriteConsoleMsg(bannerUserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_TALK)
            
114             If PersonajeExiste(UserName) Then
116                 userPriv = UserDarPrivilegioLevel(UserName)
                
118                 If (userPriv And rank) > (.flags.Privilegios And rank) Then
120                     Call WriteConsoleMsg(bannerUserIndex, "No podes banear a al alguien de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

122                 Call LogBanFromName(UserName, bannerUserIndex, Reason)
124                 Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha baneado a " & UserName & " debido a: " & LCase$(Reason) & ".", FontTypeNames.FONTTYPE_SERVER))
                    
126                 If Database_Enabled Then
128                     Call SaveBanDatabase(UserName, Reason, .name)
                    Else
                        'ponemos el flag de ban a 1
130                     Call WriteVar(CharPath & UserName & ".chr", "BAN", "Baneado", "1")
132                     Call WriteVar(CharPath & UserName & ".chr", "BAN", "BanMotivo", LCase$(Reason))
134                     Call WriteVar(CharPath & UserName & ".chr", "BAN", "BannedBy", LCase$(.name))
            
                        'ponemos la pena
136                     cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
138                     Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
140                     Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.name) & ": " & LCase$(Reason) & " " & Date & " " & Time)

                    End If
                    
142                 If (userPriv And rank) = (.flags.Privilegios And rank) Then
144                     .flags.Ban = 1
146                     Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
148                     Call CloseSocket(bannerUserIndex)

                    End If
                    
150                 Call LogGM(.name, "BAN a " & UserName)
                Else
152                 Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

154             If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
156                 Call WriteConsoleMsg(bannerUserIndex, "No podes banear a al alguien de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
158             Call LogBan(tUser, bannerUserIndex, Reason)
160             Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha baneado a " & UserList(tUser).name & ".", FontTypeNames.FONTTYPE_SERVER))
            
                'Ponemos el flag de ban a 1
162             UserList(tUser).flags.Ban = 1
            
164             If (UserList(tUser).flags.Privilegios And rank) = (.flags.Privilegios And rank) Then
166                 .flags.Ban = 1
168                 Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
170                 Call CloseSocket(bannerUserIndex)

                End If
            
172             Call LogGM(.name, "BAN a " & UserName)
            
174             If Database_Enabled Then
176                 Call SaveBanDatabase(UserName, Reason, .name)
                Else
                    'ponemos el flag de ban a 1
178                 Call WriteVar(CharPath & UserName & ".chr", "BAN", " Baneado", "1")
180                 Call WriteVar(CharPath & UserName & ".chr", "BAN", "BanMotivo", LCase$(Reason))
182                 Call WriteVar(CharPath & UserName & ".chr", "BAN", "BannedBy", LCase$(.name))
                    'ponemos la pena
184                 cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
186                 Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
188                 Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.name) & ": " & LCase$(Reason) & " " & Date & " " & Time)

                End If
            
190             Call CloseSocket(tUser)

            End If

        End With

        
        Exit Sub

BanCharacter_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.BanCharacter", Erl)
        Resume Next
        
End Sub

Public Sub BanAccount(ByVal bannerUserIndex As Integer, ByVal UserName As String, ByVal Reason As String)
        
        On Error GoTo BanAccount_Err
        

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
    
100     If InStrB(UserName, "+") Then
102         UserName = Replace(UserName, "+", " ")

        End If
    
104     tUser = NameIndex(UserName)
    
106     rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
108     With UserList(bannerUserIndex)

110         If tUser <= 0 Then
112             Call WriteConsoleMsg(bannerUserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_SERVER)
            
114             If PersonajeExiste(UserName) Then
116                 userPriv = UserDarPrivilegioLevel(UserName)
                
118                 Cuenta = ObtenerCuenta(UserName)

                    'Call LogBanFromName(UserName, bannerUserIndex, reason)
120                 Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha baneado la cuenta de " & UserName & "(" & Cuenta & ") debido a: " & LCase$(Reason) & ".", FontTypeNames.FONTTYPE_SERVER))
                
122                 Call WriteVar(CuentasPath & Cuenta & ".act", "BAN", "Baneada", "1")
124                 Call WriteVar(CuentasPath & Cuenta & ".act", "BAN", "Motivo", Reason)
126                 Call WriteVar(CuentasPath & Cuenta & ".act", "BAN", "BANEO", .name)
            
128                 Call LogGM(.name, "Baneo la cuenta de " & UserName & " por: " & Reason)

                Else
130                 Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
132             Call WriteConsoleMsg(bannerUserIndex, "Servidor> Cuenta baneada.", FontTypeNames.FONTTYPE_SERVER)
134             Cuenta = ObtenerCuenta(UserName)
            
136             Call WriteVar(CuentasPath & Cuenta & ".act", "BAN", "Baneada", "1")
138             Call WriteVar(CuentasPath & Cuenta & ".act", "BAN", "Motivo", Reason)
140             Call WriteVar(CuentasPath & Cuenta & ".act", "BAN", "BANEO", .name)
142             Call CloseSocket(tUser)

            End If

        End With

        
        Exit Sub

BanAccount_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.BanAccount", Erl)
        Resume Next
        
End Sub

Public Sub UnBanAccount(ByVal bannerUserIndex As Integer, ByVal UserName As String)
        
        On Error GoTo UnBanAccount_Err
        

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
    
100     If InStrB(UserName, "+") Then
102         UserName = Replace(UserName, "+", " ")

        End If
    
104     tUser = NameIndex(UserName)
    
106     rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
108     With UserList(bannerUserIndex)
            
110         If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                
112             Cuenta = ObtenerCuenta(UserName)

                'Call LogBanFromName(UserName, bannerUserIndex, reason)
114             Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha desbaneado la cuenta de " & UserName & "(" & Cuenta & ").", FontTypeNames.FONTTYPE_SERVER))
                
116             Call WriteVar(CuentasPath & Cuenta & ".act", "BAN", "Baneada", "0")
118             Call WriteVar(CuentasPath & Cuenta & ".act", "BAN", "Motivo", "")
120             Call WriteVar(CuentasPath & Cuenta & ".act", "BAN", "BANEO", "")
            
122             Call LogGM(.name, "Desbaneo la cuenta de " & UserName & ".")
                
            Else
124             Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

            End If

        End With

        
        Exit Sub

UnBanAccount_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.UnBanAccount", Erl)
        Resume Next
        
End Sub

Public Sub BanSerialOK(ByVal bannerUserIndex As Integer, ByVal UserName As String)
        
        On Error GoTo BanSerialOK_Err
        

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
    
100     If InStrB(UserName, "+") Then
102         UserName = Replace(UserName, "+", " ")

        End If
    
104     tUser = NameIndex(UserName)
    
106     rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
108     With UserList(bannerUserIndex)
            
110         If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                
112             Cuenta = ObtenerCuenta(UserName)
114             Serial = ObtenerHDserial(Cuenta)
116             MacAdress = ObtenerMacAdress(Cuenta)

118             Open "" & DatPath & "\BanHds.dat" For Append As #1
120             Print #1, Serial
122             Close #1
                
124             Open "" & DatPath & "\BanMacs.dat" For Append As #1
126             Print #1, MacAdress
128             Close #1

130             Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha baneado la computadora de: " & UserName & "(" & Cuenta & ").", FontTypeNames.FONTTYPE_SERVER))
            
132             Call LogGM(.name, "Baneo la computadora de " & UserName & ".")

            Else
134             Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

            End If
        
136         If tUser > 0 Then
138             Call WriteConsoleMsg(bannerUserIndex, "Servidor> Usuario expulsado.", FontTypeNames.FONTTYPE_SERVER)
                ' Call CloseSocket(tUser)

            End If

        End With

        
        Exit Sub

BanSerialOK_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.BanSerialOK", Erl)
        Resume Next
        
End Sub

Public Sub UnBanSerialOK(ByVal bannerUserIndex As Integer, ByVal UserName As String)
        
        On Error GoTo UnBanSerialOK_Err
        

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
    
100     If InStrB(UserName, "+") Then
102         UserName = Replace(UserName, "+", " ")

        End If
    
104     tUser = NameIndex(UserName)
    
106     rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
108     With UserList(bannerUserIndex)

110         If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                
112             Cuenta = ObtenerCuenta(UserName)
114             Serial = ObtenerHDserial(Cuenta)
116             MacAdress = ObtenerMacAdress(Cuenta)
            
118             Call WriteConsoleMsg(bannerUserIndex, "Solamente desbaneo manual: HDSerial:" & Serial & ". MacAdress:" & MacAdress & ".", FontTypeNames.FONTTYPE_INFO)

            Else
120             Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

            End If
        
        End With

        
        Exit Sub

UnBanSerialOK_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.UnBanSerialOK", Erl)
        Resume Next
        
End Sub

Public Sub BanTemporal(ByVal nombre As String, ByVal dias As Integer, Causa As String, Baneador As String)
        
        On Error GoTo BanTemporal_Err
        

        Dim tBan As tBaneo

100     Set tBan = New tBaneo
102     tBan.name = UCase$(nombre)
104     tBan.FechaLiberacion = (Now + dias)
106     tBan.Causa = Causa
108     tBan.Baneador = Baneador

110     Call Baneos.Add(tBan)
112     Call SaveBan(Baneos.Count)
114     Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & nombre & " fue baneado por " & Causa & " durante los próximos " & dias & " días. La medida fue tomada por: " & Baneador, FontTypeNames.FONTTYPE_SERVER))

        
        Exit Sub

BanTemporal_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.BanTemporal", Erl)
        Resume Next
        
End Sub

Sub SaveBans()
        
        On Error GoTo SaveBans_Err
        

        Dim num As Integer

100     Call WriteVar(DatPath & "baneos.dat", "INIT", "NumeroBans", Baneos.Count)

102     For num = 1 To Baneos.Count
104         Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "USER", Baneos(num).name)
106         Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "FECHA", Baneos(num).FechaLiberacion)
108         Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "BANEADOR", Baneos(num).Baneador)
110         Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "CAUSA", Baneos(num).Causa)
        Next

        
        Exit Sub

SaveBans_Err:
        Call RegistrarError(Err.Number, Err.description, "Admin.SaveBans", Erl)
        Resume Next
        
End Sub

Sub SaveBan(num As Integer)
        
        On Error GoTo SaveBan_Err
        

100     Call WriteVar(DatPath & "baneos.dat", "INIT", "NumeroBans", Baneos.Count)
102     Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "USER", Baneos(num).name)
104     Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "FECHA", Baneos(num).FechaLiberacion)
106     Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "BANEADOR", Baneos(num).Baneador)
108     Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "CAUSA", Baneos(num).Causa)
    
110     If Database_Enabled Then
112         Call SaveBanDatabase(Baneos(num).name, Baneos(num).Causa, Baneos(num).Baneador)
        Else
114         Call WriteVar(CharPath & Baneos(num).name & ".chr", "BAN", "Baneado", "1")
116         Call WriteVar(CharPath & Baneos(num).name & ".chr", "BAN", "BanMotivo", Baneos(num).Causa)
118         Call WriteVar(CharPath & Baneos(num).name & ".chr", "BAN", "BannedBy", Baneos(num).Baneador)

        End If

        
        Exit Sub

SaveBan_Err:
        Call RegistrarError(Err.Number, Err.description, "Argentum20Server.Admin.SaveBan", Erl)
        Resume Next
        
End Sub

Sub LoadBans()
        
        On Error GoTo LoadBans_Err
        

        Dim BaneosTemporales As Integer

        Dim tBan             As tBaneo, i As Integer

100     If Not FileExist(DatPath & "baneos.dat", vbNormal) Then Exit Sub

102     BaneosTemporales = val(GetVar(DatPath & "baneos.dat", "INIT", "NumeroBans"))

104     For i = 1 To BaneosTemporales
106         Set tBan = New tBaneo

108         With tBan
110             .name = GetVar(DatPath & "baneos.dat", "BANEO" & i, "USER")
112             .FechaLiberacion = GetVar(DatPath & "baneos.dat", "BANEO" & i, "FECHA")
114             .Causa = GetVar(DatPath & "baneos.dat", "BANEO" & i, "CAUSA")
116             .Baneador = GetVar(DatPath & "baneos.dat", "BANEO" & i, "BANEADOR")
        
118             Call Baneos.Add(tBan)

            End With

        Next

        
        Exit Sub

LoadBans_Err:
        Call RegistrarError(Err.Number, Err.description, "Argentum20Server.Admin.LoadBans", Erl)
        Resume Next
        
End Sub

Public Function ChangeBan(ByVal name As String, ByVal Baneado As Byte) As Boolean
        
        On Error GoTo ChangeBan_Err
        

100     If FileExist(CharPath & name & ".chr", vbNormal) Then
102         If (val(GetVar(CharPath & name & ".chr", "BAN", "BANEADO")) = 1) Then
104             Call UnBan(name)

            End If

        End If

        
        Exit Function

ChangeBan_Err:
        Call RegistrarError(Err.Number, Err.description, "Argentum20Server.Admin.ChangeBan", Erl)
        Resume Next
        
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
        
        On Error GoTo CompararPrivilegios_Err
        

        Dim PrivilegiosGM As PlayerType
        Dim Izquierda As PlayerType
        Dim Derecha As PlayerType

100     PrivilegiosGM = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero Or PlayerType.RoleMaster

        ' Obtenemos el rango de los 2 personajes.
102     Izquierda = (UserList(Personaje1).flags.Privilegios And PrivilegiosGM)
104     Derecha = (UserList(Personaje2).flags.Privilegios And PrivilegiosGM)

106     Select Case Izquierda

            Case Is > Derecha
108             CompararPrivilegios = -1

110         Case Is = Derecha
112             CompararPrivilegios = 0

114         Case Is < Derecha
116             CompararPrivilegios = 1

        End Select

        
        Exit Function

CompararPrivilegios_Err:
        Call RegistrarError(Err.Number, Err.description, "Argentum20Server.Admin.CompararPrivilegios", Erl)
        Resume Next
        
End Function
