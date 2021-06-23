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

Public AdministratorAccounts As Dictionary

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
Public IntervaloPerderStamina       As Integer
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
Public IntervaloTrabajarExtraer     As Long

Public IntervaloTrabajarConstruir   As Long

Public IntervaloCerrarConexion      As Long '[Gonzalo]

Public IntervaloUserPuedeUsarU      As Long

Public IntervaloUserPuedeUsarClic   As Long

Public IntervaloGolpeUsar           As Long

Public MargenDeIntervaloPorPing     As Long

Public IntervaloFlechasCazadores    As Long

Public TimeoutPrimerPaquete         As Long

Public TimeoutEsperandoLoggear      As Long

Public IntervaloTirar               As Long

Public IntervaloMeditar             As Long

Public IntervaloCaminar             As Long

Public IntervaloEnCombate           As Long

Public IntervaloPuedeSerAtacado     As Long

Public IntervaloGuardarUsuarios     As Long

Public LimiteGuardarUsuarios        As Integer

Public IntervaloTimerGuardarUsuarios As Integer

Public IntervaloMensajeGlobal       As Long

'BALANCE

Public PorcentajeRecuperoMana       As Integer

Public DificultadSubirSkill         As Integer

Public DesbalancePromedioVidas      As Single

Public RangoVidas                   As Single

Public ExpLevelUp(1 To STAT_MAXELV) As Long

Public InfluenciaPromedioVidas      As Single

Public ModDañoGolpeCritico          As Single

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
102     Call RegistrarError(Err.Number, Err.Description, "Admin.VersionOK", Erl)
104     Resume Next
        
End Function

Sub ReSpawnOrigPosNpcs()

        On Error GoTo Handler

        Dim i     As Integer

        Dim MiNPC As npc
   
100     For i = 1 To LastNPC

            'OJO
102         If NpcList(i).flags.NPCActive Then
        
104             If InMapBounds(NpcList(i).Orig.Map, NpcList(i).Orig.X, NpcList(i).Orig.Y) And NpcList(i).Numero = Guardias Then
106                 MiNPC = NpcList(i)
108                 Call QuitarNPC(i)
110                 Call ReSpawnNpc(MiNPC)

                End If
        
                'tildada por sugerencia de yind
                'If NpcList(i).Contadores.TiempoExistencia > 0 Then
                '        Call MuereNpc(i, 0)
                'End If
            End If
   
112     Next i

        Exit Sub
        
Handler:
114 Call RegistrarError(Err.Number, Err.Description, "Admin.ReSpawnOrigPosNpcs", Erl)
116 Resume Next

End Sub

Sub WorldSave()

        On Error GoTo Handler

        'Call LogTarea("Sub WorldSave")

        Dim LoopX As Integer

        Dim Porc  As Long

100     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor » Iniciando WorldSave", FontTypeNames.FONTTYPE_SERVER))

102     Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales

        Dim j As Integer, K As Integer

104     For j = 1 To NumMaps

106         If MapInfo(j).backup_mode = 1 Then K = K + 1
108     Next j

110     FrmStat.ProgressBar1.min = 0
112     FrmStat.ProgressBar1.max = K
114     FrmStat.ProgressBar1.Value = 0

116     For LoopX = 1 To NumMaps
            'DoEvents
    
118         If MapInfo(LoopX).backup_mode = 1 Then
    
                '  Call GrabarMapa(LoopX, App.Path & "\WorldBackUp\Mapa" & LoopX)
120             FrmStat.ProgressBar1.Value = FrmStat.ProgressBar1.Value + 1

            End If

122     Next LoopX

124     FrmStat.Visible = False

        'If FileExist(DatPath & "\bkNpc.dat", vbNormal) Then Kill (DatPath & "bkNpc.dat")
        'If FileExist(DatPath & "\bkNPCs-HOSTILES.dat", vbNormal) Then Kill (DatPath & "bkNPCs-HOSTILES.dat")

        'For LoopX = 1 To LastNPC
        '    If NpcList(LoopX).flags.BackUp = 1 Then
        '            Call BackUPnPc(LoopX)
        '    End If
        'Next

126     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor » WorldSave ha concluído", FontTypeNames.FONTTYPE_SERVER))

        Exit Sub
        
Handler:
128 Call RegistrarError(Err.Number, Err.Description, "Admin.WorldSave", Erl)
130 Resume Next

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
112                     Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
114                     Call WriteConsoleMsg(i, "Has sido liberado.", FontTypeNames.FONTTYPE_INFO)
                    End If

                End If

            End If

116     Next i

        
        Exit Sub

PurgarPenas_Err:
118     Call RegistrarError(Err.Number, Err.Description, "Admin.PurgarPenas", Erl)
120     Resume Next
        
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
134     Call RegistrarError(Err.Number, Err.Description, "Admin.PurgarScroll", Erl)
136     Resume Next
        
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
126     Call RegistrarError(Err.Number, Err.Description, "Admin.PurgarOxigeno", Erl)
128     Resume Next
        
End Sub

Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal minutos As Long, Optional ByVal GmName As String = vbNullString)
        
        On Error GoTo Encarcelar_Err
        
100     If EsGM(UserIndex) Then Exit Sub
        
102     UserList(UserIndex).Counters.Pena = minutos
        
104     Call WarpUserChar(UserIndex, Prision.Map, Prision.X, Prision.Y, True)
        
106     If LenB(GmName) = 0 Then
108         Call WriteConsoleMsg(UserIndex, "Has sido encarcelado, deberas permanecer en la carcel " & minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
        Else
110         Call WriteConsoleMsg(UserIndex, GmName & " te ha encarcelado, deberas permanecer en la carcel " & minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)

        End If
        
        
        Exit Sub

Encarcelar_Err:
112     Call RegistrarError(Err.Number, Err.Description, "Admin.Encarcelar", Erl)
114     Resume Next
        
End Sub

Public Sub BorrarUsuario(ByVal UserName As String)
        
        On Error GoTo BorrarUsuario_Err
    
        
    
100     If Database_Enabled Then
102         Call BorrarUsuarioDatabase(UserName)
    
        Else

            
        
104         If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
106             Kill CharPath & UCase$(UserName) & ".chr"

            End If

        End If
    
        
        Exit Sub

BorrarUsuario_Err:
108     Call RegistrarError(Err.Number, Err.Description, "Admin.BorrarUsuario", Erl)

        
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
106     Call RegistrarError(Err.Number, Err.Description, "Admin.BANCheck", Erl)
108     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.Description, "Admin.DonadorCheck", Erl)
108     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.Description, "Admin.CreditosDonadorCheck", Erl)
108     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.Description, "Admin.CreditosCanjeadosCheck", Erl)
108     Resume Next
        
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
110     Call RegistrarError(Err.Number, Err.Description, "Admin.DiasDonadorCheck", Erl)
112     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.Description, "Admin.ComprasDonadorCheck", Erl)
108     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.Description, "Admin.PersonajeExiste", Erl)
108     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.Description, "Admin.UnBan", Erl)
116     Resume Next
        
End Function


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
114     Call RegistrarError(Err.Number, Err.Description, "Admin.ActualizaEstadisticasWeb", Erl)
116     Resume Next
        
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
132     Call RegistrarError(Err.Number, Err.Description, "Admin.ActualizaStatsES", Erl)
134     Resume Next
        
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
118     Call RegistrarError(Err.Number, Err.Description, "Admin.UserDarPrivilegioLevel", Erl)
120     Resume Next
        
End Function

Public Sub BanTemporal(ByVal nombre As String, ByVal dias As Integer, Causa As String, Baneador As String)
        
        On Error GoTo BanTemporal_Err
        

        Dim tBan As tBaneo
100     Set tBan = New tBaneo

102     tBan.Name = UCase$(nombre)
104     tBan.FechaLiberacion = (Now + dias)
106     tBan.Causa = Causa
108     tBan.Baneador = Baneador

110     Call Baneos.Add(tBan)
112     Call SaveBan(Baneos.Count)
114     Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor » " & nombre & " fue baneado por " & Causa & " durante los próximos " & dias & " días. La medida fue tomada por: " & Baneador, FontTypeNames.FONTTYPE_SERVER))

        
        Exit Sub

BanTemporal_Err:
116     Call RegistrarError(Err.Number, Err.Description, "Admin.BanTemporal", Erl)
118     Resume Next
        
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
112     Call RegistrarError(Err.Number, Err.Description, "Admin.SaveBans", Erl)
114     Resume Next
        
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
120     Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Admin.SaveBan", Erl)
122     Resume Next
        
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
120     Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Admin.LoadBans", Erl)
122     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Admin.ChangeBan", Erl)
108     Resume Next
        
End Function

Public Function CompararUserPrivilegios(ByVal Personaje_1 As Integer, ByVal Personaje_2 As Integer) As Integer
    
    CompararUserPrivilegios = CompararPrivilegios(UserList(Personaje_1).flags.Privilegios, UserList(Personaje_2).flags.Privilegios)
        
End Function

Public Function CompararPrivilegiosUser(ByVal Personaje_1 As Integer, ByVal Personaje_2 As Integer) As Integer
        On Error GoTo CompararPrivilegiosUser_Err
        
100     CompararPrivilegiosUser = CompararPrivilegios(UserList(Personaje_1).flags.Privilegios, UserList(Personaje_2).flags.Privilegios)
        
        Exit Function

CompararPrivilegiosUser_Err:
102     Call RegistrarError(Err.Number, Err.Description, "Admin.CompararPrivilegiosUser", Erl)
104     Resume Next
        
End Function

Public Function CompararPrivilegios(ByVal Izquierda As PlayerType, ByVal Derecha As PlayerType) As Integer
    '**************************************************************************************************************************
    'Author: Jopi
    'Last Modification: 05/07/2020
    '   Funcion encargada de comparar los privilegios entre 2 Game Masters.
    '   Funciona de la misma forma que el operador spaceship de PHP.
    '       - Si los privilegios de el de la izquierda son MAYORES que el de la derecha, devuelve 1
    '       - Si los privilegios de el de la izquierda son IGUALES que el de la derecha, devuelve 0
    '       - Si los privilegios de el de la izquierda son MENORES que el de la derecha, devuelve -1
    '**************************************************************************************************************************
        
        On Error GoTo CompararPrivilegios_Err
        
        Dim PrivilegiosGM As PlayerType
100     PrivilegiosGM = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero Or PlayerType.RoleMaster

        ' Obtenemos el rango de los 2 personajes.
102     Izquierda = (Izquierda And PrivilegiosGM)
104     Derecha = (Derecha And PrivilegiosGM)

106     Select Case Izquierda

            Case Is > Derecha
108             CompararPrivilegios = 1

110         Case Is = Derecha
112             CompararPrivilegios = 0

114         Case Is < Derecha
116             CompararPrivilegios = -1

        End Select

        
        Exit Function

CompararPrivilegios_Err:
118     Call RegistrarError(Err.Number, Err.Description, "Admin.CompararPrivilegios", Erl)
120     Resume Next
        
End Function
