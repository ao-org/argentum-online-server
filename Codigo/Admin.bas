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

Function VersionOK(ByVal Ver As String) As Boolean

    VersionOK = (Ver = ULTIMAVERSION)

        
End Function

Sub ReSpawnOrigPosNpcs()

    On Error GoTo Handler

    Dim i     As Integer

    Dim MiNPC As npc
   
    For i = 1 To LastNPC

        'OJO
        If NpcList(i).flags.NPCActive Then
        
            If InMapBounds(NpcList(i).Orig.Map, NpcList(i).Orig.X, NpcList(i).Orig.Y) And NpcList(i).Numero = Guardias Then
                MiNPC = NpcList(i)
                Call QuitarNPC(i)
                Call ReSpawnNpc(MiNPC)

            End If

        End If
   
    Next i

    Exit Sub
        
Handler:
Call RegistrarError(Err.Number, Err.Description, "Admin.ReSpawnOrigPosNpcs", Erl)
Resume Next

End Sub

Sub WorldSave()

    On Error GoTo Handler

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
    '    If NpcList(LoopX).flags.BackUp = 1 Then
    '            Call BackUPnPc(LoopX)
    '    End If
    'Next

    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> WorldSave ha concluído", FontTypeNames.FONTTYPE_SERVER))

    Exit Sub
        
Handler:
Call RegistrarError(Err.Number, Err.Description, "Admin.WorldSave", Erl)
Resume Next

End Sub

Public Sub PurgarPenas()
        
    On Error GoTo PurgarPenas_Err
        

    Dim i As Long
    
    For i = 1 To LastUser

        If UserList(i).flags.UserLogged Then
            If UserList(i).Counters.Pena > 0 Then
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
                    Call WriteConsoleMsg(i, "Has sido liberado.", FontTypeNames.FONTTYPE_INFO)
                End If

            End If

        End If

    Next i

        
    Exit Sub

PurgarPenas_Err:
    Call RegistrarError(Err.Number, Err.Description, "Admin.PurgarPenas", Erl)
    Resume Next
        
End Sub

Public Sub PurgarScroll()
        
    On Error GoTo PurgarScroll_Err
        

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
                    

                End If

            End If

            If UserList(i).Counters.ScrollOro > 0 Then
                UserList(i).Counters.ScrollOro = UserList(i).Counters.ScrollOro - 1

                If UserList(i).Counters.ScrollOro < 1 Then
                    UserList(i).Counters.ScrollOro = 0
                    UserList(i).flags.ScrollOro = 1
                    Call WriteConsoleMsg(i, "Tu scroll de oro a finalizado.", FontTypeNames.FONTTYPE_New_DONADOR)
                    Call WriteContadores(i)
                    

                End If

            End If

        End If

    Next i

        
    Exit Sub

PurgarScroll_Err:
    Call RegistrarError(Err.Number, Err.Description, "Admin.PurgarScroll", Erl)
    Resume Next
        
End Sub

Public Sub PurgarOxigeno()
        
    On Error GoTo PurgarOxigeno_Err
        

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
                            

                        End If

                    End If

                End If

            End If

        End If
            
    Next i

        
    Exit Sub

PurgarOxigeno_Err:
    Call RegistrarError(Err.Number, Err.Description, "Admin.PurgarOxigeno", Erl)
    Resume Next
        
End Sub

Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal minutos As Long, Optional ByVal GmName As String = vbNullString)
        
    On Error GoTo Encarcelar_Err
        
    If EsGM(UserIndex) Then Exit Sub
        
    UserList(UserIndex).Counters.Pena = minutos
        
    Call WarpUserChar(UserIndex, Prision.Map, Prision.X, Prision.Y, True)
        
    If LenB(GmName) = 0 Then
        Call WriteConsoleMsg(UserIndex, "Has sido encarcelado, deberas permanecer en la carcel " & minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, GmName & " te ha encarcelado, deberas permanecer en la carcel " & minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)

    End If
        
        
    Exit Sub

Encarcelar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Admin.Encarcelar", Erl)
    Resume Next
        
End Sub

Public Sub BorrarUsuario(ByVal UserName As String)
        
    Call MakeQuery("UPDATE user SET name = CONCAT('DELETED_', name), deleted = TRUE WHERE UPPER(name) = ?;", True, UCase$(UserName))
        
End Sub

Public Function BANCheck(ByVal name As String) As Boolean

    BANCheck = CBool(GetUserValue(name, "is_banned"))
    
End Function

Public Function DonadorCheck(ByVal name As String) As Boolean

    DonadorCheck = GetCuentaValue(name, "is_donor")

End Function

Public Function CreditosDonadorCheck(ByVal name As String) As Long

    CreditosDonadorCheck = GetCuentaValue(name, "credits")

End Function

Public Function CreditosCanjeadosCheck(ByVal name As String) As Long

    CreditosCanjeadosCheck = GetCuentaValue(name, "credits_used")

End Function

Public Function DiasDonadorCheck(ByVal name As String) As Integer
        
    Dim DonadorExpire As Variant
        DonadorExpire = SanitizeNullValue(GetCuentaValue(name, "donor_expire"), False)
    
    If Not DonadorExpire Then Exit Function
    
    DiasDonadorCheck = DateDiff("d", Date, DonadorExpire)
        
End Function

Public Function ComprasDonadorCheck(ByVal name As String) As Long
        
    ComprasDonadorCheck = GetCuentaValue(name, "donor_purchases")
    
End Function

Public Function PersonajeExiste(ByVal name As String) As Boolean

    PersonajeExiste = GetUserValue(name, "COUNT(*)") > 0
        
End Function

Public Function UnBan(ByVal name As String) As Boolean
        
    Call MakeQuery("UPDATE user SET is_banned = FALSE WHERE UPPER(name) = ?;", True, UCase$(name))
    
End Function

Public Sub BanIpAgrega(ByVal ip As String)

    Call BanIps.Add(ip)
    
    Call BanIpGuardar

End Sub

Public Function CheckHD(ByVal hd As String) As Boolean
        
    On Error GoTo CheckHD_Err

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

    If InStr(1, Total, hd) Then
        CheckHD = True
    End If

    Exit Function

CheckHD_Err:
    Close #handle
    Call RegistrarError(Err.Number, Err.Description, "Admin.CheckHD", Erl)
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
        handle = FreeFile

    Open DatPath & "\BanMacs.dat" For Input As #handle

    Dim Linea As String, Total As String

    Do Until EOF(handle)
        Line Input #handle, Linea
        Total = Total + Linea + vbCrLf
    Loop
    
    Close #handle

    If InStr(1, Total, Mac) Then
        CheckMAC = True
    End If

    Exit Function

CheckMAC_Err:
    Close #handle
    Call RegistrarError(Err.Number, Err.Description, "Admin.CheckMAC", Erl)
    Resume Next
        
End Function

Public Function BanIpBuscar(ByVal ip As String) As Long

    Dim Dale  As Boolean: Dale = True
    Dim LoopC As Long: LoopC = 1

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
    
    Dim n As Long
    If n > 0 Then
        Call BanIps.Remove(n)
        Call BanIpGuardar
        
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

    ArchivoBanIp = DatPath & "BanIps.dat"

    ArchN = FreeFile()
    Open ArchivoBanIp For Output As #ArchN

    For LoopC = 1 To BanIps.Count
        Print #ArchN, BanIps.Item(LoopC)
    Next LoopC

    Close #ArchN

    Exit Sub

BanIpGuardar_Err:
    Close #ArchN
    Call RegistrarError(Err.Number, Err.Description, "Admin.BanIpGuardar", Erl)
    Resume Next
        
End Sub

Public Sub BanIpCargar()
        
    On Error GoTo BanIpCargar_Err

    Dim ArchN        As Long
    Dim Tmp          As String
    Dim ArchivoBanIp As String

    ArchivoBanIp = DatPath & "BanIps.dat"

    Do While BanIps.Count > 0
        Call BanIps.Remove(1)
    Loop

    ArchN = FreeFile()
    Open ArchivoBanIp For Input As #ArchN

    Do While Not EOF(ArchN)
        Line Input #ArchN, Tmp
        BanIps.Add Tmp
    Loop

    Close #ArchN

    Exit Sub

BanIpCargar_Err:
    Close #ArchN
    Call RegistrarError(Err.Number, Err.Description, "Admin.BanIpCargar", Erl)
    Resume Next
        
End Sub

Public Sub ActualizaEstadisticasWeb()
        
    On Error GoTo ActualizaEstadisticasWeb_Err
        
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

        
    Exit Sub

ActualizaEstadisticasWeb_Err:
    Call RegistrarError(Err.Number, Err.Description, "Admin.ActualizaEstadisticasWeb", Erl)
    Resume Next
        
End Sub

Public Sub ActualizaStatsES()
        
    On Error GoTo ActualizaStatsES_Err
        

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

        
    Exit Sub

ActualizaStatsES_Err:
    Call RegistrarError(Err.Number, Err.Description, "Admin.ActualizaStatsES", Erl)
    Resume Next
        
End Sub

Public Function UserDarPrivilegioLevel(ByVal name As String) As PlayerType
        
    On Error GoTo UserDarPrivilegioLevel_Err
        

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

        
    Exit Function

UserDarPrivilegioLevel_Err:
    Call RegistrarError(Err.Number, Err.Description, "Admin.UserDarPrivilegioLevel", Erl)
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
                    
                Call SaveBanDatabase(UserName, Reason, .name)

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

            Call SaveBanDatabase(UserName, Reason, .name)

            Call CloseSocket(tUser)

        End If

    End With

    Exit Sub

BanCharacter_Err:
    Call RegistrarError(Err.Number, Err.Description, "Admin.BanCharacter", Erl)
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
    Dim cantPenas As Byte
    Dim Cuenta    As String
    Dim AccountId As Integer
    
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")
    End If
    
    tUser = NameIndex(UserName)

    With UserList(bannerUserIndex)

        If tUser <= 0 Then
            Call WriteConsoleMsg(bannerUserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_SERVER)
            
            If PersonajeExiste(UserName) Then

                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor » " & .name & " ha baneado la cuenta de " & UserName & " debido a: " & Reason & ".", FontTypeNames.FONTTYPE_SERVER))
                
                AccountId = GetAccountIDDatabase(UserName)
                Call SaveBanCuentaDatabase(AccountId, Reason, .name)

                Call LogGM(.name, "Baneó la cuenta de " & UserName & " por: " & Reason)

            Else
            
                Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
                
            End If

        Else
        
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor » " & .name & " ha baneado la cuenta de " & UserName & " debido a: " & Reason & ".", FontTypeNames.FONTTYPE_SERVER))
            
            AccountId = UserList(tUser).AccountId
            Call SaveBanCuentaDatabase(AccountId, Reason, .name)
                
            Call LogGM(.name, "Baneó la cuenta de " & UserName & " por: " & Reason)

        End If
            
        ' Echo a todos los logueados en esta cuenta
        Dim i As Integer
        For i = 1 To LastUser
            If UserList(i).AccountId = AccountId Then
                Call WriteShowMessageBox(i, "Has sido baneado del servidor. Motivo: " & Reason)
                Call CloseSocket(i)
            End If
        Next

    End With
        
    Exit Sub

BanAccount_Err:
    Call RegistrarError(Err.Number, Err.Description, "Admin.BanAccount", Erl)
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
    
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")
    End If
    
    tUser = NameIndex(UserName)
    
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
    With UserList(bannerUserIndex)
        
        Cuenta = ObtenerCuenta(UserName)
        
        If LenB(Cuenta) = 0 Then
            Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        Call MakeQuery("UPDATE user SET is_banned = FALSE WHERE UPPER(name) = ?;", True, UCase$(UserName))
        
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha desbaneado la cuenta de " & UserName & "(" & Cuenta & ").", FontTypeNames.FONTTYPE_SERVER))

        Call LogGM(.name, "Desbaneo la cuenta de " & UserName & ".")

    End With

        
    Exit Sub

UnBanAccount_Err:
    Call RegistrarError(Err.Number, Err.Description, "Admin.UnBanAccount", Erl)
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
    
    If InStrB(UserName, "+") Then
        UserName = Replace$(UserName, "+", " ")
    End If
    
    tUser = NameIndex(UserName)
    
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
    With UserList(bannerUserIndex)

        Cuenta = ObtenerCuenta(UserName)
            
        If LenB(Cuenta) = 0 Then
            Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
            
        Serial = ObtenerHDserial(Cuenta)
        MacAdress = ObtenerMacAdress(Cuenta)

        Open DatPath & "\BanHds.dat" For Append As #1
            Print #1, Serial
        Close #1
                
        Open DatPath & "\BanMacs.dat" For Append As #1
            Print #1, MacAdress
        Close #1

        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha baneado la computadora de: " & UserName & "(" & Cuenta & ").", FontTypeNames.FONTTYPE_SERVER))
            
        Call LogGM(.name, "Baneo la computadora de " & UserName & ".")

        If tUser > 0 Then
            Call WriteConsoleMsg(bannerUserIndex, "Servidor> Usuario expulsado.", FontTypeNames.FONTTYPE_SERVER)
            ' Call CloseSocket(tUser)

        End If

    End With
        
    Exit Sub

BanSerialOK_Err:
    Call RegistrarError(Err.Number, Err.Description, "Admin.BanSerialOK", Erl)
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
    
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")

    End If
    
    tUser = NameIndex(UserName)
    
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
    With UserList(bannerUserIndex)
        
        Cuenta = ObtenerCuenta(UserName)
        
        If LenB(Cuenta) = 0 Then
            Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
            
        Serial = ObtenerHDserial(Cuenta)
        MacAdress = ObtenerMacAdress(Cuenta)
        
        ' TODO: Sacar MacAddress y HDSerial de los .dat
        
        Call WriteConsoleMsg(bannerUserIndex, "Solamente desbaneo manual: HDSerial:" & Serial & ". MacAdress:" & MacAdress & ".", FontTypeNames.FONTTYPE_INFO)
        
    End With
        
    Exit Sub

UnBanSerialOK_Err:
    Call RegistrarError(Err.Number, Err.Description, "Admin.UnBanSerialOK", Erl)

    Resume Next
        
End Sub

Public Sub BanTemporal(ByVal nombre As String, ByVal dias As Integer, Causa As String, Baneador As String)
        
    On Error GoTo BanTemporal_Err
        

    Dim tBan As tBaneo

    Set tBan = New tBaneo
    tBan.name = UCase$(nombre)
    tBan.FechaLiberacion = (Now + dias)
    tBan.Causa = Causa
    tBan.Baneador = Baneador

    Call Baneos.Add(tBan)
    Call SaveBan(Baneos.Count)
    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & nombre & " fue baneado por " & Causa & " durante los próximos " & dias & " días. La medida fue tomada por: " & Baneador, FontTypeNames.FONTTYPE_SERVER))

        
    Exit Sub

BanTemporal_Err:
    Call RegistrarError(Err.Number, Err.Description, "Admin.BanTemporal", Erl)
    Resume Next
        
End Sub

Sub SaveBans()
        
    On Error GoTo SaveBans_Err
        

    Dim num As Integer

    Call WriteVar(DatPath & "baneos.dat", "INIT", "NumeroBans", Baneos.Count)

    For num = 1 To Baneos.Count
        Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "USER", Baneos(num).name)
        Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "FECHA", Baneos(num).FechaLiberacion)
        Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "BANEADOR", Baneos(num).Baneador)
        Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "CAUSA", Baneos(num).Causa)
    Next

        
    Exit Sub

SaveBans_Err:
    Call RegistrarError(Err.Number, Err.Description, "Admin.SaveBans", Erl)
    Resume Next
        
End Sub

Sub SaveBan(num As Integer)
        
    On Error GoTo SaveBan_Err
        

    Call WriteVar(DatPath & "baneos.dat", "INIT", "NumeroBans", Baneos.Count)
    Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "USER", Baneos(num).name)
    Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "FECHA", Baneos(num).FechaLiberacion)
    Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "BANEADOR", Baneos(num).Baneador)
    Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "CAUSA", Baneos(num).Causa)
    

    Call SaveBanDatabase(Baneos(num).name, Baneos(num).Causa, Baneos(num).Baneador)

    Exit Sub

SaveBan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Admin.SaveBan", Erl)
    Resume Next
        
End Sub

Sub LoadBans()
        
    On Error GoTo LoadBans_Err
        

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

        
    Exit Sub

LoadBans_Err:
    Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Admin.LoadBans", Erl)
    Resume Next
        
End Sub

Public Function CompararPrivilegios(ByVal Personaje_1 As Integer, ByVal Personaje_2 As Integer) As Integer
'**************************************************************************************************************************
'Author: Jopi
'Last Modification: 05/07/2020
'   Funcion encargada de comparar los privilegios entre 2 Game Masters.
'   Funciona de la misma forma que el operador spaceship de PHP.
'       - Si los privilegios de el de la izquierda [Personaje1] son MAYORES que el de la derecha [Personaje2], devuelve 1
'       - Si los privilegios de el de la izquierda [Personaje1] son IGUALES que el de la derecha [Personaje2], devuelve 0
'       - Si los privilegios de el de la izquierda [Personaje1] son MENORES que el de la derecha [Personaje2], devuelve -1
'**************************************************************************************************************************
        
    On Error GoTo CompararPrivilegios_Err
        
    Dim PrivilegiosGM As PlayerType
    Dim Izquierda As PlayerType
    Dim Derecha As PlayerType

    PrivilegiosGM = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero Or PlayerType.RoleMaster

    ' Obtenemos el rango de los 2 personajes.
    Izquierda = (UserList(Personaje_1).flags.Privilegios And PrivilegiosGM)
    Derecha = (UserList(Personaje_2).flags.Privilegios And PrivilegiosGM)

    Select Case Izquierda

        Case Is > Derecha
            CompararPrivilegios = 1

        Case Is = Derecha
            CompararPrivilegios = 0

        Case Is < Derecha
            CompararPrivilegios = -1

    End Select

        
    Exit Function

CompararPrivilegios_Err:
    Call RegistrarError(Err.Number, Err.Description, "Admin.CompararPrivilegios", Erl)
    Resume Next
        
End Function
