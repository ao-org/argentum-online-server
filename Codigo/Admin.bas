Attribute VB_Name = "Admin"
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

Public Type t_Motd

    texto As String
    Formato As String

End Type

Public MaxLines As Integer

Public MOTD()   As t_Motd

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

Public Const IntervaloConsultaGM    As Long = 300000

'BALANCE

Public PorcentajeRecuperoMana       As Integer
Public DificultadSubirSkill         As Integer
Public RequiredSpellDisplayTime     As Integer
Public MaxInvisibleSpellDisplayTime As Integer
Public DesbalancePromedioVidas      As Single
Public RangoVidas                   As Single
Public ExpLevelUp(1 To STAT_MAXELV) As Long
Public InfluenciaPromedioVidas      As Single
Public ModDañoGolpeCritico          As Single
Public MinutosWs                    As Long
Public PlayerStunTime               As Long
Public NpcStunTime                  As Long
Public PlayerInmuneTime             As Long

Public Puerto                       As Long

Public MAXPASOS                     As Long

Public BootDelBackUp                As Byte

Public Lloviendo                    As Boolean

Public Nebando                      As Boolean

Public Nieblando                    As Boolean

Public IpList                       As New Collection

Public Baneos     As New Collection

'Public ResetThread As New clsThreading

Function VersionOK(ByVal Ver As String) As Boolean
        
        On Error GoTo VersionOK_Err
        
100     VersionOK = (Ver = ULTIMAVERSION)

        
        Exit Function

VersionOK_Err:
102     Call TraceError(Err.Number, Err.Description, "Admin.VersionOK", Erl)

        
End Function

Sub ReSpawnOrigPosNpcs()
        On Error GoTo Handler
        Dim i     As Integer
        Dim MiNPC As t_Npc
   
100     For i = 1 To LastNPC
            'OJO
102         If NpcList(i).flags.NPCActive Then
104             If InMapBounds(NpcList(i).Orig.Map, NpcList(i).Orig.X, NpcList(i).Orig.Y) And NpcList(i).Numero = Guardias Then
106                 MiNPC = NpcList(i)
108                 Call QuitarNPC(i, eResetPos)
110                 Call ReSpawnNpc(MiNPC)
                End If
            End If
112     Next i
        Exit Sub
Handler:
114 Call TraceError(Err.Number, Err.Description, "Admin.ReSpawnOrigPosNpcs", Erl)
End Sub

Sub WorldSave()

        On Error GoTo Handler

        Dim LoopX As Integer

        Dim Porc  As Long

100     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor » Iniciando WorldSave", e_FontTypeNames.FONTTYPE_SERVER))

102     Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales

        Dim j As Integer, K As Integer

104     For j = 1 To NumMaps
106         If MapInfo(j).backup_mode = 1 Then K = K + 1
108     Next j

110     FrmStat.ProgressBar1.Min = 0
112     FrmStat.ProgressBar1.max = K
114     FrmStat.ProgressBar1.Value = 0

116     For LoopX = 1 To NumMaps
            'DoEvents
    
118         If MapInfo(LoopX).backup_mode = 1 Then
120             FrmStat.ProgressBar1.Value = FrmStat.ProgressBar1.Value + 1
            End If

122     Next LoopX

124     FrmStat.Visible = False

126     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor » WorldSave ha concluído", e_FontTypeNames.FONTTYPE_SERVER))

        Exit Sub
        
Handler:
128 Call TraceError(Err.Number, Err.Description, "Admin.WorldSave", Erl)


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
114                     Call WriteConsoleMsg(i, "Has sido liberado.", e_FontTypeNames.FONTTYPE_INFO)
                    End If

                End If

            End If

116     Next i

        
        Exit Sub

PurgarPenas_Err:
118     Call TraceError(Err.Number, Err.Description, "Admin.PurgarPenas", Erl)

        
End Sub


Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal minutos As Long, Optional ByVal GmName As String = vbNullString)
        
        On Error GoTo Encarcelar_Err
        
100     If EsGM(UserIndex) Then Exit Sub
        
102     UserList(UserIndex).Counters.Pena = minutos
        
104     Call WarpUserChar(UserIndex, Prision.Map, Prision.X, Prision.Y, True)
        
106     If LenB(GmName) = 0 Then
108         Call WriteConsoleMsg(UserIndex, "Has sido encarcelado, deberas permanecer en la carcel " & minutos & " minutos.", e_FontTypeNames.FONTTYPE_INFO)
        Else
110         Call WriteConsoleMsg(UserIndex, GmName & " te ha encarcelado, deberas permanecer en la carcel " & minutos & " minutos.", e_FontTypeNames.FONTTYPE_INFO)

        End If
        
        
        Exit Sub

Encarcelar_Err:
112     Call TraceError(Err.Number, Err.Description, "Admin.Encarcelar", Erl)

        
End Sub

Public Function BANCheck(ByVal Name As String) As Boolean
        
        On Error GoTo BANCheck_Err

102         BANCheck = BANCheckDatabase(Name)
        
        Exit Function

BANCheck_Err:
106     Call TraceError(Err.Number, Err.Description, "Admin.BANCheck", Erl)

        
End Function

Public Function PersonajeExiste(ByVal Name As String) As Boolean
        
        On Error GoTo PersonajeExiste_Err

102         PersonajeExiste = GetUserValue(LCase$(Name), "COUNT(*)") > 0

        Exit Function

PersonajeExiste_Err:
106     Call TraceError(Err.Number, Err.Description, "Admin.PersonajeExiste", Erl)

        
End Function

Public Function UnBan(ByVal Name As String) As Boolean
        
        On Error GoTo UnBan_Err

102     Call UnBanDatabase(Name)
    
        'Remove it from the banned people database
110     Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "")
112     Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "")
        
        Exit Function

UnBan_Err:
114     Call TraceError(Err.Number, Err.Description, "Admin.UnBan", Erl)

        
End Function

Public Function UserDarPrivilegioLevel(ByVal Name As String) As e_PlayerType
        
        On Error GoTo UserDarPrivilegioLevel_Err
        
        '***************************************************
        'Author: Unknown
        'Last Modification: 03/02/07
        'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
        '***************************************************
100     If EsAdmin(Name) Then
102         UserDarPrivilegioLevel = e_PlayerType.Admin
104     ElseIf EsDios(Name) Then
106         UserDarPrivilegioLevel = e_PlayerType.Dios
108     ElseIf EsSemiDios(Name) Then
110         UserDarPrivilegioLevel = e_PlayerType.SemiDios
112     ElseIf EsConsejero(Name) Then
114         UserDarPrivilegioLevel = e_PlayerType.Consejero
        Else
116         UserDarPrivilegioLevel = e_PlayerType.user

        End If

        
        Exit Function

UserDarPrivilegioLevel_Err:
118     Call TraceError(Err.Number, Err.Description, "Admin.UserDarPrivilegioLevel", Erl)

        
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
114     Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor » " & nombre & " fue baneado por " & Causa & " durante los próximos " & dias & " días. La medida fue tomada por: " & Baneador, e_FontTypeNames.FONTTYPE_SERVER))

        
        Exit Sub

BanTemporal_Err:
116     Call TraceError(Err.Number, Err.Description, "Admin.BanTemporal", Erl)

        
End Sub

Sub SaveBans()
        
        On Error GoTo SaveBans_Err
        

        Dim num As Integer

100     Call WriteVar(DatPath & "baneos.dat", "INIT", "NumeroBans", Baneos.Count)

102     For num = 1 To Baneos.Count
104         Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "USER", Baneos(num).Name)
106         Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "FECHA", Baneos(num).FechaLiberacion)
108         Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "BANEADOR", Baneos(num).Baneador)
110         Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "CAUSA", Baneos(num).Causa)
        Next

        
        Exit Sub

SaveBans_Err:
112     Call TraceError(Err.Number, Err.Description, "Admin.SaveBans", Erl)

        
End Sub

Sub SaveBan(num As Integer)
        
        On Error GoTo SaveBan_Err
        

100     Call WriteVar(DatPath & "baneos.dat", "INIT", "NumeroBans", Baneos.Count)
102     Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "USER", Baneos(num).Name)
104     Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "FECHA", Baneos(num).FechaLiberacion)
106     Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "BANEADOR", Baneos(num).Baneador)
108     Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "CAUSA", Baneos(num).Causa)

112     Call SaveBanDatabase(Baneos(Num).Name, Baneos(Num).Causa, Baneos(Num).Baneador)


        
        Exit Sub

SaveBan_Err:
120     Call TraceError(Err.Number, Err.Description, "Argentum20Server.Admin.SaveBan", Erl)

        
End Sub

Sub LoadBans()
        
        On Error GoTo LoadBans_Err
        

        Dim BaneosTemporales As Integer

        Dim tBan             As tBaneo, i As Integer

100     If Not FileExist(DatPath & "baneos.dat", vbNormal) Then Exit Sub

102     BaneosTemporales = val(GetVar(DatPath & "baneos.dat", "INIT", "NumeroBans"))

        If BaneosTemporales > 0 Then
104         For i = 1 To BaneosTemporales
106             Set tBan = New tBaneo
    
108             With tBan
110                 .Name = GetVar(DatPath & "baneos.dat", "BANEO" & i, "USER")
112                 .FechaLiberacion = GetVar(DatPath & "baneos.dat", "BANEO" & i, "FECHA")
114                 .Causa = GetVar(DatPath & "baneos.dat", "BANEO" & i, "CAUSA")
116                 .Baneador = GetVar(DatPath & "baneos.dat", "BANEO" & i, "BANEADOR")
            
118                 Call Baneos.Add(tBan)
    
                End With
    
            Next
        End If

        
        Exit Sub

LoadBans_Err:
120     Call TraceError(Err.Number, Err.Description, "Argentum20Server.Admin.LoadBans", Erl)

        
End Sub

Public Function CompararUserPrivilegios(ByVal Personaje_1 As Integer, ByVal Personaje_2 As Integer) As Integer
    
100     CompararUserPrivilegios = CompararPrivilegios(UserList(Personaje_1).flags.Privilegios, UserList(Personaje_2).flags.Privilegios)
        
End Function

Public Function CompararPrivilegiosUser(ByVal Personaje_1 As Integer, ByVal Personaje_2 As Integer) As Integer
        On Error GoTo CompararPrivilegiosUser_Err
        
100     CompararPrivilegiosUser = CompararPrivilegios(UserList(Personaje_1).flags.Privilegios, UserList(Personaje_2).flags.Privilegios)
        
        Exit Function

CompararPrivilegiosUser_Err:
102     Call TraceError(Err.Number, Err.Description, "Admin.CompararPrivilegiosUser", Erl)

        
End Function

Public Function CompararPrivilegios(ByVal Izquierda As e_PlayerType, ByVal Derecha As e_PlayerType) As Integer
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
        
        Dim PrivilegiosGM As e_PlayerType
100     PrivilegiosGM = e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster

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
118     Call TraceError(Err.Number, Err.Description, "Admin.CompararPrivilegios", Erl)

        
End Function
