Attribute VB_Name = "modNuevoTimer"
' Argentum 20 Game Server
'
'    Copyright (C) 2023,2025 Noland Studios LTD
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

' All timers now use raw ticks (2^32 ring) + wrap-safe elapsed comparisons.
' Pattern:
'   nowRaw = GetTickCountRaw()
'   If TicksElapsed(lastTick, nowRaw) >= interval Then
'       If Actualizar Then lastTick = nowRaw
'       result = True
'   End If

Public Function IntervaloPermiteLanzarSpell(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteLanzarSpell_Err
    Dim nowRaw As Long: nowRaw = GetTickCountRaw()
    With UserList(UserIndex)
        If TicksElapsed(.Counters.TimerLanzarSpell, nowRaw) >= .Intervals.Magia Then
            If Actualizar Then
                .Counters.TimerLanzarSpell = nowRaw
                ' Actualizo spell-attack
                .Counters.TimerMagiaGolpe = nowRaw
            End If
            IntervaloPermiteLanzarSpell = True
        End If
    End With
    Exit Function
IntervaloPermiteLanzarSpell_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteLanzarSpell", Erl)
End Function

Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteAtacar_Err
    Dim nowRaw As Long: nowRaw = GetTickCountRaw()
    With UserList(UserIndex)
        If TicksElapsed(.Counters.TimerPuedeAtacar, nowRaw) >= .Intervals.Golpe Then
            If Actualizar Then
                .Counters.TimerPuedeAtacar = nowRaw
                .Counters.TimerGolpeMagia = nowRaw
                .Counters.TimerGolpeUsar = nowRaw
            End If
            IntervaloPermiteAtacar = True
        End If
    End With
    Exit Function
IntervaloPermiteAtacar_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteAtacar", Erl)
End Function

Public Function IntervaloPermiteTirar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteTirar_Err
    Dim nowRaw As Long: nowRaw = GetTickCountRaw()
    With UserList(UserIndex)
        If TicksElapsed(.Counters.TimerTirar, nowRaw) >= IntervaloTirar Then
            If Actualizar Then .Counters.TimerTirar = nowRaw
            IntervaloPermiteTirar = True
        End If
    End With
    Exit Function
IntervaloPermiteTirar_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteTirar", Erl)
End Function

Public Function IntervaloPermiteMagiaGolpe(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteMagiaGolpe_Err
    Dim nowRaw As Long: nowRaw = GetTickCountRaw()
    With UserList(UserIndex)
        ' NOTE: original logic compared against TimerLanzarSpell; we preserve that.
        If TicksElapsed(.Counters.TimerLanzarSpell, nowRaw) >= .Intervals.MagiaGolpe Then
            If Actualizar Then .Counters.TimerMagiaGolpe = nowRaw
            IntervaloPermiteMagiaGolpe = True
        End If
    End With
    Exit Function
IntervaloPermiteMagiaGolpe_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteMagiaGolpe", Erl)
End Function

Public Function IntervaloPermiteGolpeMagia(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteGolpeMagia_Err
    Dim nowRaw As Long: nowRaw = GetTickCountRaw()
    With UserList(UserIndex)
        If TicksElapsed(.Counters.TimerGolpeMagia, nowRaw) >= .Intervals.GolpeMagia Then
            If Actualizar Then .Counters.TimerGolpeMagia = nowRaw
            IntervaloPermiteGolpeMagia = True
        End If
    End With
    Exit Function
IntervaloPermiteGolpeMagia_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteGolpeMagia", Erl)
End Function

Public Function IntervaloPermiteGolpeUsar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteGolpeUsar_Err
    Dim nowRaw As Long: nowRaw = GetTickCountRaw()
    With UserList(UserIndex)
        If TicksElapsed(.Counters.TimerGolpeUsar, nowRaw) >= .Intervals.GolpeUsar Then
            If Actualizar Then .Counters.TimerGolpeUsar = nowRaw
            IntervaloPermiteGolpeUsar = True
        End If
    End With
    Exit Function
IntervaloPermiteGolpeUsar_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteGolpeUsar", Erl)
End Function

Public Function IntervaloPermiteTrabajarExtraer(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteTrabajar_Err
    Dim nowRaw As Long: nowRaw = GetTickCountRaw()
    With UserList(UserIndex)
        If TicksElapsed(.Counters.TimerPuedeTrabajar, nowRaw) >= .Intervals.TrabajarExtraer Then
            If Actualizar Then .Counters.TimerPuedeTrabajar = nowRaw
            IntervaloPermiteTrabajarExtraer = True
        End If
    End With
    Exit Function
IntervaloPermiteTrabajar_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteTrabajar", Erl)
End Function

Public Function IntervaloPermiteTrabajarConstruir(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteTrabajar_Err
    Dim nowRaw As Long: nowRaw = GetTickCountRaw()
    With UserList(UserIndex)
        If TicksElapsed(.Counters.TimerPuedeTrabajar, nowRaw) >= .Intervals.TrabajarConstruir Then
            If Actualizar Then .Counters.TimerPuedeTrabajar = nowRaw
            IntervaloPermiteTrabajarConstruir = True
        End If
    End With
    Exit Function
IntervaloPermiteTrabajar_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteTrabajar", Erl)
End Function

Public Function IntervaloPermiteUsar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteUsar_Err
    Dim nowRaw As Long: nowRaw = GetTickCountRaw()
    With UserList(UserIndex)
        If TicksElapsed(.Counters.TimerUsar, nowRaw) >= .Intervals.UsarU Then
            If Actualizar Then .Counters.TimerUsar = nowRaw
            IntervaloPermiteUsar = True
        End If
    End With
    Exit Function
IntervaloPermiteUsar_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteUsar", Erl)
End Function

Public Function IntervaloPermiteUsarClick(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    Dim nowRaw As Long: nowRaw = GetTickCountRaw()
    With UserList(UserIndex)
        If TicksElapsed(.Counters.TimerUsarClick, nowRaw) >= .Intervals.UsarClic Then
            If Actualizar Then .Counters.TimerUsarClick = nowRaw
            IntervaloPermiteUsarClick = True
        End If
    End With
End Function

Public Function IntervaloPermiteUsarArcos(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteUsarArcos_Err
    Dim nowRaw As Long: nowRaw = GetTickCountRaw()
    With UserList(UserIndex)
        If TicksElapsed(.Counters.TimerPuedeUsarArco, nowRaw) >= .Intervals.Arco Then
            If Actualizar Then
                .Counters.TimerPuedeUsarArco = nowRaw
                ' También actualizo los otros
                .Counters.TimerPuedeAtacar = nowRaw
                .Counters.TimerLanzarSpell = nowRaw
            End If
            IntervaloPermiteUsarArcos = True
        End If
    End With
    Exit Function
IntervaloPermiteUsarArcos_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteUsarArcos", Erl)
End Function

Public Function IntervaloPermiteCaminar(ByVal UserIndex As Integer) As Boolean
    On Error GoTo IntervaloPermiteCaminar_Err
    Dim nowRaw As Long: nowRaw = GetTickCountRaw()
    With UserList(UserIndex)
        If TicksElapsed(.Counters.TimerCaminar, nowRaw) >= .Intervals.Caminar Then
            .Counters.TimerCaminar = nowRaw
            IntervaloPermiteCaminar = True
        End If
    End With
    Exit Function
IntervaloPermiteCaminar_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteCaminar", Erl)
End Function

Public Function IntervaloPermiteMoverse(ByVal NpcIndex As Integer) As Boolean
    On Error GoTo IntervaloPermiteMoverse_Err
    Dim nowRaw As Long: nowRaw = GetTickCountRaw()
    With NpcList(NpcIndex)
        If TicksElapsed(.Contadores.IntervaloMovimiento, nowRaw) >= (.IntervaloMovimiento / GetNpcSpeedModifiers(NpcIndex)) Then
            .Contadores.IntervaloMovimiento = nowRaw
            IntervaloPermiteMoverse = True
        End If
    End With
    Exit Function
IntervaloPermiteMoverse_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteMoverse", Erl)
End Function

Public Function IntervaloPermiteLanzarHechizo(ByVal NpcIndex As Integer) As Boolean
    On Error GoTo IntervaloPermiteLanzarHechizo_Err
    With NpcList(NpcIndex)
        IntervaloPermiteLanzarHechizo = (TicksElapsed(.Contadores.IntervaloLanzarHechizo, GetTickCountRaw()) >= .IntervaloLanzarHechizo)
        If IntervaloPermiteLanzarHechizo Then .Contadores.IntervaloLanzarHechizo = GetTickCountRaw()
    End With
    Exit Function
IntervaloPermiteLanzarHechizo_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteLanzarHechizo", Erl)
End Function

Public Function IntervaloPermiteAtacarNPC(ByVal NpcIndex As Integer) As Boolean
    On Error GoTo IntervaloPermiteAtacarNPC_Err
    Dim nowRaw As Long: nowRaw = GetTickCountRaw()
    With NpcList(NpcIndex)
        If TicksElapsed(.Contadores.IntervaloAtaque, nowRaw) >= .IntervaloAtaque Then
            .Contadores.IntervaloAtaque = nowRaw
            IntervaloPermiteAtacarNPC = True
        End If
    End With
    Exit Function
IntervaloPermiteAtacarNPC_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteAtacarNPC", Erl)
End Function
