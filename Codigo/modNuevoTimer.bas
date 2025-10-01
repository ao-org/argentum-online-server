Attribute VB_Name = "modNuevoTimer"
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

'
' Las siguientes funciones devuelven TRUE o FALSE si el intervalo
' permite hacerlo. Si devuelve TRUE, setean automaticamente el
' timer para que no se pueda hacer la accion hasta el nuevo ciclo.
'
' CASTING DE HECHIZOS
Public Function IntervaloPermiteLanzarSpell(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteLanzarSpell_Err
    Dim TActual As Long
    TActual = GetTickCount()
    If TActual - UserList(UserIndex).Counters.TimerLanzarSpell >= UserList(UserIndex).Intervals.Magia Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerLanzarSpell = TActual
            ' Actualizo spell-attack
            UserList(UserIndex).Counters.TimerMagiaGolpe = TActual
        End If
        IntervaloPermiteLanzarSpell = True
    Else
        IntervaloPermiteLanzarSpell = False
    End If
    Exit Function
IntervaloPermiteLanzarSpell_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteLanzarSpell", Erl)
End Function

Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteAtacar_Err
    Dim TActual As Long
    TActual = GetTickCount()
    If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= UserList(UserIndex).Intervals.Golpe Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
            ' Actualizo attack-spell
            UserList(UserIndex).Counters.TimerGolpeMagia = TActual
            ' Actualizo attack-use
            UserList(UserIndex).Counters.TimerGolpeUsar = TActual
        End If
        IntervaloPermiteAtacar = True
    Else
        IntervaloPermiteAtacar = False
    End If
    Exit Function
IntervaloPermiteAtacar_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteAtacar", Erl)
End Function

Public Function IntervaloPermiteTirar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteTirar_Err
    Dim TActual As Long
    TActual = GetTickCount()
    If TActual - UserList(UserIndex).Counters.TimerTirar >= IntervaloTirar Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerTirar = TActual
        End If
        IntervaloPermiteTirar = True
    Else
        IntervaloPermiteTirar = False
    End If
    Exit Function
IntervaloPermiteTirar_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteTirar", Erl)
End Function

Public Function IntervaloPermiteMagiaGolpe(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteMagiaGolpe_Err
    Dim TActual As Long
    TActual = GetTickCount()
    If TActual - UserList(UserIndex).Counters.TimerLanzarSpell >= UserList(UserIndex).Intervals.MagiaGolpe Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerMagiaGolpe = TActual
        End If
        IntervaloPermiteMagiaGolpe = True
    Else
        IntervaloPermiteMagiaGolpe = False
    End If
    Exit Function
IntervaloPermiteMagiaGolpe_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteMagiaGolpe", Erl)
End Function

Public Function IntervaloPermiteGolpeMagia(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteGolpeMagia_Err
    Dim TActual As Long
    TActual = GetTickCount()
    If TActual - UserList(UserIndex).Counters.TimerGolpeMagia >= UserList(UserIndex).Intervals.GolpeMagia Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerGolpeMagia = TActual
        End If
        IntervaloPermiteGolpeMagia = True
    Else
        IntervaloPermiteGolpeMagia = False
    End If
    Exit Function
IntervaloPermiteGolpeMagia_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteGolpeMagia", Erl)
End Function

Public Function IntervaloPermiteGolpeUsar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteGolpeUsar_Err
    Dim TActual As Long
    TActual = GetTickCount()
    If TActual - UserList(UserIndex).Counters.TimerGolpeUsar >= UserList(UserIndex).Intervals.GolpeUsar Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerGolpeUsar = TActual
        End If
        IntervaloPermiteGolpeUsar = True
    Else
        IntervaloPermiteGolpeUsar = False
    End If
    Exit Function
IntervaloPermiteGolpeUsar_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteGolpeUsar", Erl)
End Function

' TRABAJO
Public Function IntervaloPermiteTrabajarExtraer(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteTrabajar_Err
    Dim TActual As Long
    TActual = GetTickCount()
    If TActual - UserList(UserIndex).Counters.TimerPuedeTrabajar >= UserList(UserIndex).Intervals.TrabajarExtraer Then
        If Actualizar Then UserList(UserIndex).Counters.TimerPuedeTrabajar = TActual
        IntervaloPermiteTrabajarExtraer = True
    Else
        IntervaloPermiteTrabajarExtraer = False
    End If
    Exit Function
IntervaloPermiteTrabajar_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteTrabajar", Erl)
End Function

Public Function IntervaloPermiteTrabajarConstruir(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteTrabajar_Err
    Dim TActual As Long
    TActual = GetTickCount()
    If TActual - UserList(UserIndex).Counters.TimerPuedeTrabajar >= UserList(UserIndex).Intervals.TrabajarConstruir Then
        If Actualizar Then UserList(UserIndex).Counters.TimerPuedeTrabajar = TActual
        IntervaloPermiteTrabajarConstruir = True
    Else
        IntervaloPermiteTrabajarConstruir = False
    End If
    Exit Function
IntervaloPermiteTrabajar_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteTrabajar", Erl)
End Function

' USAR OBJETOS CON U
Public Function IntervaloPermiteUsar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteUsar_Err
    Dim TActual As Long
    TActual = GetTickCount()
    If TActual - UserList(UserIndex).Counters.TimerUsar >= UserList(UserIndex).Intervals.UsarU Then
        If Actualizar Then UserList(UserIndex).Counters.TimerUsar = TActual
        IntervaloPermiteUsar = True
    Else
        IntervaloPermiteUsar = False
    End If
    Exit Function
IntervaloPermiteUsar_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteUsar", Erl)
End Function

' USAR OBJETOS CON CLICK
Public Function IntervaloPermiteUsarClick(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    '**
    'Author: Unknown
    'Last Modification: 25/01/2010 (ZaMa)
    '25/01/2010: ZaMa - General adjustments.
    '**
    Dim TActual As Long
    TActual = GetTickCount() And &H7FFFFFFF
    If TActual - UserList(UserIndex).Counters.TimerUsarClick >= UserList(UserIndex).Intervals.UsarClic Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerUsarClick = TActual
        End If
        IntervaloPermiteUsarClick = True
    Else
        IntervaloPermiteUsarClick = False
    End If
End Function

Public Function IntervaloPermiteUsarArcos(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    On Error GoTo IntervaloPermiteUsarArcos_Err
    Dim TActual As Long
    TActual = GetTickCount()
    If TActual - UserList(UserIndex).Counters.TimerPuedeUsarArco >= UserList(UserIndex).Intervals.Arco Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerPuedeUsarArco = TActual
            ' Tambien actualizo los otros
            UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
            UserList(UserIndex).Counters.TimerLanzarSpell = TActual
        End If
        IntervaloPermiteUsarArcos = True
    Else
        IntervaloPermiteUsarArcos = False
    End If
    Exit Function
IntervaloPermiteUsarArcos_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteUsarArcos", Erl)
End Function

Public Function IntervaloPermiteMoverse(ByVal NpcIndex As Integer) As Boolean
    On Error GoTo IntervaloPermiteMoverse_Err
    Dim TActual As Long
    TActual = GetTickCount()
    If TActual - NpcList(NpcIndex).Contadores.IntervaloMovimiento >= (NpcList(NpcIndex).IntervaloMovimiento / GetNpcSpeedModifiers(NpcIndex)) Then
        '  Call AddtoRichTextBox(frmMain.RecTxt, "Usar OK.", 255, 0, 0, True, False, False)
        NpcList(NpcIndex).Contadores.IntervaloMovimiento = TActual
        IntervaloPermiteMoverse = True
    Else
        IntervaloPermiteMoverse = False
    End If
    Exit Function
IntervaloPermiteMoverse_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteMoverse", Erl)
End Function

Public Function IntervaloPermiteLanzarHechizo(ByVal NpcIndex As Integer) As Boolean
    On Error GoTo IntervaloPermiteLanzarHechizo_Err
    With NpcList(NpcIndex)
        IntervaloPermiteLanzarHechizo = GetTickCount() - .Contadores.IntervaloLanzarHechizo >= .IntervaloLanzarHechizo
    End With
    Exit Function
IntervaloPermiteLanzarHechizo_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteLanzarHechizo", Erl)
End Function

Public Function IntervaloPermiteAtacarNPC(ByVal NpcIndex As Integer) As Boolean
    On Error GoTo IntervaloPermiteAtacarNPC_Err
    Dim TActual As Long
    TActual = GetTickCount()
    If TActual - NpcList(NpcIndex).Contadores.IntervaloAtaque >= NpcList(NpcIndex).IntervaloAtaque Then
        '  Call AddtoRichTextBox(frmMain.RecTxt, "Usar OK.", 255, 0, 0, True, False, False)
        NpcList(NpcIndex).Contadores.IntervaloAtaque = TActual
        IntervaloPermiteAtacarNPC = True
    Else
        IntervaloPermiteAtacarNPC = False
    End If
    Exit Function
IntervaloPermiteAtacarNPC_Err:
    Call TraceError(Err.Number, Err.Description, "modNuevoTimer.IntervaloPermiteAtacarNPC", Erl)
End Function
