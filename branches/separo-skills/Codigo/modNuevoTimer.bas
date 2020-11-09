Attribute VB_Name = "modNuevoTimer"
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

'
' Las siguientes funciones devuelven TRUE o FALSE si el intervalo
' permite hacerlo. Si devuelve TRUE, setean automaticamente el
' timer para que no se pueda hacer la accion hasta el nuevo ciclo.
'

' CASTING DE HECHIZOS
Public Function IntervaloPermiteLanzarSpell(ByVal UserIndex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean

    Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF

    If TActual - UserList(UserIndex).Counters.TimerLanzarSpell >= UserList(UserIndex).Intervals.magia Then
        If actualizar Then
            UserList(UserIndex).Counters.TimerLanzarSpell = TActual
            ' Actualizo spell-attack
            UserList(UserIndex).Counters.TimerMagiaGolpe = TActual

        End If

        IntervaloPermiteLanzarSpell = True
    Else
        IntervaloPermiteLanzarSpell = False

    End If

End Function

Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean

    Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF

    If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= UserList(UserIndex).Intervals.Golpe Then
        If actualizar Then
            UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
            ' Actualizo attack-spell
            UserList(UserIndex).Counters.TimerGolpeMagia = TActual

        End If

        IntervaloPermiteAtacar = True
    Else
        IntervaloPermiteAtacar = False

    End If

End Function

Public Function IntervaloPermiteTirar(ByVal UserIndex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean

    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerTirar >= IntervaloTirar Then
        If actualizar Then
            UserList(UserIndex).Counters.TimerTirar = TActual

        End If

        IntervaloPermiteTirar = True
    Else
        IntervaloPermiteTirar = False

    End If

End Function

Public Function IntervaloPermiteMagiaGolpe(ByVal UserIndex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean

    Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerLanzarSpell >= UserList(UserIndex).Intervals.MagiaGolpe Then
        If actualizar Then
            UserList(UserIndex).Counters.TimerMagiaGolpe = TActual

        End If

        IntervaloPermiteMagiaGolpe = True
    Else
        IntervaloPermiteMagiaGolpe = False

    End If

End Function

Public Function IntervaloPermiteGolpeMagia(ByVal UserIndex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean

    Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerGolpeMagia >= UserList(UserIndex).Intervals.GolpeMagia Then
        If actualizar Then
            UserList(UserIndex).Counters.TimerGolpeMagia = TActual

        End If

        IntervaloPermiteGolpeMagia = True
    Else
        IntervaloPermiteGolpeMagia = False

    End If

End Function

' ATAQUE CUERPO A CUERPO
'Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'Dim TActual As Long
'
'TActual = GetTickCount() And &H7FFFFFFF''
'
'If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= IntervaloUserPuedeAtacar Then
'    If Actualizar Then UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
'    IntervaloPermiteAtacar = True
'Else
'    IntervaloPermiteAtacar = False
'End If
'End Function

' TRABAJO
Public Function IntervaloPermiteTrabajar(ByVal UserIndex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean

    Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF

    If TActual - UserList(UserIndex).Counters.TimerPuedeTrabajar >= UserList(UserIndex).Intervals.Trabajar Then
        If actualizar Then UserList(UserIndex).Counters.TimerPuedeTrabajar = TActual
        IntervaloPermiteTrabajar = True
    Else
        IntervaloPermiteTrabajar = False

    End If

End Function

' USAR OBJETOS
Public Function IntervaloPermiteUsar(ByVal UserIndex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean

    Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF

    If TActual - UserList(UserIndex).Counters.TimerUsar >= UserList(UserIndex).Intervals.Usar Then
        If actualizar Then UserList(UserIndex).Counters.TimerUsar = TActual
        IntervaloPermiteUsar = True
    Else
        IntervaloPermiteUsar = False

    End If

End Function

Public Function IntervaloPermiteUsarArcos(ByVal UserIndex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean

    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerPuedeUsarArco >= UserList(UserIndex).Intervals.Arco Then
        If actualizar Then
            UserList(UserIndex).Counters.TimerPuedeUsarArco = TActual
            ' Tambien actualizo los otros
            UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
            UserList(UserIndex).Counters.TimerLanzarSpell = TActual

        End If

        IntervaloPermiteUsarArcos = True
    Else
        IntervaloPermiteUsarArcos = False

    End If

End Function

Public Function IntervaloPermiteCaminar(ByVal UserIndex As Integer) As Boolean

    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerCaminar >= UserList(UserIndex).Intervals.Caminar Then
        
        '  Call AddtoRichTextBox(frmMain.RecTxt, "Usar OK.", 255, 0, 0, True, False, False)
        UserList(UserIndex).Counters.TimerCaminar = TActual
        IntervaloPermiteCaminar = True
    Else
        IntervaloPermiteCaminar = False

    End If

End Function

Public Function IntervaloPermiteMoverse(ByVal NpcIndex As Integer) As Boolean

    Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF

    If TActual - Npclist(NpcIndex).Contadores.IntervaloMovimiento >= Npclist(NpcIndex).IntervaloMovimiento Then
    
        '  Call AddtoRichTextBox(frmMain.RecTxt, "Usar OK.", 255, 0, 0, True, False, False)
        Npclist(NpcIndex).Contadores.IntervaloMovimiento = TActual
        IntervaloPermiteMoverse = True
    Else
        IntervaloPermiteMoverse = False

    End If

End Function

Public Function IntervaloPermiteLanzarHechizo(ByVal NpcIndex As Integer) As Boolean

    Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF

    If TActual - Npclist(NpcIndex).Contadores.InvervaloLanzarHechizo >= Npclist(NpcIndex).InvervaloLanzarHechizo Then
    
        '  Call AddtoRichTextBox(frmMain.RecTxt, "Usar OK.", 255, 0, 0, True, False, False)
        Npclist(NpcIndex).Contadores.InvervaloLanzarHechizo = TActual
        IntervaloPermiteLanzarHechizo = True
    Else
        IntervaloPermiteLanzarHechizo = False

    End If

End Function

Public Function IntervaloPermiteAtacarNPC(ByVal NpcIndex As Integer) As Boolean

    Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF

    If TActual - Npclist(NpcIndex).Contadores.IntervaloAtaque >= Npclist(NpcIndex).IntervaloAtaque Then
    
        '  Call AddtoRichTextBox(frmMain.RecTxt, "Usar OK.", 255, 0, 0, True, False, False)
        Npclist(NpcIndex).Contadores.IntervaloAtaque = TActual
        IntervaloPermiteAtacarNPC = True
    Else
        IntervaloPermiteAtacarNPC = False

    End If

End Function

