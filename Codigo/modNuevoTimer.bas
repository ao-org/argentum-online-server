Attribute VB_Name = "modNuevoTimer"
'Argentum Online 0.11.6
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

'
' Las siguientes funciones devuelven TRUE o FALSE si el intervalo
' permite hacerlo. Si devuelve TRUE, setean automaticamente el
' timer para que no se pueda hacer la accion hasta el nuevo ciclo.
'

' CASTING DE HECHIZOS
Public Function IntervaloPermiteLanzarSpell(ByVal UserIndex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean
        
        On Error GoTo IntervaloPermiteLanzarSpell_Err
        

        Dim TActual As Long

100     TActual = GetTickCount() And &H7FFFFFFF

102     If TActual - UserList(UserIndex).Counters.TimerLanzarSpell >= UserList(UserIndex).Intervals.magia Then
104         If actualizar Then
106             UserList(UserIndex).Counters.TimerLanzarSpell = TActual
                ' Actualizo spell-attack
108             UserList(UserIndex).Counters.TimerMagiaGolpe = TActual

            End If

110         IntervaloPermiteLanzarSpell = True
        Else
112         IntervaloPermiteLanzarSpell = False

        End If

        
        Exit Function

IntervaloPermiteLanzarSpell_Err:
        Call RegistrarError(Err.Number, Err.description, "modNuevoTimer.IntervaloPermiteLanzarSpell", Erl)
        Resume Next
        
End Function

Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean
        
        On Error GoTo IntervaloPermiteAtacar_Err
        

        Dim TActual As Long

100     TActual = GetTickCount() And &H7FFFFFFF

102     If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= UserList(UserIndex).Intervals.Golpe Then
104         If actualizar Then
106             UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
                ' Actualizo attack-spell
108             UserList(UserIndex).Counters.TimerGolpeMagia = TActual

            End If

110         IntervaloPermiteAtacar = True
        Else
112         IntervaloPermiteAtacar = False

        End If

        
        Exit Function

IntervaloPermiteAtacar_Err:
        Call RegistrarError(Err.Number, Err.description, "modNuevoTimer.IntervaloPermiteAtacar", Erl)
        Resume Next
        
End Function

Public Function IntervaloPermiteTirar(ByVal UserIndex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean
        
        On Error GoTo IntervaloPermiteTirar_Err
        

        Dim TActual As Long
    
100     TActual = GetTickCount() And &H7FFFFFFF
    
102     If TActual - UserList(UserIndex).Counters.TimerTirar >= IntervaloTirar Then
104         If actualizar Then
106             UserList(UserIndex).Counters.TimerTirar = TActual

            End If

108         IntervaloPermiteTirar = True
        Else
110         IntervaloPermiteTirar = False

        End If

        
        Exit Function

IntervaloPermiteTirar_Err:
        Call RegistrarError(Err.Number, Err.description, "modNuevoTimer.IntervaloPermiteTirar", Erl)
        Resume Next
        
End Function

Public Function IntervaloPermiteMagiaGolpe(ByVal UserIndex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean
        
        On Error GoTo IntervaloPermiteMagiaGolpe_Err
        

        Dim TActual As Long

100     TActual = GetTickCount() And &H7FFFFFFF
    
102     If TActual - UserList(UserIndex).Counters.TimerLanzarSpell >= UserList(UserIndex).Intervals.MagiaGolpe Then
104         If actualizar Then
106             UserList(UserIndex).Counters.TimerMagiaGolpe = TActual

            End If

108         IntervaloPermiteMagiaGolpe = True
        Else
110         IntervaloPermiteMagiaGolpe = False

        End If

        
        Exit Function

IntervaloPermiteMagiaGolpe_Err:
        Call RegistrarError(Err.Number, Err.description, "modNuevoTimer.IntervaloPermiteMagiaGolpe", Erl)
        Resume Next
        
End Function

Public Function IntervaloPermiteGolpeMagia(ByVal UserIndex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean
        
        On Error GoTo IntervaloPermiteGolpeMagia_Err
        

        Dim TActual As Long

100     TActual = GetTickCount() And &H7FFFFFFF
    
102     If TActual - UserList(UserIndex).Counters.TimerGolpeMagia >= UserList(UserIndex).Intervals.GolpeMagia Then
104         If actualizar Then
106             UserList(UserIndex).Counters.TimerGolpeMagia = TActual

            End If

108         IntervaloPermiteGolpeMagia = True
        Else
110         IntervaloPermiteGolpeMagia = False

        End If

        
        Exit Function

IntervaloPermiteGolpeMagia_Err:
        Call RegistrarError(Err.Number, Err.description, "modNuevoTimer.IntervaloPermiteGolpeMagia", Erl)
        Resume Next
        
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
        
        On Error GoTo IntervaloPermiteTrabajar_Err
        

        Dim TActual As Long

100     TActual = GetTickCount() And &H7FFFFFFF

102     If TActual - UserList(UserIndex).Counters.TimerPuedeTrabajar >= UserList(UserIndex).Intervals.Trabajar Then
104         If actualizar Then UserList(UserIndex).Counters.TimerPuedeTrabajar = TActual
106         IntervaloPermiteTrabajar = True
        Else
108         IntervaloPermiteTrabajar = False

        End If

        
        Exit Function

IntervaloPermiteTrabajar_Err:
        Call RegistrarError(Err.Number, Err.description, "modNuevoTimer.IntervaloPermiteTrabajar", Erl)
        Resume Next
        
End Function

' USAR OBJETOS
Public Function IntervaloPermiteUsar(ByVal UserIndex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean
        
        On Error GoTo IntervaloPermiteUsar_Err
        

        Dim TActual As Long

100     TActual = GetTickCount() And &H7FFFFFFF

102     If TActual - UserList(UserIndex).Counters.TimerUsar >= UserList(UserIndex).Intervals.Usar Then
104         If actualizar Then UserList(UserIndex).Counters.TimerUsar = TActual
106         IntervaloPermiteUsar = True
        Else
108         IntervaloPermiteUsar = False

        End If

        
        Exit Function

IntervaloPermiteUsar_Err:
        Call RegistrarError(Err.Number, Err.description, "modNuevoTimer.IntervaloPermiteUsar", Erl)
        Resume Next
        
End Function

Public Function IntervaloPermiteUsarArcos(ByVal UserIndex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean
        
        On Error GoTo IntervaloPermiteUsarArcos_Err
        

        Dim TActual As Long
    
100     TActual = GetTickCount() And &H7FFFFFFF
    
102     If TActual - UserList(UserIndex).Counters.TimerPuedeUsarArco >= UserList(UserIndex).Intervals.Arco Then
104         If actualizar Then
106             UserList(UserIndex).Counters.TimerPuedeUsarArco = TActual
                ' Tambien actualizo los otros
108             UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
110             UserList(UserIndex).Counters.TimerLanzarSpell = TActual

            End If

112         IntervaloPermiteUsarArcos = True
        Else
114         IntervaloPermiteUsarArcos = False

        End If

        
        Exit Function

IntervaloPermiteUsarArcos_Err:
        Call RegistrarError(Err.Number, Err.description, "modNuevoTimer.IntervaloPermiteUsarArcos", Erl)
        Resume Next
        
End Function

Public Function IntervaloPermiteCaminar(ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo IntervaloPermiteCaminar_Err
        

        Dim TActual As Long
    
100     TActual = GetTickCount() And &H7FFFFFFF
    
102     If TActual - UserList(UserIndex).Counters.TimerCaminar >= UserList(UserIndex).Intervals.Caminar Then
        
            '  Call AddtoRichTextBox(frmMain.RecTxt, "Usar OK.", 255, 0, 0, True, False, False)
104         UserList(UserIndex).Counters.TimerCaminar = TActual
106         IntervaloPermiteCaminar = True
        Else
108         IntervaloPermiteCaminar = False

        End If

        
        Exit Function

IntervaloPermiteCaminar_Err:
        Call RegistrarError(Err.Number, Err.description, "modNuevoTimer.IntervaloPermiteCaminar", Erl)
        Resume Next
        
End Function

Public Function IntervaloPermiteMoverse(ByVal NpcIndex As Integer) As Boolean
        
        On Error GoTo IntervaloPermiteMoverse_Err
        

        Dim TActual As Long

100     TActual = GetTickCount() And &H7FFFFFFF

102     If TActual - Npclist(NpcIndex).Contadores.IntervaloMovimiento >= Npclist(NpcIndex).IntervaloMovimiento Then
    
            '  Call AddtoRichTextBox(frmMain.RecTxt, "Usar OK.", 255, 0, 0, True, False, False)
104         Npclist(NpcIndex).Contadores.IntervaloMovimiento = TActual
106         IntervaloPermiteMoverse = True
        Else
108         IntervaloPermiteMoverse = False

        End If

        
        Exit Function

IntervaloPermiteMoverse_Err:
        Call RegistrarError(Err.Number, Err.description, "modNuevoTimer.IntervaloPermiteMoverse", Erl)
        Resume Next
        
End Function

Public Function IntervaloPermiteLanzarHechizo(ByVal NpcIndex As Integer) As Boolean
        
        On Error GoTo IntervaloPermiteLanzarHechizo_Err
        

        Dim TActual As Long

100     TActual = GetTickCount() And &H7FFFFFFF

102     If TActual - Npclist(NpcIndex).Contadores.InvervaloLanzarHechizo >= Npclist(NpcIndex).InvervaloLanzarHechizo Then
    
            '  Call AddtoRichTextBox(frmMain.RecTxt, "Usar OK.", 255, 0, 0, True, False, False)
104         Npclist(NpcIndex).Contadores.InvervaloLanzarHechizo = TActual
106         IntervaloPermiteLanzarHechizo = True
        Else
108         IntervaloPermiteLanzarHechizo = False

        End If

        
        Exit Function

IntervaloPermiteLanzarHechizo_Err:
        Call RegistrarError(Err.Number, Err.description, "modNuevoTimer.IntervaloPermiteLanzarHechizo", Erl)
        Resume Next
        
End Function

Public Function IntervaloPermiteAtacarNPC(ByVal NpcIndex As Integer) As Boolean
        
        On Error GoTo IntervaloPermiteAtacarNPC_Err
        

        Dim TActual As Long

100     TActual = GetTickCount() And &H7FFFFFFF

102     If TActual - Npclist(NpcIndex).Contadores.IntervaloAtaque >= Npclist(NpcIndex).IntervaloAtaque Then
    
            '  Call AddtoRichTextBox(frmMain.RecTxt, "Usar OK.", 255, 0, 0, True, False, False)
104         Npclist(NpcIndex).Contadores.IntervaloAtaque = TActual
106         IntervaloPermiteAtacarNPC = True
        Else
108         IntervaloPermiteAtacarNPC = False

        End If

        
        Exit Function

IntervaloPermiteAtacarNPC_Err:
        Call RegistrarError(Err.Number, Err.description, "modNuevoTimer.IntervaloPermiteAtacarNPC", Erl)
        Resume Next
        
End Function

