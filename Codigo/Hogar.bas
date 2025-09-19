Attribute VB_Name = "Hogar"
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
'
Option Explicit

Public Const NUMCIUDADES    As Byte = 9

Public Ciudades(1 To NUMCIUDADES)         As t_WorldPos

Public Sub goHome(ByVal UserIndex As Integer)
      
        On Error GoTo goHome_Err
        
100     With UserList(UserIndex)

102         If .flags.Muerto = 1 Then

104             If EsGM(UserIndex) Then
106                 .Counters.TimerBarra = 5
                Else
108                 .Counters.TimerBarra = HomeTimer
                End If
110             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, e_GraphicEffects.Runa, .Counters.TimerBarra * 100, False, , .pos.x, .pos.y))
112             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(.Char.charindex, .Counters.TimerBarra, e_AccionBarra.Hogar))
                Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1994, .Counters.TimerBarra, e_FontTypeNames.FONTTYPE_New_Gris)) ' Msg1994=Volverás a tu hogar en ¬1 segundos.
                    
114             .Accion.Particula = e_GraphicEffects.Runa
116             .Accion.AccionPendiente = True
118             .Accion.TipoAccion = e_AccionBarra.Hogar
            
            Else
        
120             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1995, vbNullString, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1995=Debes estar muerto para poder utilizar este comando.

            End If
        
        End With
    
        
        Exit Sub

goHome_Err:
122     Call TraceError(Err.Number, Err.Description, "Hogar.goHome", Erl)

        
End Sub


''
' Maneja el tiempo de arrivo al hogar
'
' @param UserIndex  El index del usuario a ser afectado por el /hogar
'

Public Sub TravelingEffect(ByVal UserIndex As Integer)

        On Error GoTo TravelingEffect_Err
    
        ' Si ya paso el tiempo de penalizacion
100     If IntervaloGoHome(UserIndex) Then
102         Call HomeArrival(UserIndex)
        End If

        
        Exit Sub

TravelingEffect_Err:
104     Call TraceError(Err.Number, Err.Description, "Hogar.TravelingEffect", Erl)

        
End Sub


Public Sub HomeArrival(ByVal UserIndex As Integer)

        'Teleports user to its home.

        On Error GoTo HomeArrival_Err
    
        
    
        Dim tX   As Integer
        Dim tY   As Integer
        Dim tMap As Integer

100     With UserList(UserIndex)

            'Antes de que el pj llegue a la ciudad, lo hacemos dejar de navegar para que no se buguee.
102         If .flags.Navegando = 1 Then
104             .Char.body = iCuerpoMuerto
106             .Char.head = 0
108             .Char.ShieldAnim = NingunEscudo
110             .Char.WeaponAnim = NingunArma
112             .Char.CascoAnim = NingunCasco
            
114             .flags.Navegando = 0
115             .flags.Nadando = 0

                Call TargetUpdateTerrain(.EffectOverTime)

154             .invent.EquippedShipObjIndex = 0
156             .invent.EquippedShipSlot = 0
            
116             Call WriteNavigateToggle(UserIndex, .flags.Navegando)
117             Call WriteNadarToggle(UserIndex, False)

                'Le sacamos el navegando, pero no le mostramos a los demas porque va a ser sumoneado hasta ulla.
            End If
        
118         tX = Ciudades(.Hogar).X
120         tY = Ciudades(.Hogar).y
122         tMap = Ciudades(.Hogar).map
        
124         Call FindLegalPos(UserIndex, tMap, CByte(tX), CByte(tY))
126         Call WarpUserChar(UserIndex, tMap, tX, tY, True)
        
128         Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1996, vbNullString, e_FontTypeNames.FONTTYPE_WARNING)) ' Msg1996=Has regresado a tu ciudad de origen.
        
130         .flags.Traveling = 0
132         .Counters.goHome = 0
        
        End With
    
        
        Exit Sub

HomeArrival_Err:
134     Call TraceError(Err.Number, Err.Description, "Hogar.HomeArrival", Erl)

        
End Sub

Public Function IntervaloGoHome(ByVal UserIndex As Integer, _
                                Optional ByVal TimeInterval As Long, _
                                Optional ByVal Actualizar As Boolean = False) As Boolean
        
        On Error GoTo IntervaloGoHome_Err

        'Add the Timer which determines wether the user can be teleported to its home or not

        Dim TActual As Long
100         TActual = GetTickCount()
    
102     With UserList(UserIndex)

            ' Inicializa el timer
104         If Actualizar Then
        
106             .flags.Traveling = 1
108             .Counters.goHome = TActual + TimeInterval
            
            Else

110             If TActual >= .Counters.goHome Then
112                 IntervaloGoHome = True
                End If

            End If

        End With

        
        Exit Function

IntervaloGoHome_Err:
114     Call TraceError(Err.Number, Err.Description, "Hogar.IntervaloGoHome", Erl)

        
End Function

