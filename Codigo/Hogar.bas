Attribute VB_Name = "Hogar"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Option Explicit

'Cantidad de Ciudades
Public Const NUMCIUDADES    As Byte = 6

Public Ciudades(1 To NUMCIUDADES)         As t_WorldPos

Public Sub goHome(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Budi
        'Last Modification: 01/06/2010
        '01/06/2010: ZaMa - Ahora usa otro tipo de intervalo
        '***************************************************
        
        On Error GoTo goHome_Err
        
100     With UserList(UserIndex)

102         If .flags.Muerto = 1 Then

104             If EsGM(UserIndex) Then
106                 .Counters.TimerBarra = 5
                Else
108                 .Counters.TimerBarra = 210
                End If
110             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, e_ParticulasIndex.Runa, .Counters.TimerBarra * 100, False, , .Pos.X, .Pos.y))
112             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(.Char.charindex, .Counters.TimerBarra, e_AccionBarra.Hogar))
                Call WriteConsoleMsg(UserIndex, "Volverás a tu hogar en " & .Counters.TimerBarra & " segundos.", e_FontTypeNames.FONTTYPE_New_Gris)
                    
114             .Accion.Particula = e_ParticulasIndex.Runa
116             .Accion.AccionPendiente = True
118             .Accion.TipoAccion = e_AccionBarra.Hogar
            
            Else
        
120             Call WriteConsoleMsg(UserIndex, "Debes estar muerto para poder utilizar este comando.", e_FontTypeNames.FONTTYPE_FIGHT)

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
        '******************************************************
        'Author: ZaMa
        'Last Update: 01/06/2010 (ZaMa)
        '******************************************************
        
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
        '**************************************************************
        'Author: ZaMa
        'Last Modify by: ZaMa
        'Last Modify Date: 01/06/2010
        'Teleports user to its home.
        '**************************************************************
        
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
            
116             Call WriteNavigateToggle(UserIndex)

                'Le sacamos el navegando, pero no le mostramos a los demas porque va a ser sumoneado hasta ulla.
            End If
        
118         tX = Ciudades(.Hogar).X
120         tY = Ciudades(.Hogar).y
122         tMap = Ciudades(.Hogar).map
        
124         Call FindLegalPos(UserIndex, tMap, CByte(tX), CByte(tY))
126         Call WarpUserChar(UserIndex, tMap, tX, tY, True)
        
128         Call WriteConsoleMsg(UserIndex, "Has regresado a tu ciudad de origen.", e_FontTypeNames.FONTTYPE_WARNING)
        
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
    
        

        '**************************************************************
        'Author: ZaMa
        'Last Modify by: ZaMa
        'Last Modify Date: 01/06/2010
        '01/06/2010: ZaMa - Add the Timer which determines wether the user can be teleported to its home or not
        '**************************************************************
    
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

