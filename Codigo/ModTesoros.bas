Attribute VB_Name = "ModTesoros"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Option Explicit

Public TesoroNPC() As Integer

Public TesoroNPCMapa() As Integer

Dim TesoroMapa()     As Integer

Dim TesoroRegalo()    As t_Obj

Public BusquedaTesoroActiva As Boolean
Public BusquedaRegaloActiva As Boolean
Public BusquedaNpcActiva As Boolean
Public npc_index_evento As Integer
                           
Public TesoroNumMapa        As Integer

Public TesoroX              As Byte

Public TesoroY              As Byte

Dim RegaloMapa()     As Integer

Dim RegaloRegalo()    As t_Obj


Public RegaloNumMapa        As Integer

Public RegaloX              As Byte

Public RegaloY              As Byte

Public Sub InitTesoro()
        
        On Error GoTo InitTesoro_Err

        Dim Lector As clsIniManager
100     Set Lector = New clsIniManager
        
102     Call Lector.Initialize(DatPath & "Tesoros.dat")
        
        Dim CantidadMapas As Integer
104     CantidadMapas = val(Lector.GetValue("Tesoros", "CantidadMapas"))
        
106     If CantidadMapas <= 0 Then
108         ReDim TesoroMapa(0)
            Exit Sub
        End If
    
110     ReDim TesoroMapa(1 To CantidadMapas)
        
        Dim i As Integer
112     For i = 1 To CantidadMapas
114         TesoroMapa(i) = val(Lector.GetValue("Tesoros", "Mapa" & i))
        Next
        
        Dim TiposDeTesoros As Integer
116     TiposDeTesoros = val(Lector.GetValue("Tesoros", "TiposDeTesoros"))
        
118     If TiposDeTesoros <= 0 Then
120         ReDim TesoroRegalo(0)
            Exit Sub
        End If
    
122     ReDim TesoroRegalo(1 To TiposDeTesoros)
        
        Dim Fields() As String, str As String
124     For i = 1 To TiposDeTesoros
126         str = Lector.GetValue("Tesoros", "Tesoro" & i)
        
128         If LenB(str) Then
130             Fields = Split(str, "-", 2)
                
132             If UBound(Fields) >= 1 Then
134                 With TesoroRegalo(i)
136                     .ObjIndex = val(Fields(0))
138                     .amount = val(Fields(1))
                    End With
                End If
            End If
        Next
        
        Dim NPCs As Integer
        
140     NPCs = val(Lector.GetValue("Criatura", "NPCs"))
        
142     ReDim TesoroNPC(1 To NPCs)
        
144     For i = 1 To NPCs
146         TesoroNPC(i) = val(Lector.GetValue("Criatura", "NPC" & i))
        Next
        
148     CantidadMapas = val(Lector.GetValue("Criatura", "CantidadMapas"))
    
150     ReDim TesoroNPCMapa(1 To CantidadMapas)
        
152     For i = 1 To CantidadMapas
154         TesoroNPCMapa(i) = val(Lector.GetValue("Criatura", "Mapa" & i))
        Next
    
156     Set Lector = Nothing
        
        Exit Sub

InitTesoro_Err:
158     Call TraceError(Err.Number, Err.Description, "ModTesoros.InitTesoro", Erl)

        
End Sub

Public Sub InitRegalo()
        
        On Error GoTo InitRegalo_Err
        
        Dim Lector As clsIniManager
100     Set Lector = New clsIniManager
        
102     Call Lector.Initialize(DatPath & "Tesoros.dat")
        
        Dim CantidadMapas As Integer
104     CantidadMapas = val(Lector.GetValue("Regalos", "CantidadMapas"))
        
106     If CantidadMapas <= 0 Then
108         ReDim RegaloMapa(0)
            Exit Sub
        End If
    
110     ReDim RegaloMapa(1 To CantidadMapas)
        
        Dim i As Integer
112     For i = 1 To CantidadMapas
114         RegaloMapa(i) = val(Lector.GetValue("Regalos", "Mapa" & i))
        Next
        
        Dim TiposDeRegalos As Integer
116     TiposDeRegalos = val(Lector.GetValue("Regalos", "TiposDeRegalos"))
        
118     If TiposDeRegalos <= 0 Then
120         ReDim RegaloRegalo(0)
            Exit Sub
        End If
    
122     ReDim RegaloRegalo(1 To TiposDeRegalos)
        
        Dim Fields() As String, str As String
124     For i = 1 To TiposDeRegalos
126         str = Lector.GetValue("Regalos", "Regalo" & i)
        
128         If LenB(str) Then
130             Fields = Split(str, "-", 2)
                
132             If UBound(Fields) >= 1 Then
134                 With RegaloRegalo(i)
136                     .ObjIndex = val(Fields(0))
138                     .amount = val(Fields(1))
                    End With
                End If
            End If
        Next
    
140     Set Lector = Nothing
        
        Exit Sub

InitRegalo_Err:
142     Call TraceError(Err.Number, Err.Description, "ModTesoros.InitRegalo", Erl)

        
End Sub

Public Sub PerderTesoro()
        
        On Error GoTo PerderTesoro_Err
        

        Dim EncontreLugar As Boolean

100     TesoroNumMapa = TesoroMapa(RandomNumber(1, UBound(TesoroMapa)))
102     TesoroX = RandomNumber(20, 80)
104     TesoroY = RandomNumber(20, 80)
        
        Dim Iterations As Integer
        
106     Iterations = 0
108     Do While Not EncontreLugar
110     Iterations = Iterations + 1
112         If (MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked And e_Block.ALL_SIDES) <> e_Block.ALL_SIDES Then
114             If (MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked And FLAG_AGUA) = 0 Then
116                 EncontreLugar = True
                Else
118                 EncontreLugar = False
120                 TesoroX = RandomNumber(20, 80)
122                 TesoroY = RandomNumber(20, 80)
    
                End If
            Else
124             EncontreLugar = False
126             TesoroX = RandomNumber(20, 80)
128             TesoroY = RandomNumber(20, 80)
            End If
            'si no encuentra en 10000 salgo a la mierda
130         If Iterations >= 20 Then
132              TesoroNumMapa = TesoroMapa(RandomNumber(1, UBound(TesoroMapa)))
            End If
        Loop

        
134         BusquedaTesoroActiva = True
136         Call MakeObj(TesoroRegalo(RandomNumber(1, UBound(TesoroRegalo))), TesoroNumMapa, TesoroX, TesoroY, False)
138         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Rondan rumores que hay un tesoro enterrado en el mapa: " & get_map_name(TesoroNumMapa) & "(" & TesoroNumMapa & ") ¿Quien sera el afortunado que lo encuentre?", e_FontTypeNames.FONTTYPE_TALK))
140         Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(257, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno 257

        
        Exit Sub

PerderTesoro_Err:
142     Call TraceError(Err.Number, Err.Description, "ModTesoros.PerderTesoro", Erl)

        
End Sub

Public Sub PerderRegalo()
        
        On Error GoTo PerderRegalo_Err
        

        Dim EncontreLugar As Boolean

100     RegaloNumMapa = RegaloMapa(RandomNumber(1, UBound(RegaloMapa)))
102     RegaloX = RandomNumber(20, 80)
104     RegaloY = RandomNumber(20, 80)
        
        Dim Iterations As Integer
        
106     Iterations = 0
        If RegaloNumMapa <= 0 Then Exit Sub
108     Do While Not EncontreLugar
110     Iterations = Iterations + 1
112         If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And e_Block.ALL_SIDES) <> e_Block.ALL_SIDES Then
114             If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And FLAG_AGUA) = 0 Then
116                 EncontreLugar = True
                Else
118                 EncontreLugar = False
120                 RegaloX = RandomNumber(20, 80)
122                 RegaloY = RandomNumber(20, 80)
    
                End If
            Else
124                 EncontreLugar = False
126                 RegaloX = RandomNumber(20, 80)
128                 RegaloY = RandomNumber(20, 80)
            End If
130         If Iterations >= 20 Then
132             RegaloNumMapa = RegaloMapa(RandomNumber(1, UBound(RegaloMapa)))
            End If
        Loop
        

134     BusquedaRegaloActiva = True
136     Call MakeObj(RegaloRegalo(RandomNumber(1, UBound(RegaloRegalo))), RegaloNumMapa, RegaloX, RegaloY, False)
138     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> De repente ha surgido un item maravilloso en el mapa: " & get_map_name(RegaloNumMapa) & "(" & RegaloNumMapa & ") ¿Quien sera el valiente que lo encuentre? ¡MUCHO CUIDADO!", e_FontTypeNames.FONTTYPE_TALK))
140     Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(497, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno


        
        Exit Sub

PerderRegalo_Err:
142     Call TraceError(Err.Number, Err.Description, "ModTesoros.PerderRegalo", Erl)

        
End Sub

