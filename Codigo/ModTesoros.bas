Attribute VB_Name = "ModTesoros"
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
Public TesoroNPC()          As Integer
Public TesoroNPCMapa()      As Integer
Dim TesoroMapa()            As Integer
Dim TesoroRegalo()          As t_Obj
Public BusquedaTesoroActiva As Boolean
Public BusquedaRegaloActiva As Boolean
Public BusquedaNpcActiva    As Boolean
Public npc_index_evento     As Integer
Public TesoroNumMapa        As Integer
Public TesoroX              As Byte
Public TesoroY              As Byte
Dim RegaloMapa()            As Integer
Dim RegaloRegalo()          As t_Obj
Public RegaloNumMapa        As Integer
Public RegaloX              As Byte
Public RegaloY              As Byte

Public Sub InitTesoro()
    On Error GoTo InitTesoro_Err
    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    Call Lector.Initialize(DatPath & "Tesoros.dat")
    Dim CantidadMapas As Integer
    CantidadMapas = val(Lector.GetValue("Tesoros", "CantidadMapas"))
    If CantidadMapas <= 0 Then
        ReDim TesoroMapa(0)
        Exit Sub
    End If
    ReDim TesoroMapa(1 To CantidadMapas)
    Dim i As Integer
    For i = 1 To CantidadMapas
        TesoroMapa(i) = val(Lector.GetValue("Tesoros", "Mapa" & i))
    Next
    Dim TiposDeTesoros As Integer
    TiposDeTesoros = val(Lector.GetValue("Tesoros", "TiposDeTesoros"))
    If TiposDeTesoros <= 0 Then
        ReDim TesoroRegalo(0)
        Exit Sub
    End If
    ReDim TesoroRegalo(1 To TiposDeTesoros)
    Dim Fields() As String, str As String
    For i = 1 To TiposDeTesoros
        str = Lector.GetValue("Tesoros", "Tesoro" & i)
        If LenB(str) Then
            Fields = Split(str, "-", 2)
            If UBound(Fields) >= 1 Then
                With TesoroRegalo(i)
                    .ObjIndex = val(Fields(0))
                    .amount = val(Fields(1))
                End With
            End If
        End If
    Next
    Dim NPCs As Integer
    NPCs = val(Lector.GetValue("Criatura", "NPCs"))
    ReDim TesoroNPC(1 To NPCs)
    For i = 1 To NPCs
        TesoroNPC(i) = val(Lector.GetValue("Criatura", "NPC" & i))
    Next
    CantidadMapas = val(Lector.GetValue("Criatura", "CantidadMapas"))
    ReDim TesoroNPCMapa(1 To CantidadMapas)
    For i = 1 To CantidadMapas
        TesoroNPCMapa(i) = val(Lector.GetValue("Criatura", "Mapa" & i))
    Next
    Set Lector = Nothing
    Exit Sub
InitTesoro_Err:
    Call TraceError(Err.Number, Err.Description, "ModTesoros.InitTesoro", Erl)
End Sub

Public Sub InitRegalo()
    On Error GoTo InitRegalo_Err
    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    Call Lector.Initialize(DatPath & "Tesoros.dat")
    Dim CantidadMapas As Integer
    CantidadMapas = val(Lector.GetValue("Regalos", "CantidadMapas"))
    If CantidadMapas <= 0 Then
        ReDim RegaloMapa(0)
        Exit Sub
    End If
    ReDim RegaloMapa(1 To CantidadMapas)
    Dim i As Integer
    For i = 1 To CantidadMapas
        RegaloMapa(i) = val(Lector.GetValue("Regalos", "Mapa" & i))
    Next
    Dim TiposDeRegalos As Integer
    TiposDeRegalos = val(Lector.GetValue("Regalos", "TiposDeRegalos"))
    If TiposDeRegalos <= 0 Then
        ReDim RegaloRegalo(0)
        Exit Sub
    End If
    ReDim RegaloRegalo(1 To TiposDeRegalos)
    Dim Fields() As String, str As String
    For i = 1 To TiposDeRegalos
        str = Lector.GetValue("Regalos", "Regalo" & i)
        If LenB(str) Then
            Fields = Split(str, "-", 2)
            If UBound(Fields) >= 1 Then
                With RegaloRegalo(i)
                    .ObjIndex = val(Fields(0))
                    .amount = val(Fields(1))
                End With
            End If
        End If
    Next
    Set Lector = Nothing
    Exit Sub
InitRegalo_Err:
    Call TraceError(Err.Number, Err.Description, "ModTesoros.InitRegalo", Erl)
End Sub

Public Sub PerderTesoro()
    On Error GoTo PerderTesoro_Err
    Dim EncontreLugar As Boolean
    TesoroNumMapa = TesoroMapa(RandomNumber(1, UBound(TesoroMapa)))
    TesoroX = RandomNumber(20, 80)
    TesoroY = RandomNumber(20, 80)
    Dim Iterations As Integer
    Iterations = 0
    Do While Not EncontreLugar
        Iterations = Iterations + 1
        If (MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked And e_Block.ALL_SIDES) <> e_Block.ALL_SIDES Then
            If (MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked And FLAG_AGUA) = 0 Then
                EncontreLugar = True
            Else
                EncontreLugar = False
                TesoroX = RandomNumber(20, 80)
                TesoroY = RandomNumber(20, 80)
            End If
        Else
            EncontreLugar = False
            TesoroX = RandomNumber(20, 80)
            TesoroY = RandomNumber(20, 80)
        End If
        'si no encuentra en 10000 salgo a la mierda
        If Iterations >= 20 Then
            TesoroNumMapa = TesoroMapa(RandomNumber(1, UBound(TesoroMapa)))
        End If
    Loop
    BusquedaTesoroActiva = True
    Call MakeObj(TesoroRegalo(RandomNumber(1, UBound(TesoroRegalo))), TesoroNumMapa, TesoroX, TesoroY, False)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1567, get_map_name(TesoroNumMapa) & "¬" & TesoroNumMapa, e_FontTypeNames.FONTTYPE_TALK))  'Msg1699=Eventos> Rondan rumores que hay un tesoro enterrado en el mapa: ¬1(¬2) ¿Quien será el afortunado que lo encuentre?
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(257, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno 257
    Exit Sub
PerderTesoro_Err:
    Call TraceError(Err.Number, Err.Description, "ModTesoros.PerderTesoro", Erl)
End Sub

Public Sub PerderRegalo()
    On Error GoTo PerderRegalo_Err
    Dim EncontreLugar As Boolean
    RegaloNumMapa = RegaloMapa(RandomNumber(1, UBound(RegaloMapa)))
    RegaloX = RandomNumber(20, 80)
    RegaloY = RandomNumber(20, 80)
    Dim Iterations As Integer
    Iterations = 0
    If RegaloNumMapa <= 0 Then Exit Sub
    Do While Not EncontreLugar
        Iterations = Iterations + 1
        If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And e_Block.ALL_SIDES) <> e_Block.ALL_SIDES Then
            If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And FLAG_AGUA) = 0 Then
                EncontreLugar = True
            Else
                EncontreLugar = False
                RegaloX = RandomNumber(20, 80)
                RegaloY = RandomNumber(20, 80)
            End If
        Else
            EncontreLugar = False
            RegaloX = RandomNumber(20, 80)
            RegaloY = RandomNumber(20, 80)
        End If
        If Iterations >= 20 Then
            RegaloNumMapa = RegaloMapa(RandomNumber(1, UBound(RegaloMapa)))
        End If
    Loop
    BusquedaRegaloActiva = True
    Call MakeObj(RegaloRegalo(RandomNumber(1, UBound(RegaloRegalo))), RegaloNumMapa, RegaloX, RegaloY, False)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1568, get_map_name(RegaloNumMapa) & "¬" & RegaloNumMapa, e_FontTypeNames.FONTTYPE_TALK))  'Msg1700=Eventos> De repente ha surgido un item maravilloso en el mapa: ¬1(¬2) ¿Quien será el valiente que lo encuentre? ¡MUCHO CUIDADO!
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(497, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno
    Exit Sub
PerderRegalo_Err:
    Call TraceError(Err.Number, Err.Description, "ModTesoros.PerderRegalo", Erl)
End Sub
