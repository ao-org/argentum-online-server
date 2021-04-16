Attribute VB_Name = "ModTesoros"
Option Explicit

Public TesoroNPC() As Integer

Public TesoroNPCMapa() As Integer

Dim TesoroMapa()     As Integer

Dim TesoroRegalo()    As obj

Public BusquedaTesoroActiva As Boolean
Public BusquedaRegaloActiva As Boolean
Public BusquedaNpcActiva As Boolean
Public npc_index_evento As Integer
                           
Public TesoroNumMapa        As Integer

Public TesoroX              As Byte

Public TesoroY              As Byte

Dim RegaloMapa()     As Integer

Dim RegaloRegalo()    As obj


Public RegaloNumMapa        As Integer

Public RegaloX              As Byte

Public RegaloY              As Byte

Public Sub InitTesoro()
        
        On Error GoTo InitTesoro_Err

        Dim Lector As clsIniReader
        Set Lector = New clsIniReader
        
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
                        .Amount = val(Fields(1))
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
160     Call RegistrarError(Err.Number, Err.Description, "ModTesoros.InitTesoro", Erl)
162     Resume Next
        
End Sub

Public Sub InitRegalo()
        
        On Error GoTo InitRegalo_Err
        
        Dim Lector As clsIniReader
        Set Lector = New clsIniReader
        
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
                        .Amount = val(Fields(1))
                    End With
                End If
            End If
        Next
    
        Set Lector = Nothing
        
        Exit Sub

InitRegalo_Err:
158     Call RegistrarError(Err.Number, Err.Description, "ModTesoros.InitRegalo", Erl)
160     Resume Next
        
End Sub

Public Sub PerderTesoro()
        
        On Error GoTo PerderTesoro_Err
        

        Dim EncontreLugar As Boolean

100     TesoroNumMapa = TesoroMapa(RandomNumber(1, UBound(TesoroMapa)))
102     TesoroX = RandomNumber(20, 80)
104     TesoroY = RandomNumber(20, 80)
        
        Dim Iterations As Integer
        
        Iterations = 0
        Do While Not EncontreLugar
        Iterations = Iterations + 1
106         If (MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Then
108             If (MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked And FLAG_AGUA) = 0 Then
110                 EncontreLugar = True
                Else
112                 EncontreLugar = False
114                 TesoroX = RandomNumber(20, 80)
116                 TesoroY = RandomNumber(20, 80)
    
                End If
            Else
118             EncontreLugar = False
120             TesoroX = RandomNumber(20, 80)
122             TesoroY = RandomNumber(20, 80)
            End If
            'si no encuentra en 10000 salgo a la mierda
            If Iterations >= 20 Then
                 TesoroNumMapa = TesoroMapa(RandomNumber(1, UBound(TesoroMapa)))
            End If
        Loop

        
226         BusquedaTesoroActiva = True
228         Call MakeObj(TesoroRegalo(RandomNumber(1, UBound(TesoroRegalo))), TesoroNumMapa, TesoroX, TesoroY, False)
230         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Rondan rumores que hay un tesoro enterrado en el mapa: " & DarNameMapa(TesoroNumMapa) & "(" & TesoroNumMapa & ") ¿Quien sera el afortunado que lo encuentre?", FontTypeNames.FONTTYPE_TALK))
232         Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(257, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno 257

        
        Exit Sub

PerderTesoro_Err:
234     Call RegistrarError(Err.Number, Err.Description, "ModTesoros.PerderTesoro", Erl)
236     Resume Next
        
End Sub

Public Sub PerderRegalo()
        
        On Error GoTo PerderRegalo_Err
        

        Dim EncontreLugar As Boolean

100     RegaloNumMapa = RegaloMapa(RandomNumber(1, UBound(RegaloMapa)))
102     RegaloX = RandomNumber(20, 80)
104     RegaloY = RandomNumber(20, 80)
        
        Dim Iterations As Integer
        
        Iterations = 0
        
        Do While Not EncontreLugar
        Iterations = Iterations + 1
106         If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Then
108             If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And FLAG_AGUA) = 0 Then
110                 EncontreLugar = True
                Else
112                 EncontreLugar = False
114                 RegaloX = RandomNumber(20, 80)
116                 RegaloY = RandomNumber(20, 80)
    
                End If
            Else
118                 EncontreLugar = False
120                 RegaloX = RandomNumber(20, 80)
122                 RegaloY = RandomNumber(20, 80)
            End If
            If Iterations >= 20 Then
                RegaloNumMapa = RegaloMapa(RandomNumber(1, UBound(RegaloMapa)))
            End If
        Loop
        

246     BusquedaRegaloActiva = True
248     Call MakeObj(RegaloRegalo(RandomNumber(1, UBound(RegaloRegalo))), RegaloNumMapa, RegaloX, RegaloY, False)
250     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> De repente ha surgido un item maravilloso en el mapa: " & DarNameMapa(RegaloNumMapa) & "(" & RegaloNumMapa & ") ¿Quien sera el valiente que lo encuentre? ¡MUCHO CUIDADO!", FontTypeNames.FONTTYPE_TALK))
252     Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(497, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno


        
        Exit Sub

PerderRegalo_Err:
254     Call RegistrarError(Err.Number, Err.Description, "ModTesoros.PerderRegalo", Erl)
256     Resume Next
        
End Sub

