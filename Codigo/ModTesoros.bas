Attribute VB_Name = "ModTesoros"
Dim TesoroMapa(1 To 20) As Integer
Dim TesoroRegalo(1 To 5) As obj
Public BusquedaTesoroActiva As Boolean
Public TesoroNumMapa As Integer
Public TesoroX As Byte
Public TesoroY As Byte


Dim RegaloMapa(1 To 19) As Integer
Dim RegaloRegalo(1 To 6) As obj
Public BusquedaRegaloActiva As Boolean
Public RegaloNumMapa As Integer
Public RegaloX As Byte
Public RegaloY As Byte

Public Sub InitTesoro()
    TesoroMapa(1) = 253
    TesoroMapa(2) = 254
    TesoroMapa(3) = 265
    TesoroMapa(4) = 266
    TesoroMapa(5) = 267
    TesoroMapa(6) = 268
    TesoroMapa(7) = 250
    TesoroMapa(8) = 37
    TesoroMapa(9) = 85
    TesoroMapa(10) = 73
    TesoroMapa(11) = 42
    TesoroMapa(12) = 21
    TesoroMapa(13) = 87
    TesoroMapa(14) = 27
    TesoroMapa(15) = 28
    TesoroMapa(16) = 63
    TesoroMapa(17) = 47
    TesoroMapa(18) = 48
    TesoroMapa(19) = 252
    TesoroMapa(20) = 249
    
    TesoroRegalo(1).ObjIndex = 200
    TesoroRegalo(1).Amount = 1
    
    TesoroRegalo(2).ObjIndex = 201
    TesoroRegalo(2).Amount = 1
    
    TesoroRegalo(3).ObjIndex = 202
    TesoroRegalo(3).Amount = 1
    
    TesoroRegalo(4).ObjIndex = 203
    TesoroRegalo(4).Amount = 1
    
    TesoroRegalo(5).ObjIndex = 204
    TesoroRegalo(5).Amount = 1
    
    
End Sub
Public Sub InitRegalo()
    RegaloMapa(1) = 297
    RegaloMapa(2) = 295
    RegaloMapa(3) = 296
    RegaloMapa(4) = 276
    RegaloMapa(5) = 142
    RegaloMapa(6) = 317
    RegaloMapa(7) = 303
    RegaloMapa(8) = 302
    RegaloMapa(9) = 293
    RegaloMapa(10) = 290
    RegaloMapa(11) = 289
    RegaloMapa(12) = 294
    RegaloMapa(13) = 292
    RegaloMapa(14) = 286
    RegaloMapa(15) = 278
    RegaloMapa(16) = 277
    RegaloMapa(17) = 301
    RegaloMapa(18) = 287
    RegaloMapa(19) = 316
    
    RegaloRegalo(1).ObjIndex = 1081 'Pendiente del Sacrificio
    RegaloRegalo(1).Amount = 1
    
    RegaloRegalo(2).ObjIndex = 707 'Brazalete del Ogro (+30)
    RegaloRegalo(2).Amount = 1
    
    RegaloRegalo(3).ObjIndex = 1143 'Sortija de la Verdad
    RegaloRegalo(3).Amount = 1
    
    RegaloRegalo(4).ObjIndex = 1006 ' Anillo de las Sombras
    RegaloRegalo(4).Amount = 1
    
    RegaloRegalo(5).ObjIndex = 651 'Orbe de Inhibición
    RegaloRegalo(5).Amount = 1
    
    'TesoroRegalo(6).ObjIndex = 1181 'Báculo de Hechicero (DM +10)
   'TesoroRegalo(6).Amount = 1
    
End Sub


Public Sub PerderTesoro()




Dim EncontreLugar As Boolean
TesoroNumMapa = TesoroMapa(RandomNumber(1, 20))
TesoroX = RandomNumber(20, 80)
TesoroY = RandomNumber(20, 80)


If MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked = 0 Then
    If HayAgua(TesoroNumMapa, TesoroX, TesoroY) = False Then
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


If EncontreLugar = False Then
    If MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked = 0 Then
        If HayAgua(TesoroNumMapa, TesoroX, TesoroY) = False Then
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
End If


If EncontreLugar = False Then
    If MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked = 0 Then
        If HayAgua(TesoroNumMapa, TesoroX, TesoroY) = False Then
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
End If
        
        
        
If EncontreLugar = False Then
    If MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked = 0 Then
        If HayAgua(TesoroNumMapa, TesoroX, TesoroY) = False Then
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
End If

If EncontreLugar = False Then
    If MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked = 0 Then
        If HayAgua(TesoroNumMapa, TesoroX, TesoroY) = False Then
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
End If

If EncontreLugar = False Then
    If MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked = 0 Then
        If HayAgua(TesoroNumMapa, TesoroX, TesoroY) = False Then
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
End If
        

        
If EncontreLugar = True Then
    BusquedaTesoroActiva = True
    Call MakeObj(TesoroRegalo(RandomNumber(1, 5)), TesoroNumMapa, TesoroX, TesoroY, False)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Rondan rumores que hay un tesoro enterrado en el mapa: " & DarNameMapa(TesoroNumMapa) & "(" & TesoroNumMapa & ") ¿Quien sera el afortunado que lo encuentre?", FontTypeNames.FONTTYPE_TALK))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(257, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno 257
End If
End Sub



Public Sub PerderRegalo()




Dim EncontreLugar As Boolean
RegaloNumMapa = RegaloMapa(RandomNumber(1, 18))
RegaloX = RandomNumber(20, 80)
RegaloY = RandomNumber(20, 80)


If MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked = 0 Then
    If HayAgua(TesoroNumMapa, RegaloX, RegaloY) = False Then
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


If EncontreLugar = False Then
    If MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked = 0 Then
        If HayAgua(RegaloNumMapa, RegaloX, RegaloY) = False Then
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
End If
If EncontreLugar = False Then
    If MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked = 0 Then
        If HayAgua(RegaloNumMapa, RegaloX, RegaloY) = False Then
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
End If
If EncontreLugar = False Then
    If MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked = 0 Then
        If HayAgua(RegaloNumMapa, RegaloX, RegaloY) = False Then
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
End If
If EncontreLugar = False Then
    If MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked = 0 Then
        If HayAgua(RegaloNumMapa, RegaloX, RegaloY) = False Then
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
End If
If EncontreLugar = False Then
    If MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked = 0 Then
        If HayAgua(RegaloNumMapa, RegaloX, RegaloY) = False Then
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
End If
If EncontreLugar = False Then
    If MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked = 0 Then
        If HayAgua(RegaloNumMapa, RegaloX, RegaloY) = False Then
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
End If


        
If EncontreLugar = True Then
    BusquedaRegaloActiva = True
    Call MakeObj(RegaloRegalo(RandomNumber(1, 5)), RegaloNumMapa, RegaloX, RegaloY, False)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> De repente ha surgido un item maravilloso en el mapa: " & DarNameMapa(RegaloNumMapa) & "(" & RegaloNumMapa & ") ¿Quien sera el valiente que lo encuentre? ¡MUCHO CUIDADO!", FontTypeNames.FONTTYPE_TALK))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(497, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno
End If
End Sub





