Attribute VB_Name = "ModTesoros"

Dim TesoroMapa(1 To 20)     As Integer

Dim TesoroRegalo(1 To 5)    As obj

Public BusquedaTesoroActiva As Boolean

Public TesoroNumMapa        As Integer

Public TesoroX              As Byte

Public TesoroY              As Byte

Dim RegaloMapa(1 To 19)     As Integer

Dim RegaloRegalo(1 To 6)    As obj

Public BusquedaRegaloActiva As Boolean

Public RegaloNumMapa        As Integer

Public RegaloX              As Byte

Public RegaloY              As Byte

Public Sub InitTesoro()
        
        On Error GoTo InitTesoro_Err
        
100     TesoroMapa(1) = 253
102     TesoroMapa(2) = 254
104     TesoroMapa(3) = 265
106     TesoroMapa(4) = 266
108     TesoroMapa(5) = 267
110     TesoroMapa(6) = 268
112     TesoroMapa(7) = 250
114     TesoroMapa(8) = 37
116     TesoroMapa(9) = 85
118     TesoroMapa(10) = 73
120     TesoroMapa(11) = 42
122     TesoroMapa(12) = 21
124     TesoroMapa(13) = 87
126     TesoroMapa(14) = 27
128     TesoroMapa(15) = 28
130     TesoroMapa(16) = 63
132     TesoroMapa(17) = 47
134     TesoroMapa(18) = 48
136     TesoroMapa(19) = 252
138     TesoroMapa(20) = 249
    
140     TesoroRegalo(1).ObjIndex = 200
142     TesoroRegalo(1).Amount = 1
    
144     TesoroRegalo(2).ObjIndex = 201
146     TesoroRegalo(2).Amount = 1
    
148     TesoroRegalo(3).ObjIndex = 202
150     TesoroRegalo(3).Amount = 1
    
152     TesoroRegalo(4).ObjIndex = 203
154     TesoroRegalo(4).Amount = 1
    
156     TesoroRegalo(5).ObjIndex = 204
158     TesoroRegalo(5).Amount = 1
    
        
        Exit Sub

InitTesoro_Err:
        Call RegistrarError(Err.Number, Err.description, "ModTesoros.InitTesoro", Erl)
        Resume Next
        
End Sub

Public Sub InitRegalo()
        
        On Error GoTo InitRegalo_Err
        
100     RegaloMapa(1) = 297
102     RegaloMapa(2) = 295
104     RegaloMapa(3) = 296
106     RegaloMapa(4) = 276
108     RegaloMapa(5) = 142
110     RegaloMapa(6) = 317
112     RegaloMapa(7) = 303
114     RegaloMapa(8) = 302
116     RegaloMapa(9) = 293
118     RegaloMapa(10) = 290
120     RegaloMapa(11) = 289
122     RegaloMapa(12) = 294
124     RegaloMapa(13) = 292
126     RegaloMapa(14) = 286
128     RegaloMapa(15) = 278
130     RegaloMapa(16) = 277
132     RegaloMapa(17) = 301
134     RegaloMapa(18) = 287
136     RegaloMapa(19) = 316
    
138     RegaloRegalo(1).ObjIndex = 1081 'Pendiente del Sacrificio
140     RegaloRegalo(1).Amount = 1
    
142     RegaloRegalo(2).ObjIndex = 707 'Brazalete del Ogro (+30)
144     RegaloRegalo(2).Amount = 1
    
146     RegaloRegalo(3).ObjIndex = 1143 'Sortija de la Verdad
148     RegaloRegalo(3).Amount = 1
    
150     RegaloRegalo(4).ObjIndex = 1006 ' Anillo de las Sombras
152     RegaloRegalo(4).Amount = 1
    
154     RegaloRegalo(5).ObjIndex = 651 'Orbe de Inhibición
156     RegaloRegalo(5).Amount = 1
    
        'TesoroRegalo(6).ObjIndex = 1181 'Báculo de Hechicero (DM +10)
        'TesoroRegalo(6).Amount = 1
    
        
        Exit Sub

InitRegalo_Err:
        Call RegistrarError(Err.Number, Err.description, "ModTesoros.InitRegalo", Erl)
        Resume Next
        
End Sub

Public Sub PerderTesoro()
        
        On Error GoTo PerderTesoro_Err
        

        Dim EncontreLugar As Boolean

100     TesoroNumMapa = TesoroMapa(RandomNumber(1, 20))
102     TesoroX = RandomNumber(20, 80)
104     TesoroY = RandomNumber(20, 80)

106     If (MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Then
108         If (MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked And FLAG_AGUA) = 0 Then
110             EncontreLugar = True
            Else
112             EncontreLugar = False
114             TesoroX = RandomNumber(20, 80)
116             TesoroY = RandomNumber(20, 80)

            End If

        Else
118         EncontreLugar = False
120         TesoroX = RandomNumber(20, 80)
122         TesoroY = RandomNumber(20, 80)

        End If

124     If EncontreLugar = False Then
126         If (MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Then
128             If (MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked And FLAG_AGUA) = 0 Then
130                 EncontreLugar = True
                Else
132                 EncontreLugar = False
134                 TesoroX = RandomNumber(20, 80)
136                 TesoroY = RandomNumber(20, 80)

                End If

            Else
138             EncontreLugar = False
140             TesoroX = RandomNumber(20, 80)
142             TesoroY = RandomNumber(20, 80)

            End If

        End If

144     If EncontreLugar = False Then
146         If (MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Then
148             If (MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked And FLAG_AGUA) = 0 Then
150                 EncontreLugar = True
                Else
152                 EncontreLugar = False
154                 TesoroX = RandomNumber(20, 80)
156                 TesoroY = RandomNumber(20, 80)

                End If

            Else
158             EncontreLugar = False
160             TesoroX = RandomNumber(20, 80)
162             TesoroY = RandomNumber(20, 80)

            End If

        End If
        
164     If EncontreLugar = False Then
166         If (MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Then
168             If (MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked And FLAG_AGUA) = 0 Then
170                 EncontreLugar = True
                Else
172                 EncontreLugar = False
174                 TesoroX = RandomNumber(20, 80)
176                 TesoroY = RandomNumber(20, 80)

                End If

            Else
178             EncontreLugar = False
180             TesoroX = RandomNumber(20, 80)
182             TesoroY = RandomNumber(20, 80)

            End If

        End If

184     If EncontreLugar = False Then
186         If (MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Then
188             If (MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked And FLAG_AGUA) = 0 Then
190                 EncontreLugar = True
                Else
192                 EncontreLugar = False
194                 TesoroX = RandomNumber(20, 80)
196                 TesoroY = RandomNumber(20, 80)

                End If

            Else
198             EncontreLugar = False
200             TesoroX = RandomNumber(20, 80)
202             TesoroY = RandomNumber(20, 80)

            End If

        End If

204     If EncontreLugar = False Then
206         If (MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Then
208             If (MapData(TesoroNumMapa, TesoroX, TesoroY).Blocked And FLAG_AGUA) = 0 Then
210                 EncontreLugar = True
                Else
212                 EncontreLugar = False
214                 TesoroX = RandomNumber(20, 80)
216                 TesoroY = RandomNumber(20, 80)

                End If

            Else
218             EncontreLugar = False
220             TesoroX = RandomNumber(20, 80)
222             TesoroY = RandomNumber(20, 80)

            End If

        End If
        
224     If EncontreLugar = True Then
226         BusquedaTesoroActiva = True
228         Call MakeObj(TesoroRegalo(RandomNumber(1, 5)), TesoroNumMapa, TesoroX, TesoroY, False)
230         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Rondan rumores que hay un tesoro enterrado en el mapa: " & DarNameMapa(TesoroNumMapa) & "(" & TesoroNumMapa & ") ¿Quien sera el afortunado que lo encuentre?", FontTypeNames.FONTTYPE_TALK))
232         Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(257, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno 257

        End If

        
        Exit Sub

PerderTesoro_Err:
        Call RegistrarError(Err.Number, Err.description, "ModTesoros.PerderTesoro", Erl)
        Resume Next
        
End Sub

Public Sub PerderRegalo()
        
        On Error GoTo PerderRegalo_Err
        

        Dim EncontreLugar As Boolean

100     RegaloNumMapa = RegaloMapa(RandomNumber(1, 18))
102     RegaloX = RandomNumber(20, 80)
104     RegaloY = RandomNumber(20, 80)

106     If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Then
108         If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And FLAG_AGUA) = 0 Then
110             EncontreLugar = True
            Else
112             EncontreLugar = False
114             RegaloX = RandomNumber(20, 80)
116             RegaloY = RandomNumber(20, 80)

            End If

        Else
118         EncontreLugar = False
120         RegaloX = RandomNumber(20, 80)
122         RegaloY = RandomNumber(20, 80)

        End If

124     If EncontreLugar = False Then
126         If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Then
128             If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And FLAG_AGUA) = 0 Then
130                 EncontreLugar = True
                Else
132                 EncontreLugar = False
134                 RegaloX = RandomNumber(20, 80)
136                 RegaloY = RandomNumber(20, 80)

                End If

            Else
138             EncontreLugar = False
140             RegaloX = RandomNumber(20, 80)
142             RegaloY = RandomNumber(20, 80)

            End If

        End If

144     If EncontreLugar = False Then
146         If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Then
148             If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And FLAG_AGUA) = 0 Then
150                 EncontreLugar = True
                Else
152                 EncontreLugar = False
154                 RegaloX = RandomNumber(20, 80)
156                 RegaloY = RandomNumber(20, 80)

                End If

            Else
158             EncontreLugar = False
160             RegaloX = RandomNumber(20, 80)
162             RegaloY = RandomNumber(20, 80)

            End If

        End If

164     If EncontreLugar = False Then
166         If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Then
168             If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And FLAG_AGUA) = 0 Then
170                 EncontreLugar = True
                Else
172                 EncontreLugar = False
174                 RegaloX = RandomNumber(20, 80)
176                 RegaloY = RandomNumber(20, 80)

                End If

            Else
178             EncontreLugar = False
180             RegaloX = RandomNumber(20, 80)
182             RegaloY = RandomNumber(20, 80)

            End If

        End If

184     If EncontreLugar = False Then
186         If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Then
188             If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And FLAG_AGUA) = 0 Then
190                 EncontreLugar = True
                Else
192                 EncontreLugar = False
194                 RegaloX = RandomNumber(20, 80)
196                 RegaloY = RandomNumber(20, 80)

                End If

            Else
198             EncontreLugar = False
200             RegaloX = RandomNumber(20, 80)
202             RegaloY = RandomNumber(20, 80)

            End If

        End If

204     If EncontreLugar = False Then
206         If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Then
208             If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And FLAG_AGUA) = 0 Then
210                 EncontreLugar = True
                Else
212                 EncontreLugar = False
214                 RegaloX = RandomNumber(20, 80)
216                 RegaloY = RandomNumber(20, 80)

                End If

            Else
218             EncontreLugar = False
220             RegaloX = RandomNumber(20, 80)
222             RegaloY = RandomNumber(20, 80)

            End If

        End If

224     If EncontreLugar = False Then
226         If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Then
228             If (MapData(RegaloNumMapa, RegaloX, RegaloY).Blocked And FLAG_AGUA) = 0 Then
230                 EncontreLugar = True
                Else
232                 EncontreLugar = False
234                 RegaloX = RandomNumber(20, 80)
236                 RegaloY = RandomNumber(20, 80)

                End If

            Else
238             EncontreLugar = False
240             RegaloX = RandomNumber(20, 80)
242             RegaloY = RandomNumber(20, 80)

            End If

        End If
        
244     If EncontreLugar = True Then
246         BusquedaRegaloActiva = True
248         Call MakeObj(RegaloRegalo(RandomNumber(1, 5)), RegaloNumMapa, RegaloX, RegaloY, False)
250         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> De repente ha surgido un item maravilloso en el mapa: " & DarNameMapa(RegaloNumMapa) & "(" & RegaloNumMapa & ") ¿Quien sera el valiente que lo encuentre? ¡MUCHO CUIDADO!", FontTypeNames.FONTTYPE_TALK))
252         Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(497, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno

        End If

        
        Exit Sub

PerderRegalo_Err:
        Call RegistrarError(Err.Number, Err.description, "ModTesoros.PerderRegalo", Erl)
        Resume Next
        
End Sub

