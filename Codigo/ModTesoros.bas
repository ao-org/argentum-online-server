Attribute VB_Name = "ModTesoros"
Option Explicit

Public TesoroNPC() As Integer

Public TesoroNPCMapa() As Integer

Dim TesoroMapa()     As Integer

Dim TesoroRegalo()    As obj

Public BusquedaTesoroActiva As Boolean

Public TesoroNumMapa        As Integer

Public TesoroX              As Byte

Public TesoroY              As Byte

Dim RegaloMapa()     As Integer

Dim RegaloRegalo()    As obj

Public BusquedaRegaloActiva As Boolean

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
            ReDim TesoroTesoro(0)
            Exit Sub
        End If
    
        ReDim TesoroTesoro(1 To TiposDeTesoros)
        
        Dim Fields() As String, Str As String
        For i = 1 To TiposDeTesoros
            Str = Lector.GetValue("Tesoros", "Tesoro" & i)
        
            If LenB(Str) Then
                Fields = Split(Str, "-", 2)
                
                If UBound(Fields) >= 1 Then
                    With TesoroTesoro(i)
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
        
        Dim Fields() As String, Str As String
        For i = 1 To TiposDeRegalos
            Str = Lector.GetValue("Regalos", "Regalo" & i)
        
            If LenB(Str) Then
                Fields = Split(Str, "-", 2)
                
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
228         Call MakeObj(TesoroRegalo(RandomNumber(1, UBound(TesoroRegalo))), TesoroNumMapa, TesoroX, TesoroY, False)
230         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Rondan rumores que hay un tesoro enterrado en el mapa: " & DarNameMapa(TesoroNumMapa) & "(" & TesoroNumMapa & ") ¿Quien sera el afortunado que lo encuentre?", FontTypeNames.FONTTYPE_TALK))
232         Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(257, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno 257

        End If

        
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
248         Call MakeObj(RegaloRegalo(RandomNumber(1, UBound(RegaloRegalo))), RegaloNumMapa, RegaloX, RegaloY, False)
250         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> De repente ha surgido un item maravilloso en el mapa: " & DarNameMapa(RegaloNumMapa) & "(" & RegaloNumMapa & ") ¿Quien sera el valiente que lo encuentre? ¡MUCHO CUIDADO!", FontTypeNames.FONTTYPE_TALK))
252         Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(497, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno

        End If

        
        Exit Sub

PerderRegalo_Err:
254     Call RegistrarError(Err.Number, Err.Description, "ModTesoros.PerderRegalo", Erl)
256     Resume Next
        
End Sub

