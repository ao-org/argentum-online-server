Attribute VB_Name = "Acciones"

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

''
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

    On Error Resume Next

    '¿Rango Visión? (ToxicWaste)
    If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.X - X) > RANGO_VISION_X) Then
        Exit Sub

    End If

    '¿Posicion valida?
    If InMapBounds(Map, X, Y) Then
   
        Dim FoundChar      As Byte

        Dim FoundSomething As Byte

        Dim TempCharIndex  As Integer
       
        If MapData(Map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
            'Set the target NPC
            UserList(UserIndex).flags.TargetNPC = MapData(Map, X, Y).NpcIndex
            UserList(UserIndex).flags.TargetNpcTipo = Npclist(MapData(Map, X, Y).NpcIndex).NPCtype
        
            If Npclist(MapData(Map, X, Y).NpcIndex).Comercia = 1 Then

                '¿Esta el user muerto? Si es asi no puede comerciar
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Is it already in commerce mode??
                If UserList(UserIndex).flags.Comerciando Then
                    Exit Sub

                End If
            
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 6 Then
                    Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Iniciamos la rutina pa' comerciar.
                Call IniciarComercioNPC(UserIndex)
        
            ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Banquero Then

                '¿Esta el user muerto? Si es asi no puede comerciar
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Is it already in commerce mode??
                If UserList(UserIndex).flags.Comerciando Then
                    Exit Sub

                End If
            
                If Distancia(Npclist(MapData(Map, X, Y).NpcIndex).Pos, UserList(UserIndex).Pos) > 6 Then
                    Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del banquero.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'A depositar de una
                Call IniciarBanco(UserIndex)
            
            ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Pirata Then  'VIAJES

                '¿Esta el user muerto? Si es asi no puede comerciar
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Is it already in commerce mode??
                If UserList(UserIndex).flags.Comerciando Then
                    Exit Sub

                End If
            
                If Distancia(Npclist(MapData(Map, X, Y).NpcIndex).Pos, UserList(UserIndex).Pos) > 5 Then
                    Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del banquero.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                If Npclist(MapData(Map, X, Y).NpcIndex).SoundOpen <> 0 Then
                    Call WritePlayWave(UserIndex, Npclist(MapData(Map, X, Y).NpcIndex).SoundOpen, NO_3D_SOUND, NO_3D_SOUND)

                End If

                'A depositar de unaIniciarTransporte
                Call WriteViajarForm(UserIndex, MapData(Map, X, Y).NpcIndex)
                Exit Sub
            
            ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Revividor Or Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.ResucitadorNewbie Then

                If Distancia(UserList(UserIndex).Pos, Npclist(MapData(Map, X, Y).NpcIndex).Pos) > 5 Then
                    'Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                UserList(UserIndex).flags.Envenenado = 0
                UserList(UserIndex).flags.Incinerado = 0
      
                'Revivimos si es necesario
                If UserList(UserIndex).flags.Muerto = 1 And (Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex)) Then
                    Call WriteConsoleMsg(UserIndex, "¡Has sido resucitado!", FontTypeNames.FONTTYPE_INFO)
                    Call RevivirUsuario(UserIndex)
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Resucitar, 30, False))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("204", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                
                Else

                    'curamos totalmente
                    If UserList(UserIndex).Stats.MinHp <> UserList(UserIndex).Stats.MaxHp Then
                        UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
                        Call WritePlayWave(UserIndex, "101", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
                        'Call WriteConsoleMsg(UserIndex, "El Cura lanza unas palabras al aire. Comienzas a sentir como tu cuerpo se vuelve a formar...¡Has sido curado!", FontTypeNames.FONTTYPE_INFO)
                        Call WriteLocaleMsg(UserIndex, "83", FontTypeNames.FONTTYPE_INFOIAO)
                    
                        Call WriteUpdateUserStats(UserIndex)

                        If Status(UserIndex) = 2 Or Status(UserIndex) = 0 Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.CurarCrimi, 100, False))
                        Else
           
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Curar, 100, False))

                        End If

                    End If

                End If
            
                'Sistema Battle
            
            ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.BattleModo Then

                If Distancia(UserList(UserIndex).Pos, Npclist(MapData(Map, X, Y).NpcIndex).Pos) > 5 Then
                    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                If BattleActivado = 0 Then
                    Call WriteChatOverHead(UserIndex, "Actualmente el battle se encuentra desactivado.", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
                    Exit Sub

                End If
                        
                If UserList(UserIndex).clase = eClass.Trabajador Then
                    Call WriteConsoleMsg(UserIndex, "Los trabajadores no pueden ingresar al battle.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If
            
                If UserList(UserIndex).Stats.ELV < BattleMinNivel Then
                    Call WriteConsoleMsg(UserIndex, "Exclusivo para personajes superiores a nivel " & BattleMinNivel, FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                If UserList(UserIndex).flags.Comerciando Then
                    Call WriteConsoleMsg(UserIndex, "No podes ingresar al battle si estas comerciando.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If
            
                If UserList(UserIndex).flags.EnTorneo = True Then
                    Call WriteConsoleMsg(UserIndex, "No podes ingresar al battle estando anotado en un evento.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If
            
                If UserList(UserIndex).Accion.TipoAccion = Accion_Barra.BattleModo Then Exit Sub
                If UserList(UserIndex).donador.activo = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Runa, 400, False))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 400, Accion_Barra.BattleModo))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Runa, 50, False))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 50, Accion_Barra.BattleModo))

                End If

                UserList(UserIndex).Accion.AccionPendiente = True
                UserList(UserIndex).Accion.Particula = ParticulasIndex.Runa
                UserList(UserIndex).Accion.TipoAccion = Accion_Barra.BattleModo
            
                'Sistema Battle
         
            ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Subastador Then

                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 1 Then
                    ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del subastador.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

                Call IniciarSubasta(UserIndex)
            
            ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Quest Then

                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                Call EnviarQuest(UserIndex)
            
            ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Enlistador Then

                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 4 Then
                    Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
        
                If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                    If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                        Call EnlistarArmadaReal(UserIndex)
                    Else
                        Call RecompensaArmadaReal(UserIndex)

                    End If

                Else

                    If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                        Call EnlistarCaos(UserIndex)
                    Else
                        Call RecompensaCaos(UserIndex)

                    End If

                End If

            ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Gobernador Then

                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If
            
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                    Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del gobernador.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

                Dim DeDonde As String
            
                If UserList(UserIndex).Hogar = Npclist(UserList(UserIndex).flags.TargetNPC).GobernadorDe Then
                    Call WriteChatOverHead(UserIndex, "Ya perteneces a esta ciudad. Gracias por ser uno más de nosotros.", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
                    Exit Sub

                End If
            
                If UserList(UserIndex).Faccion.Status = 0 Or UserList(UserIndex).Faccion.Status = 2 Then
                    If Npclist(UserList(UserIndex).flags.TargetNPC).GobernadorDe = eCiudad.cBanderbill Then
                        Call WriteChatOverHead(UserIndex, "Aquí no aceptamos criminales.", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
                        Exit Sub

                    End If

                End If
            
                If UserList(UserIndex).Faccion.Status = 3 Or UserList(UserIndex).Faccion.Status = 1 Then
                    If Npclist(UserList(UserIndex).flags.TargetNPC).GobernadorDe = eCiudad.cArghal Then
                        Call WriteChatOverHead(UserIndex, "¡¡Sal de aquí ciudadano asqueroso!!", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
                        Exit Sub

                    End If

                End If
            
                If UserList(UserIndex).Hogar <> Npclist(UserList(UserIndex).flags.TargetNPC).GobernadorDe Then
            
                    UserList(UserIndex).PosibleHogar = Npclist(UserList(UserIndex).flags.TargetNPC).GobernadorDe
                
                    Select Case UserList(UserIndex).PosibleHogar

                        Case eCiudad.cUllathorpe
                            DeDonde = "Ullathorpe"
                            
                        Case eCiudad.cNix
                            DeDonde = "Nix"
                
                        Case eCiudad.cBanderbill
                            DeDonde = "Banderbill"
                        
                        Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                            DeDonde = "Lindos"
                            
                        Case eCiudad.cArghal
                            DeDonde = " Arghal"
                            
                        Case eCiudad.CHillidan
                            DeDonde = " Hillidan"
                            
                        Case Else
                            DeDonde = "Ullathorpe"

                    End Select
                    
                    UserList(UserIndex).flags.pregunta = 3
                    Call WritePreguntaBox(UserIndex, "¿Te gustaria ser ciudadano de " & DeDonde & "?")
                
                End If

            End If
        
            '¿Es un obj?
        ElseIf MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
            UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.ObjIndex
        
            Select Case ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType
            
                Case eOBJType.otPuertas 'Es una puerta
                    Call AccionParaPuerta(Map, X, Y, UserIndex)

                Case eOBJType.otCarteles 'Es un cartel
                    Call AccionParaCartel(Map, X, Y, UserIndex)

                Case eOBJType.OtCorreo 'Es un cartel
                    'Call AccionParaCorreo(Map, x, Y, UserIndex)

                Case eOBJType.otForos 'Foro
                    'Call AccionParaForo(Map, X, Y, UserIndex)
                    Call WriteConsoleMsg(UserIndex, "El foro está temporalmente deshabilitado.", FontTypeNames.FONTTYPE_EJECUCION)

                Case eOBJType.OtPozos 'Pozos
                    'Call AccionParaPozos(Map, x, Y, UserIndex)

                Case eOBJType.otArboles 'Pozos
                    'Call AccionParaArboles(Map, x, Y, UserIndex)

                Case eOBJType.otYunque 'Pozos
                    Call AccionParaYunque(Map, X, Y, UserIndex)

                Case eOBJType.otLeña    'Leña

                    If MapData(Map, X, Y).ObjInfo.ObjIndex = FOGATA_APAG And UserList(UserIndex).flags.Muerto = 0 Then
                        Call AccionParaRamita(Map, X, Y, UserIndex)

                    End If

            End Select

            '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
        ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
            UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y).ObjInfo.ObjIndex
        
            Select Case ObjData(MapData(Map, X + 1, Y).ObjInfo.ObjIndex).OBJType
            
                Case eOBJType.otPuertas 'Es una puerta
                    Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
            
            End Select

        ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
            UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex

            Select Case ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType
            
                Case eOBJType.otPuertas 'Es una puerta
                    Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)
            
            End Select

        ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
            UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y + 1).ObjInfo.ObjIndex

            Select Case ObjData(MapData(Map, X, Y + 1).ObjInfo.ObjIndex).OBJType
            
                Case eOBJType.otPuertas 'Es una puerta
                    Call AccionParaPuerta(Map, X, Y + 1, UserIndex)

            End Select

        'ElseIf HayAgua(Map, x, Y) Then
            'Call AccionParaAgua(Map, x, Y, UserIndex)

        End If

    End If

End Sub

Sub AccionParaForo(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim Pos As WorldPos

    Pos.Map = Map
    Pos.X = X
    Pos.Y = Y

    If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
        Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    '¿Hay mensajes?
    Dim f As String, tit As String, men As String, BASE As String, auxcad As String

    f = App.Path & "\foros\" & UCase$(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).ForoID) & ".for"

    If FileExist(f, vbNormal) Then

        Dim num As Integer

        num = val(GetVar(f, "INFO", "CantMSG"))
        BASE = Left$(f, Len(f) - 4)

        Dim i As Integer

        Dim n As Integer

        For i = 1 To num
            n = FreeFile
            f = BASE & i & ".for"
            Open f For Input Shared As #n
            Input #n, tit
            men = vbNullString
            auxcad = vbNullString

            Do While Not EOF(n)
                Input #n, auxcad
                men = men & vbCrLf & auxcad
            Loop
            Close #n
            Call WriteAddForumMsg(UserIndex, tit, men)
        
        Next

    End If

    Call WriteShowForumForm(UserIndex)

End Sub

Sub AccionParaPozos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim Pos As WorldPos

    Pos.Map = Map
    Pos.X = X
    Pos.Y = Y

    If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
        Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If MapData(Map, X, Y).ObjInfo.Amount <= 1 Then
        Call WriteConsoleMsg(UserIndex, "El pozo esta drenado, regresa mas tarde...", FontTypeNames.FONTTYPE_EJECUCION)
        Exit Sub

    End If

    If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Subtipo = 1 Then
        If UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN Then
            Call WriteConsoleMsg(UserIndex, "No tenes necesidad del pozo...", FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub

        End If

        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
        MapData(Map, X, Y).ObjInfo.Amount = MapData(Map, X, Y).ObjInfo.Amount - 1
        Call WriteConsoleMsg(UserIndex, "Sientes la frescura del pozo. ¡Tu maná a sido restaurada!", FontTypeNames.FONTTYPE_EJECUCION)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Call WriteUpdateUserStats(UserIndex)
        Exit Sub

    End If

    If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Subtipo = 2 Then
        If UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU Then
            Call WriteConsoleMsg(UserIndex, "No tenes necesidad del pozo...", FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub

        End If

        UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
        UserList(UserIndex).flags.Sed = 0 'Bug reparado 27/01/13
        MapData(Map, X, Y).ObjInfo.Amount = MapData(Map, X, Y).ObjInfo.Amount - 1
        Call WriteConsoleMsg(UserIndex, "Sientes la frescura del pozo. ¡Ya no sientes sed!", FontTypeNames.FONTTYPE_EJECUCION)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Call WriteUpdateHungerAndThirst(UserIndex)
        Exit Sub

    End If

End Sub

Sub AccionParaArboles(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim Pos As WorldPos

    Pos.Map = Map
    Pos.X = X
    Pos.Y = Y

    If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
        Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
        Call WriteConsoleMsg(UserIndex, "Esta prohibido manipular árboles en las ciudades.", FontTypeNames.FONTTYPE_INFOIAO)
        Call WriteWorkRequestTarget(UserIndex, 0)
        Exit Sub

    End If

    If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) < 40 Then
        Call WriteConsoleMsg(UserIndex, "No tenes suficientes conocimientos para comer del arbol. Necesitas al menos 40 skill en supervivencia.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If MapData(Map, X, Y).ObjInfo.Amount <= 1 Then
        Call WriteConsoleMsg(UserIndex, "El árbol no tiene más frutos para dar.", FontTypeNames.FONTTYPE_INFOIAO)
        Exit Sub

    End If

    If UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MaxHam Then
        Call WriteConsoleMsg(UserIndex, "No tenes hambre.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam + 5
    UserList(UserIndex).Stats.MaxHam = 100
    UserList(UserIndex).flags.Hambre = 0 'Bug reparado 27/01/13
    
    MapData(Map, X, Y).ObjInfo.Amount = MapData(Map, X, Y).ObjInfo.Amount - 1
    
    If Not UserList(UserIndex).flags.UltimoMensaje = 40 Then
        Call WriteConsoleMsg(UserIndex, "Logras conseguir algunos frutos del árbol, ya no sientes tanta hambre.", FontTypeNames.FONTTYPE_INFOIAO)
        UserList(UserIndex).flags.UltimoMensaje = 40

    End If
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(e_SoundIndex.MORFAR_MANZANA, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    Call WriteUpdateHungerAndThirst(UserIndex)

End Sub

Sub AccionParaAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim Pos As WorldPos

    Pos.Map = Map
    Pos.X = X
    Pos.Y = Y

    If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
        Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
        Call WriteConsoleMsg(UserIndex, "Esta prohibido beber agua en las orillas de las ciudades.", FontTypeNames.FONTTYPE_INFO)
        Call WriteWorkRequestTarget(UserIndex, 0)
        Exit Sub

    End If

    If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) < 30 Then
        Call WriteConsoleMsg(UserIndex, "No tenes suficientes conocimientos para beber del agua. Necesitas al menos 30 skill en supervivencia.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU Then
        Call WriteConsoleMsg(UserIndex, "No tenes sed.", FontTypeNames.FONTTYPE_INFOIAO)
        Exit Sub

    End If

    UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU + 5
    UserList(UserIndex).flags.Sed = 0 'Bug reparado 27/01/13
    
    If Not UserList(UserIndex).flags.UltimoMensaje = 41 Then
        Call WriteConsoleMsg(UserIndex, "Has bebido, ya no sientes tanta sed.", FontTypeNames.FONTTYPE_INFOIAO)
        UserList(UserIndex).flags.UltimoMensaje = 41

    End If
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    Call WriteUpdateHungerAndThirst(UserIndex)

End Sub

Sub AccionParaYunque(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim Pos As WorldPos

    Pos.Map = Map
    Pos.X = X
    Pos.Y = Y

    If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
        Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> MARTILLO_HERRERO Then
        'Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Antes debes tener equipado un martillo de herrero.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    Call EnivarArmasConstruibles(UserIndex)
    Call EnivarArmadurasConstruibles(UserIndex)
    Call WriteShowBlacksmithForm(UserIndex)

    'UserList(UserIndex).Invent.HerramientaEqpObjIndex = objindex
    'UserList(UserIndex).Invent.HerramientaEqpSlot = slot

End Sub

Sub AccionParaPuerta(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim MiObj As obj

    Dim wp    As WorldPos

    If Not (Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2) Then
        If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
            If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Cerrada = 1 Then

                'Abre la puerta
                If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
                    
                    MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexAbierta
                    
                    Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(MapData(Map, X, Y).ObjInfo.ObjIndex, X, Y))
                    
                    Call BloquearPuerta(Map, X, Y, False)
                      
                    'Sonido
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                'Cierra puerta
                MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexCerrada
                
                Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(MapData(Map, X, Y).ObjInfo.ObjIndex, X, Y))
                                
                Call BloquearPuerta(Map, X, Y, True)

                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))

            End If
        
            UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.ObjIndex
        Else
            Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)

        End If

    Else
        Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)

        ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
    End If

End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim MiObj As obj

    If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = 8 Then
  
        If Len(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).texto) > 0 Then
            Call WriteShowSignal(UserIndex, MapData(Map, X, Y).ObjInfo.ObjIndex)

        End If
  
    End If

End Sub

Sub AccionParaCorreo(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo AccionParaCorreo_Err
        

        Dim Pos As WorldPos

100     Pos.Map = Map
102     Pos.X = X
104     Pos.Y = Y

106     If UserList(UserIndex).flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto.", FontTypeNames.FONTTYPE_INFO)
108         Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

110     If Distancia(Pos, UserList(UserIndex).Pos) > 4 Then
112         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

114     If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = 47 Then
116         Call WriteListaCorreo(UserIndex, False)

        End If

        
        Exit Sub

AccionParaCorreo_Err:
        Call RegistrarError(Err.Number, Err.description, "Argentum20Server.Acciones.AccionParaCorreo", Erl)
        Resume Next
        
End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim Suerte As Byte
    Dim exito  As Byte
    Dim raise  As Integer
    
    Dim Pos    As WorldPos

    Pos.Map = Map
    Pos.X = X
    Pos.Y = Y
    
    With UserList(UserIndex)
    
        If Distancia(Pos, .Pos) > 2 Then
            Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If MapInfo(Map).lluvia And Lloviendo Then
            Call WriteConsoleMsg(UserIndex, "Esta lloviendo, no podés encender una fogata aquí.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If MapData(Map, X, Y).trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Seguro = 1 Then
            Call WriteConsoleMsg(UserIndex, "En zona segura no podés hacer fogatas.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If MapData(Map, X - 1, Y).ObjInfo.ObjIndex = FOGATA Or _
           MapData(Map, X + 1, Y).ObjInfo.ObjIndex = FOGATA Or _
           MapData(Map, X, Y - 1).ObjInfo.ObjIndex = FOGATA Or _
           MapData(Map, X, Y + 1).ObjInfo.ObjIndex = FOGATA Then
           
            Call WriteConsoleMsg(UserIndex, "Debes alejarte un poco de la otra fogata.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        If .Stats.UserSkills(Supervivencia) > 1 And .Stats.UserSkills(Supervivencia) < 6 Then
            Suerte = 3
        
        ElseIf .Stats.UserSkills(Supervivencia) >= 6 And .Stats.UserSkills(Supervivencia) <= 10 Then
            Suerte = 2
        
        ElseIf .Stats.UserSkills(Supervivencia) >= 10 And .Stats.UserSkills(Supervivencia) Then
            Suerte = 1

        End If

        exito = RandomNumber(1, Suerte)

        If exito = 1 Then
    
            If MapInfo(.Pos.Map).zone <> Ciudad Then
                
                Dim obj As obj
                obj.ObjIndex = FOGATA
                obj.Amount = 1
        
                Call WriteConsoleMsg(UserIndex, "Has prendido la fogata.", FontTypeNames.FONTTYPE_INFO)
        
                Call MakeObj(obj, Map, X, Y)

            Else
        
                Call WriteConsoleMsg(UserIndex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
            
                Exit Sub

            End If

        Else
        
            Call WriteConsoleMsg(UserIndex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)

        End If
    
    End With

    Call SubirSkill(UserIndex, Supervivencia)

End Sub
