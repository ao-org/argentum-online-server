Attribute VB_Name = "ModFamiliar"

Public Type Family

    nombre As String
    Muerto As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    Exp As Long
    ELU As Long
    nivel As Byte
    NpcIndex As Integer
    Invocado As Byte
    MinHp As Long
    MaxHp As Long
    Existe As Byte
    Id As Integer
    MinHIT As Long
    MaxHit As Long

End Type

Public Sub LimpiarMascota(Userindex)
        
        On Error GoTo LimpiarMascota_Err
        
100     Npclist(UserList(Userindex).Familiar.Id).EsFamiliar = 0
102     UserList(Userindex).Familiar.Muerto = 1
104     UserList(Userindex).Familiar.MinHp = 0
106     UserList(Userindex).Familiar.Invocado = 0
108     UserList(Userindex).Familiar.Id = 0
110     UserList(Userindex).Familiar.Paralizado = 0
112     UserList(Userindex).Familiar.Inmovilizado = 0
        'Call WriteConsoleMsg(UserIndex, "Tu familiar ha muerto, acercate al templo mas cercano para que sea resucitado.", FontTypeNames.FONTTYPE_WARNING)
114     Call WriteLocaleMsg(Userindex, "181", FontTypeNames.FONTTYPE_INFOIAO)

        
        Exit Sub

LimpiarMascota_Err:
116     Call RegistrarError(Err.Number, Err.description, "ModFamiliar.LimpiarMascota", Erl)
118     Resume Next
        
End Sub

Public Sub InvocarFamiliar(ByVal Userindex As Integer, ByVal b As Boolean)
        
        On Error GoTo InvocarFamiliar_Err
        

100     If UserList(Userindex).Familiar.Muerto = 1 Then
102         Call WriteLocaleMsg(Userindex, "345", FontTypeNames.FONTTYPE_WARNING)
104         Call WriteConsoleMsg(Userindex, "Tu familiar esta muerto, acercate al templo mas cercano para que sea resucitado.", FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub

        End If

        Dim PosCasteadaX As Byte

        Dim PosCasteadaY As Byte

        Dim PosCasteadaM As Byte

        Dim h            As Integer

        Dim TempX        As Integer

        Dim TempY        As Integer

        Dim Pos          As WorldPos

106     Pos.X = UserList(Userindex).flags.TargetX
108     Pos.Y = UserList(Userindex).flags.TargetY
110     Pos.Map = UserList(Userindex).flags.TargetMap
 
112     h = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)

        'If MapInfo(UserList(UserIndex).Pos.map).Pk = True Then

        Dim X As Long

        Dim Y As Long

114     X = Pos.X
116     Y = Pos.Y
    
118     If (MapData(UserList(Userindex).Pos.Map, X, Y).Blocked And eBlock.ALL_SIDES) = eBlock.ALL_SIDES Or MapData(UserList(Userindex).Pos.Map, X, Y).TileExit.Map > 0 Or MapData(UserList(Userindex).Pos.Map, X, Y).NpcIndex > 0 Or (MapData(UserList(Userindex).Pos.Map, X, Y).Blocked And FLAG_AGUA) <> 0 Then
120         Call WriteLocaleMsg(Userindex, "262", FontTypeNames.FONTTYPE_INFOIAO)
            'Call WriteConsoleMsg(UserIndex, "Area invalida para tirar el item.", FontTypeNames.FONTTYPE_INFO)
        Else

            'Envio Palabras magicas, wavs y fxs.
122         If UserList(Userindex).flags.NoPalabrasMagicas = 0 Then
124             Call DecirPalabrasMagicas(h, Userindex)

            End If
    
            '
    
126         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(h).wav, Pos.X, Pos.Y))  'Esta linea faltaba. Pablo (ToxicWaste)
    
128         With UserList(Userindex)

130             If .Familiar.Invocado = 0 Then
132                 .Familiar.Id = SpawnNpc(.Familiar.NpcIndex, Pos, False, True)

                    'Controlamos que se sumoneo OK
134                 If .Familiar.Id = 0 Then
                        'Call WriteConsoleMsg(UserIndex, "No hay espacio aquí para tu mascota. Se provoco un ERROR.", FontTypeNames.FONTTYPE_INFO)
136                     Call WriteLocaleMsg(Userindex, "262", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If

138                 Call CargarFamiliar(Userindex)

                    ' Call FollowAmo(.Familiar.Id)
140                 If Hechizos(h).Particle > 0 Then '¿Envio Particula?
142                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFXToFloor(UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))

                    End If

144                 If Hechizos(h).FXgrh > 0 Then 'Envio Fx?
146                     Call modSendData.SendToAreaByPos(UserList(Userindex).Pos.Map, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY))

                    End If

                Else
148                 Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageFxPiso(Hechizos(h).FXgrh, Npclist(.Familiar.Id).Pos.X, Npclist(.Familiar.Id).Pos.Y))
150                 .Familiar.Invocado = 0
152                 Call QuitarNPC(.Familiar.Id)

                End If

154             b = True

            End With

            'Else
        End If

        
        Exit Sub

InvocarFamiliar_Err:
156     Call RegistrarError(Err.Number, Err.description, "ModFamiliar.InvocarFamiliar", Erl)
158     Resume Next
        
End Sub

Public Sub RevivirFamiliar(ByVal Userindex As Integer)
        
        On Error GoTo RevivirFamiliar_Err
        

100     With UserList(Userindex)
102         .Familiar.MinHp = .Familiar.MaxHp
104         .Familiar.Muerto = 0

        End With

        'Call WriteConsoleMsg(UserIndex, "Tu familiar a sido revivido.", FontTypeNames.FONTTYPE_VIOLETA)
106     Call WriteLocaleMsg(Userindex, "159", FontTypeNames.FONTTYPE_INFOIAO)

        
        Exit Sub

RevivirFamiliar_Err:
108     Call RegistrarError(Err.Number, Err.description, "ModFamiliar.RevivirFamiliar", Erl)
110     Resume Next
        
End Sub

Public Sub CargarFamiliar(ByVal Userindex As Integer)
        
        On Error GoTo CargarFamiliar_Err
        

100     With UserList(Userindex)
102         Npclist(.Familiar.Id).name = .Familiar.nombre
104         Npclist(.Familiar.Id).Stats.MinHp = .Familiar.MinHp
106         Npclist(.Familiar.Id).Stats.MaxHp = .Familiar.MaxHp
108         Npclist(.Familiar.Id).Stats.MinHIT = .Familiar.MinHIT
110         Npclist(.Familiar.Id).Stats.MaxHit = .Familiar.MaxHit
112         Npclist(.Familiar.Id).EsFamiliar = 1

114         Npclist(.Familiar.Id).Movement = TipoAI.SigueAmo
116         Npclist(.Familiar.Id).Target = 0
118         Npclist(.Familiar.Id).TargetNPC = 0
120         .Familiar.Invocado = 1
        
        End With

        
        Exit Sub

CargarFamiliar_Err:
122     Call RegistrarError(Err.Number, Err.description, "ModFamiliar.CargarFamiliar", Erl)
124     Resume Next
        
End Sub

Public Function IndexDeFamiliar(ByVal Tipo As Byte) As Byte
        
        On Error GoTo IndexDeFamiliar_Err
        

        '**************************************************************
        'Author: Pablo Mercavides
        '**************************************************************
100     Select Case Tipo

            Case 1
102             IndexDeFamiliar = 128

104         Case 2
106             IndexDeFamiliar = 127

108         Case 3
110             IndexDeFamiliar = 129

112         Case 4
114             IndexDeFamiliar = 126

116         Case 5
118             IndexDeFamiliar = 132

120         Case 6
122             IndexDeFamiliar = 145

124         Case 7
126             IndexDeFamiliar = 130

128         Case 8
130             IndexDeFamiliar = 133

132         Case 9
134             IndexDeFamiliar = 131

        End Select

        
        Exit Function

IndexDeFamiliar_Err:
136     Call RegistrarError(Err.Number, Err.description, "ModFamiliar.IndexDeFamiliar", Erl)
138     Resume Next
        
End Function

Sub CalcularDarExpCompartida(ByVal Userindex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Integer)
        
        On Error GoTo CalcularDarExpCompartida_Err
        

        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/09/06 Nacho
        'Reescribi gran parte del Sub
        'Ahora, da toda la experiencia del npc mientras este vivo.
        '***************************************************
        Dim ExpaDar As Long

        '[Nacho] Chekeamos que las variables sean validas para las operaciones
100     If ElDaño <= 0 Then ElDaño = 0
    
102     If Npclist(NpcIndex).Stats.MaxHp <= 0 Then Exit Sub
104     If ElDaño > Npclist(NpcIndex).Stats.MinHp Then ElDaño = Npclist(NpcIndex).Stats.MinHp
    
        '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
106     ExpaDar = CLng((ElDaño) * (Npclist(NpcIndex).GiveEXP / Npclist(NpcIndex).Stats.MaxHp))
    
108     If ExpaDar <= 0 Then Exit Sub
    
        '[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
        'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
        'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
110     If ExpaDar > Npclist(NpcIndex).flags.ExpCount Then
112         ExpaDar = Npclist(NpcIndex).flags.ExpCount
114         Npclist(NpcIndex).flags.ExpCount = 0
        Else
116         Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).flags.ExpCount - ExpaDar

        End If
    
118     If ExpMult > 0 Then
120         ExpaDar = ExpaDar * ExpMult
    
        End If
    
        '[Nacho] Le damos la exp al user
        Dim ExpUser As Long

        Dim expPet  As Long

122     ExpUser = ExpaDar / 2
124     expPet = ExpaDar / 2

126     If ExpUser > 0 Then
128         If UserList(Userindex).Stats.ELV < STAT_MAXELV Then
130             UserList(Userindex).Stats.Exp = UserList(Userindex).Stats.Exp + ExpUser

132             If UserList(Userindex).Stats.Exp > MAXEXP Then UserList(Userindex).Stats.Exp = MAXEXP
                ' Call WriteConsoleMsg(UserIndex, "ID*140*" & ExpUser, FontTypeNames.FONTTYPE_EXP)
134             Call WriteLocaleMsg(Userindex, "140", FontTypeNames.FONTTYPE_EXP, ExpUser)
136             Call CheckUserLevel(Userindex)

            End If

        End If
    
138     If expPet > 0 Then
140         UserList(Userindex).Familiar.Exp = UserList(Userindex).Familiar.Exp + expPet

142         If UserList(Userindex).Familiar.Exp > MAXEXP Then UserList(Userindex).Familiar.Exp = MAXEXP
             
            ' Call WriteConsoleMsg(UserIndex, "ID*52*" & UserList(UserIndex).Familiar.Nombre & "*" & expPet & "*", FontTypeNames.FONTTYPE_EXP)
144         Call CheckFamiliarLevel(Userindex)

        End If

        
        Exit Sub

CalcularDarExpCompartida_Err:
146     Call RegistrarError(Err.Number, Err.description, "ModFamiliar.CalcularDarExpCompartida", Erl)
148     Resume Next
        
End Sub

Sub CheckFamiliarLevel(ByVal Userindex As Integer)

        '*************************************************

        On Error GoTo ErrHandler

        '¿Alcanzo el maximo nivel?
100     If UserList(Userindex).Familiar.nivel >= STAT_MAXELV Then
102         UserList(Userindex).Familiar.ELU = 0
104         UserList(Userindex).Familiar.Exp = 0
            Exit Sub

        End If

106     If UserList(Userindex).Familiar.Exp >= UserList(Userindex).Familiar.ELU Then
    
            'Checkea otra vez, esto sucede si tiene mas EXP y puede saltarse el maximo
            'nivel
108         If UserList(Userindex).Familiar.nivel >= STAT_MAXELV Then
110             UserList(Userindex).Familiar.Exp = 0
112             UserList(Userindex).Familiar.ELU = 0
                Exit Sub

            End If
    
114         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_NIVEL, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

116         UserList(Userindex).Familiar.nivel = UserList(Userindex).Familiar.nivel + 1
118         Call WriteConsoleMsg(Userindex, UserList(Userindex).Familiar.nombre & " a subido de nivel!", FontTypeNames.FONTTYPE_INFOBOLD)
    
120         UserList(Userindex).Familiar.Exp = UserList(Userindex).Familiar.Exp - UserList(Userindex).Familiar.ELU
    
            'Nueva subida de exp x lvl. Pablo (ToxicWaste)
122         If UserList(Userindex).Familiar.nivel < 15 Then
124             UserList(Userindex).Familiar.ELU = UserList(Userindex).Familiar.ELU * 1.4
126         ElseIf UserList(Userindex).Familiar.nivel < 21 Then
128             UserList(Userindex).Familiar.ELU = UserList(Userindex).Familiar.ELU * 1.35
130         ElseIf UserList(Userindex).Familiar.nivel < 33 Then
132             UserList(Userindex).Familiar.ELU = UserList(Userindex).Familiar.ELU * 1.3
134         ElseIf UserList(Userindex).Familiar.nivel < 41 Then
136             UserList(Userindex).Familiar.ELU = UserList(Userindex).Familiar.ELU * 1.225
            Else
138             UserList(Userindex).Familiar.ELU = UserList(Userindex).Familiar.ELU * 1.25

            End If

140         UserList(Userindex).Familiar.MaxHp = UserList(Userindex).Familiar.MaxHp + 8
    
142         UserList(Userindex).Familiar.MinHIT = UserList(Userindex).Familiar.MinHIT + 5
144         UserList(Userindex).Familiar.MaxHit = UserList(Userindex).Familiar.MaxHit + 5
    
146         Npclist(UserList(Userindex).Familiar.Id).Stats.MaxHit = UserList(Userindex).Familiar.MaxHit
148         Npclist(UserList(Userindex).Familiar.Id).Stats.MinHIT = UserList(Userindex).Familiar.MinHIT
     
150         Npclist(UserList(Userindex).Familiar.Id).Stats.MaxHp = UserList(Userindex).Familiar.MaxHp
152         Npclist(UserList(Userindex).Familiar.Id).Stats.MinHp = UserList(Userindex).Familiar.MaxHp

            '    Select Case UserList(UserIndex).clase
            '        Case eClass.Warrior
            '
            '            AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 2, 3)
            '            AumentoSTA = AumentoSTDef
            '
            '
            '        Case Else
            '
            '            AumentoHIT = 2
            '            AumentoSTA = AumentoSTDef
            ''
            '
            '   End Select

        End If

        Exit Sub

ErrHandler:
154     Call LogError("Error en la subrutina de check mascota nivel - Error : " & Err.Number & " - Description : " & Err.description)

End Sub
