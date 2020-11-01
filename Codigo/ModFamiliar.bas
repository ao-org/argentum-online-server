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

Public Sub LimpiarMascota(UserIndex)
    Npclist(UserList(UserIndex).Familiar.Id).EsFamiliar = 0
    UserList(UserIndex).Familiar.Muerto = 1
    UserList(UserIndex).Familiar.MinHp = 0
    UserList(UserIndex).Familiar.Invocado = 0
    UserList(UserIndex).Familiar.Id = 0
    UserList(UserIndex).Familiar.Paralizado = 0
    UserList(UserIndex).Familiar.Inmovilizado = 0
    'Call WriteConsoleMsg(UserIndex, "Tu familiar ha muerto, acercate al templo mas cercano para que sea resucitado.", FontTypeNames.FONTTYPE_WARNING)
    Call WriteLocaleMsg(UserIndex, "181", FontTypeNames.FONTTYPE_INFOIAO)
End Sub
Public Sub InvocarFamiliar(ByVal UserIndex As Integer, ByVal b As Boolean)

If UserList(UserIndex).Familiar.Muerto = 1 Then
    Call WriteLocaleMsg(UserIndex, "345", FontTypeNames.FONTTYPE_WARNING)
    Call WriteConsoleMsg(UserIndex, "Tu familiar esta muerto, acercate al templo mas cercano para que sea resucitado.", FontTypeNames.FONTTYPE_INFOIAO)
    Exit Sub
End If
Dim PosCasteadaX As Byte
Dim PosCasteadaY As Byte
Dim PosCasteadaM As Byte
Dim h As Integer
Dim TempX As Integer
Dim TempY As Integer
Dim Pos As WorldPos

    Pos.x = UserList(UserIndex).flags.TargetX
    Pos.Y = UserList(UserIndex).flags.TargetY
    Pos.Map = UserList(UserIndex).flags.TargetMap
 
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)



'If MapInfo(UserList(UserIndex).Pos.map).Pk = True Then

    Dim x As Long
    Dim Y As Long
    x = Pos.x
    Y = Pos.Y
    
    
    
    If MapData(UserList(UserIndex).Pos.Map, x, Y).Blocked = 1 Or MapData(UserList(UserIndex).Pos.Map, x, Y).TileExit.Map > 0 Or MapData(UserList(UserIndex).Pos.Map, x, Y).NpcIndex > 0 Or HayAgua(UserList(UserIndex).Pos.Map, x, Y) Then
                Call WriteLocaleMsg(UserIndex, "262", FontTypeNames.FONTTYPE_INFOIAO)
                'Call WriteConsoleMsg(UserIndex, "Area invalida para tirar el item.", FontTypeNames.FONTTYPE_INFO)
            Else
    'Envio Palabras magicas, wavs y fxs.
    If UserList(UserIndex).flags.NoPalabrasMagicas = 0 Then
        Call DecirPalabrasMagicas(h, UserIndex)
    End If
    
    '
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).wav, Pos.x, Pos.Y))  'Esta linea faltaba. Pablo (ToxicWaste)
    
With UserList(UserIndex)


    If .Familiar.Invocado = 0 Then
           .Familiar.Id = SpawnNpc(.Familiar.NpcIndex, Pos, False, True)
            'Controlamos que se sumoneo OK
            If .Familiar.Id = 0 Then
                'Call WriteConsoleMsg(UserIndex, "No hay espacio aquí para tu mascota. Se provoco un ERROR.", FontTypeNames.FONTTYPE_INFO)
                Call WriteLocaleMsg(UserIndex, "262", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub
            End If
            Call CargarFamiliar(UserIndex)
           ' Call FollowAmo(.Familiar.Id)
            If Hechizos(h).Particle > 0 Then '¿Envio Particula?
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))
            End If
            If Hechizos(h).FXgrh > 0 Then 'Envio Fx?
                Call modSendData.SendToAreaByPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
            End If
    Else
             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageFxPiso(Hechizos(h).FXgrh, Npclist(.Familiar.Id).Pos.x, Npclist(.Familiar.Id).Pos.Y))
        .Familiar.Invocado = 0
        Call QuitarNPC(.Familiar.Id)
    End If

    b = True
End With
'Else
End If


End Sub
Public Sub RevivirFamiliar(ByVal UserIndex As Integer)
With UserList(UserIndex)
.Familiar.MinHp = .Familiar.MaxHp
.Familiar.Muerto = 0
End With
'Call WriteConsoleMsg(UserIndex, "Tu familiar a sido revivido.", FontTypeNames.FONTTYPE_VIOLETA)
Call WriteLocaleMsg(UserIndex, "159", FontTypeNames.FONTTYPE_INFOIAO)
End Sub
Public Sub CargarFamiliar(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Npclist(.Familiar.Id).name = .Familiar.nombre
        Npclist(.Familiar.Id).Stats.MinHp = .Familiar.MinHp
        Npclist(.Familiar.Id).Stats.MaxHp = .Familiar.MaxHp
        Npclist(.Familiar.Id).Stats.MinHIT = .Familiar.MinHIT
        Npclist(.Familiar.Id).Stats.MaxHit = .Familiar.MaxHit
        Npclist(.Familiar.Id).EsFamiliar = 1

        Npclist(.Familiar.Id).Movement = TipoAI.SigueAmo
        Npclist(.Familiar.Id).Target = 0
        Npclist(.Familiar.Id).TargetNPC = 0
        .Familiar.Invocado = 1
        
    End With
End Sub
Public Function IndexDeFamiliar(ByVal Tipo As Byte) As Byte
'**************************************************************
'Author: Pablo Mercavides
'**************************************************************
    Select Case Tipo
        Case 1
            IndexDeFamiliar = 128
        Case 2
            IndexDeFamiliar = 127
        Case 3
            IndexDeFamiliar = 129
        Case 4
            IndexDeFamiliar = 126
        Case 5
            IndexDeFamiliar = 132
        Case 6
            IndexDeFamiliar = 145
        Case 7
            IndexDeFamiliar = 130
        Case 8
            IndexDeFamiliar = 133
        Case 9
            IndexDeFamiliar = 131
    End Select

End Function


Sub CalcularDarExpCompartida(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Integer)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/09/06 Nacho
'Reescribi gran parte del Sub
'Ahora, da toda la experiencia del npc mientras este vivo.
'***************************************************
Dim ExpaDar As Long






'[Nacho] Chekeamos que las variables sean validas para las operaciones
If ElDaño <= 0 Then ElDaño = 0

If Npclist(NpcIndex).Stats.MaxHp <= 0 Then Exit Sub
If ElDaño > Npclist(NpcIndex).Stats.MinHp Then ElDaño = Npclist(NpcIndex).Stats.MinHp

'[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
ExpaDar = CLng((ElDaño) * (Npclist(NpcIndex).GiveEXP / Npclist(NpcIndex).Stats.MaxHp))

If ExpaDar <= 0 Then Exit Sub

'[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
        'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
        'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
If ExpaDar > Npclist(NpcIndex).flags.ExpCount Then
    ExpaDar = Npclist(NpcIndex).flags.ExpCount
    Npclist(NpcIndex).flags.ExpCount = 0
Else
    Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).flags.ExpCount - ExpaDar
End If


If ExpMult > 0 Then
    ExpaDar = ExpaDar * ExpMult

End If

'[Nacho] Le damos la exp al user
Dim ExpUser As Long
Dim expPet As Long
ExpUser = ExpaDar / 2
expPet = ExpaDar / 2
If ExpUser > 0 Then
        UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpUser
        If UserList(UserIndex).Stats.Exp > MAXEXP Then _
            UserList(UserIndex).Stats.Exp = MAXEXP
       ' Call WriteConsoleMsg(UserIndex, "ID*140*" & ExpUser, FontTypeNames.FONTTYPE_EXP)
        Call WriteLocaleMsg(UserIndex, "140", FontTypeNames.FONTTYPE_EXP, ExpUser)
        Call CheckUserLevel(UserIndex)
End If

If expPet > 0 Then
        UserList(UserIndex).Familiar.Exp = UserList(UserIndex).Familiar.Exp + expPet
        If UserList(UserIndex).Familiar.Exp > MAXEXP Then _
            UserList(UserIndex).Familiar.Exp = MAXEXP
            
       ' Call WriteConsoleMsg(UserIndex, "ID*52*" & UserList(UserIndex).Familiar.Nombre & "*" & expPet & "*", FontTypeNames.FONTTYPE_EXP)
        Call CheckFamiliarLevel(UserIndex)
End If

End Sub


Sub CheckFamiliarLevel(ByVal UserIndex As Integer)

'*************************************************

On Error GoTo Errhandler




'¿Alcanzo el maximo nivel?
If UserList(UserIndex).Familiar.nivel >= STAT_MAXELV Then
    UserList(UserIndex).Familiar.ELU = 0
    UserList(UserIndex).Familiar.Exp = 0
    Exit Sub
End If
    



If UserList(UserIndex).Familiar.Exp >= UserList(UserIndex).Familiar.ELU Then
    
    'Checkea otra vez, esto sucede si tiene mas EXP y puede saltarse el maximo
    'nivel
    If UserList(UserIndex).Familiar.nivel >= STAT_MAXELV Then
        UserList(UserIndex).Familiar.Exp = 0
        UserList(UserIndex).Familiar.ELU = 0
        Exit Sub
    End If
    

    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
    
    

    UserList(UserIndex).Familiar.nivel = UserList(UserIndex).Familiar.nivel + 1
    Call WriteConsoleMsg(UserIndex, UserList(UserIndex).Familiar.nombre & " a subido de nivel!", FontTypeNames.FONTTYPE_INFOBOLD)
    
    UserList(UserIndex).Familiar.Exp = UserList(UserIndex).Familiar.Exp - UserList(UserIndex).Familiar.ELU
    
    'Nueva subida de exp x lvl. Pablo (ToxicWaste)
    If UserList(UserIndex).Familiar.nivel < 15 Then
        UserList(UserIndex).Familiar.ELU = UserList(UserIndex).Familiar.ELU * 1.4
    ElseIf UserList(UserIndex).Familiar.nivel < 21 Then
        UserList(UserIndex).Familiar.ELU = UserList(UserIndex).Familiar.ELU * 1.35
    ElseIf UserList(UserIndex).Familiar.nivel < 33 Then
        UserList(UserIndex).Familiar.ELU = UserList(UserIndex).Familiar.ELU * 1.3
    ElseIf UserList(UserIndex).Familiar.nivel < 41 Then
        UserList(UserIndex).Familiar.ELU = UserList(UserIndex).Familiar.ELU * 1.225
    Else
        UserList(UserIndex).Familiar.ELU = UserList(UserIndex).Familiar.ELU * 1.25
    End If

    UserList(UserIndex).Familiar.MaxHp = UserList(UserIndex).Familiar.MaxHp + 8
    
    UserList(UserIndex).Familiar.MinHIT = UserList(UserIndex).Familiar.MinHIT + 5
    UserList(UserIndex).Familiar.MaxHit = UserList(UserIndex).Familiar.MaxHit + 5
    
    
    
    Npclist(UserList(UserIndex).Familiar.Id).Stats.MaxHit = UserList(UserIndex).Familiar.MaxHit
    Npclist(UserList(UserIndex).Familiar.Id).Stats.MinHIT = UserList(UserIndex).Familiar.MinHIT
     
    Npclist(UserList(UserIndex).Familiar.Id).Stats.MaxHp = UserList(UserIndex).Familiar.MaxHp
    Npclist(UserList(UserIndex).Familiar.Id).Stats.MinHp = UserList(UserIndex).Familiar.MaxHp


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

Errhandler:
    Call LogError("Error en la subrutina de check mascota nivel - Error : " & Err.Number & " - Description : " & Err.description)
End Sub
