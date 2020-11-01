Attribute VB_Name = "MoLogros"
'Bueno hoy empezamos a crear un nuevo sistema de logros  20/04/2015

'Logro 1, Matar 20 NPCS

'Logro 2, Matar 60 npcs

'Logro 3, Matar 100 Npcs

'Logro 4, Matar 500 Npcs

'Logro 5, Matar 2000 Npcs
Public NPcLogros() As TLogros
Public UserLogros() As TLogros
Public LevelLogros() As TLogros
Private Const ARCHIVOCONFIG = "logros.ini"

Private CantNPcLogros As Byte
Private CantUserLogros As Byte
Private CantLevelLogros As Byte




Type TLogros
    nombre As String
    Desc As String
    cant As Long
    QueNpc As Integer
    TipoRecompensa As Byte
    ObjRecompensa As String
    OroRecompensa As Long
    ExpRecompensa As Long
    HechizoRecompensa As Byte
    Finalizada As Boolean
End Type
Public Sub CargarLogros()
Dim i As Integer

    CantNPcLogros = val(GetVar(DatPath & ARCHIVOCONFIG, "INIT", "NPcLogros"))
    CantUserLogros = val(GetVar(DatPath & ARCHIVOCONFIG, "INIT", "UserLogros"))
    CantLevelLogros = val(GetVar(DatPath & ARCHIVOCONFIG, "INIT", "LevelLogros"))
    

ReDim NPcLogros(1 To CantNPcLogros)
ReDim UserLogros(1 To CantUserLogros)
ReDim LevelLogros(1 To CantLevelLogros)

i = 1
    If CantNPcLogros > 0 Then
        For i = 1 To CantNPcLogros
            NPcLogros(i).nombre = GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "Nombre")
            NPcLogros(i).Desc = GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "Desc")
            NPcLogros(i).cant = val(GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "Cant"))
            NPcLogros(i).TipoRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "TipoRecompensa"))
            NPcLogros(i).ExpRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "ExpRecompensa"))
            NPcLogros(i).HechizoRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "HechizoRecompensa"))
            NPcLogros(i).OroRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "OroRecompensa"))
            NPcLogros(i).ObjRecompensa = GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "ObjRecompensa")
            NPcLogros(i).QueNpc = val(GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "QueNPC"))
        Next i
    End If
    
    
    If CantUserLogros > 0 Then
        For i = 1 To CantUserLogros
            UserLogros(i).nombre = GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "Nombre")
            UserLogros(i).Desc = GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "Desc")
            UserLogros(i).cant = val(GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "Cant"))
            UserLogros(i).TipoRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "TipoRecompensa"))
            UserLogros(i).ExpRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "ExpRecompensa"))
            UserLogros(i).HechizoRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "HechizoRecompensa"))
            UserLogros(i).OroRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "OroRecompensa"))
            UserLogros(i).ObjRecompensa = GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "ObjRecompensa")
            'Debug.Print i & ":" & UserLogros(i).ObjRecompensa
        Next i
    End If
    
    If CantLevelLogros > 0 Then
        For i = 1 To CantLevelLogros
            LevelLogros(i).nombre = GetVar(DatPath & ARCHIVOCONFIG, "LevelLogros" & i, "Nombre")
            LevelLogros(i).Desc = GetVar(DatPath & ARCHIVOCONFIG, "LevelLogros" & i, "Desc")
            LevelLogros(i).cant = val(GetVar(DatPath & ARCHIVOCONFIG, "LevelLogros" & i, "Cant"))
            LevelLogros(i).TipoRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "LevelLogros" & i, "TipoRecompensa"))
            LevelLogros(i).ExpRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "LevelLogros" & i, "ExpRecompensa"))
            LevelLogros(i).HechizoRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "LevelLogros" & i, "HechizoRecompensa"))
            LevelLogros(i).OroRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "LevelLogros" & i, "OroRecompensa"))
            LevelLogros(i).ObjRecompensa = GetVar(DatPath & ARCHIVOCONFIG, "LevelLogros" & i, "ObjRecompensa")
          '  Debug.Print i & ":" & LevelLogros(i).ObjRecompensa
        Next i
    End If
End Sub
    
    
Public Sub EnviarRecompensaStat(ByVal UserIndex As Integer)

If UserList(UserIndex).flags.BattleModo = 1 Then
    Call WriteConsoleMsg(UserIndex, "Aquí no podés utilizar el sistema de recompensas.", FontTypeNames.FONTTYPE_EXP)
    Exit Sub
End If


Call WriteRecompensas(UserIndex)



End Sub
Public Sub CheckearRecompesas(ByVal UserIndex As Integer, ByVal Index As Byte)

If UserList(UserIndex).flags.BattleModo = 1 Then Exit Sub

Select Case Index

Case 1
    If UserList(UserIndex).Stats.NPCsMuertos >= NPcLogros(UserList(UserIndex).NPcLogros + 1).cant Then
        'Call WriteConsoleMsg(UserIndex, "¡Felicitaciones! Ya podes reclamar una nueva recompensa.", FontTypeNames.FONTTYPE_EXP)
        Call WriteTrofeoToggleOn(UserIndex)
    End If
Case 2
    If UserList(UserIndex).Stats.UsuariosMatados >= UserLogros(UserList(UserIndex).UserLogros + 1).cant Then
        'Call WriteConsoleMsg(UserIndex, "¡Felicitaciones! Ya podes reclamar una nueva recompensa.", FontTypeNames.FONTTYPE_EXP)
        Call WriteTrofeoToggleOn(UserIndex)
    End If
Case 3
    If UserList(UserIndex).Stats.ELV >= LevelLogros(UserList(UserIndex).LevelLogros + 1).cant Then
        'Call WriteConsoleMsg(UserIndex, "¡Felicitaciones! Ya podes reclamar una nueva recompensa.", FontTypeNames.FONTTYPE_EXP)
        Call WriteTrofeoToggleOn(UserIndex)
    End If
End Select

End Sub


Public Sub EntregarRecompensas(ByVal UserIndex As Integer, ByVal Index As Byte)


If UserList(UserIndex).flags.BattleModo = 1 Then
    Call WriteConsoleMsg(UserIndex, "Aquí no podés utilizar el sistema de recompensas.", FontTypeNames.FONTTYPE_EXP)
    Exit Sub
End If

Select Case Index

Case 1
    If UserList(UserIndex).Stats.NPCsMuertos >= NPcLogros(UserList(UserIndex).NPcLogros + 1).cant Then
    
        Call WriteConsoleMsg(UserIndex, "Acá tenes tu recompensa por este logro. ¡Que lo disfrutes y seguí participando!", FontTypeNames.FONTTYPE_EXP)
        UserList(UserIndex).NPcLogros = UserList(UserIndex).NPcLogros + 1
        Call WriteRecompensas(UserIndex)
        Call WriteTrofeoToggleOff(UserIndex)
        Exit Sub
    Else
        Call WriteConsoleMsg(UserIndex, "Aún no has terminado este logro ¡Continua luchando!", FontTypeNames.FONTTYPE_EXP)
        Exit Sub
    End If
Case 2
    If UserList(UserIndex).Stats.UsuariosMatados >= UserLogros(UserList(UserIndex).UserLogros + 1).cant Then
    
        Call WriteConsoleMsg(UserIndex, "Acá tenes tu recompensa por este logro. ¡Que lo disfrutes y seguí participando!", FontTypeNames.FONTTYPE_EXP)
        UserList(UserIndex).UserLogros = UserList(UserIndex).UserLogros + 1
        Call WriteRecompensas(UserIndex)
        Call WriteTrofeoToggleOff(UserIndex)
        Exit Sub
    Else
        Call WriteConsoleMsg(UserIndex, "Aún no has terminado este logro ¡Continua luchando!", FontTypeNames.FONTTYPE_EXP)
        Exit Sub
    End If

Case 3
    If UserList(UserIndex).Stats.ELV >= LevelLogros(UserList(UserIndex).LevelLogros + 1).cant Then
    
        Call WriteConsoleMsg(UserIndex, "Acá tenes tu recompensa por este logro. ¡Que lo disfrutes y seguí participando!", FontTypeNames.FONTTYPE_EXP)
        UserList(UserIndex).LevelLogros = UserList(UserIndex).LevelLogros + 1
        Call WriteRecompensas(UserIndex)
        Call WriteTrofeoToggleOff(UserIndex)
        Exit Sub
    Else
        Call WriteConsoleMsg(UserIndex, "Aún no has terminado este logro ¡Continua luchando!", FontTypeNames.FONTTYPE_EXP)
        Exit Sub
    End If
End Select






End Sub
