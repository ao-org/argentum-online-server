Attribute VB_Name = "MoLogros"
'Bueno hoy empezamos a crear un nuevo sistema de logros  20/04/2015

'Logro 1, Matar 20 NPCS

'Logro 2, Matar 60 npcs

'Logro 3, Matar 100 Npcs

'Logro 4, Matar 500 Npcs

'Logro 5, Matar 2000 Npcs
Public NPcLogros()   As TLogros

Public UserLogros()  As TLogros

Public LevelLogros() As TLogros

Private Const ARCHIVOCONFIG = "logros.ini"

Private CantNPcLogros   As Byte

Private CantUserLogros  As Byte

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
        
        On Error GoTo CargarLogros_Err
        

        Dim i As Integer

100     CantNPcLogros = val(GetVar(DatPath & ARCHIVOCONFIG, "INIT", "NPcLogros"))
102     CantUserLogros = val(GetVar(DatPath & ARCHIVOCONFIG, "INIT", "UserLogros"))
104     CantLevelLogros = val(GetVar(DatPath & ARCHIVOCONFIG, "INIT", "LevelLogros"))

106     ReDim NPcLogros(1 To CantNPcLogros)
108     ReDim UserLogros(1 To CantUserLogros)
110     ReDim LevelLogros(1 To CantLevelLogros)

112     i = 1

114     If CantNPcLogros > 0 Then

116         For i = 1 To CantNPcLogros
118             NPcLogros(i).nombre = GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "Nombre")
120             NPcLogros(i).Desc = GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "Desc")
122             NPcLogros(i).cant = val(GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "Cant"))
124             NPcLogros(i).TipoRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "TipoRecompensa"))
126             NPcLogros(i).ExpRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "ExpRecompensa"))
128             NPcLogros(i).HechizoRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "HechizoRecompensa"))
130             NPcLogros(i).OroRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "OroRecompensa"))
132             NPcLogros(i).ObjRecompensa = GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "ObjRecompensa")
134             NPcLogros(i).QueNpc = val(GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "QueNPC"))
136         Next i

        End If
    
138     If CantUserLogros > 0 Then

140         For i = 1 To CantUserLogros
142             UserLogros(i).nombre = GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "Nombre")
144             UserLogros(i).Desc = GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "Desc")
146             UserLogros(i).cant = val(GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "Cant"))
148             UserLogros(i).TipoRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "TipoRecompensa"))
150             UserLogros(i).ExpRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "ExpRecompensa"))
152             UserLogros(i).HechizoRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "HechizoRecompensa"))
154             UserLogros(i).OroRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "OroRecompensa"))
156             UserLogros(i).ObjRecompensa = GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "ObjRecompensa")
                'Debug.Print i & ":" & UserLogros(i).ObjRecompensa
158         Next i

        End If
    
160     If CantLevelLogros > 0 Then

162         For i = 1 To CantLevelLogros
164             LevelLogros(i).nombre = GetVar(DatPath & ARCHIVOCONFIG, "LevelLogros" & i, "Nombre")
166             LevelLogros(i).Desc = GetVar(DatPath & ARCHIVOCONFIG, "LevelLogros" & i, "Desc")
168             LevelLogros(i).cant = val(GetVar(DatPath & ARCHIVOCONFIG, "LevelLogros" & i, "Cant"))
170             LevelLogros(i).TipoRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "LevelLogros" & i, "TipoRecompensa"))
172             LevelLogros(i).ExpRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "LevelLogros" & i, "ExpRecompensa"))
174             LevelLogros(i).HechizoRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "LevelLogros" & i, "HechizoRecompensa"))
176             LevelLogros(i).OroRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "LevelLogros" & i, "OroRecompensa"))
178             LevelLogros(i).ObjRecompensa = GetVar(DatPath & ARCHIVOCONFIG, "LevelLogros" & i, "ObjRecompensa")
                '  Debug.Print i & ":" & LevelLogros(i).ObjRecompensa
180         Next i

        End If

        
        Exit Sub

CargarLogros_Err:
182     Call RegistrarError(Err.Number, Err.description, "MoLogros.CargarLogros", Erl)
184     Resume Next
        
End Sub
    
Public Sub EnviarRecompensaStat(ByVal UserIndex As Integer)
        
        On Error GoTo EnviarRecompensaStat_Err
        

100     If UserList(UserIndex).flags.BattleModo = 1 Then
102         Call WriteConsoleMsg(UserIndex, "Aquí no podés utilizar el sistema de recompensas.", FontTypeNames.FONTTYPE_EXP)
            Exit Sub

        End If

104     Call WriteRecompensas(UserIndex)

        
        Exit Sub

EnviarRecompensaStat_Err:
106     Call RegistrarError(Err.Number, Err.description, "MoLogros.EnviarRecompensaStat", Erl)
108     Resume Next
        
End Sub

Public Sub CheckearRecompesas(ByVal UserIndex As Integer, ByVal index As Byte)
        
        On Error GoTo CheckearRecompesas_Err
        

100     If UserList(UserIndex).flags.BattleModo = 1 Then Exit Sub

102     Select Case index

            Case 1

104             If UserList(UserIndex).Stats.NPCsMuertos >= NPcLogros(UserList(UserIndex).NPcLogros + 1).cant Then
                    'Call WriteConsoleMsg(UserIndex, "¡Felicitaciones! Ya podes reclamar una nueva recompensa.", FontTypeNames.FONTTYPE_EXP)
106                 Call WriteTrofeoToggleOn(UserIndex)

                End If

108         Case 2

110             If UserList(UserIndex).Stats.UsuariosMatados >= UserLogros(UserList(UserIndex).UserLogros + 1).cant Then
                    'Call WriteConsoleMsg(UserIndex, "¡Felicitaciones! Ya podes reclamar una nueva recompensa.", FontTypeNames.FONTTYPE_EXP)
112                 Call WriteTrofeoToggleOn(UserIndex)

                End If

114         Case 3

116             If UserList(UserIndex).Stats.ELV >= LevelLogros(UserList(UserIndex).LevelLogros + 1).cant Then
                    'Call WriteConsoleMsg(UserIndex, "¡Felicitaciones! Ya podes reclamar una nueva recompensa.", FontTypeNames.FONTTYPE_EXP)
118                 Call WriteTrofeoToggleOn(UserIndex)

                End If

        End Select

        
        Exit Sub

CheckearRecompesas_Err:
120     Call RegistrarError(Err.Number, Err.description, "MoLogros.CheckearRecompesas", Erl)
122     Resume Next
        
End Sub

Public Sub EntregarRecompensas(ByVal UserIndex As Integer, ByVal index As Byte)
        
        On Error GoTo EntregarRecompensas_Err
        

100     If UserList(UserIndex).flags.BattleModo = 1 Then
102         Call WriteConsoleMsg(UserIndex, "Aquí no podés utilizar el sistema de recompensas.", FontTypeNames.FONTTYPE_EXP)
            Exit Sub

        End If

104     Select Case index

            Case 1

106             If UserList(UserIndex).Stats.NPCsMuertos >= NPcLogros(UserList(UserIndex).NPcLogros + 1).cant Then
    
108                 Call WriteConsoleMsg(UserIndex, "Acá tenes tu recompensa por este logro. ¡Que lo disfrutes y seguí participando!", FontTypeNames.FONTTYPE_EXP)
110                 UserList(UserIndex).NPcLogros = UserList(UserIndex).NPcLogros + 1
112                 Call WriteRecompensas(UserIndex)
114                 Call WriteTrofeoToggleOff(UserIndex)
                    Exit Sub
                Else
116                 Call WriteConsoleMsg(UserIndex, "Aún no has terminado este logro ¡Continua luchando!", FontTypeNames.FONTTYPE_EXP)
                    Exit Sub

                End If

118         Case 2

120             If UserList(UserIndex).Stats.UsuariosMatados >= UserLogros(UserList(UserIndex).UserLogros + 1).cant Then
    
122                 Call WriteConsoleMsg(UserIndex, "Acá tenes tu recompensa por este logro. ¡Que lo disfrutes y seguí participando!", FontTypeNames.FONTTYPE_EXP)
124                 UserList(UserIndex).UserLogros = UserList(UserIndex).UserLogros + 1
126                 Call WriteRecompensas(UserIndex)
128                 Call WriteTrofeoToggleOff(UserIndex)
                    Exit Sub
                Else
130                 Call WriteConsoleMsg(UserIndex, "Aún no has terminado este logro ¡Continua luchando!", FontTypeNames.FONTTYPE_EXP)
                    Exit Sub

                End If

132         Case 3

134             If UserList(UserIndex).Stats.ELV >= LevelLogros(UserList(UserIndex).LevelLogros + 1).cant Then
    
136                 Call WriteConsoleMsg(UserIndex, "Acá tenes tu recompensa por este logro. ¡Que lo disfrutes y seguí participando!", FontTypeNames.FONTTYPE_EXP)
138                 UserList(UserIndex).LevelLogros = UserList(UserIndex).LevelLogros + 1
140                 Call WriteRecompensas(UserIndex)
142                 Call WriteTrofeoToggleOff(UserIndex)
                    Exit Sub
                Else
144                 Call WriteConsoleMsg(UserIndex, "Aún no has terminado este logro ¡Continua luchando!", FontTypeNames.FONTTYPE_EXP)
                    Exit Sub

                End If

        End Select

        
        Exit Sub

EntregarRecompensas_Err:
146     Call RegistrarError(Err.Number, Err.description, "MoLogros.EntregarRecompensas", Erl)
148     Resume Next
        
End Sub
