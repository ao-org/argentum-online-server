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

106     i = 1

108     If CantNPcLogros > 0 Then
110         ReDim NPcLogros(1 To CantNPcLogros)
112         For i = 1 To CantNPcLogros
114             NPcLogros(i).nombre = GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "Nombre")
116             NPcLogros(i).Desc = GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "Desc")
118             NPcLogros(i).cant = val(GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "Cant"))
120             NPcLogros(i).TipoRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "TipoRecompensa"))
122             NPcLogros(i).ExpRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "ExpRecompensa"))
124             NPcLogros(i).HechizoRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "HechizoRecompensa"))
126             NPcLogros(i).OroRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "OroRecompensa"))
128             NPcLogros(i).ObjRecompensa = GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "ObjRecompensa")
130             NPcLogros(i).QueNpc = val(GetVar(DatPath & ARCHIVOCONFIG, "NPcLogros" & i, "QueNPC"))
132         Next i

        End If
    
134     If CantUserLogros > 0 Then
136         ReDim UserLogros(1 To CantUserLogros)
138         For i = 1 To CantUserLogros
140             UserLogros(i).nombre = GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "Nombre")
142             UserLogros(i).Desc = GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "Desc")
144             UserLogros(i).cant = val(GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "Cant"))
146             UserLogros(i).TipoRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "TipoRecompensa"))
148             UserLogros(i).ExpRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "ExpRecompensa"))
150             UserLogros(i).HechizoRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "HechizoRecompensa"))
152             UserLogros(i).OroRecompensa = val(GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "OroRecompensa"))
154             UserLogros(i).ObjRecompensa = GetVar(DatPath & ARCHIVOCONFIG, "UserLogros" & i, "ObjRecompensa")
                'Debug.Print i & ":" & UserLogros(i).ObjRecompensa
156         Next i

        End If
    
158     If CantLevelLogros > 0 Then
160         ReDim LevelLogros(1 To CantLevelLogros)
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
182     Call TraceError(Err.Number, Err.Description, "MoLogros.CargarLogros", Erl)
184
        
End Sub
    
Public Sub EnviarRecompensaStat(ByVal UserIndex As Integer)
        
        On Error GoTo EnviarRecompensaStat_Err
        

100     Call WriteRecompensas(UserIndex)

        
        Exit Sub

EnviarRecompensaStat_Err:
102     Call TraceError(Err.Number, Err.Description, "MoLogros.EnviarRecompensaStat", Erl)
104
        
End Sub

Public Sub CheckearRecompesas(ByVal UserIndex As Integer, ByVal Index As Byte)
        
        On Error GoTo CheckearRecompesas_Err
        
100     Select Case Index

            Case 1

102             If UserList(UserIndex).Stats.NPCsMuertos >= NPcLogros(UserList(UserIndex).NPcLogros + 1).cant Then
                    'Call WriteConsoleMsg(UserIndex, "¡Felicitaciones! Ya podes reclamar una nueva recompensa.", FontTypeNames.FONTTYPE_EXP)
104                 Call WriteTrofeoToggleOn(UserIndex)

                End If

106         Case 2

108             If UserList(UserIndex).Stats.UsuariosMatados >= UserLogros(UserList(UserIndex).UserLogros + 1).cant Then
                    'Call WriteConsoleMsg(UserIndex, "¡Felicitaciones! Ya podes reclamar una nueva recompensa.", FontTypeNames.FONTTYPE_EXP)
110                 Call WriteTrofeoToggleOn(UserIndex)

                End If

112         Case 3

114             If UserList(UserIndex).Stats.ELV >= LevelLogros(UserList(UserIndex).LevelLogros + 1).cant Then
                    'Call WriteConsoleMsg(UserIndex, "¡Felicitaciones! Ya podes reclamar una nueva recompensa.", FontTypeNames.FONTTYPE_EXP)
116                 Call WriteTrofeoToggleOn(UserIndex)

                End If

        End Select

        
        Exit Sub

CheckearRecompesas_Err:
118     Call TraceError(Err.Number, Err.Description, "MoLogros.CheckearRecompesas", Erl)
120
        
End Sub

Public Sub EntregarRecompensas(ByVal UserIndex As Integer, ByVal Index As Byte)
        
        On Error GoTo EntregarRecompensas_Err


100     Select Case Index

            Case 1

102             If UserList(UserIndex).Stats.NPCsMuertos >= NPcLogros(UserList(UserIndex).NPcLogros + 1).cant Then
    
104                 Call WriteConsoleMsg(UserIndex, "Acá tenes tu recompensa por este logro. ¡Que lo disfrutes y seguí participando!", FontTypeNames.FONTTYPE_EXP)
106                 UserList(UserIndex).NPcLogros = UserList(UserIndex).NPcLogros + 1
108                 Call WriteRecompensas(UserIndex)
110                 Call WriteTrofeoToggleOff(UserIndex)
                    Exit Sub
                Else
112                 Call WriteConsoleMsg(UserIndex, "Aún no has terminado este logro ¡Continua luchando!", FontTypeNames.FONTTYPE_EXP)
                    Exit Sub

                End If

114         Case 2

116             If UserList(UserIndex).Stats.UsuariosMatados >= UserLogros(UserList(UserIndex).UserLogros + 1).cant Then
    
118                 Call WriteConsoleMsg(UserIndex, "Acá tenes tu recompensa por este logro. ¡Que lo disfrutes y seguí participando!", FontTypeNames.FONTTYPE_EXP)
120                 UserList(UserIndex).UserLogros = UserList(UserIndex).UserLogros + 1
122                 Call WriteRecompensas(UserIndex)
124                 Call WriteTrofeoToggleOff(UserIndex)
                    Exit Sub
                Else
126                 Call WriteConsoleMsg(UserIndex, "Aún no has terminado este logro ¡Continua luchando!", FontTypeNames.FONTTYPE_EXP)
                    Exit Sub

                End If

128         Case 3

130             If UserList(UserIndex).Stats.ELV >= LevelLogros(UserList(UserIndex).LevelLogros + 1).cant Then
    
132                 Call WriteConsoleMsg(UserIndex, "Acá tenes tu recompensa por este logro. ¡Que lo disfrutes y seguí participando!", FontTypeNames.FONTTYPE_EXP)
134                 UserList(UserIndex).LevelLogros = UserList(UserIndex).LevelLogros + 1
136                 Call WriteRecompensas(UserIndex)
138                 Call WriteTrofeoToggleOff(UserIndex)
                    Exit Sub
                Else
140                 Call WriteConsoleMsg(UserIndex, "Aún no has terminado este logro ¡Continua luchando!", FontTypeNames.FONTTYPE_EXP)
                    Exit Sub

                End If

        End Select

        
        Exit Sub

EntregarRecompensas_Err:
142     Call TraceError(Err.Number, Err.Description, "MoLogros.EntregarRecompensas", Erl)
144
        
End Sub
