Attribute VB_Name = "ModRanking"
'Author: Nhelk(Santiago)
'Date: 21/11/2014

Option Explicit

Public Type tUserRanking '' Estructura de datos para cada puesto del ranking

    Nick As String
    Value As Long

End Type

Private Type tRanking '' Estructura de 10 usuarios, cada tipo de ranking esta declarado con esta estructura

    user(1 To 10) As tUserRanking

End Type

Public Enum eRankings '' Cada ranking tiene un identificador.

    Battle = 1
    asesino = 2
    Muertes = 3
    NPCs = 4

End Enum

Public Const NumRanks          As Byte = 1 ''Cuantos tipos de rankings existen (r1vs1, r2vs2, nivel, etc)

Public Rankings(1 To NumRanks) As tRanking ''Array con todos los tipos de ranking, para identificar cada uno se usa el enum eRankings

Public Sub CheckRanking(ByVal Tipo As eRankings, ByVal UserIndex As Integer, ByVal Value As Long)
        ''CheckRanking
        ''Cada vez que se cambia algun valor de cualquier usuario, se verifica si puede ingresar al ranking, _
          cambiar de posicion o solamente actualizar el valor.
        
        On Error GoTo CheckRanking_Err
        
                                                   
        Dim FindPos As Byte, LoopC As Long, InRank As Byte, backup As tUserRanking

100     InRank = isRank(UserList(UserIndex).name, Tipo) ''Verificamos si esta en el ranking y si esta, en que posicion.

102     With Rankings(Tipo)

104         If InRank > 1 Then  ''Si no es el primero del ranking
106             .user(InRank).Value = Value ''Actualizamos el valor ANTES de reordenarlo

108             Do While .user(InRank - 1).Value < Value ''Mientras que el usuario que esta arriba en el ranking tenga menos puntos, va a seguir subiendo de posiciones.
110                 backup = .user(InRank) ''Guardamos el personaje en cuestion ya que vamos a cambiar los datos
112                 .user(InRank) = .user(InRank - 1) ''Reemplazamos al personaje, por el que estaba un puesto arriba
114                 .user(InRank - 1) = backup ''En ese puesto, ponemos el personaje que ascendio un puesto
116                 InRank = InRank - 1 ''Actualizamos la variable temporal que esta guardando la posicion de el pj que esta actualizando su posicion

118                 If InRank = 1 Then ''Si llego al primer puesto

                        Exit Do ''Salimos, ya no puede seguir subiendo.

                    End If

                Loop
120         ElseIf InRank = 1 Then ''Si es el primero del ranking
122             .user(InRank).Value = Value ''Actualizamos el valor.
124         ElseIf InRank = 0 Then ''Si no esta en el ranking

126             For LoopC = 10 To 1 Step -1 ''Recorremos todos los usuarios del ranking a ver si puede entrar

128                 If .user(LoopC).Value < Value Then ''El valor del personaje es mayor al del puesto del ranking?
130                     FindPos = LoopC ''Encontramos una posicion, pero seguimos el bucle para ver si puede seguir subiendo.

                    End If

132             Next LoopC

134             If FindPos > 0 Then ''Encontro alguna posicion?
136                 If Not FindPos = 10 Then ''Excepto que sea el ultimo puesto, tenemos que reordenar el ranking.

138                     For LoopC = 10 To FindPos + 1 Step -1 ''Recorremos desde el ultimo puesto hasta un puesto abajo de donde va a ingresar el pj
140                         .user(LoopC).Nick = .user(LoopC - 1).Nick ''Actualizamos los valores para dejarle el lugar
142                         .user(LoopC).Value = .user(LoopC - 1).Value
144                     Next LoopC

                    End If

146                 .user(FindPos).Nick = UCase$(UserList(UserIndex).name) ''Ingresa el pj al ranking en el puesto que encontramos.
148                 .user(FindPos).Value = Value

                End If

            End If

        End With

150     Call GuardarRanking

        
        Exit Sub

CheckRanking_Err:
        Call RegistrarError(Err.Number, Err.description, "ModRanking.CheckRanking", Erl)
        Resume Next
        
End Sub

Private Function isRank(ByVal Nick As String, ByVal Tipo As eRankings) As Byte
        
        On Error GoTo isRank_Err
        

        'Funcion que devuelve el puesto del ranking si es que esta en el mismo, devuelve 0 si no esta en el ranking.
        Dim x As Long

100     For x = 1 To 10 ''Recorremos el ranking

102         If UCase$(Nick) = UCase$(Rankings(Tipo).user(x).Nick) Then ''Esta en este puesto?
104             isRank = CByte(x) ''Devolvemos el valor que encontramos

                Exit Function ''Salimos, ya no hay nada mas que hacer.

            End If

            ''No esta en este puesto, seguimos buscando
106     Next x

        ''No esta en el ranking, devolvemos 0 como valor.
108     isRank = 0

        
        Exit Function

isRank_Err:
        Call RegistrarError(Err.Number, Err.description, "ModRanking.isRank", Erl)
        Resume Next
        
End Function

Public Sub GuardarRanking()
        
        On Error GoTo GuardarRanking_Err
        

        Dim Tipo     As Long

        Dim x        As Long

        Dim rankfile As String

100     rankfile = App.Path & "\Ranking.ini"

102     For Tipo = 1 To NumRanks

104         With Rankings(Tipo)

106             For x = 1 To 10 ''Recorremos el ranking
108                 Call WriteVar(rankfile, Tipo, x, .user(x).Nick & "*" & .user(x).Value)
110             Next x

            End With

112     Next Tipo

        
        Exit Sub

GuardarRanking_Err:
        Call RegistrarError(Err.Number, Err.description, "ModRanking.GuardarRanking", Erl)
        Resume Next
        
End Sub

Public Sub CargarRanking()
        
        On Error GoTo CargarRanking_Err
        

        Dim Tipo     As Long

        Dim x        As Long

        Dim rankfile As String

100     rankfile = App.Path & "\Ranking.ini"

        Dim tmpstring As String

102     For Tipo = 1 To NumRanks

104         With Rankings(Tipo)

106             For x = 1 To 10 ''Recorremos el ranking
108                 tmpstring = GetVar(rankfile, Tipo, x)
110                 .user(x).Nick = ReadField(1, tmpstring, Asc("*"))
112                 .user(x).Value = ReadField(2, tmpstring, Asc("*"))
114             Next x

            End With

116     Next Tipo

        
        Exit Sub

CargarRanking_Err:
        Call RegistrarError(Err.Number, Err.description, "ModRanking.CargarRanking", Erl)
        Resume Next
        
End Sub
