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
                                                   
    Dim FindPos As Byte, LoopC As Long, InRank As Byte, backup As tUserRanking

    InRank = isRank(UserList(UserIndex).name, Tipo) ''Verificamos si esta en el ranking y si esta, en que posicion.

    With Rankings(Tipo)

        If InRank > 1 Then  ''Si no es el primero del ranking
            .user(InRank).Value = Value ''Actualizamos el valor ANTES de reordenarlo

            Do While .user(InRank - 1).Value < Value ''Mientras que el usuario que esta arriba en el ranking tenga menos puntos, va a seguir subiendo de posiciones.
                backup = .user(InRank) ''Guardamos el personaje en cuestion ya que vamos a cambiar los datos
                .user(InRank) = .user(InRank - 1) ''Reemplazamos al personaje, por el que estaba un puesto arriba
                .user(InRank - 1) = backup ''En ese puesto, ponemos el personaje que ascendio un puesto
                InRank = InRank - 1 ''Actualizamos la variable temporal que esta guardando la posicion de el pj que esta actualizando su posicion

                If InRank = 1 Then ''Si llego al primer puesto

                    Exit Do ''Salimos, ya no puede seguir subiendo.

                End If

            Loop
        ElseIf InRank = 1 Then ''Si es el primero del ranking
            .user(InRank).Value = Value ''Actualizamos el valor.
        ElseIf InRank = 0 Then ''Si no esta en el ranking

            For LoopC = 10 To 1 Step -1 ''Recorremos todos los usuarios del ranking a ver si puede entrar

                If .user(LoopC).Value < Value Then ''El valor del personaje es mayor al del puesto del ranking?
                    FindPos = LoopC ''Encontramos una posicion, pero seguimos el bucle para ver si puede seguir subiendo.

                End If

            Next LoopC

            If FindPos > 0 Then ''Encontro alguna posicion?
                If Not FindPos = 10 Then ''Excepto que sea el ultimo puesto, tenemos que reordenar el ranking.

                    For LoopC = 10 To FindPos + 1 Step -1 ''Recorremos desde el ultimo puesto hasta un puesto abajo de donde va a ingresar el pj
                        .user(LoopC).Nick = .user(LoopC - 1).Nick ''Actualizamos los valores para dejarle el lugar
                        .user(LoopC).Value = .user(LoopC - 1).Value
                    Next LoopC

                End If

                .user(FindPos).Nick = UCase$(UserList(UserIndex).name) ''Ingresa el pj al ranking en el puesto que encontramos.
                .user(FindPos).Value = Value

            End If

        End If

    End With

    Call GuardarRanking

End Sub

Private Function isRank(ByVal Nick As String, ByVal Tipo As eRankings) As Byte

    'Funcion que devuelve el puesto del ranking si es que esta en el mismo, devuelve 0 si no esta en el ranking.
    Dim x As Long

    For x = 1 To 10 ''Recorremos el ranking

        If UCase$(Nick) = UCase$(Rankings(Tipo).user(x).Nick) Then ''Esta en este puesto?
            isRank = CByte(x) ''Devolvemos el valor que encontramos

            Exit Function ''Salimos, ya no hay nada mas que hacer.

        End If

        ''No esta en este puesto, seguimos buscando
    Next x

    ''No esta en el ranking, devolvemos 0 como valor.
    isRank = 0

End Function

Public Sub GuardarRanking()

    Dim Tipo     As Long

    Dim x        As Long

    Dim rankfile As String

    rankfile = App.Path & "\Ranking.ini"

    For Tipo = 1 To NumRanks

        With Rankings(Tipo)

            For x = 1 To 10 ''Recorremos el ranking
                Call WriteVar(rankfile, Tipo, x, .user(x).Nick & "*" & .user(x).Value)
            Next x

        End With

    Next Tipo

End Sub

Public Sub CargarRanking()

    Dim Tipo     As Long

    Dim x        As Long

    Dim rankfile As String

    rankfile = App.Path & "\Ranking.ini"

    Dim tmpstring As String

    For Tipo = 1 To NumRanks

        With Rankings(Tipo)

            For x = 1 To 10 ''Recorremos el ranking
                tmpstring = GetVar(rankfile, Tipo, x)
                .user(x).Nick = ReadField(1, tmpstring, Asc("*"))
                .user(x).Value = ReadField(2, tmpstring, Asc("*"))
            Next x

        End With

    Next Tipo

End Sub
