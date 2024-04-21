Attribute VB_Name = "TCP"
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit

Sub DarCuerpo(ByVal UserIndex As Integer)
        
        On Error GoTo DarCuerpo_Err
        

        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 14/03/2007
        'Elije una cabeza para el usuario y le da un body
        '*************************************************
        Dim NewBody    As Integer

        Dim UserRaza   As Byte

        Dim UserGenero As Byte

100     UserGenero = UserList(UserIndex).genero
102     UserRaza = UserList(UserIndex).raza

104     Select Case UserGenero

            Case e_Genero.Hombre

106             Select Case UserRaza

                    Case e_Raza.Humano
108                     NewBody = 1

110                 Case e_Raza.Elfo
112                     NewBody = 2

114                 Case e_Raza.Drow
116                     NewBody = 3

118                 Case e_Raza.Enano
120                     NewBody = 300

122                 Case e_Raza.Gnomo
124                     NewBody = 300

126                 Case e_Raza.Orco
128                     NewBody = 582

                End Select

130         Case e_Genero.Mujer

132             Select Case UserRaza

                    Case e_Raza.Humano
134                     NewBody = 1

136                 Case e_Raza.Elfo
138                     NewBody = 2

140                 Case e_Raza.Drow
142                     NewBody = 3

144                 Case e_Raza.Gnomo
146                     NewBody = 300

148                 Case e_Raza.Enano
150                     NewBody = 300

152                 Case e_Raza.Orco
154                     NewBody = 581

                End Select

        End Select

156     UserList(UserIndex).Char.Body = NewBody

        
        Exit Sub

DarCuerpo_Err:
158     Call TraceError(Err.Number, Err.Description, "TCP.DarCuerpo", Erl)

        
End Sub

Sub RellenarInventario(ByVal UserIndex As String)
        
        On Error GoTo RellenarInventario_Err
        

100     With UserList(UserIndex)
        
            Dim NumItems As Integer

102         NumItems = 1
    
            ' Todos reciben pociones rojas
104         .invent.Object(NumItems).ObjIndex = 4335 'Pocion Roja
106         .invent.Object(NumItems).amount = 350
108         NumItems = NumItems + 1
        
            ' Magicas puras reciben más azules
110         Select Case .clase

            Case e_Class.Mage, e_Class.Druid
                 .invent.Object(NumItems).ObjIndex = 4336 ' Pocion Azul
                 .invent.Object(NumItems).amount = 550
                 NumItems = NumItems + 1

                    Case e_Class.Bard, e_Class.Cleric
                 .invent.Object(NumItems).ObjIndex = 4336 ' Pocion Azul
                 .invent.Object(NumItems).amount = 450
                 NumItems = NumItems + 1

                    Case e_Class.Paladin, e_Class.Assasin, e_Class.Bandit
                 .invent.Object(NumItems).ObjIndex = 4336 ' Pocion Azul
                 .invent.Object(NumItems).amount = 350
                 NumItems = NumItems + 1
         
            End Select


     

            ' Hechizos
126         Select Case .clase

                 Case e_Class.Mage, e_Class.Cleric, e_Class.Druid, e_Class.Bard, e_Class.Paladin, e_Class.Bandit, e_Class.Assasin
128                 .Stats.UserHechizos(1) = 291 ' Onda mágica


            End Select
        
            ' Pociones amarillas y verdes
134         Select Case .clase

            Case e_Class.Assasin, e_Class.Bard, e_Class.Cleric, e_Class.Hunter, e_Class.Paladin, e_Class.Trabajador, e_Class.Warrior, e_Class.Bandit, e_Class.Pirat, e_Class.Thief

                 .invent.Object(NumItems).ObjIndex = 4337 ' Pocion Amarilla
                 .invent.Object(NumItems).amount = 100
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 4338 ' Pocion Verde
                 .invent.Object(NumItems).amount = 100
                 NumItems = NumItems + 1

             Case e_Class.Mage, e_Class.Druid
                 .invent.Object(NumItems).ObjIndex = 4337 ' Pocion Amarilla
                 .invent.Object(NumItems).amount = 60
                 NumItems = NumItems + 1


            End Select
            
            ' Poción violeta
148         .invent.Object(NumItems).ObjIndex = 4334 ' Pocion violeta
150         .invent.Object(NumItems).amount = 50
152         NumItems = NumItems + 1
        
            ' Armas
154         Select Case .clase
                Case e_Class.Cleric
                 .invent.Object(NumItems).ObjIndex = 3686 ' Daga del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1

                 .invent.Object(NumItems).ObjIndex = 3487 ' Espada del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1

                 .invent.Object(NumItems).ObjIndex = 3488 ' Escudo de Principiante
                .invent.Object(NumItems).amount = 1
                NumItems = NumItems + 1

                 .invent.Object(NumItems).ObjIndex = 3489 ' Casco de Principiante
                 .invent.Object(NumItems).amount = 1
                NumItems = NumItems + 1

                .invent.Object(NumItems).ObjIndex = 3490 ' Anillo del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1

               Case e_Class.Paladin
                 .invent.Object(NumItems).ObjIndex = 3487 ' Espada del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3488 ' Escudo de Principiante
                .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                .invent.Object(NumItems).ObjIndex = 3489 ' Casco de Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3490 ' Anillo del Principiante
                .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1

             Case e_Class.Hunter
                 .invent.Object(NumItems).ObjIndex = 3491 ' Arco del Principiante
                 .invent.Object(NumItems).amount = 1
                NumItems = NumItems + 1

                 .invent.Object(NumItems).ObjIndex = 3492 ' Flecha del Principiante
                 .invent.Object(NumItems).amount = 650
                 NumItems = NumItems + 1

                 .invent.Object(NumItems).ObjIndex = 3489  ' Casco de Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1

                 .invent.Object(NumItems).ObjIndex = 3490 ' Anillo del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 
                  .invent.Object(NumItems).objIndex = 3488 ' Escudo de Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1

                Case e_Class.Trabajador
                 .invent.Object(NumItems).ObjIndex = 3487 ' Espada del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3488 ' Escudo de Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3489 ' Casco de Principiante
                .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3491 ' Arco del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3492 ' Flecha del Principiante
                 .invent.Object(NumItems).amount = 300
                 NumItems = NumItems + 1

                Case e_Class.Pirat
                 .invent.Object(NumItems).ObjIndex = 3487 ' Espada del Principiante
                .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3488 ' Escudo de Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3489 ' Casco de Principiante
                 .invent.Object(NumItems).amount = 1
                NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3490 ' Anillo del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3497 ' Pistola del Principiante
                .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3498 ' Balas del Principiante
                 .invent.Object(NumItems).amount = 350
                NumItems = NumItems + 1

                Case e_Class.Warrior
                 .invent.Object(NumItems).ObjIndex = 3487 ' Espada del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3488 ' Escudo de Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3489 ' Casco de Principiante
                .invent.Object(NumItems).amount = 1
                NumItems = NumItems + 1
                .invent.Object(NumItems).ObjIndex = 3491 ' Arco del Principiante
                .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3492 ' Flecha del Principiante
                 .invent.Object(NumItems).amount = 300
                 NumItems = NumItems + 1

            Case e_Class.Thief
                 .invent.Object(NumItems).ObjIndex = 3686 ' Daga del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 1353 ' Nudillos del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3488 ' Escudo de Principiante
                .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                .invent.Object(NumItems).ObjIndex = 3489  ' Casco de Principiante
                .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 

             Case e_Class.Bandit
                 .invent.Object(NumItems).ObjIndex = 1353 ' Nudillos del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3487 ' Espada del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3488 ' Escudo de Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3489 ' Casco de Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3490 ' Anillo del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1

             Case e_Class.Mage
                 .invent.Object(NumItems).ObjIndex = 3495 ' Bastón del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3493 ' Sombrero del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1

             Case e_Class.Assasin
                 .invent.Object(NumItems).ObjIndex = 3686 ' Daga del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                .invent.Object(NumItems).ObjIndex = 3488 ' Escudo de Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3489 ' Casco de Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3490 ' Anillo del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3487 ' Espada del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1

             Case e_Class.Druid
                 .invent.Object(NumItems).ObjIndex = 3686 ' Daga del Principiante
                 .invent.Object(NumItems).amount = 1
                NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3494 ' Flauta del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3504  'Casco de Lobo (Resistencia Magica 1)
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1

             Case e_Class.Bard
                 .invent.Object(NumItems).ObjIndex = 3686 ' Daga del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3488 ' Escudo de Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3489 ' Casco de Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3490 ' Anillo del Principiante
                 .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1
                 .invent.Object(NumItems).ObjIndex = 3496 ' Laúd del Principiante
                .invent.Object(NumItems).amount = 1
                 NumItems = NumItems + 1


                  End Select
        
            
220             If .raza = Enano Or .raza = Gnomo Then
222                Select Case .clase
                          Case e_Class.Trabajador, e_Class.Thief, e_Class.Paladin, e_Class.Cleric, e_Class.Assasin, e_Class.Bandit, e_Class.Pirat, e_Class.Warrior, e_Class.Hunter
                          .invent.Object(NumItems).objIndex = 3500 ' Armadura de Principiante
                          Case e_Class.Mage, e_Class.Druid, e_Class.Bard
                         .invent.Object(NumItems).objIndex = 3502 ' Túnica del Principiante
                   End Select


                 
                Else
                Select Case .clase
                         Case e_Class.Trabajador, e_Class.Thief, e_Class.Paladin, e_Class.Cleric, e_Class.Assasin, e_Class.Bandit, e_Class.Pirat, e_Class.Warrior, e_Class.Hunter
                          .invent.Object(NumItems).ObjIndex = 3500 ' Armadura de Principiante
                         Case e_Class.Mage, e_Class.Druid, e_Class.Bard
                          .invent.Object(NumItems).ObjIndex = 3502 ' Túnica del Principiante
                   End Select
                    End If

            
            .Invent.Object(NumItems).Equipped = 0
            Call EquiparInvItem(UserIndex, NumItems)
                        
232         .Invent.Object(NumItems).amount = 1
234         .Invent.Object(NumItems).Equipped = 1
236         .Invent.ArmourEqpSlot = NumItems
238         .Invent.ArmourEqpObjIndex = .Invent.Object(NumItems).ObjIndex
240          NumItems = NumItems + 1

            ' Animación según raza
242          .Char.Body = ObtenerRopaje(UserIndex, ObjData(.Invent.ArmourEqpObjIndex))
        
            ' Comida y bebida
244         .invent.Object(NumItems).ObjIndex = 3684 ' Manzana
246         .invent.Object(NumItems).amount = 50
248         NumItems = NumItems + 1

250         .invent.Object(NumItems).ObjIndex = 3685 ' Agua
252         .invent.Object(NumItems).amount = 50
254         NumItems = NumItems + 1

            ' Seteo la cantidad de items
256         .Invent.NroItems = NumItems
            
            .flags.ModificoInventario = True
            .flags.ModificoHechizos = True
            
        End With
   
        
        Exit Sub

RellenarInventario_Err:
258     Call TraceError(Err.Number, Err.Description, "TCP.RellenarInventario", Erl)

        
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
        
        On Error GoTo AsciiValidos_Err
        

        Dim car As Byte

        Dim i   As Integer

100     cad = LCase$(cad)

102     For i = 1 To Len(cad)
104         car = Asc(mid$(cad, i, 1))
    
106         If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
108             AsciiValidos = False
                Exit Function

            End If
    
110     Next i

112     AsciiValidos = True

        
        Exit Function

AsciiValidos_Err:
114     Call TraceError(Err.Number, Err.Description, "TCP.AsciiValidos", Erl)

        
End Function

Function DescripcionValida(ByVal cad As String) As Boolean
        
        On Error GoTo AsciiValidos_Err
        

        Dim car As Byte

        Dim i   As Integer

100     cad = LCase$(cad)

102     For i = 1 To Len(cad)
104         car = Asc(mid$(cad, i, 1))
    
106         If car < 32 Or car >= 126 Then
108             DescripcionValida = False
                Exit Function

            End If
    
110     Next i

112     DescripcionValida = True

        
        Exit Function

AsciiValidos_Err:
114     Call TraceError(Err.Number, Err.Description, "TCP.DescripcionValida", Erl)

        
End Function

Function Numeric(ByVal cad As String) As Boolean
        
        On Error GoTo Numeric_Err
        

        Dim car As Byte

        Dim i   As Integer

100     cad = LCase$(cad)

102     For i = 1 To Len(cad)
104         car = Asc(mid$(cad, i, 1))
    
106         If (car < 48 Or car > 57) Then
108             Numeric = False
                Exit Function

            End If
    
110     Next i

112     Numeric = True

        
        Exit Function

Numeric_Err:
114     Call TraceError(Err.Number, Err.Description, "TCP.Numeric", Erl)

        
End Function

Function NombrePermitido(ByVal nombre As String) As Boolean
        
        On Error GoTo NombrePermitido_Err
        

        Dim i As Integer

100     For i = 1 To UBound(ForbidenNames)

102         If LCase$(nombre) = ForbidenNames(i) Then
104             NombrePermitido = False
                Exit Function

            End If

106     Next i

108     NombrePermitido = True

        
        Exit Function

NombrePermitido_Err:
110     Call TraceError(Err.Number, Err.Description, "TCP.NombrePermitido", Erl)

        
End Function

Function Validate_Skills(ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo Validate_Skills_Err
        

        Dim LoopC As Integer

100     For LoopC = 1 To NUMSKILLS

102         If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then
                Exit Function

104             If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100

            End If

106     Next LoopC

108     Validate_Skills = True
    
        
        Exit Function

Validate_Skills_Err:
110     Call TraceError(Err.Number, Err.Description, "TCP.Validate_Skills", Erl)

        
End Function

Function ConnectNewUser(ByVal userindex As Integer, ByRef name As String, ByVal UserRaza As e_Raza, ByVal UserSexo As e_Genero, ByVal UserClase As e_Class, ByVal Head As Integer, ByVal Hogar As e_Ciudad) As Boolean
        
        On Error GoTo ConnectNewUser_Err
        
100     With UserList(UserIndex)
        
            Dim LoopC As Long
        
102         If .flags.UserLogged Then
104             Call LogSecurity("El usuario " & .name & " ha intentado crear a " & name & " desde la IP " & .ConnectionDetails.IP)
106             Call CloseSocketSL(UserIndex)
108             Call Cerrar_Usuario(UserIndex)
                Exit Function
            End If
            
            ' Nombre válido
            If Not ValidarNombre(name) Then
                Call LogSecurity("ValidarNombre failed in ConnectNewUser for " & name & " desde la IP " & .ConnectionDetails.IP)
                Call CloseSocketSL(UserIndex)
                Exit Function
            End If
            
112         If Not NombrePermitido(name) Then
114             Call WriteShowMessageBox(UserIndex, "El nombre no está permitido.")
                Exit Function
            End If
    
            '¿Existe el personaje?
116         If PersonajeExiste(name) Then
118             Call WriteShowMessageBox(UserIndex, "Ya existe el personaje.")
                Exit Function
            End If
            
            ' Raza válida
120         If UserRaza <= 0 Or UserRaza > NUMRAZAS Then Exit Function
            
            ' Género válido
122         If UserSexo < Hombre Or UserSexo > Mujer Then Exit Function
            
            ' Ciudad válida
124         If Hogar <= 0 Or Hogar > NUMCIUDADES Then Exit Function
            
            ' Cabeza válida
126         If Not ValidarCabeza(UserRaza, UserSexo, Head) Then Exit Function
            
            'Prevenimos algun bug con dados inválidos
128         'If .Stats.UserAtributos(e_Atributos.Fuerza) = 0 Then Exit Function
        
130         .Stats.UserAtributos(e_Atributos.Fuerza) = 18 + ModRaza(UserRaza).Fuerza
132         .Stats.UserAtributos(e_Atributos.Agilidad) = 18 + ModRaza(UserRaza).Agilidad
134         .Stats.UserAtributos(e_Atributos.Inteligencia) = 18 + ModRaza(UserRaza).Inteligencia
136         .Stats.UserAtributos(e_Atributos.Constitucion) = 18 + ModRaza(UserRaza).Constitucion
138         .Stats.UserAtributos(e_Atributos.Carisma) = 18 + ModRaza(UserRaza).Carisma
            
            .Stats.UserAtributosBackUP(e_Atributos.Fuerza) = .Stats.UserAtributos(e_Atributos.Fuerza)
            .Stats.UserAtributosBackUP(e_Atributos.Agilidad) = .Stats.UserAtributos(e_Atributos.Agilidad)
            .Stats.UserAtributosBackUP(e_Atributos.Inteligencia) = .Stats.UserAtributos(e_Atributos.Inteligencia)
            .Stats.UserAtributosBackUP(e_Atributos.Constitucion) = .Stats.UserAtributos(e_Atributos.Constitucion)
            .Stats.UserAtributosBackUP(e_Atributos.Carisma) = .Stats.UserAtributos(e_Atributos.Carisma)
            
140         .flags.Muerto = 0
142         .flags.Escondido = 0
    
144         .flags.Casado = 0
146         .flags.SpouseId = 0
    
148         .name = name

150         .clase = Min(max(0, UserClase), NUMCLASES)
152         .raza = UserRaza
        
154         .Char.Head = Head
        
156         .genero = UserSexo
158         .Hogar = Hogar
        
            '%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%
160         .Stats.SkillPts = 10
        
162         .Char.Heading = e_Heading.SOUTH
        
164         Call DarCuerpo(UserIndex) 'Ladder REVISAR
        
166         .OrigChar = .Char
    
168         Call ClearClothes(.char)

            ' WyroX: Vida inicial
174         .Stats.MaxHp = .Stats.UserAtributos(e_Atributos.Constitucion)
176         .Stats.MinHp = .Stats.MaxHp
177         .Stats.Shield = 0

            ' WyroX: Maná inicial
178         .Stats.MaxMAN = .Stats.UserAtributos(e_Atributos.Inteligencia) * ModClase(.clase).ManaInicial
180         .Stats.MinMAN = .Stats.MaxMAN
        
            Dim MiInt As Integer
182         MiInt = RandomNumber(1, .Stats.UserAtributos(e_Atributos.Agilidad) \ 6)
    
184         If MiInt = 1 Then MiInt = 2
        
186         .Stats.MaxSta = 20 * MiInt
188         .Stats.MinSta = 20 * MiInt
        
190         .Stats.MaxAGU = 100
192         .Stats.MinAGU = 100
        
194         .Stats.MaxHam = 100
196         .Stats.MinHam = 100
    
202         .flags.VecesQueMoriste = 0
204         .flags.Montado = 0
    
206         .Stats.MaxHit = 2
208         .Stats.MinHIT = 1
        
210         .Stats.GLD = 0
        
212         .Stats.Exp = 0
214         .Stats.ELV = 1
        
216         Call RellenarInventario(UserIndex)
    
            #If ConUpTime Then
218             .LogOnTime = Now
220             .UpTime = 0
            #End If
            Call ResetCd(UserList(UserIndex))
            'Valores Default de facciones al Activar nuevo usuario
222         Call ResetFacciones(UserIndex)
        
224         .Faccion.Status = 1
        
226         .ChatCombate = 1
228         .ChatGlobal = 1
            
            Call UpdateUserTelemetryKey(UserIndex)
            
            Select Case .Hogar
                Case e_Ciudad.cUllathorpe
                    .Pos.map = 1
                    .Pos.X = 56
                    .Pos.Y = 44
                Case e_Ciudad.cArghal
                    .Pos.map = 151
                    .Pos.X = 46
                    .Pos.Y = 34
                Case e_Ciudad.cNix
                    .Pos.map = 34
                    .Pos.X = 40
                    .Pos.Y = 86
                Case e_Ciudad.cLindos
                    .Pos.map = 62
                    .Pos.X = 62
                    .Pos.Y = 44
                Case e_Ciudad.cBanderbill
                    .Pos.map = 59
                    .Pos.X = 54
                    .Pos.Y = 42
                Case e_Ciudad.cArkhein
                    .Pos.map = 196
                    .Pos.X = 49
                    .Pos.Y = 64
            End Select
        
254         UltimoChar = UCase$(name)
        
256         Call SaveNewUser(UserIndex)
    
258         ConnectNewUser = True
#If PYMMO = 1 Then
260         Call ConnectUser(userindex, name, True)
#ElseIf PYMMO = 0 Then
260         Call ConnectUser(userindex, name, False)
#End If
        End With
        
        Exit Function

ConnectNewUser_Err:
262     Call TraceError(Err.Number, Err.Description, "TCP.ConnectNewUser", Erl)

        
End Function

Sub CloseSocket(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

102     If UserIndex = LastUser Then

104         Do Until UserList(LastUser).flags.UserLogged
106             LastUser = LastUser - 1
108             If LastUser < 1 Then Exit Do
            Loop

        End If
    
110     With UserList(UserIndex)

112         If .ConnectionDetails.ConnIDValida Then Call CloseSocketSL(UserIndex)
    
            'mato los comercios seguros
116         If IsValidUserRef(.ComUsu.DestUsu) Then
        
118             If UserList(.ComUsu.DestUsu.ArrayIndex).flags.UserLogged Then
            
120                 If UserList(.ComUsu.DestUsu.ArrayIndex).ComUsu.DestUsu.ArrayIndex = userIndex Then
                
122                     Call WriteConsoleMsg(.ComUsu.DestUsu.ArrayIndex, "Comercio cancelado por el otro usuario", e_FontTypeNames.FONTTYPE_TALK)
124                     Call FinComerciarUsu(.ComUsu.DestUsu.ArrayIndex)
                    
                    End If
    
                End If
    
            End If
    
128         If .flags.UserLogged Then
130             Call CloseUser(UserIndex)
        
132             If NumUsers > 0 Then NumUsers = NumUsers - 1
        
            Else
136             Call ResetUserSlot(UserIndex)
            End If
    
140         .ConnectionDetails.ConnIDValida = False

        End With
    

        Exit Sub

ErrHandler:

144     UserList(UserIndex).ConnectionDetails.ConnIDValida = False
146     Call ResetUserSlot(UserIndex)
148     Call TraceError(Err.Number, Err.Description, "TCP.CloseSocket", Erl)


End Sub

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal UserIndex As Integer)
        
        On Error GoTo CloseSocketSL_Err

100     If UserList(UserIndex).ConnectionDetails.ConnIDValida Then
102         Call modNetwork.Kick(UserList(UserIndex).ConnectionDetails.ConnID)

106         UserList(UserIndex).ConnectionDetails.ConnIDValida = False
        End If
        
        Exit Sub

CloseSocketSL_Err:
108     Call TraceError(Err.Number, Err.Description, "TCP.CloseSocketSL", Erl)

        
End Sub

Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean
        
        On Error GoTo EstaPCarea_Err
        

        Dim X As Integer, y As Integer

100     For y = UserList(Index).Pos.y - MinYBorder + 1 To UserList(Index).Pos.y + MinYBorder - 1
102         For X = UserList(Index).Pos.X - MinXBorder + 1 To UserList(Index).Pos.X + MinXBorder - 1

104             If MapData(UserList(Index).Pos.map, X, y).UserIndex = Index2 Then
106                 EstaPCarea = True
                    Exit Function

                End If
        
108         Next X
110     Next y

112     EstaPCarea = False

        
        Exit Function

EstaPCarea_Err:
114     Call TraceError(Err.Number, Err.Description, "TCP.EstaPCarea", Erl)

        
End Function

Function HayPCarea(ByVal map As Integer, ByVal X As Integer, ByVal y As Integer, ByVal ignoreUserMuerto As Boolean) As Boolean
        
        On Error GoTo HayPCarea_Err
        

        Dim tX As Integer, tY As Integer

100     For tY = y - MinYBorder + 1 To y + MinYBorder - 1
102         For tX = X - MinXBorder + 1 To X + MinXBorder - 1

104             If InMapBounds(map, tX, tY) Then
106                 If MapData(map, tX, tY).UserIndex > 0 Then
                        If Not ignoreUserMuerto Then
                            HayPCarea = True
                        Else
                            If UserList(MapData(map, tX, tY).userindex).flags.Muerto = 0 Then HayPCarea = True
                        End If
108                     Exit Function
                    End If

                End If

            Next
        Next

110     HayPCarea = False

        
        Exit Function

HayPCarea_Err:
112     Call TraceError(Err.Number, Err.Description, "TCP.HayPCarea", Erl)

        
End Function

Function HayOBJarea(Pos As t_WorldPos, ObjIndex As Integer) As Boolean
        
        On Error GoTo HayOBJarea_Err
        

        Dim X As Integer, y As Integer

100     For y = Pos.y - MinYBorder + 1 To Pos.y + MinYBorder - 1
102         For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1

104             If MapData(Pos.map, X, y).ObjInfo.ObjIndex = ObjIndex Then
106                 HayOBJarea = True
                    Exit Function

                End If
        
108         Next X
110     Next y

112     HayOBJarea = False

        
        Exit Function

HayOBJarea_Err:
114     Call TraceError(Err.Number, Err.Description, "TCP.HayOBJarea", Erl)

        
End Function

Function ValidateChr(ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo ValidateChr_Err
        

100     ValidateChr = UserList(UserIndex).Char.Body <> 0 And Validate_Skills(UserIndex)

        
        Exit Function

ValidateChr_Err:
102     Call TraceError(Err.Number, Err.Description, "TCP.ValidateChr", Erl)

        
End Function

Function EntrarCuenta(ByVal UserIndex As Integer, ByVal CuentaEmail As String, ByVal MD5 As String) As Boolean
        
        On Error GoTo EntrarCuenta_Err
        
        Dim adminIdx As Integer
        Dim laCuentaEsDeAdmin As Boolean
        
100     If ServerSoloGMs > 0 Then
102         laCuentaEsDeAdmin = False

104         For adminIdx = 0 To AdministratorAccounts.Count - 1
                ' Si el e-mail está declarado junto al nick de la cuenta donde esta el PJ GM en el Server.ini te dejo entrar.
106             If UCase$(AdministratorAccounts.Items(adminIdx)) = UCase$(CuentaEmail) Then
108                 laCuentaEsDeAdmin = True
                End If
110         Next adminIdx
            
112         If Not laCuentaEsDeAdmin Then
114             Call WriteShowMessageBox(UserIndex, "El servidor se encuentra habilitado solo para administradores por el momento.")
                Exit Function
            End If

        End If

        #If DEBUGGING = 0 Then
124         If LCase$(Md5Cliente) <> LCase$(MD5) Then
126             Call WriteShowMessageBox(UserIndex, "Error al comprobar el cliente del juego, por favor reinstale y vuelva a intentar.")
                Exit Function
            End If
        #End If

128     If Not CheckMailString(CuentaEmail) Then
130         Call WriteShowMessageBox(UserIndex, "Email inválido.")
            Exit Function
        End If
    
132     EntrarCuenta = EnterAccountDatabase(UserIndex, CuentaEmail)
        
        Exit Function

EntrarCuenta_Err:
134     Call TraceError(Err.Number, Err.Description, "TCP.EntrarCuenta", Erl)


        
End Function
Function ConnectUser(ByVal userIndex As Integer, ByRef name As String, Optional ByVal newUser As Boolean = False) As Boolean
On Error GoTo ErrHandler
    ConnectUser = False
    With UserList(userIndex)
        If Not ConnectUser_Check(UserIndex, Name) Then
            Call LogSecurity("ConnectUser_Check " & Name & " failed.")
            Exit Function
        End If
        Call ConnectUser_Prepare(userIndex, name)
        If LoadUser(userIndex) Then
            If ConnectUser_Complete(userIndex, name, newUser) Then
                ConnectUser = True
                Exit Function
            End If
        Else
            Call WriteShowMessageBox(userIndex, "Cannot load character")
            Call CloseSocket(userIndex)
        End If
    End With

    Exit Function

    
ErrHandler:
     Call TraceError(Err.Number, Err.Description, "TCP.ConnectUser", Erl)
     Call WriteShowMessageBox(UserIndex, "El personaje contiene un error. Comuníquese con un miembro del staff.")
     Call CloseSocket(UserIndex)

End Function

Sub SendMOTD(ByVal UserIndex As Integer)
        
        On Error GoTo SendMOTD_Err
        

        Dim j As Long

100     For j = 1 To MaxLines
102         Call WriteConsoleMsg(UserIndex, MOTD(j).texto, e_FontTypeNames.FONTTYPE_EXP)
104     Next j
    
        
        Exit Sub

SendMOTD_Err:
106     Call TraceError(Err.Number, Err.Description, "TCP.SendMOTD", Erl)

        
End Sub

Sub ResetFacciones(ByVal UserIndex As Integer)
        
        On Error GoTo ResetFacciones_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 23/01/2007
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, MatadosIngreso y NextRecompensa.
        '*************************************************
100     With UserList(UserIndex).Faccion
            If .status = e_Facciones.Armada Or .status = e_Facciones.concilio Then
                .status = e_Facciones.Ciudadano
            Else
108             .status = e_Facciones.Criminal
            End If
112         .RecibioArmaduraCaos = 0
114         .RecibioArmaduraReal = 0
120         .RecompensasCaos = 0
122         .RecompensasReal = 0
126         .NivelIngreso = 0
128         .MatadosIngreso = 0
            .FactionScore = 0
        End With

        
        Exit Sub

ResetFacciones_Err:
132     Call TraceError(Err.Number, Err.Description, "TCP.ResetFacciones", Erl)

        
End Sub

Sub ResetContadores(ByVal UserIndex As Integer)
        
        On Error GoTo ResetContadores_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '05/20/2007 Integer - Agregue todas las variables que faltaban.
        '*************************************************
100     With UserList(UserIndex).Counters
102         .AGUACounter = 0
104         .AttackCounter = 0
106         .Ceguera = 0
108         .COMCounter = 0
110         .Estupidez = 0
112         .Frio = 0
114         .HPCounter = 0
116         .IdleCount = 0
118         .Invisibilidad = 0
            .DisabledInvisibility = 0
120         .Paralisis = 0
122         .Inmovilizado = 0
124         .Pasos = 0
126         .Pena = 0
128         .PiqueteC = 0
130         .STACounter = 0
132         .Veneno = 0
134         .Trabajando = 0
            .LastTrabajo = 0
136         .Ocultando = 0
138         .Lava = 0
140         .Maldicion = 0
142         .Saliendo = False
144         .Salir = 0
146         .TiempoOculto = 0
148         .TimerMagiaGolpe = 0
150         .TimerGolpeMagia = 0
152         .TimerLanzarSpell = 0
154         .TimerPuedeAtacar = 0
156         .TimerPuedeUsarArco = 0
158         .TimerPuedeTrabajar = 0
160         .TimerUsar = 0
161         .TimerUsarClick = 0
            'Ladder
162         .Incineracion = 0
            'Ladder
170         .TiempoParaSubastar = 0
172         .TimerPerteneceNpc = 0
174         .TimerPuedeSerAtacado = 0
176         .TiempoDeInmunidad = 0
178         .RepetirMensaje = 0
180         .MensajeGlobal = 0
182         .CuentaRegresiva = -1
184         .SpeedHackCounter = 0
186         .LastStep = 0
188         .TimerBarra = 0
            .LastResetTick = 0
            .CounterGmMessages = 0
            .LastTransferGold = 0
            .controlHechizos.HechizosCasteados = 0
            .controlHechizos.HechizosTotales = 0
            .timeChat = 0
            .timeFx = 0
            .timeGuildChat = 0
        End With

        
        Exit Sub

ResetContadores_Err:
190     Call TraceError(Err.Number, Err.Description, "TCP.ResetContadores", Erl)

        
End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)
        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '*************************************************
        
        On Error GoTo ResetCharInfo_Err
        

100     With UserList(UserIndex).Char
102         .Body = 0
104         .CascoAnim = 0
106         .CharIndex = 0
108         .FX = 0
110         .Head = 0
112         .loops = 0
114         .Heading = 0
116         .loops = 0
118         .ShieldAnim = 0
120         .WeaponAnim = 0
121         .CartAnim = 0
122         .Arma_Aura = ""
124         .Body_Aura = ""
126         .Head_Aura = ""
128         .Otra_Aura = ""
130         .DM_Aura = ""
132         .RM_Aura = ""
134         .Escudo_Aura = ""
136         .ParticulaFx = 0
138         .speeding = 0

        End With

        
        Exit Sub

ResetCharInfo_Err:
144     Call TraceError(Err.Number, Err.Description, "TCP.ResetCharInfo", Erl)

        
End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)
        
        On Error GoTo ResetBasicUserInfo_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        'Agregue que se resetee el maná
        '*************************************************
        Dim LoopC As Integer

100     With UserList(UserIndex)
102         .name = vbNullString
104         .Cuenta = vbNullString
106         .ID = -1
108         .AccountID = -1
110         .Desc = vbNullString
112         .DescRM = vbNullString
114         .Pos.map = 0
116         .Pos.X = 0
118         .Pos.y = 0
120         .ConnectionDetails.IP = vbNullString
122         .clase = 0
124         .Email = vbNullString
126         .genero = 0
128         .Hogar = 0
130         .raza = 0
132         .EmpoCont = 0

154         With .Stats
156             .InventLevel = 0
158             .Banco = 0
160             .ELV = 0
162             .Exp = 0
164             .def = 0
                '.CriminalesMatados = 0
166             .NPCsMuertos = 0
168             .UsuariosMatados = 0
                .PuntosPesca = 0
                .Creditos = 0
170             .SkillPts = 0
172             .GLD = 0
174             .UserAtributos(1) = 0
176             .UserAtributos(2) = 0
178             .UserAtributos(3) = 0
180             .UserAtributos(4) = 0
182             .UserAtributosBackUP(1) = 0
184             .UserAtributosBackUP(2) = 0
186             .UserAtributosBackUP(3) = 0
188             .UserAtributosBackUP(4) = 0
190             .MaxMAN = 0
192             .MinMAN = 0
                .tipoUsuario = e_TipoUsuario.tNormal
            End With
            
194         .NroMascotas = 0
            Dim i As Integer
            For i = LBound(.MascotasType) To UBound(.MascotasType)
                .MascotasType(i) = 0
            Next i
            .LastTransportNetwork.Map = -1
        End With

        
        Exit Sub

ResetBasicUserInfo_Err:
200     Call TraceError(Err.Number, Err.Description, "TCP.ResetBasicUserInfo", Erl)

        
End Sub

Sub ResetGuildInfo(ByVal UserIndex As Integer)
        
        On Error GoTo ResetGuildInfo_Err
        

100     If UserList(UserIndex).EscucheClan > 0 Then
102         Call modGuilds.GMDejaDeEscucharClan(UserIndex, UserList(UserIndex).EscucheClan)
104         UserList(UserIndex).EscucheClan = 0

        End If

106     If UserList(UserIndex).GuildIndex > 0 Then
108         Call modGuilds.m_DesconectarMiembroDelClan(UserIndex, UserList(UserIndex).GuildIndex)

        End If

110     UserList(UserIndex).GuildIndex = 0
    
        
        Exit Sub

ResetGuildInfo_Err:
112     Call TraceError(Err.Number, Err.Description, "TCP.ResetGuildInfo", Erl)

        
End Sub

Sub ResetPacketRateData(ByVal UserIndex As Integer)

        On Error GoTo ResetPacketRateData_Err

        Dim i As Long
        
        With UserList(UserIndex)
        
            For i = 1 To MAX_PACKET_COUNTERS
                .MacroIterations(i) = 0
                .PacketTimers(i) = 0
                .PacketCounters(i) = 0
            Next i
            
        End With
        
        Exit Sub
        
ResetPacketRateData_Err:
282     Call TraceError(Err.Number, Err.Description, "TCP.ResetPacketRateData", Erl)

End Sub

Sub ResetUserFlags(ByVal UserIndex As Integer)
        '*************************************************
        'Author: Unknown
        'Last modified: 03/29/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '*************************************************
        
        On Error GoTo ResetUserFlags_Err
        

100     With UserList(UserIndex).flags
102         .LevelBackup = 0
104         .Comerciando = False
106         .Ban = 0
108         .Escondido = 0
110         .DuracionEfecto = 0
116         .NpcInv = 0
118         .StatsChanged = 0
120         Call ClearNpcRef(.TargetNPC)
122         .TargetNpcTipo = e_NPCType.Comun
124         .TargetObj = 0
126         .TargetObjMap = 0
128         .TargetObjX = 0
130         .TargetObjY = 0
132         Call SetUserRef(.targetUser, 0)
134         .TipoPocion = 0
136         .TomoPocion = False
138         .Descuento = vbNullString
144         .Descansar = False
146         .Navegando = 0
148         .Oculto = 0
150         .Envenenado = 0
154         .invisible = 0
156         .Paralizado = 0
158         .Inmovilizado = 0
160         .Maldicion = 0
164         .Meditando = 0
168         .Privilegios = 0
170         .PuedeMoverse = 0
172         .OldBody = 0
174         .OldHead = 0
176         .AdminInvisible = 0
178         .ValCoDe = 0
180         .Hechizo = 0
182         .Silenciado = 0
186         .AdminPerseguible = False
188         .VecesQueMoriste = 0
190         .MinutosRestantes = 0
192         .SegundosPasados = 0
196         .Montado = 0
198         .Incinerado = 0
199         .ActiveTransform = 0
200         .Casado = 0
202         .SpouseId = 0
204         Call SetUserRef(.Candidato, 0)
206         .UsandoMacro = False
208         .pregunta = 0

210         .Subastando = False
212         .Paraliza = 0
214         .Envenena = 0
216         .NoPalabrasMagicas = 0
218         .NoMagiaEfecto = 0
220         .incinera = 0
222         .Estupidiza = 0
224         .GolpeCertero = 0
226         .PendienteDelExperto = 0
230         .PendienteDelSacrificio = 0
232         .AnilloOcultismo = 0
234         .RegeneracionMana = 0
236         .RegeneracionHP = 0
            .StatusMask = 0

244         .LastKillerIndex = 0
        
246         .UserLogged = False
248         .FirstPacket = False
250         .Inmunidad = 0
            
252         .Mimetizado = e_EstadoMimetismo.Desactivado
254         .MascotasGuardadas = 0
255         .Cleave = 0
256         .EnConsulta = False
258         .YaGuardo = False
            .ModificoAttributos = False
            .ModificoHechizos = False
            .ModificoInventario = False
            .ModificoInventarioBanco = False
            .ModificoSkills = False
            .ModificoMascotas = False
            .ModificoQuests = False
            .ModificoQuestsHechas = False
            .RespondiendoPregunta = False
            Call ClearUserRef(.LastAttacker)
            .LastAttackedByUserTime = 0
            Call ClearUserRef(.LastHelpUser)
            .LastHelpByTime = 0

            Dim i As Integer
266         For i = LBound(.ChatHistory) To UBound(.ChatHistory)
268             .ChatHistory(i) = vbNullString
            Next

270         .EnReto = False
272         .SolicitudReto.estado = e_SolicitudRetoEstado.Libre
274         Call SetUserRef(.AceptoReto, 0)
276         .LastPos.map = 0
278         .ReturnPos.map = 0
            
280         .Crafteando = 0

            'HarThaoS: Captura de bandera
            .jugando_captura = 0
            .CurrentTeam = 0
            .jugando_captura_timer = 0
            .jugando_captura_muertes = 0
            Call SetUserRef(.SigueUsuario, 0)
            Call SetUserRef(.GMMeSigue, 0)
        End With

        
        Exit Sub

ResetUserFlags_Err:
282     Call TraceError(Err.Number, Err.Description, "TCP.ResetUserFlags", Erl)

        
End Sub

Sub ResetAccionesPendientes(ByVal UserIndex As Integer)
        
        On Error GoTo ResetAccionesPendientes_Err
        

        '*************************************************
        '*************************************************
100     With UserList(UserIndex).Accion
102         .AccionPendiente = False
104         .HechizoPendiente = 0
106         .RunaObj = 0
108         .Particula = 0
110         .TipoAccion = 0
112         .ObjSlot = 0

        End With

        
        Exit Sub

ResetAccionesPendientes_Err:
114     Call TraceError(Err.Number, Err.Description, "TCP.ResetAccionesPendientes", Erl)

        
End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)
        
        On Error GoTo ResetUserSpells_Err
        

        Dim LoopC As Long

100     For LoopC = 1 To MAXUSERHECHIZOS
102         UserList(UserIndex).Stats.UserHechizos(LoopC) = 0
            ' UserList(UserIndex).Stats.UserHechizosInterval(LoopC) = 0
104     Next LoopC

        
        Exit Sub

ResetUserSpells_Err:
106     Call TraceError(Err.Number, Err.Description, "TCP.ResetUserSpells", Erl)

        
End Sub

Sub ResetUserSkills(ByVal UserIndex As Integer)
        
        On Error GoTo ResetUserSkills_Err
        

        Dim LoopC As Long

100     For LoopC = 1 To NUMSKILLS
102         UserList(UserIndex).Stats.UserSkills(LoopC) = 0
104     Next LoopC

        
        Exit Sub

ResetUserSkills_Err:
106     Call TraceError(Err.Number, Err.Description, "TCP.ResetUserSkills", Erl)

        
End Sub

Sub ResetUserBanco(ByVal UserIndex As Integer)
        
        On Error GoTo ResetUserBanco_Err
        

        Dim LoopC As Long
    
100     For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
102         UserList(UserIndex).BancoInvent.Object(LoopC).amount = 0
104         UserList(UserIndex).BancoInvent.Object(LoopC).Equipped = 0
106         UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = 0
108     Next LoopC
    
110     UserList(UserIndex).BancoInvent.NroItems = 0

        
        Exit Sub

ResetUserBanco_Err:
112     Call TraceError(Err.Number, Err.Description, "TCP.ResetUserBanco", Erl)

        
End Sub

Sub ResetUserKeys(ByVal UserIndex As Integer)
        
        On Error GoTo ResetUserKeys_Err
    
        
100     With UserList(UserIndex)
            Dim i As Integer
        
102         For i = 1 To MAXKEYS
104             .Keys(i) = 0
            Next
        End With
        
        Exit Sub

ResetUserKeys_Err:
106     Call TraceError(Err.Number, Err.Description, "TCP.ResetUserKeys", Erl)

        
End Sub

Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)
        
        On Error GoTo LimpiarComercioSeguro_Err
        

100     With UserList(UserIndex).ComUsu

102         If IsValidUserRef(.DestUsu) Then
104             Call FinComerciarUsu(.DestUsu.ArrayIndex)
106             Call FinComerciarUsu(UserIndex)

            End If

        End With

        
        Exit Sub

LimpiarComercioSeguro_Err:
108     Call TraceError(Err.Number, Err.Description, "TCP.LimpiarComercioSeguro", Erl)

        
End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)
        On Error GoTo ResetUserSlot_Err
        Call SaveDCUserCache(UserIndex)
        Call AntiCheat.OnPlayerDisconnect(UserIndex)
        With UserList(UserIndex)
100         .ConnectionDetails.ConnIDValida = False
102         .ConnectionDetails.ConnID = 0
104         .Stats.Shield = 0

106         If .Grupo.EnGrupo Then
108             If .Grupo.Lider.ArrayIndex = UserIndex Then
110                 Call FinalizarGrupo(UserIndex)
                Else
111                 Call SalirDeGrupoForzado(UserIndex)
                End If
            End If
        
            If m_NameIndex.Exists(UCase(.name)) Then
                Call m_NameIndex.Remove(UCase(.name))
            End If
        
112         .Grupo.CantidadMiembros = 0
114         .Grupo.EnGrupo = False
115         .Grupo.Id = -1
116         Call SetUserRef(.Grupo.Lider, 0)
118         Call SetUserRef(.Grupo.PropuestaDe, 0)
120         Call SetUserRef(.Grupo.Miembros(6), 0)
122         Call SetUserRef(.Grupo.Miembros(1), 0)
124         Call SetUserRef(.Grupo.Miembros(2), 0)
126         Call SetUserRef(.Grupo.Miembros(3), 0)
128         Call SetUserRef(.Grupo.Miembros(4), 0)
130         Call SetUserRef(.Grupo.Miembros(5), 0)
131         Call ClearEffectList(.EffectOverTime)
132         Call ClearModifiers(.Modifiers)
        End With
133     Call ResetQuestStats(UserIndex)
134     Call ResetGuildInfo(UserIndex)
136     Call LimpiarComercioSeguro(UserIndex)
138     Call ResetFacciones(UserIndex)
140     Call ResetContadores(UserIndex)
141     Call ResetPacketRateData(UserIndex)
142     Call ResetCharInfo(UserIndex)
144     Call ResetBasicUserInfo(UserIndex)
146     Call ResetUserFlags(UserIndex)
148     Call ResetAccionesPendientes(UserIndex)
152     Call LimpiarInventario(UserIndex)
154     Call ResetUserSpells(UserIndex)
156     Call ResetUserBanco(UserIndex)
158     Call ResetUserSkills(UserIndex)
160     Call ResetUserKeys(UserIndex)
161     Call ResetCd(UserList(UserIndex))
162     With UserList(UserIndex).ComUsu
164         .Acepto = False
166         .cant = 0
168         .DestNick = vbNullString
170         Call SetUserRef(.DestUsu, 0)
172         .Objeto = 0
        End With
174     UserList(UserIndex).InUse = False
176     Call IncreaseVersionId(UserIndex)
178     Call ReleaseUser(UserIndex)
        Exit Sub
ResetUserSlot_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.ResetUserSlot", Erl)
End Sub

Sub ClearAndSaveUser(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler
    
    Dim errordesc As String
    Dim map As Integer
    Dim aN  As Integer
    Dim i   As Integer

100 With UserList(UserIndex)
102         errordesc = "ERROR AL SETEAR NPC"
        
104         Call ClearAttackerNpc(UserIndex)
        
128         errordesc = "ERROR AL DESMONTAR"
    
130         If .flags.Montado > 0 Then
132             Call DoMontar(UserIndex, ObjData(.Invent.MonturaObjIndex), .Invent.MonturaSlot)
            End If
            
134         errordesc = "ERROR AL CANCELAR SOLICITUD DE RETO"
            
136         If .flags.EnReto Then
138             Call AbandonarReto(UserIndex, True)

140         ElseIf .flags.SolicitudReto.estado <> e_SolicitudRetoEstado.Libre Then
142             Call CancelarSolicitudReto(UserIndex, .name & " se ha desconectado.")
            
144         ElseIf IsValidUserRef(.flags.AceptoReto) Then
146             Call CancelarSolicitudReto(.flags.AceptoReto.ArrayIndex, .name & " se ha desconectado.")
            End If
            
            
            'Se desconecta un usuario seguido
            If IsValidUserRef(.flags.GMMeSigue) Then
                Call WriteCancelarSeguimiento(.flags.GMMeSigue.ArrayIndex)
                Call SetUserRef(UserList(.flags.GMMeSigue.ArrayIndex).flags.SigueUsuario, 0)
                UserList(.flags.GMMeSigue.ArrayIndex).Invent = UserList(.flags.GMMeSigue.ArrayIndex).Invent_bk
                UserList(.flags.GMMeSigue.ArrayIndex).Stats = UserList(.flags.GMMeSigue.ArrayIndex).Stats_bk
                'UserList(.flags.GMMeSigue).Char.charindex = UserList(.flags.GMMeSigue).Char.charindex_bk
                Call WriteUserCharIndexInServer(.flags.GMMeSigue.ArrayIndex)
                Call UpdateUserInv(True, .flags.GMMeSigue.ArrayIndex, 1)
                Call WriteUpdateUserStats(.flags.GMMeSigue.ArrayIndex)
                Call WriteConsoleMsg(.flags.GMMeSigue.ArrayIndex, "El usuario " & UserList(UserIndex).name & " que estabas siguiendo se desconectó.", e_FontTypeNames.FONTTYPE_INFO)
                Call SetUserRef(.flags.GMMeSigue, 0)
                'Falta revertir inventario del GM
            End If
                
            If IsValidUserRef(.flags.SigueUsuario) Then
                'Para que el usuario deje de mandar el floodeo de paquetes
                Call WriteNotificarClienteSeguido(.flags.SigueUsuario.ArrayIndex, 0)
                Call SetUserRef(UserList(.flags.SigueUsuario.ArrayIndex).flags.GMMeSigue, 0)
                UserList(UserIndex).Invent = UserList(UserIndex).Invent_bk
                UserList(UserIndex).Stats = UserList(UserIndex).Stats_bk
                Call SetUserRef(.flags.SigueUsuario, 0)
            End If
            
        
148         errordesc = "ERROR AL SACAR MIMETISMO"
150         If .flags.Mimetizado > 0 Then

152             .Char.Body = .CharMimetizado.Body
154             .Char.Head = .CharMimetizado.Head
156             .Char.CascoAnim = .CharMimetizado.CascoAnim
158             .Char.ShieldAnim = .CharMimetizado.ShieldAnim
160             .Char.WeaponAnim = .CharMimetizado.WeaponAnim
161             .char.CartAnim = .CharMimetizado.CartAnim
162             .Counters.Mimetismo = 0
164             .flags.Mimetizado = e_EstadoMimetismo.Desactivado

            End If
            Call ClearEffectList(.EffectOverTime, e_EffectType.eAny, False)
166         errordesc = "ERROR AL LIMPIAR INVENTARIO DE CRAFTEO"
168         If .flags.Crafteando <> 0 Then
170             Call ReturnCraftingItems(UserIndex)
            End If
        
172         errordesc = "ERROR AL ENVIAR PARTICULA"
        
174         .Char.FX = 0
176         .Char.loops = 0
178         .Char.ParticulaFx = 0
180         Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, 0, 0, True))
182         Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 0, 0))

186         errordesc = "ERROR AL ENVIAR INVI"
        
            'Le devolvemos el body y head originales
188         If .flags.AdminInvisible = 1 Then Call DoAdminInvisible(UserIndex)
        
190         errordesc = "ERROR AL CANCELAR SUBASTA"
    
192         If .flags.Subastando = True Then
194             Call CancelarSubasta
    
            End If
        
196         errordesc = "ERROR AL BORRAR INDEX DE TORNEO"
    
198         If .flags.EnTorneo = True Then
200             Call BorrarIndexInTorneo(UserIndex)
202             .flags.EnTorneo = False
    
            End If
        
            'Save statistics
            'Call Statistics.UserDisconnected(UserIndex)
        
            ' Grabamos el personaje del usuario
        
204         errordesc = "ERROR AL GRABAR PJ"

206         Call SaveUser(UserIndex, True)

    End With
    
    Exit Sub
    
ErrHandler:
        'Call LogError("Error en CloseUser. Número " & Err.Number & ". Descripción: " & Err.Description & ". Detalle:" & errordesc)
208     Call TraceError(Err.Number, Err.Description & ". Detalle:" & errordesc, Erl)
210     Resume Next ' TODO: Provisional hasta solucionar bugs graves

End Sub

Sub CloseUser(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler
    
        Dim errordesc As String
        Dim map As Integer
        Dim aN  As Integer
        Dim i   As Integer
        
100     With UserList(UserIndex)
            map = .pos.map
104         If Not .flags.YaGuardo Then
106             Call ClearAndSaveUser(UserIndex)
            End If

108         errordesc = "ERROR AL DESCONTAR USER DE MAPA"
    
110         If MapInfo(map).NumUsers > 0 Then
112             Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
    
            End If
    
114         errordesc = "ERROR AL ERASEUSERCHAR"
        
            'Borrar el personaje
116         Call EraseUserChar(UserIndex, True)
        
118         errordesc = "ERROR AL BORRAR MASCOTAS"
        
            'Borrar mascotas
120         For i = 1 To MAXMASCOTAS
122             If IsValidNpcRef(.MascotasIndex(i)) Then
124                 If NpcList(.MascotasIndex(i).ArrayIndex).flags.NPCActive Then _
                        Call QuitarNPC(.MascotasIndex(i).ArrayIndex, eClearPlayerPets)
                End If
                Call ClearNpcRef(.MascotasIndex(i))
126         Next i
        
128         errordesc = "ERROR Update Map Users map: " & map
        
            'Update Map Users
130         MapInfo(map).NumUsers = MapInfo(map).NumUsers - 1
            Call Execute("update user set is_logged = 0 where id = ?;", UserList(UserIndex).ID)
        
132         If MapInfo(map).NumUsers < 0 Then MapInfo(map).NumUsers = 0
    
            ' Si el usuario habia dejado un msg en la gm's queue lo borramos
            'If Ayuda.Existe(.Name) Then Call Ayuda.Quitar(.Name)
        
134         errordesc = "ERROR AL m_NameIndex.Remove() Name:" & .name & " cuenta:" & .Cuenta
            
136         Call m_NameIndex.Remove(UCase$(.name))
        
138         errordesc = "ERROR AL RESETSLOT Name:" & .name & " cuenta:" & .Cuenta

140         .flags.UserLogged = False

            .Counters.Saliendo = False
                
142         Call ResetUserSlot(UserIndex)
    

        End With
    
        Exit Sub
    
ErrHandler:
        'Call LogError("Error en CloseUser. Número " & Err.Number & ". Descripción: " & Err.Description & ". Detalle:" & errordesc)
144     Call TraceError(Err.Number, Err.Description & ". Detalle:" & errordesc, Erl)
146     Resume Next ' TODO: Provisional hasta solucionar bugs graves

End Sub

Public Sub EcharPjsNoPrivilegiados()
        
        On Error GoTo EcharPjsNoPrivilegiados_Err
        

        Dim LoopC As Long

100     For LoopC = 1 To LastUser

102         If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnectionDetails.ConnIDValida Then
104             If UserList(LoopC).flags.Privilegios And e_PlayerType.user Then
106                 Call CloseSocket(LoopC)

                End If

            End If

108     Next LoopC

        
        Exit Sub

EcharPjsNoPrivilegiados_Err:
110     Call TraceError(Err.Number, Err.Description, "TCP.EcharPjsNoPrivilegiados", Erl)

        
End Sub

Function ValidarCabeza(ByVal UserRaza As e_Raza, ByVal UserSexo As e_Genero, ByVal Head As Integer) As Boolean

100     Select Case UserSexo
    
            Case e_Genero.Hombre
        
102             Select Case UserRaza
                
                    Case e_Raza.Humano
104                     ValidarCabeza = head >= 1 And head <= 41 Or head >= 778 And head <= 791
                    
106                 Case e_Raza.Elfo
108                     ValidarCabeza = head >= 101 And head <= 132 Or head >= 531 And head <= 545
                    
110                 Case e_Raza.Drow
112                     ValidarCabeza = Head >= 200 And Head <= 229
                    
114                 Case e_Raza.Enano
116                     ValidarCabeza = head >= 300 And head <= 344
                    
118                 Case e_Raza.Gnomo
120                     ValidarCabeza = Head >= 400 And Head <= 429
                    
122                 Case e_Raza.Orco
124                     ValidarCabeza = Head >= 500 And Head <= 529
                
                End Select
        
126         Case e_Genero.Mujer
        
128             Select Case UserRaza
                
                    Case e_Raza.Humano
130                     ValidarCabeza = head >= 50 And head <= 80 Or head >= 187 And head <= 190 Or head >= 230 And head <= 246
                    
132                 Case e_Raza.Elfo
134                     ValidarCabeza = head >= 150 And head <= 179 Or head >= 758 And head <= 777
                    
136                 Case e_Raza.Drow
138                     ValidarCabeza = Head >= 250 And Head <= 279
                    
140                 Case e_Raza.Enano
142                     ValidarCabeza = Head >= 350 And Head <= 379
                    
144                 Case e_Raza.Gnomo
146                     ValidarCabeza = Head >= 450 And Head <= 479
                    
148                 Case e_Raza.Orco
150                     ValidarCabeza = Head >= 550 And Head <= 579
                
                End Select
    
        End Select

End Function

Function ValidarNombre(nombre As String) As Boolean
    
100     If Len(nombre) < 3 Or Len(nombre) > 18 Then Exit Function
    
        Dim Temp As String
102     Temp = UCase$(nombre)
    
        Dim i As Long, Char As Integer, LastChar As Integer
104     For i = 1 To Len(Temp)
106         Char = Asc(mid$(Temp, i, 1))
        
108         If (Char < 65 Or Char > 90) And Char <> 32 Then
                Exit Function
        
110         ElseIf Char = 32 And LastChar = 32 Then
                Exit Function
            End If
        
112         LastChar = Char
        Next

114     If Asc(mid$(Temp, 1, 1)) = 32 Or Asc(mid$(Temp, Len(Temp), 1)) = 32 Then
            Exit Function
        End If
    
116     ValidarNombre = True

End Function

Function ContarUsuariosMismaCuenta(ByVal AccountID As Long) As Integer

        Dim i As Integer
    
100     For i = 1 To LastUser
        
102         If UserList(i).flags.UserLogged And UserList(i).AccountID = AccountID Then
104             ContarUsuariosMismaCuenta = ContarUsuariosMismaCuenta + 1
            End If
        
        Next

End Function

Sub ResetCd(ByRef user As t_User)
    Dim i As Integer
    For i = 0 To e_CdTypes.CDCount - 1
        user.CdTimes(i) = 0
    Next i
End Sub

Sub VaciarInventario(ByVal UserIndex As Integer)

    Dim i As Long

    With UserList(UserIndex)
        For i = 1 To MAX_INVENTORY_SLOTS
            .Invent.Object(i).amount = 0
            .Invent.Object(i).Equipped = 0
            .Invent.Object(i).ObjIndex = 0
        Next i
    End With
End Sub
