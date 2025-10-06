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
    UserGenero = UserList(UserIndex).genero
    UserRaza = UserList(UserIndex).raza
    Select Case UserGenero
        Case e_Genero.Hombre
            Select Case UserRaza
                Case e_Raza.Humano
                    NewBody = 1
                Case e_Raza.Elfo
                    NewBody = 2
                Case e_Raza.Drow
                    NewBody = 3
                Case e_Raza.Enano
                    NewBody = 300
                Case e_Raza.Gnomo
                    NewBody = 300
                Case e_Raza.Orco
                    NewBody = 582
            End Select
        Case e_Genero.Mujer
            Select Case UserRaza
                Case e_Raza.Humano
                    NewBody = 1
                Case e_Raza.Elfo
                    NewBody = 2
                Case e_Raza.Drow
                    NewBody = 3
                Case e_Raza.Gnomo
                    NewBody = 300
                Case e_Raza.Enano
                    NewBody = 300
                Case e_Raza.Orco
                    NewBody = 581
            End Select
    End Select
    UserList(UserIndex).Char.body = NewBody
    Exit Sub
DarCuerpo_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.DarCuerpo", Erl)
End Sub

Sub RellenarInventario(ByVal UserIndex As String)
    On Error GoTo RellenarInventario_Err
    With UserList(UserIndex)
        Dim NumItems As Integer
        NumItems = 1
        ' Todos reciben pociones rojas
        .invent.Object(NumItems).ObjIndex = 4335 'Pocion Roja
        .invent.Object(NumItems).amount = 350
        NumItems = NumItems + 1
        ' Magicas puras reciben más azules
        Select Case .clase
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
        Select Case .clase
            Case e_Class.Mage, e_Class.Cleric, e_Class.Druid, e_Class.Bard, e_Class.Paladin, e_Class.Bandit, e_Class.Assasin
                .Stats.UserHechizos(1) = 1 ' Dardo mágico
        End Select
        ' Pociones amarillas y verdes
        Select Case .clase
            Case e_Class.Assasin, e_Class.Bard, e_Class.Cleric, e_Class.Hunter, e_Class.Paladin, e_Class.Trabajador, e_Class.Warrior, e_Class.Bandit, e_Class.Pirat, e_Class.Thief
                .invent.Object(NumItems).ObjIndex = 4337 ' Pocion Amarilla
                .invent.Object(NumItems).amount = 60
                NumItems = NumItems + 1
                .invent.Object(NumItems).ObjIndex = 4338 ' Pocion Verde
                .invent.Object(NumItems).amount = 60
                NumItems = NumItems + 1
            Case e_Class.Mage, e_Class.Druid
                .invent.Object(NumItems).ObjIndex = 4337 ' Pocion Amarilla
                .invent.Object(NumItems).amount = 60
                NumItems = NumItems + 1
        End Select
        ' Poción violeta
        .invent.Object(NumItems).ObjIndex = 4334 ' Pocion violeta
        .invent.Object(NumItems).amount = 15
        NumItems = NumItems + 1
        .invent.Object(NumItems).ObjIndex = 3791 ' Pasaje a Jourmut
        .invent.Object(NumItems).amount = 2
        NumItems = NumItems + 1
        ' Armas
        Select Case .clase
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
            Case e_Class.Hunter
                .invent.Object(NumItems).ObjIndex = 3686 ' Daga del principiante
                .invent.Object(NumItems).amount = 1
                NumItems = NumItems + 1
                .invent.Object(NumItems).ObjIndex = 3491 ' Arco del principiante
                .invent.Object(NumItems).amount = 1
                NumItems = NumItems + 1
                .invent.Object(NumItems).ObjIndex = 3492 ' Flecha del Principiante
                .invent.Object(NumItems).amount = 1050
                NumItems = NumItems + 1
                .invent.Object(NumItems).ObjIndex = 3489  ' Casco de Principiante
                .invent.Object(NumItems).amount = 1
                NumItems = NumItems + 1
                .invent.Object(NumItems).ObjIndex = 3490 ' Anillo del Principiante
                .invent.Object(NumItems).amount = 1
                NumItems = NumItems + 1
                .invent.Object(NumItems).ObjIndex = 3488 ' Escudo de Principiante
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
                .invent.Object(NumItems).amount = 600
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
                .invent.Object(NumItems).ObjIndex = 3490 ' Anillo del Principiante
                .invent.Object(NumItems).amount = 1
                NumItems = NumItems + 1
                .invent.Object(NumItems).ObjIndex = 3491 ' Arco del Principiante
                .invent.Object(NumItems).amount = 1
                NumItems = NumItems + 1
                .invent.Object(NumItems).ObjIndex = 3492 ' Flecha del Principiante
                .invent.Object(NumItems).amount = 600
                NumItems = NumItems + 1
                .invent.Object(NumItems).ObjIndex = 3686 ' Daga del Principiante
                .invent.Object(NumItems).amount = 1
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
                .invent.Object(NumItems).ObjIndex = 3487 ' Espada del Principiante
                .invent.Object(NumItems).amount = 1
                NumItems = NumItems + 1
                .invent.Object(NumItems).ObjIndex = 3494 ' Flauta del Principiante
                .invent.Object(NumItems).amount = 1
                NumItems = NumItems + 1
                .invent.Object(NumItems).ObjIndex = 1778  'Casco de Lobo (Resistencia Magica 1)
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
        ' Armadura o túnica de principiante
        Select Case .clase
                ' Todas menos mago, druida y bardo:
            Case e_Class.Trabajador, e_Class.Thief, e_Class.Paladin, e_Class.Cleric, e_Class.Assasin, e_Class.Bandit, e_Class.Pirat, e_Class.Warrior, e_Class.Hunter
                .invent.Object(NumItems).ObjIndex = 3500 ' Armadura de Principiante
                ' Mago, druida y bardo:
            Case e_Class.Mage, e_Class.Druid, e_Class.Bard
                .invent.Object(NumItems).ObjIndex = 3502 ' Túnica del Principiante
        End Select
        .invent.Object(NumItems).Equipped = 0
        Call EquiparInvItem(UserIndex, NumItems)
        .invent.Object(NumItems).amount = 1
        .invent.Object(NumItems).Equipped = 1
        .invent.EquippedArmorSlot = NumItems
        .invent.EquippedArmorObjIndex = .invent.Object(NumItems).ObjIndex
        NumItems = NumItems + 1
        ' Animación según raza
        .Char.body = ObtenerRopaje(UserIndex, ObjData(.invent.EquippedArmorObjIndex))
        ' Comida y bebida
        .invent.Object(NumItems).ObjIndex = 3684 ' Manzana
        .invent.Object(NumItems).amount = 50
        NumItems = NumItems + 1
        .invent.Object(NumItems).ObjIndex = 3685 ' Agua
        .invent.Object(NumItems).amount = 50
        NumItems = NumItems + 1
        ' Seteo la cantidad de items
        .invent.NroItems = NumItems
        .flags.ModificoInventario = True
        .flags.ModificoHechizos = True
    End With
    Exit Sub
RellenarInventario_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.RellenarInventario", Erl)
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    On Error GoTo AsciiValidos_Err
    Dim car As Byte
    Dim i   As Integer
    cad = LCase$(cad)
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
            AsciiValidos = False
            Exit Function
        End If
    Next i
    AsciiValidos = True
    Exit Function
AsciiValidos_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.AsciiValidos", Erl)
End Function

Function Numeric(ByVal cad As String) As Boolean
    On Error GoTo Numeric_Err
    Dim car As Byte
    Dim i   As Integer
    cad = LCase$(cad)
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        If (car < 48 Or car > 57) Then
            Numeric = False
            Exit Function
        End If
    Next i
    Numeric = True
    Exit Function
Numeric_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.Numeric", Erl)
End Function

Function NombrePermitido(ByVal nombre As String) As Boolean
    On Error GoTo NombrePermitido_Err
    Dim i As Integer
    For i = 1 To UBound(ForbidenNames)
        If LCase$(nombre) = ForbidenNames(i) Then
            NombrePermitido = False
            Exit Function
        End If
    Next i
    NombrePermitido = True
    Exit Function
NombrePermitido_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.NombrePermitido", Erl)
End Function

Function Validate_Skills(ByVal UserIndex As Integer) As Boolean
    On Error GoTo Validate_Skills_Err
    Dim LoopC As Integer
    For LoopC = 1 To NUMSKILLS
        If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then
            Exit Function
            If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100
        End If
    Next LoopC
    Validate_Skills = True
    Exit Function
Validate_Skills_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.Validate_Skills", Erl)
End Function

Function ConnectNewUser(ByVal UserIndex As Integer, _
                        ByRef name As String, _
                        ByVal UserRaza As e_Raza, _
                        ByVal UserSexo As e_Genero, _
                        ByVal UserClase As e_Class, _
                        ByVal head As Integer, _
                        ByVal Hogar As e_Ciudad) As Boolean
    On Error GoTo ConnectNewUser_Err
    With UserList(UserIndex)
        Dim LoopC As Long
        If .flags.UserLogged Then
            Call LogSecurity("El usuario " & .name & " ha intentado crear a " & name & " desde la IP " & .ConnectionDetails.IP)
            Call CloseSocketSL(UserIndex)
            Call Cerrar_Usuario(UserIndex)
            Exit Function
        End If
        
#If LOGIN_STRESS_TEST = 0 Then
        ' Nombre válido
        If Not ValidarNombre(name) Then
            Call LogSecurity("ValidarNombre failed in ConnectNewUser for " & name & " desde la IP " & .ConnectionDetails.IP)
            Call CloseSocketSL(UserIndex)
            Exit Function
        End If
        If Not NombrePermitido(name) Then
            Call WriteShowMessageBox(UserIndex, 1768, vbNullString) 'Msg1768=El nombre no está permitido.
            Exit Function
        End If
#End If
        '¿Existe el personaje?
        If PersonajeExiste(name) Then
            Call WriteShowMessageBox(UserIndex, 1769, vbNullString) 'Msg1769=Ya existe el personaje.
            Exit Function
        End If
        ' Raza válida
        If UserRaza <= 0 Or UserRaza > NUMRAZAS Then Exit Function
        ' Género válido
        If UserSexo < Hombre Or UserSexo > Mujer Then Exit Function
        ' Ciudad válida
        If Hogar <= 0 Or Hogar > NUMCIUDADES Then Exit Function
        ' Cabeza válida
#If LOGIN_STRESS_TEST = 0 Then
        If Not ValidarCabeza(UserRaza, UserSexo, head) Then Exit Function
#Else
        head = GetRandomHead(UserRaza, UserSexo)
#End If
        'Prevenimos algun bug con dados inválidos
        'If .Stats.UserAtributos(e_Atributos.Fuerza) = 0 Then Exit Function
        .Stats.UserAtributos(e_Atributos.Fuerza) = 18 + ModRaza(UserRaza).Fuerza
        .Stats.UserAtributos(e_Atributos.Agilidad) = 18 + ModRaza(UserRaza).Agilidad
        .Stats.UserAtributos(e_Atributos.Inteligencia) = 18 + ModRaza(UserRaza).Inteligencia
        .Stats.UserAtributos(e_Atributos.Constitucion) = 18 + ModRaza(UserRaza).Constitucion
        .Stats.UserAtributos(e_Atributos.Carisma) = 18 + ModRaza(UserRaza).Carisma
        .Stats.UserAtributosBackUP(e_Atributos.Fuerza) = .Stats.UserAtributos(e_Atributos.Fuerza)
        .Stats.UserAtributosBackUP(e_Atributos.Agilidad) = .Stats.UserAtributos(e_Atributos.Agilidad)
        .Stats.UserAtributosBackUP(e_Atributos.Inteligencia) = .Stats.UserAtributos(e_Atributos.Inteligencia)
        .Stats.UserAtributosBackUP(e_Atributos.Constitucion) = .Stats.UserAtributos(e_Atributos.Constitucion)
        .Stats.UserAtributosBackUP(e_Atributos.Carisma) = .Stats.UserAtributos(e_Atributos.Carisma)
        .flags.Muerto = 0
        .flags.Escondido = 0
        .flags.Casado = 0
        .flags.SpouseId = 0
        .name = name
        .clase = Min(max(0, UserClase), NUMCLASES)
        .raza = UserRaza
        .Char.head = head
        .genero = UserSexo
        .Hogar = cForgat
        '%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%
        .Stats.SkillPts = 10
        .Char.Heading = e_Heading.SOUTH
        Call DarCuerpo(UserIndex) 'Ladder REVISAR
        .OrigChar = .Char
        Call ClearClothes(.Char)
        '  Vida inicial
        .Stats.MaxHp = .Stats.UserAtributos(e_Atributos.Constitucion)
        .Stats.MinHp = .Stats.MaxHp
        .Stats.shield = 0
        '  Maná inicial
        .Stats.MaxMAN = .Stats.UserAtributos(e_Atributos.Inteligencia) * ModClase(.clase).ManaInicial
        .Stats.MinMAN = .Stats.MaxMAN
        Dim MiInt As Integer
        MiInt = RandomNumber(1, .Stats.UserAtributos(e_Atributos.Agilidad) \ 6)
        If MiInt = 1 Then MiInt = 2
        .Stats.MaxSta = 20 * MiInt
        .Stats.MinSta = 20 * MiInt
        .Stats.MaxAGU = 100
        .Stats.MinAGU = 100
        .Stats.MaxHam = 100
        .Stats.MinHam = 100
        .flags.VecesQueMoriste = 0
        .flags.Montado = 0
        .Stats.MaxHit = 2
        .Stats.MinHIT = 1
        .Stats.GLD = 0
        .Stats.Exp = 0
        .Stats.ELV = 1
        Call RellenarInventario(UserIndex)
        #If ConUpTime Then
            .LogOnTime = Now
            .UpTime = 0
        #End If
        Call ResetCd(UserList(UserIndex))
        'Valores Default de facciones al Activar nuevo usuario
        Call ResetFacciones(UserIndex)
        .Faccion.Status = 1
        .ChatCombate = 1
        .ChatGlobal = 1
        Select Case .Hogar
            Case e_Ciudad.cUllathorpe
                .pos.Map = 1
                .pos.x = 56
                .pos.y = 44
            Case e_Ciudad.cArghal
                .pos.Map = 151
                .pos.x = 52
                .pos.y = 36
            Case e_Ciudad.cForgat
                .pos.Map = 517
                .pos.x = 48
                .pos.y = 64
            Case e_Ciudad.cNix
                .pos.Map = 34
                .pos.x = 40
                .pos.y = 86
            Case e_Ciudad.cLindos
                .pos.Map = 408
                .pos.x = 63
                .pos.y = 39
            Case e_Ciudad.cBanderbill
                .pos.Map = 59
                .pos.x = 47
                .pos.y = 41
            Case e_Ciudad.cArkhein
                .pos.Map = 196
                .pos.x = 43
                .pos.y = 58
            Case e_Ciudad.cEldoria
                .pos.Map = 440
                .pos.x = 50
                .pos.y = 88
            Case e_Ciudad.cPenthar
                .pos.Map = 560
                .pos.x = 40
                .pos.y = 69
        End Select
        UltimoChar = UCase$(name)
        Call SaveNewUser(UserIndex)
        ConnectNewUser = True
        #If PYMMO = 1 Then
            Call ConnectUser(UserIndex, name, True)
        #ElseIf PYMMO = 0 Then
            Call ConnectUser(UserIndex, name, False)
        #End If
    End With
    Exit Function
ConnectNewUser_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.ConnectNewUser", Erl)
End Function

Sub CloseSocket(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    If UserIndex = LastUser Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser < 1 Then Exit Do
        Loop
    End If
    With UserList(UserIndex)
        If .ConnectionDetails.ConnIDValida Then Call CloseSocketSL(UserIndex)
        'mato los comercios seguros
        If IsValidUserRef(.ComUsu.DestUsu) Then
            If UserList(.ComUsu.DestUsu.ArrayIndex).flags.UserLogged Then
                If UserList(.ComUsu.DestUsu.ArrayIndex).ComUsu.DestUsu.ArrayIndex = UserIndex Then
                    Call WriteConsoleMsg(.ComUsu.DestUsu.ArrayIndex, PrepareMessageLocaleMsg(1844, vbNullString, e_FontTypeNames.FONTTYPE_TALK)) ' Msg1844=Comercio cancelado por el otro usuario.
                    Call FinComerciarUsu(.ComUsu.DestUsu.ArrayIndex)
                End If
            End If
        End If
        If .flags.UserLogged Then
            Call CloseUser(UserIndex)
            If NumUsers > 0 Then NumUsers = NumUsers - 1
        Else
            Call ResetUserSlot(UserIndex)
        End If
        .ConnectionDetails.ConnIDValida = False
    End With
    Exit Sub
ErrHandler:
    UserList(UserIndex).ConnectionDetails.ConnIDValida = False
    Call ResetUserSlot(UserIndex)
    Call TraceError(Err.Number, Err.Description, "TCP.CloseSocket", Erl)
End Sub

Sub CloseSocketSL(ByVal UserIndex As Integer)
    On Error GoTo CloseSocketSL_Err
    If UserList(UserIndex).ConnectionDetails.ConnIDValida Then
        Call modNetwork.Kick(UserList(UserIndex).ConnectionDetails.ConnID)
        UserList(UserIndex).ConnectionDetails.ConnIDValida = False
    End If
    Exit Sub
CloseSocketSL_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.CloseSocketSL", Erl)
End Sub

Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean
    On Error GoTo EstaPCarea_Err
    Dim x As Integer, y As Integer
    For y = UserList(Index).pos.y - MinYBorder + 1 To UserList(Index).pos.y + MinYBorder - 1
        For x = UserList(Index).pos.x - MinXBorder + 1 To UserList(Index).pos.x + MinXBorder - 1
            If MapData(UserList(Index).pos.Map, x, y).UserIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        Next x
    Next y
    EstaPCarea = False
    Exit Function
EstaPCarea_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.EstaPCarea", Erl)
End Function

Function HayPCarea(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal ignoreUserMuerto As Boolean) As Boolean
    On Error GoTo HayPCarea_Err
    Dim tX As Integer, tY As Integer
    For tY = y - MinYBorder + 1 To y + MinYBorder - 1
        For tX = x - MinXBorder + 1 To x + MinXBorder - 1
            If InMapBounds(Map, tX, tY) Then
                If MapData(Map, tX, tY).UserIndex > 0 Then
                    If Not ignoreUserMuerto Then
                        HayPCarea = True
                    Else
                        If UserList(MapData(Map, tX, tY).UserIndex).flags.Muerto = 0 Then HayPCarea = True
                    End If
                    Exit Function
                End If
            End If
        Next
    Next
    HayPCarea = False
    Exit Function
HayPCarea_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.HayPCarea", Erl)
End Function

Function HayOBJarea(pos As t_WorldPos, ObjIndex As Integer) As Boolean
    On Error GoTo HayOBJarea_Err
    Dim x As Integer, y As Integer
    For y = pos.y - MinYBorder + 1 To pos.y + MinYBorder - 1
        For x = pos.x - MinXBorder + 1 To pos.x + MinXBorder - 1
            If MapData(pos.Map, x, y).ObjInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function
            End If
        Next x
    Next y
    HayOBJarea = False
    Exit Function
HayOBJarea_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.HayOBJarea", Erl)
End Function

Function ValidateChr(ByVal UserIndex As Integer) As Boolean
    On Error GoTo ValidateChr_Err
    ValidateChr = UserList(UserIndex).Char.body <> 0 And Validate_Skills(UserIndex)
    Exit Function
ValidateChr_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.ValidateChr", Erl)
End Function

Function EntrarCuenta(ByVal UserIndex As Integer, ByVal CuentaEmail As String, ByVal md5 As String) As Boolean
    On Error GoTo EntrarCuenta_Err
    Dim adminIdx          As Integer
    Dim laCuentaEsDeAdmin As Boolean
    If ServerSoloGMs > 0 Then
        laCuentaEsDeAdmin = False
        For adminIdx = 0 To AdministratorAccounts.count - 1
            ' Si el e-mail está declarado junto al nick de la cuenta donde esta el PJ GM en el Server.ini te dejo entrar.
            If UCase$(AdministratorAccounts.Items(adminIdx)) = UCase$(CuentaEmail) Then
                laCuentaEsDeAdmin = True
            End If
        Next adminIdx
        If Not laCuentaEsDeAdmin Then
            Call WriteShowMessageBox(UserIndex, 1770, vbNullString) 'Msg1770=El servidor se encuentra habilitado solo para administradores por el momento.
            Exit Function
        End If
    End If
    #If DEBUGGING = 0 Then
        If LCase$(Md5Cliente) <> LCase$(md5) Then
            Call WriteShowMessageBox(UserIndex, 1771, vbNullString) 'Msg1771=Error al comprobar el cliente del juego, por favor reinstale y vuelva a intentar.
            Exit Function
        End If
    #End If
    If Not CheckMailString(CuentaEmail) Then
        Call WriteShowMessageBox(UserIndex, 1772, vbNullString) 'Msg1772=Email inválido.
        Exit Function
    End If
    EntrarCuenta = EnterAccountDatabase(UserIndex, CuentaEmail)
    Exit Function
EntrarCuenta_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.EntrarCuenta", Erl)
End Function

Function ConnectUser(ByVal UserIndex As Integer, ByRef name As String, Optional ByVal newUser As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    ConnectUser = False
    With UserList(UserIndex)
        If Not ConnectUser_Check(UserIndex, name) Then
            Call LogSecurity("ConnectUser_Check " & name & " failed.")
            Exit Function
        End If
        Call ConnectUser_Prepare(UserIndex, name)
        If LoadCharacterFromDB(UserIndex) Then
            If ConnectUser_Complete(UserIndex, name, newUser) Then
                ConnectUser = True
                Exit Function
            End If
        Else
            Call WriteShowMessageBox(UserIndex, 1773, vbNullString) 'Msg1773=No se puede cargar el personaje.
            Call CloseSocket(UserIndex)
        End If
    End With
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "TCP.ConnectUser", Erl)
    Call WriteShowMessageBox(UserIndex, "El personaje contiene un error. Comuníquese con un miembro del staff.")
    Call CloseSocket(UserIndex)
End Function

Private Sub SendWelcomeUptime(ByVal UserIndex As Integer)
    Dim Msg As String
    Msg = "Server Uptime: " & FormatUptime()
    ' Pick the font/type you prefer. Examples used in this codebase include FONTTYPE_INFO or FONTTYPE_GUILD.
    ' If your helper uses a different enum or function name, keep the same idea:
    '   SendData(ToUser, UserIndex, PrepareMessageConsoleMsg(msg, e_FontTypeNames.FONTTYPE_INFO))
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageConsoleMsg(Msg, e_FontTypeNames.FONTTYPE_INFO))
End Sub

Sub SendMOTD(ByVal UserIndex As Integer)
    On Error GoTo SendMOTD_Err
    Dim j As Long
    For j = 1 To MaxLines
        Call WriteConsoleMsg(UserIndex, MOTD(j).texto, e_FontTypeNames.FONTTYPE_EXP)
    Next j
    Call SendWelcomeUptime(UserIndex)
    Exit Sub
SendMOTD_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.SendMOTD", Erl)
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
    With UserList(UserIndex).Faccion
        If .Status = e_Facciones.Armada Or .Status = e_Facciones.concilio Then
            .Status = e_Facciones.Ciudadano
        Else
            .Status = e_Facciones.Criminal
        End If
        .RecibioArmaduraCaos = 0
        .RecibioArmaduraReal = 0
        .RecompensasCaos = 0
        .RecompensasReal = 0
        .NivelIngreso = 0
        .MatadosIngreso = 0
        .FactionScore = 0
    End With
    Exit Sub
ResetFacciones_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.ResetFacciones", Erl)
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
    With UserList(UserIndex).Counters
        .AGUACounter = 0
        .AttackCounter = 0
        .Ceguera = 0
        .COMCounter = 0
        .Estupidez = 0
        .Frio = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .DisabledInvisibility = 0
        .Paralisis = 0
        .Inmovilizado = 0
        .pasos = 0
        .Pena = 0
        .PiqueteC = 0
        .STACounter = 0
        .Veneno = 0
        .Trabajando = 0
        .LastTrabajo = 0
        .Ocultando = 0
        .Lava = 0
        .Maldicion = 0
        .Saliendo = False
        .Salir = 0
        .TiempoOculto = 0
        .TimerMagiaGolpe = 0
        .TimerGolpeMagia = 0
        .TimerLanzarSpell = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeUsarArco = 0
        .TimerPuedeTrabajar = 0
        .TimerUsar = 0
        .TimerUsarClick = 0
        'Ladder
        .Incineracion = 0
        'Ladder
        .TiempoParaSubastar = 0
        .TimerPerteneceNpc = 0
        .TimerPuedeSerAtacado = 0
        .TiempoDeInmunidad = 0
        .RepetirMensaje = 0
        .MensajeGlobal = 0
        .CuentaRegresiva = -1
        .SpeedHackCounter = 0
        .LastStep = 0
        .TimerBarra = 0
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
    Call TraceError(Err.Number, Err.Description, "TCP.ResetContadores", Erl)
End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)
    '*************************************************
    'Author: Unknown
    'Last modified: 03/15/2006
    'Resetea todos los valores generales y las stats
    '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
    '*************************************************
    On Error GoTo ResetCharInfo_Err
    With UserList(UserIndex).Char
        .body = 0
        .CascoAnim = 0
        .charindex = 0
        .FX = 0
        .head = 0
        .loops = 0
        .Heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
        .CartAnim = 0
        .Arma_Aura = ""
        .Body_Aura = ""
        .Head_Aura = ""
        .Otra_Aura = ""
        .DM_Aura = ""
        .RM_Aura = ""
        .Escudo_Aura = ""
        .ParticulaFx = 0
        .speeding = 0
    End With
    Exit Sub
ResetCharInfo_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.ResetCharInfo", Erl)
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
    With UserList(UserIndex)
        .name = vbNullString
        .Cuenta = vbNullString
        .Id = -1
        .AccountID = -1
        .Desc = vbNullString
        .DescRM = vbNullString
        .pos.Map = 0
        .pos.x = 0
        .pos.y = 0
        .ConnectionDetails.IP = vbNullString
        .clase = 0
        .Email = vbNullString
        .genero = 0
        .Hogar = 0
        .raza = 0
        .EmpoCont = 0
        With .Stats
            .Banco = 0
            .ELV = 0
            .Exp = 0
            .def = 0
            '.CriminalesMatados = 0
            .NPCsMuertos = 0
            .UsuariosMatados = 0
            .PuntosPesca = 0
            .Creditos = 0
            .SkillPts = 0
            .GLD = 0
            .UserAtributos(1) = 0
            .UserAtributos(2) = 0
            .UserAtributos(3) = 0
            .UserAtributos(4) = 0
            .UserAtributosBackUP(1) = 0
            .UserAtributosBackUP(2) = 0
            .UserAtributosBackUP(3) = 0
            .UserAtributosBackUP(4) = 0
            .MaxMAN = 0
            .MinMAN = 0
            .tipoUsuario = e_TipoUsuario.tNormal
        End With
        .NroMascotas = 0
        Dim i As Integer
        For i = LBound(.MascotasType) To UBound(.MascotasType)
            .MascotasType(i) = 0
        Next i
        .LastTransportNetwork.Map = -1
    End With
    Exit Sub
ResetBasicUserInfo_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.ResetBasicUserInfo", Erl)
End Sub

Sub ResetGuildInfo(ByVal UserIndex As Integer)
    On Error GoTo ResetGuildInfo_Err
    If UserList(UserIndex).EscucheClan > 0 Then
        Call modGuilds.GMDejaDeEscucharClan(UserIndex, UserList(UserIndex).EscucheClan)
        UserList(UserIndex).EscucheClan = 0
    End If
    If UserList(UserIndex).GuildIndex > 0 Then
        Call modGuilds.m_DesconectarMiembroDelClan(UserIndex, UserList(UserIndex).GuildIndex)
    End If
    UserList(UserIndex).GuildIndex = 0
    Exit Sub
ResetGuildInfo_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.ResetGuildInfo", Erl)
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
    Call TraceError(Err.Number, Err.Description, "TCP.ResetPacketRateData", Erl)
End Sub

Sub ResetUserFlags(ByVal UserIndex As Integer)
    '*************************************************
    'Author: Unknown
    'Last modified: 03/29/2006
    'Resetea todos los valores generales y las stats
    '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
    '*************************************************
    On Error GoTo ResetUserFlags_Err
    With UserList(UserIndex).flags
        .LevelBackup = 0
        .Comerciando = False
        .Ban = 0
        .Escondido = 0
        .DuracionEfecto = 0
        .NpcInv = 0
        .StatsChanged = 0
        Call ClearNpcRef(.TargetNPC)
        .TargetNpcTipo = e_NPCType.Comun
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        Call SetUserRef(.TargetUser, 0)
        .TipoPocion = 0
        .TomoPocion = False
        .Descuento = vbNullString
        .Descansar = False
        .Navegando = 0
        .Oculto = 0
        .Envenenado = 0
        .invisible = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Maldicion = 0
        .Meditando = 0
        .Privilegios = 0
        .PuedeMoverse = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .ValCoDe = 0
        .Hechizo = 0
        .Silenciado = 0
        .AdminPerseguible = False
        .VecesQueMoriste = 0
        .MinutosRestantes = 0
        .SegundosPasados = 0
        .Montado = 0
        .Incinerado = 0
        .ActiveTransform = 0
        .Casado = 0
        .SpouseId = 0
        Call SetUserRef(.Candidato, 0)
        .UsandoMacro = False
        .pregunta = 0
        .DivineBlood = 0
        .Subastando = False
        .Paraliza = 0
        .Envenena = 0
        .NoPalabrasMagicas = 0
        .NoMagiaEfecto = 0
        .incinera = 0
        .Estupidiza = 0
        .GolpeCertero = 0
        .PendienteDelExperto = 0
        .PendienteDelSacrificio = 0
        .AnilloOcultismo = 0
        .RegeneracionMana = 0
        .RegeneracionHP = 0
        .StatusMask = 0
        .LastKillerIndex = 0
        .UserLogged = False
        .FirstPacket = False
        .Inmunidad = 0
        .Mimetizado = e_EstadoMimetismo.Desactivado
        .MascotasGuardadas = 0
        .Cleave = 0
        .EnConsulta = False
        .YaGuardo = False
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
        For i = LBound(.ChatHistory) To UBound(.ChatHistory)
            .ChatHistory(i) = vbNullString
        Next
        .EnReto = False
        .SolicitudReto.Estado = e_SolicitudRetoEstado.Libre
        Call SetUserRef(.AceptoReto, 0)
        .LastPos.Map = 0
        .ReturnPos.Map = 0
        .Crafteando = 0
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
    Call TraceError(Err.Number, Err.Description, "TCP.ResetUserFlags", Erl)
End Sub

Sub ResetAccionesPendientes(ByVal UserIndex As Integer)
    On Error GoTo ResetAccionesPendientes_Err
    '*************************************************
    '*************************************************
    With UserList(UserIndex).Accion
        .AccionPendiente = False
        .HechizoPendiente = 0
        .RunaObj = 0
        .Particula = 0
        .TipoAccion = 0
        .ObjSlot = 0
    End With
    Exit Sub
ResetAccionesPendientes_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.ResetAccionesPendientes", Erl)
End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)
    On Error GoTo ResetUserSpells_Err
    Dim LoopC As Long
    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(UserIndex).Stats.UserHechizos(LoopC) = 0
        ' UserList(UserIndex).Stats.UserHechizosInterval(LoopC) = 0
    Next LoopC
    Exit Sub
ResetUserSpells_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.ResetUserSpells", Erl)
End Sub

Sub ResetUserSkills(ByVal UserIndex As Integer)
    On Error GoTo ResetUserSkills_Err
    Dim LoopC As Long
    For LoopC = 1 To NUMSKILLS
        UserList(UserIndex).Stats.UserSkills(LoopC) = 0
    Next LoopC
    Exit Sub
ResetUserSkills_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.ResetUserSkills", Erl)
End Sub

Sub ResetUserBanco(ByVal UserIndex As Integer)
    On Error GoTo ResetUserBanco_Err
    Dim LoopC As Long
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
        UserList(UserIndex).BancoInvent.Object(LoopC).amount = 0
        UserList(UserIndex).BancoInvent.Object(LoopC).Equipped = 0
        UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = 0
        UserList(UserIndex).BancoInvent.Object(LoopC).ElementalTags = 0
    Next LoopC
    UserList(UserIndex).BancoInvent.NroItems = 0
    Exit Sub
ResetUserBanco_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.ResetUserBanco", Erl)
End Sub

Sub ResetUserKeys(ByVal UserIndex As Integer)
    On Error GoTo ResetUserKeys_Err
    With UserList(UserIndex)
        Dim i As Integer
        For i = 1 To MAXKEYS
            .Keys(i) = 0
        Next
    End With
    Exit Sub
ResetUserKeys_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.ResetUserKeys", Erl)
End Sub

Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)
    On Error GoTo LimpiarComercioSeguro_Err
    With UserList(UserIndex).ComUsu
        If IsValidUserRef(.DestUsu) Then
            Call FinComerciarUsu(.DestUsu.ArrayIndex)
            Call FinComerciarUsu(UserIndex)
        End If
    End With
    Exit Sub
LimpiarComercioSeguro_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.LimpiarComercioSeguro", Erl)
End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)
    On Error GoTo ResetUserSlot_Err
    Call SaveDCUserCache(UserIndex)
    Call AntiCheat.OnPlayerDisconnect(UserIndex)
    With UserList(UserIndex)
        .ConnectionDetails.ConnIDValida = False
        .ConnectionDetails.ConnID = 0
        .Stats.shield = 0
        If .Grupo.EnGrupo Then
            If .Grupo.Lider.ArrayIndex = UserIndex Then
                Call FinalizarGrupo(UserIndex)
            Else
                Call SalirDeGrupoForzado(UserIndex)
            End If
        End If
        If m_NameIndex.Exists(UCase(.name)) Then
            Call m_NameIndex.Remove(UCase(.name))
        End If
        .Grupo.CantidadMiembros = 0
        .Grupo.EnGrupo = False
        .Grupo.Id = -1
        Call SetUserRef(.Grupo.Lider, 0)
        Call SetUserRef(.Grupo.PropuestaDe, 0)
        Call SetUserRef(.Grupo.Miembros(6), 0)
        Call SetUserRef(.Grupo.Miembros(1), 0)
        Call SetUserRef(.Grupo.Miembros(2), 0)
        Call SetUserRef(.Grupo.Miembros(3), 0)
        Call SetUserRef(.Grupo.Miembros(4), 0)
        Call SetUserRef(.Grupo.Miembros(5), 0)
        Call ClearEffectList(.EffectOverTime)
        Call ClearModifiers(.Modifiers)
    End With
    Call ResetQuestStats(UserIndex)
    Call ResetGuildInfo(UserIndex)
    Call LimpiarComercioSeguro(UserIndex)
    Call ResetFacciones(UserIndex)
    Call ResetContadores(UserIndex)
    Call ResetPacketRateData(UserIndex)
    Call ResetCharInfo(UserIndex)
    Call ResetBasicUserInfo(UserIndex)
    Call ResetUserFlags(UserIndex)
    Call ResetAccionesPendientes(UserIndex)
    Call LimpiarInventario(UserIndex)
    Call ResetUserSpells(UserIndex)
    Call ResetUserBanco(UserIndex)
    Call ResetUserSkills(UserIndex)
    Call ResetUserKeys(UserIndex)
    Call ResetCd(UserList(UserIndex))
    With UserList(UserIndex).ComUsu
        .Acepto = False
        .cant = 0
        .DestNick = vbNullString
        Call SetUserRef(.DestUsu, 0)
        .Objeto = 0
    End With
    UserList(UserIndex).InUse = False
    Call IncreaseVersionId(UserIndex)
    Call ReleaseUser(UserIndex)
    Exit Sub
ResetUserSlot_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.ResetUserSlot", Erl)
End Sub

Sub ClearAndSaveUser(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Dim errordesc As String
    Dim Map       As Integer
    Dim aN        As Integer
    Dim i         As Integer
    With UserList(UserIndex)
        errordesc = "ERROR AL SETEAR NPC"
        Call ClearAttackerNpc(UserIndex)
        errordesc = "ERROR AL DESMONTAR"
        If .flags.Montado > 0 Then
            Call DoMontar(UserIndex, ObjData(.invent.EquippedSaddleObjIndex), .invent.EquippedSaddleSlot)
        End If
        errordesc = "ERROR AL CANCELAR SOLICITUD DE RETO"
        If .flags.EnReto Then
            Call AbandonarReto(UserIndex, True)
        ElseIf .flags.SolicitudReto.Estado <> e_SolicitudRetoEstado.Libre Then
            Call CancelarSolicitudReto(UserIndex, .name & " se ha desconectado.")
        ElseIf IsValidUserRef(.flags.AceptoReto) Then
            Call CancelarSolicitudReto(.flags.AceptoReto.ArrayIndex, .name & " se ha desconectado.")
        End If
        'Se desconecta un usuario seguido
        If IsValidUserRef(.flags.GMMeSigue) Then
            Call WriteCancelarSeguimiento(.flags.GMMeSigue.ArrayIndex)
            Call SetUserRef(UserList(.flags.GMMeSigue.ArrayIndex).flags.SigueUsuario, 0)
            UserList(.flags.GMMeSigue.ArrayIndex).invent = UserList(.flags.GMMeSigue.ArrayIndex).Invent_bk
            UserList(.flags.GMMeSigue.ArrayIndex).Stats = UserList(.flags.GMMeSigue.ArrayIndex).Stats_bk
            'UserList(.flags.GMMeSigue).Char.charindex = UserList(.flags.GMMeSigue).Char.charindex_bk
            Call WriteUserCharIndexInServer(.flags.GMMeSigue.ArrayIndex)
            Call UpdateUserInv(True, .flags.GMMeSigue.ArrayIndex, 1)
            Call WriteUpdateUserStats(.flags.GMMeSigue.ArrayIndex)
            Call WriteConsoleMsg(.flags.GMMeSigue.ArrayIndex, PrepareMessageLocaleMsg(1866, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1866=El usuario ¬1 que estabas siguiendo se desconectó.
            Call SetUserRef(.flags.GMMeSigue, 0)
            'Falta revertir inventario del GM
        End If
        If IsValidUserRef(.flags.SigueUsuario) Then
            'Para que el usuario deje de mandar el floodeo de paquetes
            Call WriteNotificarClienteSeguido(.flags.SigueUsuario.ArrayIndex, 0)
            Call SetUserRef(UserList(.flags.SigueUsuario.ArrayIndex).flags.GMMeSigue, 0)
            UserList(UserIndex).invent = UserList(UserIndex).Invent_bk
            UserList(UserIndex).Stats = UserList(UserIndex).Stats_bk
            Call SetUserRef(.flags.SigueUsuario, 0)
        End If
        errordesc = "ERROR AL SACAR MIMETISMO"
        If .flags.Mimetizado > 0 Then
            .Char.body = .CharMimetizado.body
            .Char.head = .CharMimetizado.head
            .Char.CascoAnim = .CharMimetizado.CascoAnim
            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            .Char.CartAnim = .CharMimetizado.CartAnim
            .Counters.Mimetismo = 0
            .flags.Mimetizado = e_EstadoMimetismo.Desactivado
        End If
        Call ClearEffectList(.EffectOverTime, e_EffectType.eAny, False)
        errordesc = "ERROR AL LIMPIAR INVENTARIO DE CRAFTEO"
        If .flags.Crafteando <> 0 Then
            Call ReturnCraftingItems(UserIndex)
        End If
        errordesc = "ERROR AL ENVIAR PARTICULA"
        .Char.FX = 0
        .Char.loops = 0
        .Char.ParticulaFx = 0
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, 0, 0, True))
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 0, 0))
        errordesc = "ERROR AL ENVIAR INVI"
        'Le devolvemos el body y head originales
        If .flags.AdminInvisible = 1 Then Call DoAdminInvisible(UserIndex)
        errordesc = "ERROR AL CANCELAR SUBASTA"
        If .flags.Subastando = True Then
            Call CancelarSubasta
        End If
        errordesc = "ERROR AL BORRAR INDEX DE TORNEO"
        If .flags.EnTorneo = True Then
            Call BorrarIndexInTorneo(UserIndex)
            .flags.EnTorneo = False
        End If
        'Save statistics
        'Call Statistics.UserDisconnected(UserIndex)
        ' Grabamos el personaje del usuario
        errordesc = "ERROR AL GRABAR PJ"
        Call SaveUser(UserIndex, True)
    End With
    Exit Sub
ErrHandler:
    'Call LogError("Error en CloseUser. Número " & Err.Number & ". Descripción: " & Err.Description & ". Detalle:" & errordesc)
    Call TraceError(Err.Number, Err.Description & ". Detalle:" & errordesc, Erl)
    Resume Next ' TODO: Provisional hasta solucionar bugs graves
End Sub

Sub CloseUser(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Dim errordesc As String
    Dim Map       As Integer
    Dim aN        As Integer
    Dim i         As Integer
    With UserList(UserIndex)
        Map = .pos.Map
        If Not .flags.YaGuardo Then
            Call ClearAndSaveUser(UserIndex)
        End If
        errordesc = "ERROR AL DESCONTAR USER DE MAPA"
        If MapInfo(Map).NumUsers > 0 Then
            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageRemoveCharDialog(.Char.charindex))
        End If
        errordesc = "ERROR AL ERASEUSERCHAR"
        'Borrar el personaje
        Call EraseUserChar(UserIndex, True)
        errordesc = "ERROR AL BORRAR MASCOTAS"
        'Borrar mascotas
        For i = 1 To MAXMASCOTAS
            If IsValidNpcRef(.MascotasIndex(i)) Then
                If NpcList(.MascotasIndex(i).ArrayIndex).flags.NPCActive Then Call QuitarNPC(.MascotasIndex(i).ArrayIndex, eClearPlayerPets)
            End If
            Call ClearNpcRef(.MascotasIndex(i))
        Next i
        errordesc = "ERROR Update Map Users map: " & Map
        'Update Map Users
        MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1
        Call Execute("update user set is_logged = 0 where id = ?;", UserList(UserIndex).Id)
        If MapInfo(Map).NumUsers < 0 Then MapInfo(Map).NumUsers = 0
        ' Si el usuario habia dejado un msg en la gm's queue lo borramos
        'If Ayuda.Existe(.Name) Then Call Ayuda.Quitar(.Name)
        errordesc = "ERROR AL m_NameIndex.Remove() Name:" & .name & " cuenta:" & .Cuenta
        Call m_NameIndex.Remove(UCase$(.name))
        errordesc = "ERROR AL RESETSLOT Name:" & .name & " cuenta:" & .Cuenta
        .flags.UserLogged = False
        .Counters.Saliendo = False
        Call ResetUserSlot(UserIndex)
    End With
    Exit Sub
ErrHandler:
    'Call LogError("Error en CloseUser. Número " & Err.Number & ". Descripción: " & Err.Description & ". Detalle:" & errordesc)
    Call TraceError(Err.Number, Err.Description & ". Detalle:" & errordesc, Erl)
    Resume Next ' TODO: Provisional hasta solucionar bugs graves
End Sub

Public Sub EcharPjsNoPrivilegiados()
    On Error GoTo EcharPjsNoPrivilegiados_Err
    Dim LoopC As Long
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnectionDetails.ConnIDValida Then
            If UserList(LoopC).flags.Privilegios And e_PlayerType.User Then
                Call CloseSocket(LoopC)
            End If
        End If
    Next LoopC
    Exit Sub
EcharPjsNoPrivilegiados_Err:
    Call TraceError(Err.Number, Err.Description, "TCP.EcharPjsNoPrivilegiados", Erl)
End Sub

Function ValidarCabeza(ByVal UserRaza As e_Raza, ByVal UserSexo As e_Genero, ByVal head As Integer) As Boolean
    Select Case UserSexo
        Case e_Genero.Hombre
            Select Case UserRaza
                Case e_Raza.Humano
                    ValidarCabeza = head >= 1 And head <= 41 Or head >= 778 And head <= 791
                Case e_Raza.Elfo
                    ValidarCabeza = head >= 101 And head <= 132 Or head >= 531 And head <= 545
                Case e_Raza.Drow
                    ValidarCabeza = head >= 200 And head <= 229 Or head >= 792 And head <= 810
                Case e_Raza.Enano
                    ValidarCabeza = head >= 300 And head <= 344
                Case e_Raza.Gnomo
                    ValidarCabeza = head >= 400 And head <= 429
                Case e_Raza.Orco
                    ValidarCabeza = head >= 500 And head <= 529
            End Select
        Case e_Genero.Mujer
            Select Case UserRaza
                Case e_Raza.Humano
                    ValidarCabeza = head >= 50 And head <= 80 Or head >= 187 And head <= 190 Or head >= 230 And head <= 246
                Case e_Raza.Elfo
                    ValidarCabeza = head >= 150 And head <= 179 Or head >= 758 And head <= 777
                Case e_Raza.Drow
                    ValidarCabeza = head >= 250 And head <= 279
                Case e_Raza.Enano
                    ValidarCabeza = head >= 350 And head <= 379
                Case e_Raza.Gnomo
                    ValidarCabeza = head >= 450 And head <= 479
                Case e_Raza.Orco
                    ValidarCabeza = head >= 550 And head <= 579
            End Select
    End Select
End Function

Function ValidarNombre(nombre As String) As Boolean
    If Len(nombre) < 3 Or Len(nombre) > 18 Then Exit Function
    Dim Temp As String
    Temp = UCase$(nombre)
    Dim i As Long, Char As Integer, LastChar As Integer
    For i = 1 To Len(Temp)
        Char = Asc(mid$(Temp, i, 1))
        If (Char < 65 Or Char > 90) And Char <> 32 Then
            Exit Function
        ElseIf Char = 32 And LastChar = 32 Then
            Exit Function
        End If
        LastChar = Char
    Next
    If Asc(mid$(Temp, 1, 1)) = 32 Or Asc(mid$(Temp, Len(Temp), 1)) = 32 Then
        Exit Function
    End If
    ValidarNombre = True
End Function

Function ContarUsuariosMismaCuenta(ByVal AccountID As Long) As Integer
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged And UserList(i).AccountID = AccountID Then
            ContarUsuariosMismaCuenta = ContarUsuariosMismaCuenta + 1
        End If
    Next
End Function

Sub ResetCd(ByRef User As t_User)
    Dim i As Integer
    For i = 0 To e_CdTypes.CDCount - 1
        User.CdTimes(i) = 0
    Next i
End Sub

Sub VaciarInventario(ByVal UserIndex As Integer)
    Dim i As Long
    With UserList(UserIndex)
        For i = 1 To MAX_INVENTORY_SLOTS
            .invent.Object(i).amount = 0
            .invent.Object(i).Equipped = 0
            .invent.Object(i).ObjIndex = 0
        Next i
    End With
End Sub
