Attribute VB_Name = "TCP"
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

            Case eGenero.Hombre

106             Select Case UserRaza

                    Case eRaza.Humano
108                     NewBody = 1

110                 Case eRaza.Elfo
112                     NewBody = 2

114                 Case eRaza.Drow
116                     NewBody = 3

118                 Case eRaza.Enano
120                     NewBody = 300

122                 Case eRaza.Gnomo
124                     NewBody = 300

126                 Case eRaza.Orco
128                     NewBody = 582

                End Select

130         Case eGenero.Mujer

132             Select Case UserRaza

                    Case eRaza.Humano
134                     NewBody = 1

136                 Case eRaza.Elfo
138                     NewBody = 2

140                 Case eRaza.Drow
142                     NewBody = 3

144                 Case eRaza.Gnomo
146                     NewBody = 300

148                 Case eRaza.Enano
150                     NewBody = 300

152                 Case eRaza.Orco
154                     NewBody = 581

                End Select

        End Select

156     UserList(UserIndex).Char.Body = NewBody

        
        Exit Sub

DarCuerpo_Err:
158     Call RegistrarError(Err.Number, Err.Description, "TCP.DarCuerpo", Erl)
160     Resume Next
        
End Sub

Sub AsignarAtributos(ByVal UserIndex As String)
        
        On Error GoTo AsignarAtributos_Err
        

100     Select Case UserList(UserIndex).raza

            Case eRaza.Humano
102             UserList(UserIndex).Stats.UserAtributos(1) = 19
104             UserList(UserIndex).Stats.UserAtributos(2) = 19
106             UserList(UserIndex).Stats.UserAtributos(3) = 19
108             UserList(UserIndex).Stats.UserAtributos(4) = 20

110         Case eRaza.Elfo
112             UserList(UserIndex).Stats.UserAtributos(1) = 18
114             UserList(UserIndex).Stats.UserAtributos(2) = 20
116             UserList(UserIndex).Stats.UserAtributos(3) = 21
118             UserList(UserIndex).Stats.UserAtributos(4) = 18

120         Case eRaza.Drow
122             UserList(UserIndex).Stats.UserAtributos(1) = 20
124             UserList(UserIndex).Stats.UserAtributos(2) = 18
126             UserList(UserIndex).Stats.UserAtributos(3) = 20
128             UserList(UserIndex).Stats.UserAtributos(4) = 19

130         Case eRaza.Gnomo
132             UserList(UserIndex).Stats.UserAtributos(1) = 13
134             UserList(UserIndex).Stats.UserAtributos(2) = 21
136             UserList(UserIndex).Stats.UserAtributos(3) = 22
138             UserList(UserIndex).Stats.UserAtributos(4) = 17

140         Case eRaza.Enano
142             UserList(UserIndex).Stats.UserAtributos(1) = 21
144             UserList(UserIndex).Stats.UserAtributos(2) = 17
146             UserList(UserIndex).Stats.UserAtributos(3) = 12
148             UserList(UserIndex).Stats.UserAtributos(4) = 22

150         Case eRaza.Orco
152             UserList(UserIndex).Stats.UserAtributos(1) = 23
154             UserList(UserIndex).Stats.UserAtributos(2) = 17
156             UserList(UserIndex).Stats.UserAtributos(3) = 12
158             UserList(UserIndex).Stats.UserAtributos(4) = 21

        End Select

        
        Exit Sub

AsignarAtributos_Err:
160     Call RegistrarError(Err.Number, Err.Description, "TCP.AsignarAtributos", Erl)
162     Resume Next
        
End Sub

Sub RellenarInventario(ByVal UserIndex As String)
        
        On Error GoTo RellenarInventario_Err
        

100     With UserList(UserIndex)
        
            Dim NumItems As Integer

102         NumItems = 1
    
            ' Todos reciben pociones rojas
104         .Invent.Object(NumItems).ObjIndex = 1616 'Pocion Roja
106         .Invent.Object(NumItems).amount = 100
108         NumItems = NumItems + 1
        
            ' Magicas puras reciben más azules
110         Select Case .clase

                Case eClass.Mage, eClass.Druid
112                 .Invent.Object(NumItems).ObjIndex = 1617 ' Pocion Azul
114                 .Invent.Object(NumItems).amount = 100
116                 NumItems = NumItems + 1

            End Select
        
            ' Semi mágicas reciben menos
118         Select Case .clase

                Case eClass.Bard, eClass.Cleric, eClass.Paladin, eClass.Assasin, eClass.Bandit
120                 .Invent.Object(NumItems).ObjIndex = 1617 ' Pocion Azul
122                 .Invent.Object(NumItems).amount = 50
124                 NumItems = NumItems + 1

            End Select

            ' Arma y hechizos
126         Select Case .clase

                Case eClass.Mage, eClass.Cleric, eClass.Druid, eClass.Bard
128                 .Stats.UserHechizos(1) = 1 ' Dardo mágico

            End Select
        
            ' Pociones amarillas y verdes
130         Select Case .clase

                Case eClass.Assasin, eClass.Bard, eClass.Cleric, eClass.Hunter, eClass.Paladin, eClass.Trabajador, eClass.Warrior, eClass.Bandit, eClass.Pirat, eClass.Thief
132                 .Invent.Object(NumItems).ObjIndex = 1618 ' Pocion Amarilla
134                 .Invent.Object(NumItems).amount = 25
136                 NumItems = NumItems + 1

138                 .Invent.Object(NumItems).ObjIndex = 1619 ' Pocion Verde
140                 .Invent.Object(NumItems).amount = 25
142                 NumItems = NumItems + 1

            End Select
            
            ' Poción violeta
144         .Invent.Object(NumItems).ObjIndex = 2332 ' Pocion violeta
146         .Invent.Object(NumItems).amount = 10
148         NumItems = NumItems + 1
        
            ' Equipo el arma
150         .Invent.Object(NumItems).ObjIndex = 460 ' Daga (Newbies)
152         .Invent.Object(NumItems).amount = 1
154         .Invent.Object(NumItems).Equipped = 1
156         .Invent.WeaponEqpSlot = NumItems
158         .Invent.WeaponEqpObjIndex = .Invent.Object(NumItems).ObjIndex
160         .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
162         NumItems = NumItems + 1
        
            
164         If .genero = eGenero.Hombre Then
166             If .raza = Enano Or .raza = Gnomo Then
168                 .Invent.Object(NumItems).ObjIndex = 466 'Vestimentas de Bajo (Newbies)
                Else
                
170                 .Invent.Object(NumItems).ObjIndex = RandomNumber(463, 465) ' Vestimentas comunes (Newbies)
                End If
            Else
172             If .raza = Enano Or .raza = Gnomo Then
174                 .Invent.Object(NumItems).ObjIndex = 563 'Vestimentas de Baja (Newbies)
                Else
176                 .Invent.Object(NumItems).ObjIndex = RandomNumber(1283, 1285) ' Vestimentas de Mujer (Newbies)
                End If
            End If
                        
178         .Invent.Object(NumItems).amount = 1
180         .Invent.Object(NumItems).Equipped = 1
182         .Invent.ArmourEqpSlot = NumItems
184         .Invent.ArmourEqpObjIndex = .Invent.Object(NumItems).ObjIndex
186          NumItems = NumItems + 1

            ' Animación según raza

188          .Char.Body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
        
            ' Comida y bebida
190         .Invent.Object(NumItems).ObjIndex = 573 ' Manzana
192         .Invent.Object(NumItems).amount = 100
194         NumItems = NumItems + 1

196         .Invent.Object(NumItems).ObjIndex = 572 ' Agua
198         .Invent.Object(NumItems).amount = 100
200         NumItems = NumItems + 1

202         .Invent.Object(NumItems).ObjIndex = 200 ' Cofre Inicial - TODO: Remover
204         .Invent.Object(NumItems).amount = 1
206         NumItems = NumItems + 1

            ' Seteo la cantidad de items
208         .Invent.NroItems = NumItems

        End With
   
        
        Exit Sub

RellenarInventario_Err:
210     Call RegistrarError(Err.Number, Err.Description, "TCP.RellenarInventario", Erl)
212     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.Description, "TCP.AsciiValidos", Erl)
116     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.Description, "TCP.DescripcionValida", Erl)
116     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.Description, "TCP.Numeric", Erl)
116     Resume Next
        
End Function

Function NombrePermitido(ByVal nombre As String) As Boolean
        
        On Error GoTo NombrePermitido_Err
        

        Dim i As Integer

100     For i = 1 To UBound(ForbidenNames)

102         If InStr(nombre, ForbidenNames(i)) Then
104             NombrePermitido = False
                Exit Function

            End If

106     Next i

108     NombrePermitido = True

        
        Exit Function

NombrePermitido_Err:
110     Call RegistrarError(Err.Number, Err.Description, "TCP.NombrePermitido", Erl)
112     Resume Next
        
End Function

Function ValidateSkills(ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo ValidateSkills_Err
        

        Dim LoopC As Integer

100     For LoopC = 1 To NUMSKILLS

102         If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then
                Exit Function

104             If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100

            End If

106     Next LoopC

108     ValidateSkills = True
    
        
        Exit Function

ValidateSkills_Err:
110     Call RegistrarError(Err.Number, Err.Description, "TCP.ValidateSkills", Erl)
112     Resume Next
        
End Function

Function ConnectNewUser(ByVal UserIndex As Integer, ByRef name As String, ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, ByVal UserClase As eClass, ByVal Head As Integer, ByRef UserCuenta As String, ByVal Hogar As eCiudad) As Boolean
        '*************************************************
        'Author: Unknown
        'Last modified: 20/4/2007
        'Conecta un nuevo Usuario
        '23/01/2007 Pablo (ToxicWaste) - Agregué ResetFaccion al crear usuario
        '24/01/2007 Pablo (ToxicWaste) - Agregué el nuevo mana inicial de los magos.
        '12/02/2007 Pablo (ToxicWaste) - Puse + 1 de const al Elfo normal.
        '20/04/2007 Pablo (ToxicWaste) - Puse -1 de fuerza al Elfo.
        '09/01/2008 Pablo (ToxicWaste) - Ahora los modificadores de Raza se controlan desde Balance.dat
        '*************************************************
        
        On Error GoTo ConnectNewUser_Err
        
100     With UserList(UserIndex)
        
            Dim LoopC As Long
        
102         If .flags.UserLogged Then
104             Call LogCheating("El usuario " & .name & " ha intentado crear a " & name & " desde la IP " & .ip)
106             Call CloseSocketSL(UserIndex)
108             Call Cerrar_Usuario(UserIndex)
                Exit Function
            End If
            
            ' Nombre válido
110         If Not ValidarNombre(name) Then Exit Function
            
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
128         If .Stats.UserAtributos(eAtributos.Fuerza) = 0 Then Exit Function
        
130         .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + ModRaza(UserRaza).Fuerza
132         .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + ModRaza(UserRaza).Agilidad
134         .Stats.UserAtributos(eAtributos.Inteligencia) = .Stats.UserAtributos(eAtributos.Inteligencia) + ModRaza(UserRaza).Inteligencia
136         .Stats.UserAtributos(eAtributos.Constitucion) = .Stats.UserAtributos(eAtributos.Constitucion) + ModRaza(UserRaza).Constitucion
138         .Stats.UserAtributos(eAtributos.Carisma) = .Stats.UserAtributos(eAtributos.Carisma) + ModRaza(UserRaza).Carisma
        
140         .flags.Muerto = 0
142         .flags.Escondido = 0
    
144         .flags.Casado = 0
146         .flags.Pareja = ""
    
148         .name = name

150         .clase = UserClase
152         .raza = UserRaza
        
154         .Char.Head = Head
        
156         .genero = UserSexo
158         .Hogar = Hogar
        
            '%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%
160         .Stats.SkillPts = 10
        
162         .Char.Heading = eHeading.SOUTH
        
164         Call DarCuerpo(UserIndex) 'Ladder REVISAR
        
166         .OrigChar = .Char
    
168         .Char.WeaponAnim = NingunArma
170         .Char.ShieldAnim = NingunEscudo
172         .Char.CascoAnim = NingunCasco

            ' WyroX: Vida inicial
174         .Stats.MaxHp = .Stats.UserAtributos(eAtributos.Constitucion)
176         .Stats.MinHp = .Stats.MaxHp

            ' WyroX: Maná inicial
178         .Stats.MaxMAN = .Stats.UserAtributos(eAtributos.Inteligencia) * ModClase(.clase).ManaInicial
180         .Stats.MinMAN = .Stats.MaxMAN
        
            Dim MiInt As Integer
182         MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Agilidad) \ 6)
    
184         If MiInt = 1 Then MiInt = 2
        
186         .Stats.MaxSta = 20 * MiInt
188         .Stats.MinSta = 20 * MiInt
        
190         .Stats.MaxAGU = 100
192         .Stats.MinAGU = 100
        
194         .Stats.MaxHam = 100
196         .Stats.MinHam = 100
    
198         .flags.ScrollExp = 1
200         .flags.ScrollOro = 1
    
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
        
            'Valores Default de facciones al Activar nuevo usuario
222         Call ResetFacciones(UserIndex)
        
224         .Faccion.Status = 1
        
226         .ChatCombate = 1
228         .ChatGlobal = 1
        
            'Resetamos CORREO
230         .Correo.CantCorreo = 0
232         .Correo.NoLeidos = 0
            'Resetamos CORREO
        
234         .Pos.Map = 37
236         .Pos.X = 76
238         .Pos.Y = 82
        
240         UltimoChar = UCase$(name)
        
242         Call SaveNewUser(UserIndex)
    
244         ConnectNewUser = True
    
246         Call ConnectUser(UserIndex, name, UserCuenta)

        End With
        
        Exit Function

ConnectNewUser_Err:
248     Call RegistrarError(Err.Number, Err.Description, "TCP.ConnectNewUser", Erl)
250     Resume Next
        
End Function

Sub CloseSocket(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

100     Call FlushBuffer(UserIndex)

102     If UserIndex = LastUser Then

104         Do Until UserList(LastUser).flags.UserLogged
106             LastUser = LastUser - 1

108             If LastUser < 1 Then Exit Do
            Loop

        End If
    
110     With UserList(UserIndex)
    
            'Call SecurityIp.IpRestarConexion(GetLongIp(.ip))

112         If .ConnID <> -1 Then Call CloseSocketSL(UserIndex)
    
            'Es el mismo user al que está revisando el centinela??
            'IMPORTANTE!!! hacerlo antes de resetear así todavía sabemos el nombre del user
            ' y lo podemos loguear
114         If Centinela.RevisandoUserIndex = UserIndex Then Call modCentinela.CentinelaUserLogout
    
            'mato los comercios seguros
116         If .ComUsu.DestUsu > 0 Then
        
118             If UserList(.ComUsu.DestUsu).flags.UserLogged Then
            
120                 If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                
122                     Call WriteConsoleMsg(.ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
124                     Call FinComerciarUsu(.ComUsu.DestUsu)
                    
                    End If
    
                End If
    
            End If
    
            'Empty buffer for reuse
126         Call .incomingData.ReadASCIIStringFixed(.incomingData.Length)
    
128         If .flags.UserLogged Then
130             Call CloseUser(UserIndex)
        
132             If NumUsers > 0 Then NumUsers = NumUsers - 1
134             Call MostrarNumUsers
        
            Else
136             Call ResetUserSlot(UserIndex)
    
            End If
    
138         .ConnID = -1
140         .ConnIDValida = False
142         .NumeroPaquetesPorMiliSec = 0
    
        End With
    

        Exit Sub

ErrHandler:
144     UserList(UserIndex).ConnID = -1
146     UserList(UserIndex).ConnIDValida = False
148     UserList(UserIndex).NumeroPaquetesPorMiliSec = 0

150     Call ResetUserSlot(UserIndex)

152     Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.Description & " - UserIndex = " & UserIndex)
154     Resume Next

End Sub

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal UserIndex As Integer)
        
        On Error GoTo CloseSocketSL_Err

100     If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
102         Call BorraSlotSock(UserList(UserIndex).ConnID)
104         Call WSApiCloseSocket(UserList(UserIndex).ConnID)
106         UserList(UserIndex).ConnIDValida = False

        End If
        
        Exit Sub

CloseSocketSL_Err:
108     Call RegistrarError(Err.Number, Err.Description, "TCP.CloseSocketSL", Erl)

110     Resume Next
        
End Sub

''
' Send an string to a Slot
'
' @param userIndex The index of the User
' @param Datos The string that will be send

Public Sub EnviarDatosASlot(ByVal UserIndex As Integer, ByRef Datos As String)
            '***************************************************
            'Author: Unknown
            'Last Modification: 09/11/20
            'Modified By: Jopi
            'Last Modified by: WyroX - Si no hay espacio, flusheo el buffer e intento de nuevo
            'Se agrega el paquete a la cola, para prevenir errores.
            '***************************************************
        
            On Error GoTo EnviarDatosASlot_Err
        

100         Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(Datos)

            Exit Sub

EnviarDatosASlot_Err:
102         If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
104             Call FlushBuffer(UserIndex)
106             Resume
            End If
        
End Sub

Function EstaPCarea(index As Integer, Index2 As Integer) As Boolean
        
        On Error GoTo EstaPCarea_Err
        

        Dim X As Integer, Y As Integer

100     For Y = UserList(index).Pos.Y - MinYBorder + 1 To UserList(index).Pos.Y + MinYBorder - 1
102         For X = UserList(index).Pos.X - MinXBorder + 1 To UserList(index).Pos.X + MinXBorder - 1

104             If MapData(UserList(index).Pos.Map, X, Y).UserIndex = Index2 Then
106                 EstaPCarea = True
                    Exit Function

                End If
        
108         Next X
110     Next Y

112     EstaPCarea = False

        
        Exit Function

EstaPCarea_Err:
114     Call RegistrarError(Err.Number, Err.Description, "TCP.EstaPCarea", Erl)
116     Resume Next
        
End Function

Function HayPCarea(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        
        On Error GoTo HayPCarea_Err
        

        Dim tX As Integer, tY As Integer

100     For tY = Y - MinYBorder + 1 To Y + MinYBorder - 1
102         For tX = X - MinXBorder + 1 To X + MinXBorder - 1

104             If InMapBounds(Map, tX, tY) Then
106                 If MapData(Map, tX, tY).UserIndex > 0 Then
108                     HayPCarea = True
                        Exit Function

                    End If

                End If

            Next
        Next

110     HayPCarea = False

        
        Exit Function

HayPCarea_Err:
112     Call RegistrarError(Err.Number, Err.Description, "TCP.HayPCarea", Erl)
114     Resume Next
        
End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean
        
        On Error GoTo HayOBJarea_Err
        

        Dim X As Integer, Y As Integer

100     For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
102         For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1

104             If MapData(Pos.Map, X, Y).ObjInfo.ObjIndex = ObjIndex Then
106                 HayOBJarea = True
                    Exit Function

                End If
        
108         Next X
110     Next Y

112     HayOBJarea = False

        
        Exit Function

HayOBJarea_Err:
114     Call RegistrarError(Err.Number, Err.Description, "TCP.HayOBJarea", Erl)
116     Resume Next
        
End Function

Function ValidateChr(ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo ValidateChr_Err
        

100     ValidateChr = UserList(UserIndex).Char.Body <> 0 And ValidateSkills(UserIndex)

        
        Exit Function

ValidateChr_Err:
102     Call RegistrarError(Err.Number, Err.Description, "TCP.ValidateChr", Erl)
104     Resume Next
        
End Function

Function EntrarCuenta(ByVal UserIndex As Integer, CuentaEmail As String, CuentaPassword As String, MacAddress As String, ByVal HDserial As Long, MD5 As String) As Boolean
        
        On Error GoTo EntrarCuenta_Err
        
        Dim adminIdx As Long
        Dim laCuentaEsDeAdmin As Boolean
        
        
100     If ServerSoloGMs > 0 Then
102         laCuentaEsDeAdmin = False
            
104         For adminIdx = 1 To AdministratorAccounts.Count
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

116     If CheckMAC(MacAddress) Then
118         Call WriteShowMessageBox(UserIndex, "Su cuenta se encuentra bajo tolerancia 0. Tiene prohibido el acceso. Cod: #0001")
            Exit Function
        End If
    
120     If CheckHD(HDserial) Then
122         Call WriteShowMessageBox(UserIndex, "Su cuenta se encuentra bajo tolerancia 0. Tiene prohibido el acceso. Cod: #0002")
            Exit Function
        End If

124     If Md5Cliente <> vbNullString And LCase$(Md5Cliente) <> LCase$(MD5) Then
126         Call WriteShowMessageBox(UserIndex, "Error al comprobar el cliente del juego, por favor reinstale y vuelva a intentar.")
            Exit Function
        End If

128     If Not CheckMailString(CuentaEmail) Then
130         Call WriteShowMessageBox(UserIndex, "Email inválido.")
            Exit Function
        End If
    
132     EntrarCuenta = EnterAccountDatabase(UserIndex, CuentaEmail, SDesencriptar(CuentaPassword), MacAddress, HDserial, UserList(UserIndex).ip)
        
        Exit Function

EntrarCuenta_Err:
134     Call RegistrarError(Err.Number, Err.Description, "TCP.EntrarCuenta", Erl)

136     Resume Next
        
End Function

Sub ConnectUser(ByVal UserIndex As Integer, _
                ByRef name As String, _
                ByRef UserCuenta As String)

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim n    As Integer

            Dim tStr As String
        
102         If .flags.UserLogged Then
104             Call LogCheating("El usuario " & .name & " ha intentado loguear a " & name & " desde la IP " & .ip)
            
                'Kick player ( and leave character inside :D )!
106             Call CloseSocketSL(UserIndex)
108             Call Cerrar_Usuario(UserIndex)
            
                Exit Sub

            End If
            
            '¿Ya esta conectado el personaje?
            Dim tIndex As Integer

110         tIndex = NameIndex(name)

112         If tIndex > 0 And tIndex <> UserIndex Then
114             If UserList(tIndex).Counters.Saliendo Then
116                 Call WriteShowMessageBox(UserIndex, "El personaje está saliendo.")
                Else
118                 Call WriteShowMessageBox(UserIndex, "El personaje ya está conectado. Espere mientras es desconectado.")

                    ' Le avisamos al usuario que está jugando, en caso de que haya uno
120                 Call WriteShowMessageBox(tIndex, "Alguien está ingresando con tu personaje. Si no has sido tú, por favor cambia la contraseña de tu cuenta.")
122                 Call Cerrar_Usuario(tIndex)

                End If
            
124             Call CloseSocket(UserIndex)
                Exit Sub

            End If
        
            '¿Supera el máximo de usuarios por cuenta?
126         If MaxUsersPorCuenta > 0 Then
128             If ContarUsuariosMismaCuenta(.AccountId) >= MaxUsersPorCuenta Then
130                 If MaxUsersPorCuenta = 1 Then
132                     Call WriteShowMessageBox(UserIndex, "Ya hay un usuario conectado con esta cuenta.")
                    Else
134                     Call WriteShowMessageBox(UserIndex, "La cuenta ya alcanzó el máximo de " & MaxUsersPorCuenta & " usuarios conectados.")

                    End If

136                 Call CloseSocket(UserIndex)
                    Exit Sub

                End If

            End If
        
            'Reseteamos los FLAGS
138         .flags.Escondido = 0
140         .flags.TargetNPC = 0
142         .flags.TargetNpcTipo = eNPCType.Comun
144         .flags.TargetObj = 0
146         .flags.TargetUser = 0
148         .Char.FX = 0
150         .Counters.CuentaRegresiva = -1
        
            'Controlamos no pasar el maximo de usuarios
152         If NumUsers >= MaxUsers Then
154             Call WriteShowMessageBox(UserIndex, "El servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
156             Call CloseSocket(UserIndex)
                Exit Sub

            End If
        
            '¿Este IP ya esta conectado?
158         If MaxConexionesIP > 0 Then
160             If ContarMismaIP(UserIndex, .ip) >= MaxConexionesIP Then
162                 Call WriteShowMessageBox(UserIndex, "Has alcanzado el límite de conexiones por IP.")
164                 Call CloseSocket(UserIndex)
                    Exit Sub

                End If

            End If

            'Le damos los privilegios
166         .flags.Privilegios = UserDarPrivilegioLevel(name)

            'Add RM flag if needed
168         If EsRolesMaster(name) Then
170             .flags.Privilegios = .flags.Privilegios Or PlayerType.RoleMaster

            End If
        
172         If EsGM(UserIndex) Then
174             Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & name & " se conecto al juego.", FontTypeNames.FONTTYPE_INFOBOLD))
176             Call LogGM(.name, "Se conectó con IP: " & .ip)

            Else

178             If ServerSoloGMs > 0 Then
                    ' Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
180                 Call WriteShowMessageBox(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
182                 Call CloseSocket(UserIndex)
                    Exit Sub

                End If

            End If
        
184         If EnPausa Then
186             Call WritePauseToggle(UserIndex)
188             Call WriteConsoleMsg(UserIndex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER)
190             Call CloseSocket(UserIndex)
                Exit Sub

            End If
    
            'Donador
192         If DonadorCheck(UserCuenta) Then

                Dim LoopC As Integer

194             For LoopC = 1 To Donadores.Count

196                 If UCase$(Donadores(LoopC).name) = UCase$(UserCuenta) Then
198                     .donador.activo = 1
200                     .donador.FechaExpiracion = Donadores(LoopC).FechaExpiracion
                        Exit For

                    End If

202             Next LoopC

            End If
        
            ' Seteamos el nombre
204         .name = name
            
206         m_NameIndex(UCase$(name)) = UserIndex

208         .showName = True
        
            ' Cargamos el personaje
210         Call LoadUser(UserIndex)

212         If Not ValidateChr(UserIndex) Then
214             Call WriteShowMessageBox(UserIndex, "Error en el personaje. Comuniquese con el staff.")
216             Call CloseSocket(UserIndex)
                Exit Sub

            End If
    
218         If UCase$(.Cuenta) <> UCase$(UserCuenta) Then
220             Call WriteShowMessageBox(UserIndex, "El personaje no corresponde a su cuenta.")
222             Call CloseSocket(UserIndex)
                Exit Sub

            End If
        
224         If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
226         If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
228         If .Invent.WeaponEqpSlot = 0 And .Invent.NudilloSlot = 0 And .Invent.HerramientaEqpSlot = 0 Then .Char.WeaponAnim = NingunArma
        
230         .flags.SeguroParty = True
232         .flags.SeguroClan = True
234         .flags.SeguroResu = True
        
236         .CurrentInventorySlots = getMaxInventorySlots(UserIndex)
        
238         Call WriteInventoryUnlockSlots(UserIndex)
        
240         Call LoadUserIntervals(UserIndex)
242         Call WriteIntervals(UserIndex)
        
244         Call UpdateUserInv(True, UserIndex, 0)
246         Call UpdateUserHechizos(True, UserIndex, 0)
        
248         Call EnviarLlaves(UserIndex)

250         If .Correo.NoLeidos > 0 Then
252             Call WriteCorreoPicOn(UserIndex)

            End If

254         If .flags.Paralizado Then
256             Call WriteParalizeOK(UserIndex)

            End If
        
258         If .flags.Inmovilizado Then
260             Call WriteInmovilizaOK(UserIndex)

            End If
        
            ''
            'TODO : Feo, esto tiene que ser parche cliente
262         If .flags.Estupidez = 0 Then
264             Call WriteDumbNoMore(UserIndex)

            End If
        
            'Ladder Inmunidad
266         .flags.Inmunidad = 1
268         .Counters.TiempoDeInmunidad = IntervaloPuedeSerAtacado
            'Ladder Inmunidad
        
            'Mapa válido
270         If Not MapaValido(.Pos.Map) Then
272             Call WriteErrorMsg(UserIndex, "EL PJ se encuenta en un mapa invalido.")
274             Call CloseSocket(UserIndex)
                Exit Sub

            End If
        
            'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
            'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martin Sotuyo Dodero (Maraxus)
276         If MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex <> 0 Or MapData(.Pos.Map, .Pos.X, .Pos.Y).NpcIndex <> 0 Then

                Dim FoundPlace As Boolean

                Dim esAgua     As Boolean

                Dim tX         As Long

                Dim tY         As Long
        
278             FoundPlace = False
280             esAgua = (MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked And FLAG_AGUA) <> 0
        
282             For tY = .Pos.Y - 1 To .Pos.Y + 1
284                 For tX = .Pos.X - 1 To .Pos.X + 1

286                     If esAgua Then

                            'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
288                         If LegalPos(.Pos.Map, tX, tY, True, False) Then
290                             FoundPlace = True
                                Exit For

                            End If

                        Else

                            'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
292                         If LegalPos(.Pos.Map, tX, tY, False, True) Then
294                             FoundPlace = True
                                Exit For

                            End If

                        End If

296                 Next tX
            
298                 If FoundPlace Then Exit For
300             Next tY
        
302             If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
304                 .Pos.X = tX
306                 .Pos.Y = tY
                Else

                    'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
308                 If MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex <> 0 Then

                        'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
310                     If UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu > 0 Then

                            'Le avisamos al que estaba comerciando que se tuvo que ir.
312                         If UserList(UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu).flags.UserLogged Then
314                             Call FinComerciarUsu(UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu)
316                             Call WriteConsoleMsg(UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_WARNING)

                            End If

                            'Lo sacamos.
318                         If UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).flags.UserLogged Then
320                             Call FinComerciarUsu(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex)
322                             Call WriteErrorMsg(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconectate...")

                            End If

                        End If
                
324                     Call CloseSocket(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex)

                    End If

                End If

            End If
        
            'If in the water, and has a boat, equip it!
326         If .Invent.BarcoObjIndex > 0 And (MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked And FLAG_AGUA) <> 0 Then
328             .flags.Navegando = 1
330             Call EquiparBarco(UserIndex)

            End If
            
332         If .Invent.MagicoObjIndex <> 0 Then
334             If ObjData(.Invent.MagicoObjIndex).EfectoMagico = 11 Then .flags.Paraliza = 1
            End If

336         Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
338         .flags.NecesitaOxigeno = RequiereOxigeno(.Pos.Map)
        
340         Call WriteHora(UserIndex)
342         Call WriteChangeMap(UserIndex, .Pos.Map) 'Carga el mapa
        
            'If .flags.Privilegios <> PlayerType.user And .flags.Privilegios <> (PlayerType.user Or PlayerType.ChaosCouncil) And .flags.Privilegios <> (PlayerType.user Or PlayerType.RoyalCouncil) And .flags.Privilegios <> (PlayerType.user Or PlayerType.Admin) And .flags.Privilegios <> (PlayerType.user Or PlayerType.Dios) Then
            ' .flags.ChatColor = RGB(2, 161, 38)
            'ElseIf .flags.Privilegios = (PlayerType.user Or PlayerType.RoyalCouncil) Then
            ' .flags.ChatColor = RGB(0, 255, 255)
344         If .flags.Privilegios = PlayerType.Admin Then
346             .flags.ChatColor = RGB(217, 164, 32)
348         ElseIf .flags.Privilegios = PlayerType.Dios Then
350             .flags.ChatColor = RGB(217, 164, 32)
352         ElseIf .flags.Privilegios = PlayerType.SemiDios Then
354             .flags.ChatColor = RGB(2, 161, 38)
356         ElseIf .flags.Privilegios = PlayerType.Consejero Then
358             .flags.ChatColor = RGB(2, 161, 38)
            Else
360             .flags.ChatColor = vbWhite

            End If
        
            ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
            #If ConUpTime Then
362             .LogOnTime = Now
            #End If
        
            'Crea  el personaje del usuario
364         Call MakeUserChar(True, .Pos.Map, UserIndex, .Pos.Map, .Pos.X, .Pos.Y, 1)

366         Call WriteUserCharIndexInServer(UserIndex)

368         Call ActualizarVelocidadDeUsuario(UserIndex)
        
370         If (.flags.Privilegios And PlayerType.user) = 0 Then
372             Call DoAdminInvisible(UserIndex)

            End If
        
374         Call WriteUpdateUserStats(UserIndex)
        
376         Call WriteUpdateHungerAndThirst(UserIndex)
        
378         Call WriteUpdateDM(UserIndex)
380         Call WriteUpdateRM(UserIndex)
        
382         Call SendMOTD(UserIndex)
        
384         Call SetUserLogged(UserIndex)
        
            'Actualiza el Num de usuarios
386         NumUsers = NumUsers + 1
388         .flags.UserLogged = True
390         .Counters.LastSave = GetTickCount
        
392         MapInfo(.Pos.Map).NumUsers = MapInfo(.Pos.Map).NumUsers + 1
        
394         If .Stats.SkillPts > 0 Then
396             Call WriteSendSkills(UserIndex)
398             Call WriteLevelUp(UserIndex, .Stats.SkillPts)

            End If
        
400         If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers
        
402         If NumUsers > RecordUsuarios Then
404             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Record de usuarios conectados simultáneamente: " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_INFO))
406             RecordUsuarios = NumUsers
            
408             If Database_Enabled Then
410                 Call SetRecordUsersDatabase(RecordUsuarios)
                Else
412                 Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(RecordUsuarios))

                End If

            End If
        
414         Call WriteFYA(UserIndex)
416         Call WriteBindKeys(UserIndex)
        
418         If .NroMascotas > 0 And MapInfo(.Pos.Map).Seguro = 0 And .flags.MascotasGuardadas = 0 Then

                Dim i As Integer

420             For i = 1 To MAXMASCOTAS

422                 If .MascotasType(i) > 0 Then
424                     .MascotasIndex(i) = SpawnNpc(.MascotasType(i), .Pos, False, False, False, UserIndex)
                    
426                     If .MascotasIndex(i) > 0 Then
428                         NpcList(.MascotasIndex(i)).MaestroUser = UserIndex
430                         Call FollowAmo(.MascotasIndex(i))
                        Else
432                         .MascotasIndex(i) = 0

                        End If

                    End If

434             Next i

            End If
        
436         If .flags.Navegando = 1 Then
438             Call WriteNavigateToggle(UserIndex)
440             Call EquiparBarco(UserIndex)

            End If
        
442         If .flags.Montado = 1 Then
444             Call WriteEquiteToggle(UserIndex)

            End If

446         Call ActualizarVelocidadDeUsuario(UserIndex)
        
448         If .GuildIndex > 0 Then

                'welcome to the show baby...
450             If Not modGuilds.m_ConectarMiembroAClan(UserIndex, .GuildIndex) Then
452                 Call WriteConsoleMsg(UserIndex, "Tu estado no te permite entrar al clan.", FontTypeNames.FONTTYPE_GUILD)

                End If

            End If
        
454         tStr = modGuilds.a_ObtenerRechazoDeChar(.name)
    
456         If LenB(tStr) <> 0 Then
458             Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)

            End If

460         If Lloviendo Then
462             Call WriteRainToggle(UserIndex)

            End If
        
464         If ServidorNublado Then
466             Call WriteNubesToggle(UserIndex)

            End If

468         Call WriteLoggedMessage(UserIndex)
        
470         If .Stats.ELV = 1 Then
472             Call WriteConsoleMsg(UserIndex, "¡Bienvenido a las tierras de AO20! ¡" & .name & " que tengas buen viaje y mucha suerte!", FontTypeNames.FONTTYPE_GUILD)
474         ElseIf .Stats.ELV < 14 Then
476             Call WriteConsoleMsg(UserIndex, "¡Bienvenido de nuevo " & .name & "! Actualmente estas en el nivel " & .Stats.ELV & " en " & DarNameMapa(.Pos.Map) & ", ¡buen viaje y mucha suerte!", FontTypeNames.FONTTYPE_GUILD)

            End If

478         If Status(UserIndex) = 2 Or Status(UserIndex) = 0 Then
480             Call WriteSafeModeOff(UserIndex)
482             .flags.Seguro = False
            Else
484             .flags.Seguro = True
486             Call WriteSafeModeOn(UserIndex)

            End If
        
            'Call modGuilds.SendGuildNews(UserIndex)
        
488         If .MENSAJEINFORMACION <> vbNullString Then
490             Call WriteConsoleMsg(UserIndex, .MENSAJEINFORMACION, FontTypeNames.FONTTYPE_CENTINELA)
492             .MENSAJEINFORMACION = vbNullString

            End If

494         tStr = modGuilds.a_ObtenerRechazoDeChar(.name)
        
496         If LenB(tStr) <> 0 Then
498             Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)

            End If

500         If EventoActivo Then
502             Call WriteConsoleMsg(UserIndex, PublicidadEvento & ". Tiempo restante: " & TiempoRestanteEvento & " minuto(s).", FontTypeNames.FONTTYPE_New_Eventos)

            End If
        
504         Call WriteContadores(UserIndex)
506         Call WriteOxigeno(UserIndex)

            'Call SendData(UserIndex, UserIndex, PrepareMessageParticleFXToFloor(.Pos.x, .Pos.y, 209, 10))
            'Call SendData(UserIndex, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, 209, 40, False))
        
            'Load the user statistics
            'Call Statistics.UserConnected(UserIndex)

508         Call MostrarNumUsers
            'Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageParticleFXToFloor(.Pos.X, .Pos.y, ParticulasIndex.LogeoLevel1, 400))
            'Call SaveUser(UserIndex, CharPath & UCase$(.name) & ".chr")
        
            ' n = FreeFile
            ' Open App.Path & "\logs\numusers.log" For Output As n
            'Print #n, NumUsers
            ' Close #n

        End With
    
        Exit Sub
    
ErrHandler:
510     Call RegistrarError(Err.Number, Err.Description, "TCP.ConnectUser", Erl)
512     Call WriteShowMessageBox(UserIndex, "El personaje contiene un error. Comuníquese con un miembro del staff.")
514     Call CloseSocket(UserIndex)

End Sub

Sub SendMOTD(ByVal UserIndex As Integer)
        
        On Error GoTo SendMOTD_Err
        

        Dim j As Long

100     For j = 1 To MaxLines
102         Call WriteConsoleMsg(UserIndex, MOTD(j).texto, FontTypeNames.FONTTYPE_EXP)
104     Next j
    
        
        Exit Sub

SendMOTD_Err:
106     Call RegistrarError(Err.Number, Err.Description, "TCP.SendMOTD", Erl)
108     Resume Next
        
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
102         .ArmadaReal = 0
104         .ciudadanosMatados = 0
106         .CriminalesMatados = 0
108         .Status = 0
110         .FuerzasCaos = 0
112         .RecibioArmaduraCaos = 0
114         .RecibioArmaduraReal = 0
116         .RecibioExpInicialCaos = 0
118         .RecibioExpInicialReal = 0
120         .RecompensasCaos = 0
122         .RecompensasReal = 0
124         .Reenlistadas = 0
126         .NivelIngreso = 0
128         .MatadosIngreso = 0
130         .NextRecompensa = 0

        End With

        
        Exit Sub

ResetFacciones_Err:
132     Call RegistrarError(Err.Number, Err.Description, "TCP.ResetFacciones", Erl)
134     Resume Next
        
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
120         .Paralisis = 0
122         .Inmovilizado = 0
124         .Pasos = 0
126         .Pena = 0
128         .PiqueteC = 0
130         .STACounter = 0
132         .Veneno = 0
134         .Trabajando = 0
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
            'Ladder
162         .Incineracion = 0
            'Ladder
164         .ScrollExperiencia = 0
166         .ScrollOro = 0
168         .Oxigeno = 0
170         .TiempoParaSubastar = 0
172         .TimerPerteneceNpc = 0
174         .TimerPuedeSerAtacado = 0
176         .TiempoDeInmunidad = 0
178         .RepetirMensaje = 0
180         .MensajeGlobal = 0
182         .CuentaRegresiva = -1
184         .SpeedHackCounter = 0
186         .LastStep = 0
        End With

        
        Exit Sub

ResetContadores_Err:
188     Call RegistrarError(Err.Number, Err.Description, "TCP.ResetContadores", Erl)
190     Resume Next
        
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
140     Call RegistrarError(Err.Number, Err.Description, "TCP.ResetCharInfo", Erl)
142     Resume Next
        
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
106         .Id = -1
108         .AccountId = -1
110         .Desc = vbNullString
112         .DescRM = vbNullString
114         .Pos.Map = 0
116         .Pos.X = 0
118         .Pos.Y = 0
120         .ip = vbNullString
122         .clase = 0
124         .email = vbNullString
126         .genero = 0
128         .Hogar = 0
130         .raza = 0
132         .EmpoCont = 0
        
            'Ladder     Reseteo de Correos
134         .Correo.CantCorreo = 0
136         .Correo.NoLeidos = 0
        
138         For LoopC = 1 To MAX_CORREOS_SLOTS
140             .Correo.Mensaje(LoopC).Remitente = ""
142             .Correo.Mensaje(LoopC).Mensaje = ""
144             .Correo.Mensaje(LoopC).Item = 0
146             .Correo.Mensaje(LoopC).ItemCount = 0
148             .Correo.Mensaje(LoopC).Fecha = ""
150             .Correo.Mensaje(LoopC).Leido = 0
152         Next LoopC

            'Ladder     Reseteo de Correos
        
154         With .Stats
156             .InventLevel = 0
158             .Banco = 0
160             .ELV = 0
162             .Exp = 0
164             .def = 0
                '.CriminalesMatados = 0
166             .NPCsMuertos = 0
168             .UsuariosMatados = 0
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
            
            End With
            
194         .NroMascotas = 0

            #If AntiExternos Then
196             .Redundance = 0
            #End If

        End With

        
        Exit Sub

ResetBasicUserInfo_Err:
198     Call RegistrarError(Err.Number, Err.Description, "TCP.ResetBasicUserInfo", Erl)
200     Resume Next
        
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
112     Call RegistrarError(Err.Number, Err.Description, "TCP.ResetGuildInfo", Erl)
114     Resume Next
        
End Sub

Sub ResetUserFlags(ByVal UserIndex As Integer)
        '*************************************************
        'Author: Unknown
        'Last modified: 03/29/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '03/29/2006 Maraxus - Reseteo el CentinelaOK también.
        '*************************************************
        
        On Error GoTo ResetUserFlags_Err
        

100     With UserList(UserIndex).flags
102         .LevelBackup = 0
104         .Comerciando = False
106         .Ban = 0
108         .Escondido = 0
110         .DuracionEfecto = 0
112         .ScrollExp = 1
114         .ScrollOro = 1
116         .NpcInv = 0
118         .StatsChanged = 0
120         .TargetNPC = 0
122         .TargetNpcTipo = eNPCType.Comun
124         .TargetObj = 0
126         .TargetObjMap = 0
128         .TargetObjX = 0
130         .TargetObjY = 0
132         .TargetUser = 0
134         .TipoPocion = 0
136         .TomoPocion = False
138         .Descuento = vbNullString
140         .Hambre = 0
142         .Sed = 0
144         .Descansar = False
146         .Navegando = 0
148         .Oculto = 0
150         .Envenenado = 0
152         .Ahogandose = 0
154         .invisible = 0
156         .Paralizado = 0
158         .Inmovilizado = 0
160         .Maldicion = 0
162         .Bendicion = 0
164         .Meditando = 0
166         .Escribiendo = 0
168         .Privilegios = 0
170         .PuedeMoverse = 0
172         .OldBody = 0
174         .OldHead = 0
176         .AdminInvisible = 0
178         .ValCoDe = 0
180         .Hechizo = 0
182         .Silenciado = 0
184         .CentinelaOK = False
186         .AdminPerseguible = False
            'Ladder
188         .VecesQueMoriste = 0
190         .MinutosRestantes = 0
192         .SegundosPasados = 0
194         .CarroMineria = 0
196         .Montado = 0
198         .Incinerado = 0
200         .Casado = 0
202         .Pareja = ""
204         .Candidato = 0
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
228         .CarroMineria = 0
230         .PendienteDelSacrificio = 0
232         .AnilloOcultismo = 0
234         .RegeneracionMana = 0
236         .RegeneracionHP = 0
238         .RegeneracionSta = 0
240         .NecesitaOxigeno = False

242         .LastCrimMatado = vbNullString
244         .LastCiudMatado = vbNullString
        
246         .UserLogged = False
248         .FirstPacket = False
250         .Inmunidad = 0
            
252         .Mimetizado = e_EstadoMimetismo.Desactivado
254         .MascotasGuardadas = 0

256         .EnConsulta = False
            
258         .ProcesosPara = vbNullString
260         .ScreenShotPara = vbNullString
262         Set .ScreenShot = Nothing

            Dim i As Integer
264         For i = LBound(.ChatHistory) To UBound(.ChatHistory)
266             .ChatHistory(i) = vbNullString
            Next

268         .EnReto = False
270         .SolicitudReto.estado = SolicitudRetoEstado.Libre
272         .AceptoReto = 0
274         .LastPos.Map = 0

        End With

        
        Exit Sub

ResetUserFlags_Err:
276     Call RegistrarError(Err.Number, Err.Description, "TCP.ResetUserFlags", Erl)
278     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.Description, "TCP.ResetAccionesPendientes", Erl)
116     Resume Next
        
End Sub

Sub ResetDonadorFlag(ByVal UserIndex As Integer)
        
        On Error GoTo ResetDonadorFlag_Err
        

        '*************************************************
        '*************************************************
100     With UserList(UserIndex).donador
102         .activo = 0
104         .CreditoDonador = 0
106         .FechaExpiracion = 0

        End With

        
        Exit Sub

ResetDonadorFlag_Err:
108     Call RegistrarError(Err.Number, Err.Description, "TCP.ResetDonadorFlag", Erl)
110     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.Description, "TCP.ResetUserSpells", Erl)
108     Resume Next
        
End Sub

Sub ResetUserSkills(ByVal UserIndex As Integer)
        
        On Error GoTo ResetUserSkills_Err
        

        Dim LoopC As Long

100     For LoopC = 1 To NUMSKILLS
102         UserList(UserIndex).Stats.UserSkills(LoopC) = 0
104     Next LoopC

        
        Exit Sub

ResetUserSkills_Err:
106     Call RegistrarError(Err.Number, Err.Description, "TCP.ResetUserSkills", Erl)
108     Resume Next
        
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
112     Call RegistrarError(Err.Number, Err.Description, "TCP.ResetUserBanco", Erl)
114     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.Description, "TCP.ResetUserKeys", Erl)

        
End Sub

Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)
        
        On Error GoTo LimpiarComercioSeguro_Err
        

100     With UserList(UserIndex).ComUsu

102         If .DestUsu > 0 Then
104             Call FinComerciarUsu(.DestUsu)
106             Call FinComerciarUsu(UserIndex)

            End If

        End With

        
        Exit Sub

LimpiarComercioSeguro_Err:
108     Call RegistrarError(Err.Number, Err.Description, "TCP.LimpiarComercioSeguro", Erl)
110     Resume Next
        
End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)
        
        On Error GoTo ResetUserSlot_Err
        

100     UserList(UserIndex).ConnIDValida = False
102     UserList(UserIndex).ConnID = -1

104     If UserList(UserIndex).Grupo.Lider = UserIndex Then
106         Call FinalizarGrupo(UserIndex)

        End If

108     If UserList(UserIndex).Grupo.EnGrupo Then
110         Call SalirDeGrupoForzado(UserIndex)

        End If

112     UserList(UserIndex).Grupo.CantidadMiembros = 0
114     UserList(UserIndex).Grupo.EnGrupo = False
116     UserList(UserIndex).Grupo.Lider = 0
118     UserList(UserIndex).Grupo.PropuestaDe = 0
120     UserList(UserIndex).Grupo.Miembros(6) = 0
122     UserList(UserIndex).Grupo.Miembros(1) = 0
124     UserList(UserIndex).Grupo.Miembros(2) = 0
126     UserList(UserIndex).Grupo.Miembros(3) = 0
128     UserList(UserIndex).Grupo.Miembros(4) = 0
130     UserList(UserIndex).Grupo.Miembros(5) = 0

132     Call ResetQuestStats(UserIndex)
134     Call ResetGuildInfo(UserIndex)
136     Call LimpiarComercioSeguro(UserIndex)
138     Call ResetFacciones(UserIndex)
140     Call ResetContadores(UserIndex)
142     Call ResetCharInfo(UserIndex)
144     Call ResetBasicUserInfo(UserIndex)
146     Call ResetUserFlags(UserIndex)
148     Call ResetAccionesPendientes(UserIndex)
150     Call ResetDonadorFlag(UserIndex)
152     Call LimpiarInventario(UserIndex)
154     Call ResetUserSpells(UserIndex)
        'Call ResetUserPets(UserIndex)
156     Call ResetUserBanco(UserIndex)
158     Call ResetUserSkills(UserIndex)
160     Call ResetUserKeys(UserIndex)

162     With UserList(UserIndex).ComUsu
164         .Acepto = False
166         .cant = 0
168         .DestNick = vbNullString
170         .DestUsu = 0
172         .Objeto = 0

        End With

        
        Exit Sub

ResetUserSlot_Err:
174     Call RegistrarError(Err.Number, Err.Description, "TCP.ResetUserSlot", Erl)
176     Resume Next
        
End Sub

Sub CloseUser(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler
    
        Dim errordesc As String
        Dim Map As Integer
        Dim aN  As Integer
        Dim i   As Integer
        
100     With UserList(UserIndex)
            
102         Map = .Pos.Map
        
104         errordesc = "ERROR AL SETEAR NPC"
        
106         aN = .flags.AtacadoPorNpc
    
108         If aN > 0 Then
110             NpcList(aN).Movement = NpcList(aN).flags.OldMovement
112             NpcList(aN).Hostile = NpcList(aN).flags.OldHostil
114             NpcList(aN).flags.AttackedBy = vbNullString
116             NpcList(aN).Target = 0
    
            End If
    
118         aN = .flags.NPCAtacado
    
120         If aN > 0 Then
122             If NpcList(aN).flags.AttackedFirstBy = .name Then
124                 NpcList(aN).flags.AttackedFirstBy = vbNullString
                End If
            End If
    
126         .flags.AtacadoPorNpc = 0
128         .flags.NPCAtacado = 0
        
130         errordesc = "ERROR AL DESMONTAR"
    
132         If .flags.Montado > 0 Then
134             Call DoMontar(UserIndex, ObjData(.Invent.MonturaObjIndex), .Invent.MonturaSlot)
            End If
            
136         errordesc = "ERROR AL CANCELAR SOLICITUD DE RETO"
            
138         If .flags.EnReto Then
140             Call AbandonarReto(UserIndex, True)

142         ElseIf .flags.SolicitudReto.estado <> SolicitudRetoEstado.Libre Then
144             Call CancelarSolicitudReto(UserIndex, .name & " se ha desconectado.")
            
146         ElseIf .flags.AceptoReto > 0 Then
148             Call CancelarSolicitudReto(.flags.AceptoReto, .name & " se ha desconectado.")
            End If
        
150         errordesc = "ERROR AL SACAR MIMETISMO"
152         If .flags.Mimetizado > 0 Then

154             .Char.Body = .CharMimetizado.Body
156             .Char.Head = .CharMimetizado.Head
158             .Char.CascoAnim = .CharMimetizado.CascoAnim
160             .Char.ShieldAnim = .CharMimetizado.ShieldAnim
162             .Char.WeaponAnim = .CharMimetizado.WeaponAnim
164             .Counters.Mimetismo = 0
166             .flags.Mimetizado = e_EstadoMimetismo.Desactivado

            End If
        
168         errordesc = "ERROR AL ENVIAR PARTICULA"
        
170         .Char.FX = 0
172         .Char.loops = 0
174         .Char.ParticulaFx = 0
176         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, 0, 0, True))
178         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
        
180         .flags.UserLogged = False
182         .Counters.Saliendo = False
        
184         errordesc = "ERROR AL ENVIAR INVI"
        
            'Le devolvemos el body y head originales
186         If .flags.AdminInvisible = 1 Then Call DoAdminInvisible(UserIndex)
        
188         errordesc = "ERROR AL CANCELAR SUBASTA"
    
190         If .flags.Subastando = True Then
192             Call CancelarSubasta
    
            End If
        
194         errordesc = "ERROR AL BORRAR INDEX DE TORNEO"
    
196         If .flags.EnTorneo = True Then
198             Call BorrarIndexInTorneo(UserIndex)
200             .flags.EnTorneo = False
    
            End If
        
            'Save statistics
            'Call Statistics.UserDisconnected(UserIndex)
        
            ' Grabamos el personaje del usuario
        
202         errordesc = "ERROR AL GRABAR PJ"
        
204         Call SaveUser(UserIndex, True)
    
206         errordesc = "ERROR AL DESCONTAR USER DE MAPA"
    
208         If MapInfo(Map).NumUsers > 0 Then
210             Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
    
            End If
    
212         errordesc = "ERROR AL ERASEUSERCHAR"
        
            'Borrar el personaje
214         Call EraseUserChar(UserIndex, True)
        
216         errordesc = "ERROR AL BORRAR MASCOTAS"
        
            'Borrar mascotas
218         For i = 1 To MAXMASCOTAS
220             If .MascotasIndex(i) > 0 Then
222                 If NpcList(.MascotasIndex(i)).flags.NPCActive Then _
                        Call QuitarNPC(.MascotasIndex(i))
                End If
224         Next i
        
226         errordesc = "ERROR Update Map Users"
        
            'Update Map Users
228         MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1
        
230         If MapInfo(Map).NumUsers < 0 Then MapInfo(Map).NumUsers = 0
    
            ' Si el usuario habia dejado un msg en la gm's queue lo borramos
            'If Ayuda.Existe(.Name) Then Call Ayuda.Quitar(.Name)
        
232         errordesc = "ERROR AL m_NameIndex.Remove() Name:" & .name & " cuenta:" & .Cuenta
            
234         Call m_NameIndex.Remove(UCase$(.name))
        
236         errordesc = "ERROR AL RESETSLOT Name:" & .name & " cuenta:" & .Cuenta
        
238         Call ResetUserSlot(UserIndex)

        End With
    
        Exit Sub
    
ErrHandler:
        'Call LogError("Error en CloseUser. Número " & Err.Number & ". Descripción: " & Err.Description & ". Detalle:" & errordesc)
240     Call RegistrarError(Err.Number, Err.Description & ". Detalle:" & errordesc, Erl)
242     Resume Next ' TODO: Provisional hasta solucionar bugs graves

End Sub

Sub ReloadSokcet()

        On Error GoTo ErrHandler

100     Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    
102     If NumUsers <= 0 Then
104         Call WSApiReiniciarSockets
        Else
            'Call apiclosesocket(SockListen)
            'SockListen = ListenForConnect(Puerto, hWndMsg, "")
        End If

        Exit Sub
ErrHandler:
106     Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.Description)

End Sub

Public Sub EcharPjsNoPrivilegiados()
        
        On Error GoTo EcharPjsNoPrivilegiados_Err
        

        Dim LoopC As Long

100     For LoopC = 1 To LastUser

102         If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
104             If UserList(LoopC).flags.Privilegios And PlayerType.user Then
106                 Call CloseSocket(LoopC)

                End If

            End If

108     Next LoopC

        
        Exit Sub

EcharPjsNoPrivilegiados_Err:
110     Call RegistrarError(Err.Number, Err.Description, "TCP.EcharPjsNoPrivilegiados", Erl)
112     Resume Next
        
End Sub

Function ValidarCabeza(ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, ByVal Head As Integer) As Boolean

100     Select Case UserSexo
    
            Case eGenero.Hombre
        
102             Select Case UserRaza
                
                    Case eRaza.Humano
104                     ValidarCabeza = Head >= 1 And Head <= 41
                    
106                 Case eRaza.Elfo
108                     ValidarCabeza = Head >= 101 And Head <= 132
                    
110                 Case eRaza.Drow
112                     ValidarCabeza = Head >= 200 And Head <= 229
                    
114                 Case eRaza.Enano
116                     ValidarCabeza = Head >= 300 And Head <= 329
                    
118                 Case eRaza.Gnomo
120                     ValidarCabeza = Head >= 400 And Head <= 429
                    
122                 Case eRaza.Orco
124                     ValidarCabeza = Head >= 500 And Head <= 529
                
                End Select
        
126         Case eGenero.Mujer
        
128             Select Case UserRaza
                
                    Case eRaza.Humano
130                     ValidarCabeza = Head >= 50 And Head <= 80
                    
132                 Case eRaza.Elfo
134                     ValidarCabeza = Head >= 150 And Head <= 179
                    
136                 Case eRaza.Drow
138                     ValidarCabeza = Head >= 250 And Head <= 279
                    
140                 Case eRaza.Enano
142                     ValidarCabeza = Head >= 350 And Head <= 379
                    
144                 Case eRaza.Gnomo
146                     ValidarCabeza = Head >= 450 And Head <= 479
                    
148                 Case eRaza.Orco
150                     ValidarCabeza = Head >= 550 And Head <= 579
                
                End Select
    
        End Select

End Function

Function ValidarNombre(nombre As String) As Boolean
    
100     If Len(nombre) < 1 Or Len(nombre) > 18 Then Exit Function
    
        Dim temp As String
102     temp = UCase$(nombre)
    
        Dim i As Long, Char As Integer, LastChar As Integer
104     For i = 1 To Len(temp)
106         Char = Asc(mid$(temp, i, 1))
        
108         If (Char < 65 Or Char > 90) And Char <> 32 Then
                Exit Function
        
110         ElseIf Char = 32 And LastChar = 32 Then
                Exit Function
            End If
        
112         LastChar = Char
        Next

114     If Asc(mid$(temp, 1, 1)) = 32 Or Asc(mid$(temp, Len(temp), 1)) = 32 Then
            Exit Function
        End If
    
116     ValidarNombre = True

End Function

Function ContarUsuariosMismaCuenta(ByVal AccountId As Long) As Integer

        Dim i As Integer
    
100     For i = 1 To LastUser
        
102         If UserList(i).flags.UserLogged And UserList(i).AccountId = AccountId Then
104             ContarUsuariosMismaCuenta = ContarUsuariosMismaCuenta + 1
            End If
        
        Next

End Function
