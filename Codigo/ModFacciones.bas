Attribute VB_Name = "ModFacciones"
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

Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)
        
        On Error GoTo EnlistarArmadaReal_Err
        

        '***************************************************
        'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
        'Last Modification: 23/01/2007
        'Handles the entrance of users to the "Armada Real"
        '***************************************************
100     If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
102         Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a las tropas reales!!! Ve a combatir criminales", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

104     If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
106         Call WriteChatOverHead(UserIndex, "¡¡¡Maldito insolente!!! vete de aqui seguidor de las sombras", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

108     If Status(UserIndex) = 2 Then
110         Call WriteChatOverHead(UserIndex, "¡¡¡No se permiten criminales en el ejercito imperial!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

112     If Status(UserIndex) = 0 Then
114         Call WriteChatOverHead(UserIndex, "¡¡¡No se permiten neutrales en el ejercito imperial, antes debes solicitar el perdón con un sacerdote!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

116     If UserList(UserIndex).Stats.ELV < 35 Then
118         Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel 35!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

120     If UserList(UserIndex).Faccion.CriminalesMatados < 30 Then
122         Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 30 criminales, solo has matado " & UserList(UserIndex).Faccion.CriminalesMatados, str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
 
124     If UserList(UserIndex).Faccion.CiudadanosMatados > 0 Then
126         Call WriteChatOverHead(UserIndex, "¡Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

128     If UserList(UserIndex).Faccion.Reenlistadas > 4 Then
130         Call WriteChatOverHead(UserIndex, "¡Has sido expulsado de las fuerzas reales demasiadas veces!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

132     UserList(UserIndex).Faccion.ArmadaReal = 1
134     UserList(UserIndex).Faccion.Reenlistadas = UserList(UserIndex).Faccion.Reenlistadas + 1

136     Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al Ejercito Imperial!!!, aqui tienes tus vestimentas. Cumple bien tu labor exterminando Criminales y me encargaré de recompensarte.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
138     UserList(UserIndex).Faccion.Status = 3
140     Call WriteConsoleMsg(UserIndex, "¡Ahora perteneces al ejercito revolucionario!", FontTypeNames.FONTTYPE_INFOIAO)
142     Call RefreshCharStatus(UserIndex)

144     If UserList(UserIndex).Faccion.RecibioArmaduraReal = 0 Then

            Dim MiObj As obj

146         MiObj.Amount = 1

148         Select Case UserList(UserIndex).clase

                Case eClass.Cleric, eClass.Assasin ' ok
150                 MiObj.ObjIndex = 1544

152             Case eClass.Paladin, eClass.Warrior, eClass.Trabajador, eClass.Hunter ' ok
154                 MiObj.ObjIndex = 1548

156             Case eClass.Mage, eClass.Druid, eClass.Bard ' ok
158                 MiObj.ObjIndex = 1546

            End Select
    
160         If Not MeterItemEnInventario(UserIndex, MiObj) Then
162             Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

            End If
    
164         UserList(UserIndex).Faccion.RecibioArmaduraReal = 1
166         UserList(UserIndex).Faccion.NivelIngreso = UserList(UserIndex).Stats.ELV
168         UserList(UserIndex).Faccion.FechaIngreso = Date
            'Esto por ahora es inútil, siempre va a ser cero, pero bueno, despues va a servir.
170         UserList(UserIndex).Faccion.MatadosIngreso = UserList(UserIndex).Faccion.CiudadanosMatados

        End If

172     UserList(UserIndex).Faccion.RecompensasReal = 1
174     UserList(UserIndex).Faccion.NextRecompensa = 60

        'Call LogEjercitoReal(UserList(UserIndex).Name & " ingresó el " & Date & " cuando era nivel " & UserList(UserIndex).Stats.ELV)

        
        Exit Sub

EnlistarArmadaReal_Err:
176     Call RegistrarError(Err.Number, Err.description, "ModFacciones.EnlistarArmadaReal", Erl)
178     Resume Next
        
End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)
        
        On Error GoTo RecompensaArmadaReal_Err
        

        '***************************************************
        'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
        'Last Modification: 23/01/2007
        'Handles the way of gaining new ranks in the "Armada Real"
        '***************************************************
        Dim Crimis    As Long

        Dim Lvl       As Byte

        Dim NextRecom As Long

        Dim Nobleza   As Long

100     Lvl = UserList(UserIndex).Stats.ELV
102     Crimis = UserList(UserIndex).Faccion.CriminalesMatados
104     NextRecom = UserList(UserIndex).Faccion.NextRecompensa

106     If Crimis < NextRecom Then
108         Call WriteChatOverHead(UserIndex, "Mata " & NextRecom - Crimis & " Criminales más para recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

110     Select Case NextRecom

            Case 30:
    
112             If Lvl < 27 Then
114                 Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & 40 - Lvl & " Niveles para poder recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

116             UserList(UserIndex).Faccion.RecompensasReal = 1
118             UserList(UserIndex).Faccion.NextRecompensa = 60
    
            Case 60:
120             UserList(UserIndex).Faccion.RecompensasReal = 2
122             UserList(UserIndex).Faccion.NextRecompensa = 120
    
            Case 120:
124             UserList(UserIndex).Faccion.RecompensasReal = 3
126             UserList(UserIndex).Faccion.NextRecompensa = 180
    
            Case 180:

128             If Lvl < 45 Then
130                 Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & 40 - Lvl & " Niveles para poder recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

132             UserList(UserIndex).Faccion.RecompensasReal = 4
134             UserList(UserIndex).Faccion.NextRecompensa = 220
    
            Case 220:

136             If Lvl < 50 Then
138                 Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & 40 - Lvl & " Niveles para poder recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

140             UserList(UserIndex).Faccion.RecompensasReal = 5
142             UserList(UserIndex).Faccion.NextRecompensa = 640

            Case 5000:

144             If Lvl < 37 Then
146                 Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & 37 - Lvl & " Niveles para poder recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

148             UserList(UserIndex).Faccion.RecompensasReal = 14
150             UserList(UserIndex).Faccion.NextRecompensa = 10000
    
            Case 10000:
152             Call WriteChatOverHead(UserIndex, "Eres uno de mis mejores Soldados. Mataste " & Crimis & ", sigue asi. Ya no tengo más recompensa para darte que mi agradescimiento. ¡Felicidades!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
    
            Case Else:
                Exit Sub

        End Select

154     Call WriteChatOverHead(UserIndex, "¡¡¡Aqui tienes tu recompensa " + TituloReal(UserIndex) + "!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        'UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpX100

        
        Exit Sub

RecompensaArmadaReal_Err:
156     Call RegistrarError(Err.Number, Err.description, "ModFacciones.RecompensaArmadaReal", Erl)
158     Resume Next
        
End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer)
        
        On Error GoTo ExpulsarFaccionReal_Err
        

100     UserList(UserIndex).Faccion.ArmadaReal = 0
102     UserList(UserIndex).Faccion.Status = 1
104     Call RefreshCharStatus(UserIndex)
        'Call PerderItemsFaccionarios(UserIndex)
106     Call WriteConsoleMsg(UserIndex, "Has sido expulsado del Ejertico Revolucionario.", FontTypeNames.FONTTYPE_INFOIAO)

        'Desequipamos la armadura real si está equipada
        'If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
        
        Exit Sub

ExpulsarFaccionReal_Err:
108     Call RegistrarError(Err.Number, Err.description, "ModFacciones.ExpulsarFaccionReal", Erl)
110     Resume Next
        
End Sub

Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer)
        
        On Error GoTo ExpulsarFaccionCaos_Err
        

100     UserList(UserIndex).Faccion.FuerzasCaos = 0
    
102     UserList(UserIndex).Faccion.Status = 0
104     Call RefreshCharStatus(UserIndex)
    
        'Call PerderItemsFaccionarios(UserIndex)
106     Call WriteConsoleMsg(UserIndex, "Has sido expulsado de la Legión Oscura.", FontTypeNames.FONTTYPE_INFOIAO)
        'Desequipamos la armadura real si está equipada
    
        'If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
        'If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
        
        Exit Sub

ExpulsarFaccionCaos_Err:
108     Call RegistrarError(Err.Number, Err.description, "ModFacciones.ExpulsarFaccionCaos", Erl)
110     Resume Next
        
End Sub

Public Function TituloReal(ByVal UserIndex As Integer) As String
        
        On Error GoTo TituloReal_Err
        

        '***************************************************
        'Autor: Unknown
        'Last Modification: 23/01/2007 Pablo (ToxicWaste)
        'Handles the titles of the members of the "Armada Real"
        '***************************************************
100     Select Case UserList(UserIndex).Faccion.RecompensasReal

            Case 1
102             TituloReal = "Aprendiz"

104         Case 2
106             TituloReal = "Soldado"

108         Case 3
110             TituloReal = "Sargento"

112         Case 4
114             TituloReal = "Protector del Reino"

116         Case 5
118             TituloReal = "Campeón de la Luz"

        End Select

        
        Exit Function

TituloReal_Err:
120     Call RegistrarError(Err.Number, Err.description, "ModFacciones.TituloReal", Erl)
122     Resume Next
        
End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)
        '***************************************************
        'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
        'Last Modification: 23/01/2007
        'Handles the entrance of users to the "Legión Oscura"
        '***************************************************
        
        On Error GoTo EnlistarCaos_Err
        

100     If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
102         Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a la legión oscura!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

104     If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
106         Call WriteChatOverHead(UserIndex, "Las sombras reinarán en Argentum. ¡¡¡Fuera de aqui insecto Real!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

        '[Barrin 17-12-03] Si era miembro de la Armada Real no se puede enlistar
108     If UserList(UserIndex).Faccion.RecibioExpInicialReal = 1 Then 'Tomamos el valor de ahí: ¿Recibio la experiencia para entrar?
110         Call WriteChatOverHead(UserIndex, "No permitiré que ningún insecto real ingrese a mis tropas.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

        '[/Barrin]

112     If Status(UserIndex) = 1 Then
114         Call WriteChatOverHead(UserIndex, "¡¡Ja ja ja!! Tu no eres bienvenido aqui asqueroso Ciudadano", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

116     If UserList(UserIndex).Faccion.CiudadanosMatados + UserList(UserIndex).Faccion.CriminalesMatados < 70 Then
118         Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 70 usuarios, solo has matado " & UserList(UserIndex).Faccion.CiudadanosMatados + UserList(UserIndex).Faccion.CriminalesMatados, str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

120     If UserList(UserIndex).Stats.ELV < 35 Then
122         Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel 35!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

124     If UserList(UserIndex).Faccion.Reenlistadas > 4 Then
126         If UserList(UserIndex).Faccion.Reenlistadas = 200 Then
128             Call WriteChatOverHead(UserIndex, "Has sido expulsado de las fuerzas oscuras y durante tu rebeldía has atacado a mi ejército. ¡Vete de aquí!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Else
130             Call WriteChatOverHead(UserIndex, "¡Has sido expulsado de las fuerzas oscuras demasiadas veces!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)

            End If

            Exit Sub

        End If

132     UserList(UserIndex).Faccion.Reenlistadas = UserList(UserIndex).Faccion.Reenlistadas + 1
134     UserList(UserIndex).Faccion.FuerzasCaos = 1

136     UserList(UserIndex).Faccion.Status = 2
138     Call WriteConsoleMsg(UserIndex, "¡Ahora perteneces al caos!", FontTypeNames.FONTTYPE_INFOIAO)
140     Call RefreshCharStatus(UserIndex)

142     Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al lado oscuro!!! Aqui tienes tus armaduras. Derrama sangre Ciudadana y Real y serás recompensado, lo prometo.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)

144     If UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0 Then

            Dim MiObj As obj

146         MiObj.Amount = 1
    
148         Select Case UserList(UserIndex).clase

                Case eClass.Cleric, eClass.Assasin
150                 MiObj.ObjIndex = 1500

152             Case eClass.Paladin, eClass.Warrior, eClass.Trabajador, eClass.Hunter
154                 MiObj.ObjIndex = 1504

156             Case eClass.Mage, eClass.Druid, eClass.Bard
158                 MiObj.ObjIndex = 1502

            End Select
    
160         If Not MeterItemEnInventario(UserIndex, MiObj) Then
162             Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

            End If
    
164         UserList(UserIndex).Faccion.RecibioArmaduraCaos = 1
166         UserList(UserIndex).Faccion.NivelIngreso = UserList(UserIndex).Stats.ELV
168         UserList(UserIndex).Faccion.FechaIngreso = Date

        End If

170     UserList(UserIndex).Faccion.RecompensasCaos = 1
172     UserList(UserIndex).Faccion.NextRecompensa = 140

        'Call LogEjercitoCaos(UserList(UserIndex).Name & " ingresó el " & Date & " cuando era nivel " & UserList(UserIndex).Stats.ELV)

        
        Exit Sub

EnlistarCaos_Err:
174     Call RegistrarError(Err.Number, Err.description, "ModFacciones.EnlistarCaos", Erl)
176     Resume Next
        
End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)
        
        On Error GoTo RecompensaCaos_Err
        

        '***************************************************
        'Author: Pablo (ToxicWaste) & Unknown (orginal version)
        'Last Modification: 23/01/2007
        'Handles the way of gaining new ranks in the "Legión Oscura"
        '***************************************************
        Dim Ciudas    As Long

        Dim Lvl       As Byte

        Dim NextRecom As Long

100     Lvl = UserList(UserIndex).Stats.ELV
102     Ciudas = UserList(UserIndex).Faccion.CiudadanosMatados
104     NextRecom = UserList(UserIndex).Faccion.NextRecompensa

106     If Ciudas < NextRecom Then
108         Call WriteChatOverHead(UserIndex, "Mata " & NextRecom - Ciudas & " Cuidadanos más para recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

        Dim huborecompensa As Boolean

110     Select Case NextRecom

            Case 140:
112             UserList(UserIndex).Faccion.RecompensasCaos = 2
114             UserList(UserIndex).Faccion.NextRecompensa = 250
116             huborecompensa = True
    
            Case 250:
118             UserList(UserIndex).Faccion.RecompensasCaos = 3
120             UserList(UserIndex).Faccion.NextRecompensa = 400
122             huborecompensa = True
    
            Case 400:
124             UserList(UserIndex).Faccion.RecompensasCaos = 4
126             UserList(UserIndex).Faccion.NextRecompensa = 800
128             huborecompensa = True
    
            Case 800:
130             UserList(UserIndex).Faccion.RecompensasCaos = 5
132             UserList(UserIndex).Faccion.NextRecompensa = 1500
134             huborecompensa = True

            Case 1500:
                ' If Lvl < 27 Then
                '  Call WriteChatOverHead(UserIndex, "Mataste Suficientes Ciudadanos, pero te faltan " & 27 - Lvl & " Niveles para poder recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                '  Exit Sub
                ' End If
136             Call WriteChatOverHead(UserIndex, "¡Felicitaciones, eres de mis mejores guerreros, estas en lo más alto!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub

            Case Else:
                Exit Sub
        
        End Select

138     If huborecompensa Then
140         Call WriteChatOverHead(UserIndex, "¡¡¡Bien hecho " + TituloCaos(UserIndex) + ", aquí tienes tu recompensa!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)

        End If

        
        Exit Sub

RecompensaCaos_Err:
142     Call RegistrarError(Err.Number, Err.description, "ModFacciones.RecompensaCaos", Erl)
144     Resume Next
        
End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String
        
        On Error GoTo TituloCaos_Err
        

100     Select Case UserList(UserIndex).Faccion.RecompensasCaos

            Case 1
102             TituloCaos = "Acólito"

104         Case 2
106             TituloCaos = "Alma Corrupta"

108         Case 3
110             TituloCaos = "Corruptor"

112         Case 4
114             TituloCaos = "Protector del Averno"

116         Case 5
118             TituloCaos = "Campeón de la Oscuridad"

        End Select

        
        Exit Function

TituloCaos_Err:
120     Call RegistrarError(Err.Number, Err.description, "ModFacciones.TituloCaos", Erl)
122     Resume Next
        
End Function
