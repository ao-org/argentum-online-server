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
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Handles the entrance of users to the "Armada Real"
'***************************************************
If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a las tropas reales!!! Ve a combatir criminales", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call WriteChatOverHead(UserIndex, "¡¡¡Maldito insolente!!! vete de aqui seguidor de las sombras", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If Status(UserIndex) = 2 Then
    Call WriteChatOverHead(UserIndex, "¡¡¡No se permiten criminales en el ejercito imperial!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If Status(UserIndex) = 0 Then
    Call WriteChatOverHead(UserIndex, "¡¡¡No se permiten neutrales en el ejercito imperial, antes debes solicitar el perdón con un sacerdote!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 35 Then
    Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel 35!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.CriminalesMatados < 30 Then
    Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 30 criminales, solo has matado " & UserList(UserIndex).Faccion.CriminalesMatados, str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If
 
If UserList(UserIndex).Faccion.CiudadanosMatados > 0 Then
    Call WriteChatOverHead(UserIndex, "¡Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.Reenlistadas > 4 Then
    Call WriteChatOverHead(UserIndex, "¡Has sido expulsado de las fuerzas reales demasiadas veces!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If


UserList(UserIndex).Faccion.ArmadaReal = 1
UserList(UserIndex).Faccion.Reenlistadas = UserList(UserIndex).Faccion.Reenlistadas + 1

Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al Ejercito Imperial!!!, aqui tienes tus vestimentas. Cumple bien tu labor exterminando Criminales y me encargaré de recompensarte.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
UserList(UserIndex).Faccion.Status = 3
Call WriteConsoleMsg(UserIndex, "¡Ahora perteneces al ejercito revolucionario!", FontTypeNames.FONTTYPE_INFOIAO)
Call RefreshCharStatus(UserIndex)


If UserList(UserIndex).Faccion.RecibioArmaduraReal = 0 Then
    Dim MiObj As obj
    MiObj.Amount = 1

        Select Case UserList(UserIndex).clase
            Case eClass.Cleric, eClass.Assasin ' ok
                MiObj.ObjIndex = 1544
            Case eClass.Paladin, eClass.Warrior, eClass.Trabajador, eClass.Hunter ' ok
                MiObj.ObjIndex = 1548
            Case eClass.Mage, eClass.Druid, eClass.Bard ' ok
                MiObj.ObjIndex = 1546
        End Select
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If

    
    UserList(UserIndex).Faccion.RecibioArmaduraReal = 1
    UserList(UserIndex).Faccion.NivelIngreso = UserList(UserIndex).Stats.ELV
    UserList(UserIndex).Faccion.FechaIngreso = Date
    'Esto por ahora es inútil, siempre va a ser cero, pero bueno, despues va a servir.
    UserList(UserIndex).Faccion.MatadosIngreso = UserList(UserIndex).Faccion.CiudadanosMatados

End If

    UserList(UserIndex).Faccion.RecompensasReal = 1
    UserList(UserIndex).Faccion.NextRecompensa = 60


'Call LogEjercitoReal(UserList(UserIndex).Name & " ingresó el " & Date & " cuando era nivel " & UserList(UserIndex).Stats.ELV)

End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Handles the way of gaining new ranks in the "Armada Real"
'***************************************************
Dim Crimis As Long
Dim Lvl As Byte
Dim NextRecom As Long
Dim Nobleza As Long
Lvl = UserList(UserIndex).Stats.ELV
Crimis = UserList(UserIndex).Faccion.CriminalesMatados
NextRecom = UserList(UserIndex).Faccion.NextRecompensa

If Crimis < NextRecom Then
    Call WriteChatOverHead(UserIndex, "Mata " & NextRecom - Crimis & " Criminales más para recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

Select Case NextRecom
    Case 30:
    
        If Lvl < 27 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & 40 - Lvl & " Niveles para poder recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        UserList(UserIndex).Faccion.RecompensasReal = 1
        UserList(UserIndex).Faccion.NextRecompensa = 60
    
    Case 60:
        UserList(UserIndex).Faccion.RecompensasReal = 2
        UserList(UserIndex).Faccion.NextRecompensa = 120
    
    Case 120:
        UserList(UserIndex).Faccion.RecompensasReal = 3
        UserList(UserIndex).Faccion.NextRecompensa = 180
    
    Case 180:
        If Lvl < 45 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & 40 - Lvl & " Niveles para poder recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        UserList(UserIndex).Faccion.RecompensasReal = 4
        UserList(UserIndex).Faccion.NextRecompensa = 220
    
    Case 220:
        If Lvl < 50 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & 40 - Lvl & " Niveles para poder recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        UserList(UserIndex).Faccion.RecompensasReal = 5
        UserList(UserIndex).Faccion.NextRecompensa = 640
    Case 5000:
        If Lvl < 37 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & 37 - Lvl & " Niveles para poder recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        UserList(UserIndex).Faccion.RecompensasReal = 14
        UserList(UserIndex).Faccion.NextRecompensa = 10000
    
    Case 10000:
        Call WriteChatOverHead(UserIndex, "Eres uno de mis mejores Soldados. Mataste " & Crimis & ", sigue asi. Ya no tengo más recompensa para darte que mi agradescimiento. ¡Felicidades!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    
    Case Else:
        Exit Sub
End Select

Call WriteChatOverHead(UserIndex, "¡¡¡Aqui tienes tu recompensa " + TituloReal(UserIndex) + "!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
'UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpX100


End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer)

    UserList(UserIndex).Faccion.ArmadaReal = 0
    UserList(UserIndex).Faccion.Status = 1
    Call RefreshCharStatus(UserIndex)
    'Call PerderItemsFaccionarios(UserIndex)
    Call WriteConsoleMsg(UserIndex, "Has sido expulsado del Ejertico Revolucionario.", FontTypeNames.FONTTYPE_INFOIAO)
    'Desequipamos la armadura real si está equipada
    'If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
End Sub

Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer)

    UserList(UserIndex).Faccion.FuerzasCaos = 0
    
        UserList(UserIndex).Faccion.Status = 0
    Call RefreshCharStatus(UserIndex)
    
    'Call PerderItemsFaccionarios(UserIndex)
    Call WriteConsoleMsg(UserIndex, "Has sido expulsado de la Legión Oscura.", FontTypeNames.FONTTYPE_INFOIAO)
    'Desequipamos la armadura real si está equipada
    
    'If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    'If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
End Sub

Public Function TituloReal(ByVal UserIndex As Integer) As String
'***************************************************
'Autor: Unknown
'Last Modification: 23/01/2007 Pablo (ToxicWaste)
'Handles the titles of the members of the "Armada Real"
'***************************************************
Select Case UserList(UserIndex).Faccion.RecompensasReal
    Case 1
        TituloReal = "Aprendiz"
    Case 2
        TituloReal = "Soldado"
    Case 3
        TituloReal = "Sargento"
    Case 4
        TituloReal = "Protector del Reino"
    Case 5
        TituloReal = "Campeón de la Luz"
End Select


End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Handles the entrance of users to the "Legión Oscura"
'***************************************************

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a la legión oscura!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call WriteChatOverHead(UserIndex, "Las sombras reinarán en Argentum. ¡¡¡Fuera de aqui insecto Real!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

'[Barrin 17-12-03] Si era miembro de la Armada Real no se puede enlistar
If UserList(UserIndex).Faccion.RecibioExpInicialReal = 1 Then 'Tomamos el valor de ahí: ¿Recibio la experiencia para entrar?
    Call WriteChatOverHead(UserIndex, "No permitiré que ningún insecto real ingrese a mis tropas.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If
'[/Barrin]

If Status(UserIndex) = 1 Then
    Call WriteChatOverHead(UserIndex, "¡¡Ja ja ja!! Tu no eres bienvenido aqui asqueroso Ciudadano", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.CiudadanosMatados + UserList(UserIndex).Faccion.CriminalesMatados < 70 Then
    Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 70 usuarios, solo has matado " & UserList(UserIndex).Faccion.CiudadanosMatados + UserList(UserIndex).Faccion.CriminalesMatados, str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 35 Then
    Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel 35!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If


If UserList(UserIndex).Faccion.Reenlistadas > 4 Then
    If UserList(UserIndex).Faccion.Reenlistadas = 200 Then
        Call WriteChatOverHead(UserIndex, "Has sido expulsado de las fuerzas oscuras y durante tu rebeldía has atacado a mi ejército. ¡Vete de aquí!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Else
        Call WriteChatOverHead(UserIndex, "¡Has sido expulsado de las fuerzas oscuras demasiadas veces!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    End If
    Exit Sub
End If

UserList(UserIndex).Faccion.Reenlistadas = UserList(UserIndex).Faccion.Reenlistadas + 1
UserList(UserIndex).Faccion.FuerzasCaos = 1

UserList(UserIndex).Faccion.Status = 2
Call WriteConsoleMsg(UserIndex, "¡Ahora perteneces al caos!", FontTypeNames.FONTTYPE_INFOIAO)
Call RefreshCharStatus(UserIndex)

Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al lado oscuro!!! Aqui tienes tus armaduras. Derrama sangre Ciudadana y Real y serás recompensado, lo prometo.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)

If UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0 Then
    Dim MiObj As obj
    MiObj.Amount = 1
    
    Select Case UserList(UserIndex).clase
        Case eClass.Cleric, eClass.Assasin
            MiObj.ObjIndex = 1500
        Case eClass.Paladin, eClass.Warrior, eClass.Trabajador, eClass.Hunter
            MiObj.ObjIndex = 1504
        Case eClass.Mage, eClass.Druid, eClass.Bard
            MiObj.ObjIndex = 1502
    End Select
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    UserList(UserIndex).Faccion.RecibioArmaduraCaos = 1
    UserList(UserIndex).Faccion.NivelIngreso = UserList(UserIndex).Stats.ELV
    UserList(UserIndex).Faccion.FechaIngreso = Date

End If


    UserList(UserIndex).Faccion.RecompensasCaos = 1
    UserList(UserIndex).Faccion.NextRecompensa = 140

'Call LogEjercitoCaos(UserList(UserIndex).Name & " ingresó el " & Date & " cuando era nivel " & UserList(UserIndex).Stats.ELV)

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Handles the way of gaining new ranks in the "Legión Oscura"
'***************************************************
Dim Ciudas As Long
Dim Lvl As Byte
Dim NextRecom As Long
Lvl = UserList(UserIndex).Stats.ELV
Ciudas = UserList(UserIndex).Faccion.CiudadanosMatados
NextRecom = UserList(UserIndex).Faccion.NextRecompensa

If Ciudas < NextRecom Then
    Call WriteChatOverHead(UserIndex, "Mata " & NextRecom - Ciudas & " Cuidadanos más para recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

Dim huborecompensa As Boolean

Select Case NextRecom
    Case 140:
        UserList(UserIndex).Faccion.RecompensasCaos = 2
        UserList(UserIndex).Faccion.NextRecompensa = 250
        huborecompensa = True
    
    Case 250:
        UserList(UserIndex).Faccion.RecompensasCaos = 3
        UserList(UserIndex).Faccion.NextRecompensa = 400
        huborecompensa = True
    
    Case 400:
        UserList(UserIndex).Faccion.RecompensasCaos = 4
        UserList(UserIndex).Faccion.NextRecompensa = 800
        huborecompensa = True
    
    Case 800:
        UserList(UserIndex).Faccion.RecompensasCaos = 5
        UserList(UserIndex).Faccion.NextRecompensa = 1500
        huborecompensa = True
    Case 1500:
       ' If Lvl < 27 Then
          '  Call WriteChatOverHead(UserIndex, "Mataste Suficientes Ciudadanos, pero te faltan " & 27 - Lvl & " Niveles para poder recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
          '  Exit Sub
       ' End If
       Call WriteChatOverHead(UserIndex, "¡Felicitaciones, eres de mis mejores guerreros, estas en lo más alto!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
    Case Else:
        Exit Sub
        
End Select

If huborecompensa Then
Call WriteChatOverHead(UserIndex, "¡¡¡Bien hecho " + TituloCaos(UserIndex) + ", aquí tienes tu recompensa!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)

End If

End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String
Select Case UserList(UserIndex).Faccion.RecompensasCaos
    Case 1
        TituloCaos = "Acólito"
    Case 2
        TituloCaos = "Alma Corrupta"
    Case 3
        TituloCaos = "Corruptor"
    Case 4
        TituloCaos = "Protector del Averno"
    Case 5
        TituloCaos = "Campeón de la Oscuridad"
End Select

End Function
