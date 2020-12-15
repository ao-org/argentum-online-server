Attribute VB_Name = "PraetoriansCoopNPC"
''**************************************************************
'' PraetoriansCoopNPC.bas - Handles the Praeorians NPCs.
''
'' Implemented by Mariano Barrou (El Oso)
''**************************************************************
'
''**************************************************************************
''This program is free software; you can redistribute it and/or modify
''it under the terms of the Affero General Public License;
''either version 1 of the License, or any later version.
''
''This program is distributed in the hope that it will be useful,
''but WITHOUT ANY WARRANTY; without even the implied warranty of
''MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
''Affero General Public License for more details.
''
''You should have received a copy of the Affero General Public License
''along with this program; if not, you can find it at http://www.affero.org/oagpl.html
''**************************************************************************
'
Option Explicit

'''''''
'' Pretorianos
'''''''
Public ClanPretoriano() As clsClanPretoriano

''''''''''''''''''''''''''''''''''''''''''''''
''Esta constante identifica en que mapa esta
''la fortaleza pretoriana (no es lo mismo de
''donde estan los NPCs!).
''Se extrae el dato del server.ini en sub LoadSIni
Public MAPA_PRETORIANO          As Integer
Public PRETORIANO_X             As Byte
Public PRETORIANO_Y             As Byte
Public PRETORIANO_RESPAWNEA     As Boolean

''''''''''''''''''''''''''''''''''''''''''''''
''Estos numeros son necesarios por cuestiones de
''sonido. Son los numeros de los wavs del cliente.
Public Const SONIDO_DRAGON_VIVO As Integer = 30

Public Enum ePretorianAI

    King = 1
    Healer
    SpellCaster
    SwordMaster
    Shooter
    Thief
    Last

End Enum

' Contains all the pretorian's combinations, and its the offsets
Public PretorianAIOffset(1 To 7) As Integer
Public PretorianDatNumbers()     As Integer
'
''Added by Nacho
''Cuantos pretorianos vivos quedan. Uno por cada alcoba
'Public pretorianosVivos As Integer
'
Private FileReader As clsIniReader

Public Sub LoadPretorianData()

        On Error GoTo LoadPretorianData_Err
 
100     Set FileReader = New clsIniReader
102     Call FileReader.Initialize(DatPath & "Pretorianos.dat")
        
        'Ubicación predeterminada de los pretorianos.
104     MAPA_PRETORIANO = val(FileReader.GetValue("UBICACION", "Mapa"))
106     PRETORIANO_X = val(FileReader.GetValue("UBICACION", "X"))
108     PRETORIANO_Y = val(FileReader.GetValue("UBICACION", "Y"))
110     PRETORIANO_RESPAWNEA = IIf(val(FileReader.GetValue("UBICACION", "Respawn")) = 1, True, False)

        'Configuracion de los NPC's
        Dim NroCombinaciones As Integer
112         NroCombinaciones = val(FileReader.GetValue("MAIN", "Combinaciones"))

114     ReDim PretorianDatNumbers(1 To NroCombinaciones)

        Dim TempInt        As Integer
        Dim counter        As Long
        Dim PretorianIndex As Integer

116     PretorianIndex = 1

        ' KINGS
118     TempInt = val(FileReader.GetValue("KING", "Cantidad"))
120     PretorianAIOffset(ePretorianAI.King) = 1

122     For counter = 1 To TempInt

            ' Alto
124         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("KING", "Alto" & counter))
126         PretorianIndex = PretorianIndex + 1
            ' Bajo
128         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("KING", "Bajo" & counter))
130         PretorianIndex = PretorianIndex + 1

132     Next counter

        ' HEALERS
134     TempInt = val(FileReader.GetValue("HEALER", "Cantidad"))
136     PretorianAIOffset(ePretorianAI.Healer) = PretorianIndex

138     For counter = 1 To TempInt

            ' Alto
140         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("HEALER", "Alto" & counter))
142         PretorianIndex = PretorianIndex + 1
            ' Bajo
144         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("HEALER", "Bajo" & counter))
146         PretorianIndex = PretorianIndex + 1

148     Next counter

        ' SPELLCASTER
150     TempInt = val(FileReader.GetValue("SPELLCASTER", "Cantidad"))
152     PretorianAIOffset(ePretorianAI.SpellCaster) = PretorianIndex

154     For counter = 1 To TempInt

            ' Alto
156         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("SPELLCASTER", "Alto" & counter))
158         PretorianIndex = PretorianIndex + 1
            ' Bajo
160         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("SPELLCASTER", "Bajo" & counter))
162         PretorianIndex = PretorianIndex + 1

164     Next counter

        ' SWORDSWINGER
166     TempInt = val(FileReader.GetValue("SWORDSWINGER", "Cantidad"))
168     PretorianAIOffset(ePretorianAI.SwordMaster) = PretorianIndex

170     For counter = 1 To TempInt

            ' Alto
172         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("SWORDSWINGER", "Alto" & counter))
174         PretorianIndex = PretorianIndex + 1
            ' Bajo
176         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("SWORDSWINGER", "Bajo" & counter))
178         PretorianIndex = PretorianIndex + 1

180     Next counter

        ' LONGRANGE
182     TempInt = val(FileReader.GetValue("LONGRANGE", "Cantidad"))
184     PretorianAIOffset(ePretorianAI.Shooter) = PretorianIndex

186     For counter = 1 To TempInt

            ' Alto
188         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("LONGRANGE", "Alto" & counter))
190         PretorianIndex = PretorianIndex + 1
            ' Bajo
192         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("LONGRANGE", "Bajo" & counter))
194         PretorianIndex = PretorianIndex + 1

196     Next counter

        ' THIEF
198     TempInt = val(FileReader.GetValue("THIEF", "Cantidad"))
200     PretorianAIOffset(ePretorianAI.Thief) = PretorianIndex

202     For counter = 1 To TempInt

            ' Alto
204         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("THIEF", "Alto" & counter))
206         PretorianIndex = PretorianIndex + 1
            ' Bajo
208         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("THIEF", "Bajo" & counter))
210         PretorianIndex = PretorianIndex + 1

212     Next counter

        ' Last
214     PretorianAIOffset(ePretorianAI.Last) = PretorianIndex

        ' Inicializa los clanes pretorianos
216     ReDim ClanPretoriano(ePretorianType.Default To ePretorianType.Custom) As clsClanPretoriano
218     Set ClanPretoriano(ePretorianType.Default) = New clsClanPretoriano ' Clan default
220     Set ClanPretoriano(ePretorianType.Custom) = New clsClanPretoriano ' Invocable por gms
        
        'Invocamos al Clan Pretoriano en su respectivo mapa.
        'Activando su respawn automatico.
222     If Not ClanPretoriano(ePretorianType.Default).SpawnClan(MAPA_PRETORIANO, PRETORIANO_X, PRETORIANO_Y, ePretorianType.Default, PRETORIANO_RESPAWNEA) Then
224         Call LogError("No se pudo invocar al Clan Pretoriano.")
            Exit Sub
        End If
        
226     Set FileReader = Nothing
        
        Exit Sub

LoadPretorianData_Err:
        
228     Set FileReader = Nothing
        
230     Call RegistrarError(Err.Number, Err.description, "PraetoriansCoopNPC.LoadPretorianData", Erl)
        
232     Resume Next
        
End Sub

Public Sub EliminarPretorianos(ByVal Mapa As Integer)

        On Error GoTo EliminarPretorianos_Err
        
        Dim index As Byte
100     For index = 1 To UBound(ClanPretoriano)
                 
            ' Search for the clan to be deleted
102         If ClanPretoriano(index).ClanMap = Mapa Then
104             Call ClanPretoriano(index).DeleteClan
                Exit For
        
            End If
                
        Next

        Exit Sub

EliminarPretorianos_Err:
106     Call RegistrarError(Err.Number, Err.description, "PraetoriansCoopNPC.EliminarPretorianos", Erl)
108     Resume Next
        
End Sub
