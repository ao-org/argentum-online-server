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
        MAPA_PRETORIANO = val(FileReader.GetValue("UBICACION", "Mapa"))
        PRETORIANO_X = val(FileReader.GetValue("UBICACION", "X"))
        PRETORIANO_Y = val(FileReader.GetValue("UBICACION", "Y"))
        PRETORIANO_RESPAWNEA = IIf(val(FileReader.GetValue("UBICACION", "Respawn")) = 1, True, False)

        'Configuracion de los NPC's
        Dim NroCombinaciones As Integer
104         NroCombinaciones = val(FileReader.GetValue("MAIN", "Combinaciones"))

106     ReDim PretorianDatNumbers(1 To NroCombinaciones)

        Dim TempInt        As Integer
        Dim counter        As Long
        Dim PretorianIndex As Integer

108     PretorianIndex = 1

        ' KINGS
110     TempInt = val(FileReader.GetValue("KING", "Cantidad"))
112     PretorianAIOffset(ePretorianAI.King) = 1

114     For counter = 1 To TempInt

            ' Alto
116         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("KING", "Alto" & counter))
118         PretorianIndex = PretorianIndex + 1
            ' Bajo
120         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("KING", "Bajo" & counter))
122         PretorianIndex = PretorianIndex + 1

124     Next counter

        ' HEALERS
126     TempInt = val(FileReader.GetValue("HEALER", "Cantidad"))
128     PretorianAIOffset(ePretorianAI.Healer) = PretorianIndex

130     For counter = 1 To TempInt

            ' Alto
132         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("HEALER", "Alto" & counter))
134         PretorianIndex = PretorianIndex + 1
            ' Bajo
136         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("HEALER", "Bajo" & counter))
138         PretorianIndex = PretorianIndex + 1

140     Next counter

        ' SPELLCASTER
142     TempInt = val(FileReader.GetValue("SPELLCASTER", "Cantidad"))
144     PretorianAIOffset(ePretorianAI.SpellCaster) = PretorianIndex

146     For counter = 1 To TempInt

            ' Alto
148         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("SPELLCASTER", "Alto" & counter))
150         PretorianIndex = PretorianIndex + 1
            ' Bajo
152         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("SPELLCASTER", "Bajo" & counter))
154         PretorianIndex = PretorianIndex + 1

156     Next counter

        ' SWORDSWINGER
158     TempInt = val(FileReader.GetValue("SWORDSWINGER", "Cantidad"))
160     PretorianAIOffset(ePretorianAI.SwordMaster) = PretorianIndex

162     For counter = 1 To TempInt

            ' Alto
164         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("SWORDSWINGER", "Alto" & counter))
166         PretorianIndex = PretorianIndex + 1
            ' Bajo
168         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("SWORDSWINGER", "Bajo" & counter))
170         PretorianIndex = PretorianIndex + 1

172     Next counter

        ' LONGRANGE
174     TempInt = val(FileReader.GetValue("LONGRANGE", "Cantidad"))
176     PretorianAIOffset(ePretorianAI.Shooter) = PretorianIndex

178     For counter = 1 To TempInt

            ' Alto
180         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("LONGRANGE", "Alto" & counter))
182         PretorianIndex = PretorianIndex + 1
            ' Bajo
184         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("LONGRANGE", "Bajo" & counter))
186         PretorianIndex = PretorianIndex + 1

188     Next counter

        ' THIEF
190     TempInt = val(FileReader.GetValue("THIEF", "Cantidad"))
192     PretorianAIOffset(ePretorianAI.Thief) = PretorianIndex

194     For counter = 1 To TempInt

            ' Alto
196         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("THIEF", "Alto" & counter))
198         PretorianIndex = PretorianIndex + 1
            ' Bajo
200         PretorianDatNumbers(PretorianIndex) = val(FileReader.GetValue("THIEF", "Bajo" & counter))
202         PretorianIndex = PretorianIndex + 1

204     Next counter

        ' Last
206     PretorianAIOffset(ePretorianAI.Last) = PretorianIndex

        ' Inicializa los clanes pretorianos
208     ReDim ClanPretoriano(ePretorianType.Default To ePretorianType.Custom) As clsClanPretoriano
210     Set ClanPretoriano(ePretorianType.Default) = New clsClanPretoriano ' Clan default
212     Set ClanPretoriano(ePretorianType.Custom) = New clsClanPretoriano ' Invocable por gms
        
        'Invocamos al Clan Pretoriano en su respectivo mapa.
        'Activando su respawn automatico.
        If Not ClanPretoriano(ePretorianType.Default).SpawnClan(MAPA_PRETORIANO, PRETORIANO_X, PRETORIANO_Y, ePretorianType.Default, PRETORIANO_RESPAWNEA) Then
            Call LogError("No se pudo invocar al Clan Pretoriano.")
            Exit Sub
        End If
        
        Set FileReader = Nothing
        
        Exit Sub

LoadPretorianData_Err:
        
        Set FileReader = Nothing
        
        Call RegistrarError(Err.Number, Err.description, "PraetoriansCoopNPC.LoadPretorianData", Erl)
        
        Resume Next
        
End Sub

Public Sub EliminarPretorianos(ByVal Mapa As Integer)

        On Error GoTo EliminarPretorianos_Err
        
        Dim Index As Byte
100     For Index = 1 To UBound(ClanPretoriano)
                 
            ' Search for the clan to be deleted
102         If ClanPretoriano(Index).ClanMap = Mapa Then
104             Call ClanPretoriano(Index).DeleteClan
                Exit For
        
            End If
                
        Next

        Exit Sub

EliminarPretorianos_Err:
        Call RegistrarError(Err.Number, Err.description, "PraetoriansCoopNPC.EliminarPretorianos", Erl)
        Resume Next
        
End Sub
