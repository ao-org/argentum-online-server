Attribute VB_Name = "Statistics"
'**************************************************************
' modStatistics.bas - Takes statistics on the game for later study.
'
' Implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

Private Type fragLvlRace

    matrix(1 To 50, 1 To 6) As Long

End Type

Private Type fragLvlLvl

    matrix(1 To 50, 1 To 50) As Long

End Type

Private fragLvlRaceData(1 To 7)               As fragLvlRace

Private fragLvlLvlData(1 To 7)                As fragLvlLvl

Private fragAlignmentLvlData(1 To 50, 1 To 4) As Long

'Currency just in case.... chats are way TOO often...
Private keyOcurrencies(255)                   As Currency

Public Sub StoreFrag(ByVal killer As Integer, ByVal victim As Integer)
        
        On Error GoTo StoreFrag_Err
        

        Dim clase     As Integer

        Dim raza      As Integer

        Dim alignment As Integer
    
100     If UserList(victim).Stats.ELV > 50 Or UserList(killer).Stats.ELV > 50 Then Exit Sub
    
102     Select Case UserList(killer).clase

            Case eClass.Assasin
104             clase = 1
        
106         Case eClass.Bard
108             clase = 2
        
110         Case eClass.Mage
112             clase = 3
        
114         Case eClass.Paladin
116             clase = 4
        
118         Case eClass.Warrior
120             clase = 5
        
122         Case eClass.Cleric
124             clase = 6
        
126         Case eClass.Hunter
128             clase = 7
        
130         Case Else
                Exit Sub

        End Select
    
132     Select Case UserList(killer).raza

            Case eRaza.Elfo
134             raza = 1
        
136         Case eRaza.Drow
138             raza = 2
        
140         Case eRaza.Enano
142             raza = 3
        
144         Case eRaza.Gnomo
146             raza = 4
        
148         Case eRaza.Humano
150             raza = 5
            
152         Case eRaza.Orco
154             raza = 6
        
156         Case Else
                Exit Sub

        End Select
    
158     If UserList(killer).Faccion.ArmadaReal Then
160         alignment = 1
162     ElseIf UserList(killer).Faccion.FuerzasCaos Then
164         alignment = 2
166     ElseIf Status(killer) = 2 Then
168         alignment = 3
        Else
170         alignment = 4

        End If
    
172     fragLvlRaceData(clase).matrix(UserList(killer).Stats.ELV, raza) = fragLvlRaceData(clase).matrix(UserList(killer).Stats.ELV, raza) + 1
    
174     fragLvlLvlData(clase).matrix(UserList(killer).Stats.ELV, UserList(victim).Stats.ELV) = fragLvlLvlData(clase).matrix(UserList(killer).Stats.ELV, UserList(victim).Stats.ELV) + 1
    
176     fragAlignmentLvlData(UserList(killer).Stats.ELV, alignment) = fragAlignmentLvlData(UserList(killer).Stats.ELV, alignment) + 1

        
        Exit Sub

StoreFrag_Err:
178     Call TraceError(Err.Number, Err.Description, "Statistics.StoreFrag", Erl)
180
        
End Sub

Public Sub ParseChat(ByRef S As String)
        
        On Error GoTo ParseChat_Err
        

        Dim i   As Long

        Dim Key As Integer
    
100     For i = 1 To Len(S)
102         Key = Asc(mid$(S, i, 1))
        
104         keyOcurrencies(Key) = keyOcurrencies(Key) + 1
106     Next i
    
        'Add a NULL-terminated to consider that possibility too....
108     keyOcurrencies(0) = keyOcurrencies(0) + 1

        
        Exit Sub

ParseChat_Err:
110     Call TraceError(Err.Number, Err.Description, "Statistics.ParseChat", Erl)
112
        
End Sub
