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
        Call RegistrarError(Err.Number, Err.description, "Statistics.StoreFrag", Erl)
        Resume Next
        
End Sub

Public Sub DumpStatistics()
        
        On Error GoTo DumpStatistics_Err
        

        Dim handle As Integer

100     handle = FreeFile()
    
        Dim line As String

        Dim i    As Long

        Dim j    As Long
    
102     Open App.Path & "\logs\frags.txt" For Output As handle
    
        'Save lvl vs lvl frag matrix for each class - we use GNU Octave's ASCII file format
    
104     Print #handle, "# name: fragLvlLvl_Ase"
106     Print #handle, "# type: matrix"
108     Print #handle, "# rows: 50"
110     Print #handle, "# columns: 50"
    
112     For j = 1 To 50
114         For i = 1 To 50
116             line = line & " " & CStr(fragLvlLvlData(1).matrix(i, j))
118         Next i
        
120         Print #handle, line
122         line = vbNullString
124     Next j
    
126     Print #handle, "# name: fragLvlLvl_Bar"
128     Print #handle, "# type: matrix"
130     Print #handle, "# rows: 50"
132     Print #handle, "# columns: 50"
    
134     For j = 1 To 50
136         For i = 1 To 50
138             line = line & " " & CStr(fragLvlLvlData(2).matrix(i, j))
140         Next i
        
142         Print #handle, line
144         line = vbNullString
146     Next j
    
148     Print #handle, "# name: fragLvlLvl_Mag"
150     Print #handle, "# type: matrix"
152     Print #handle, "# rows: 50"
154     Print #handle, "# columns: 50"
    
156     For j = 1 To 50
158         For i = 1 To 50
160             line = line & " " & CStr(fragLvlLvlData(3).matrix(i, j))
162         Next i
        
164         Print #handle, line
166         line = vbNullString
168     Next j
    
170     Print #handle, "# name: fragLvlLvl_Pal"
172     Print #handle, "# type: matrix"
174     Print #handle, "# rows: 50"
176     Print #handle, "# columns: 50"
    
178     For j = 1 To 50
180         For i = 1 To 50
182             line = line & " " & CStr(fragLvlLvlData(4).matrix(i, j))
184         Next i
        
186         Print #handle, line
188         line = vbNullString
190     Next j
    
192     Print #handle, "# name: fragLvlLvl_Gue"
194     Print #handle, "# type: matrix"
196     Print #handle, "# rows: 50"
198     Print #handle, "# columns: 50"
    
200     For j = 1 To 50
202         For i = 1 To 50
204             line = line & " " & CStr(fragLvlLvlData(5).matrix(i, j))
206         Next i
        
208         Print #handle, line
210         line = vbNullString
212     Next j
    
214     Print #handle, "# name: fragLvlLvl_Cle"
216     Print #handle, "# type: matrix"
218     Print #handle, "# rows: 50"
220     Print #handle, "# columns: 50"
    
222     For j = 1 To 50
224         For i = 1 To 50
226             line = line & " " & CStr(fragLvlLvlData(6).matrix(i, j))
228         Next i
        
230         Print #handle, line
232         line = vbNullString
234     Next j
    
236     Print #handle, "# name: fragLvlLvl_Caz"
238     Print #handle, "# type: matrix"
240     Print #handle, "# rows: 50"
242     Print #handle, "# columns: 50"
    
244     For j = 1 To 50
246         For i = 1 To 50
248             line = line & " " & CStr(fragLvlLvlData(7).matrix(i, j))
250         Next i
        
252         Print #handle, line
254         line = vbNullString
256     Next j
    
        'Save lvl vs race frag matrix for each class - we use GNU Octave's ASCII file format
    
258     Print #handle, "# name: fragLvlRace_Ase"
260     Print #handle, "# type: matrix"
262     Print #handle, "# rows: 5"
264     Print #handle, "# columns: 50"
    
266     For j = 1 To 5
268         For i = 1 To 50
270             line = line & " " & CStr(fragLvlRaceData(1).matrix(i, j))
272         Next i
        
274         Print #handle, line
276         line = vbNullString
278     Next j
    
280     Print #handle, "# name: fragLvlRace_Bar"
282     Print #handle, "# type: matrix"
284     Print #handle, "# rows: 5"
286     Print #handle, "# columns: 50"
    
288     For j = 1 To 5
290         For i = 1 To 50
292             line = line & " " & CStr(fragLvlRaceData(2).matrix(i, j))
294         Next i
        
296         Print #handle, line
298         line = vbNullString
300     Next j
    
302     Print #handle, "# name: fragLvlRace_Mag"
304     Print #handle, "# type: matrix"
306     Print #handle, "# rows: 5"
308     Print #handle, "# columns: 50"
    
310     For j = 1 To 5
312         For i = 1 To 50
314             line = line & " " & CStr(fragLvlRaceData(3).matrix(i, j))
316         Next i
        
318         Print #handle, line
320         line = vbNullString
322     Next j
    
324     Print #handle, "# name: fragLvlRace_Pal"
326     Print #handle, "# type: matrix"
328     Print #handle, "# rows: 5"
330     Print #handle, "# columns: 50"
    
332     For j = 1 To 5
334         For i = 1 To 50
336             line = line & " " & CStr(fragLvlRaceData(4).matrix(i, j))
338         Next i
        
340         Print #handle, line
342         line = vbNullString
344     Next j
    
346     Print #handle, "# name: fragLvlRace_Gue"
348     Print #handle, "# type: matrix"
350     Print #handle, "# rows: 5"
352     Print #handle, "# columns: 50"
    
354     For j = 1 To 5
356         For i = 1 To 50
358             line = line & " " & CStr(fragLvlRaceData(5).matrix(i, j))
360         Next i
        
362         Print #handle, line
364         line = vbNullString
366     Next j
    
368     Print #handle, "# name: fragLvlRace_Cle"
370     Print #handle, "# type: matrix"
372     Print #handle, "# rows: 5"
374     Print #handle, "# columns: 50"
    
376     For j = 1 To 5
378         For i = 1 To 50
380             line = line & " " & CStr(fragLvlRaceData(6).matrix(i, j))
382         Next i
        
384         Print #handle, line
386         line = vbNullString
388     Next j
    
390     Print #handle, "# name: fragLvlRace_Caz"
392     Print #handle, "# type: matrix"
394     Print #handle, "# rows: 5"
396     Print #handle, "# columns: 50"
    
398     For j = 1 To 5
400         For i = 1 To 50
402             line = line & " " & CStr(fragLvlRaceData(7).matrix(i, j))
404         Next i
        
406         Print #handle, line
408         line = vbNullString
410     Next j
    
        'Save lvl vs class frag matrix for each race - we use GNU Octave's ASCII file format
    
412     Print #handle, "# name: fragLvlClass_Elf"
414     Print #handle, "# type: matrix"
416     Print #handle, "# rows: 7"
418     Print #handle, "# columns: 50"
    
420     For j = 1 To 7
422         For i = 1 To 50
424             line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 1))
426         Next i
        
428         Print #handle, line
430         line = vbNullString
432     Next j
    
434     Print #handle, "# name: fragLvlClass_Dar"
436     Print #handle, "# type: matrix"
438     Print #handle, "# rows: 7"
440     Print #handle, "# columns: 50"
    
442     For j = 1 To 7
444         For i = 1 To 50
446             line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 2))
448         Next i
        
450         Print #handle, line
452         line = vbNullString
454     Next j
    
456     Print #handle, "# name: fragLvlClass_Dwa"
458     Print #handle, "# type: matrix"
460     Print #handle, "# rows: 7"
462     Print #handle, "# columns: 50"
    
464     For j = 1 To 7
466         For i = 1 To 50
468             line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 3))
470         Next i
        
472         Print #handle, line
474         line = vbNullString
476     Next j
    
478     Print #handle, "# name: fragLvlClass_Gno"
480     Print #handle, "# type: matrix"
482     Print #handle, "# rows: 7"
484     Print #handle, "# columns: 50"
    
486     For j = 1 To 7
488         For i = 1 To 50
490             line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 4))
492         Next i
        
494         Print #handle, line
496         line = vbNullString
498     Next j
    
500     Print #handle, "# name: fragLvlClass_Hum"
502     Print #handle, "# type: matrix"
504     Print #handle, "# rows: 7"
506     Print #handle, "# columns: 50"
    
508     For j = 1 To 7
510         For i = 1 To 50
512             line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 5))
514         Next i
        
516         Print #handle, line
518         line = vbNullString
520     Next j
    
        'Save lvl vs alignment frag matrix for each race - we use GNU Octave's ASCII file format
    
522     Print #handle, "# name: fragAlignmentLvl"
524     Print #handle, "# type: matrix"
526     Print #handle, "# rows: 4"
528     Print #handle, "# columns: 50"
    
530     For j = 1 To 4
532         For i = 1 To 50
534             line = line & " " & CStr(fragAlignmentLvlData(i, j))
536         Next i
        
538         Print #handle, line
540         line = vbNullString
542     Next j
    
544     Close handle
    
        'Dump Chat statistics
546     handle = FreeFile()
    
548     Open App.Path & "\logs\huffman.log" For Output As handle
    
        Dim Total As Currency
    
        'Compute total characters
550     For i = 0 To 255
552         Total = Total + keyOcurrencies(i)
554     Next i
    
        'Show each character's ocurrencies
556     If Total <> 0 Then

558         For i = 0 To 255
560             Print #handle, CStr(i) & "    " & CStr(Round(keyOcurrencies(i) / Total, 8))
562         Next i

        End If
    
564     Print #handle, "TOTAL =    " & CStr(Total)
    
566     Close handle

        
        Exit Sub

DumpStatistics_Err:
        Call RegistrarError(Err.Number, Err.description, "Statistics.DumpStatistics", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "Statistics.ParseChat", Erl)
        Resume Next
        
End Sub
