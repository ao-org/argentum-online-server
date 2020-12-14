VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsIniManager.cls - Loads INI files into memory and applies Binary Search to get values at high speed.
' Use it instead of GetVar when reading several values form the same file at once, otherwise it's not usefull.
' Based in the idea of AlejoLP and his clsLeerInis.
'
' Designed and implemented by Juan Martin Sotuyo Dodero (Maraxus)
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

''
'Loads a complete INI file into memory and sorts it's data and keys for faster searches.
'It is MUCH faster than GetPrivateProfileStringA if you search for several values within a file, otherwise stick to the API.
'It's particularly usefull when you can keep a whole file in memory such as NPCs.dat'
' Based in the idea of AlejoLP and his clsLeerInis.
'
' @author Juan Martin Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.1.0
' @date 20060501

'01/05/2006 - Juan Martin Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - First Release
'
'01/04/2008 - Juan Martin Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - Add: KeyExists method allows to check for valid section keys.
'
'10/10/2010 - Torres Patricio(Pato)
'   - Add:  AddNode method to add a main node.
'           AddKey method to add key into the refered main node.
'           DumpFile method to dump the data into a file.

Option Explicit
Option Base 0

''
'Structure that contains a value and it's key in a INI file
'
' @param    key String containing the key associated to the value.
' @param    value String containing the value of the INI entry.
' @see      MainNode
'

Private Type ChildNode

    Key As String
    Value As String

End Type

''
'Structure that contains all info under a tag in a INI file.
'Such tags are indicated with the "[" and "]" characters.
'
' @param    name String containing the text within the "[" and "]" characters.
'It's the key used when searching for a main section of the INI data.
' @param    values Array of ChildNodes, each containing a value entry along with it's key.
' @param    numValues Number of entrys in the main node.

Private Type MainNode

    name As String
    values() As ChildNode
    numValues As Integer

End Type

''
'Containts all Main sections of the loaded INI file
Private fileData() As MainNode

''
'Stores the total number of main sections in the loaded INI file
Private MainNodes  As Long

''
'Default constructor. Does nothing.

Private Sub Class_Initialize()

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero
    'Last Modify Date: 5/01/2006
    '
    '**************************************************************
End Sub

''
'Destroy every array and deallocates al memory.
'

Private Sub Class_Terminate()
        
    On Error GoTo Class_Terminate_Err
        

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero
    'Last Modify Date: 5/01/2006
    '
    '**************************************************************
    Dim i As Long
    
    'Clean up
    If MainNodes Then

        For i = 1 To MainNodes - 1
            Erase fileData(i).values
        Next i
        
        Erase fileData

    End If
    
    MainNodes = 0

        
    Exit Sub

Class_Terminate_Err:
    Call RegistrarError(Err.Number, Err.description, "clsIniReader.Class_Terminate", Erl)
    Resume Next
        
End Sub

''
'Loads a INI file so it's values can be read. Must be called before being able to use GetValue.
'
' @param    file Complete path of the INI file to be loaded.
' @see      GetValue

Public Sub Initialize(ByVal File As String)
        '**************************************************************
        'Author: Juan Martin Sotuyo Dodero
        'Last Modify Date: 27/07/2006
        'Opens the requested file and loads it's data into memory
        '**************************************************************
        Dim handle As Integer

        Dim Text   As String

        Dim Pos    As Long
    
        'Prevent memory losses if we are attempting to reload a file....
100     Call Class_Terminate
    
        'Get a free handle and start reading line by line until the end
102     handle = FreeFile
    
104     Open File For Input As handle
    
106     Do Until EOF(handle)
108         Line Input #handle, Text
        
            'Is it null??
110         If Len(Text) Then

                'If it starts with '[' it is a main node or nothing (GetPrivateProfileStringA works this way), otherwise it's a value
112             If Left$(Text, 1) = "[" Then
                    'If it has an ending ']' it's a main node, otherwise it's nothing
114                 Pos = InStr(2, Text, "]")

116                 If Pos Then
                        'Add a main node
118                     ReDim Preserve fileData(MainNodes) As MainNode
                    
120                     fileData(MainNodes).name = UCase$(Trim$(mid$(Text, 2, Pos - 2)))
                    
122                     MainNodes = MainNodes + 1

                    End If

                Else
                    'So it's a value. Check if it has a '=', otherwise it's nothing
124                 Pos = InStr(2, Text, "=")

126                 If Pos Then

                        'Is it under any main node??
128                     If MainNodes Then

130                         With fileData(MainNodes - 1)
                                'Add it to the main node's value
132                             ReDim Preserve .values(.numValues) As ChildNode
                            
134                             .values(.numValues).Value = Right$(Text, Len(Text) - Pos)
136                             .values(.numValues).Key = UCase$(Left$(Text, Pos - 1))
                            
138                             .numValues = .numValues + 1

                            End With

                        End If

                    End If

                End If

            End If

        Loop
    
140     Close handle
    
        Dim i As Long
    
142     If MainNodes Then
            'Sort main nodes to allow binary search
144         Call SortMainNodes(0, MainNodes - 1)
        
            'Sort values of each node to allow binary search
146         For i = 0 To MainNodes - 1

148             If fileData(i).numValues Then Call SortChildNodes(fileData(i), 0, fileData(i).numValues - 1)
150         Next i

        End If
        
End Sub

''
'Sorts all child nodes within the given MainNode alphabetically by their keys. Uses quicksort.
'
' @param    Node The MainNode whose values are to be sorted.
' @param    first The first index to consider when sorting.
' @param    last The last index to be considered when sorting.

Private Sub SortChildNodes(ByRef Node As MainNode, ByVal First As Integer, ByVal Last As Integer)
        
        On Error GoTo SortChildNodes_Err
        

        '**************************************************************
        'Author: Juan Martin Sotuyo Dodero
        'Last Modify Date: 5/01/2006
        'Sorts the list of values in a given MainNode using quicksort,
        'this allows the use of Binary Search for faster searches
        '**************************************************************
        Dim min  As Integer      'First item in the list

        Dim max  As Integer      'Last item in the list

        Dim comp As String      'Item used to compare

        Dim temp As ChildNode
    
100     min = First
102     max = Last
    
104     With Node
106         comp = .values((min + max) \ 2).Key
    
108         Do While min <= max
110             Do While .values(min).Key < comp And min < Last
112                 min = min + 1
                Loop

114             Do While .values(max).Key > comp And max > First
116                 max = max - 1
                Loop

118             If min <= max Then
120                 temp = .values(min)
122                 .values(min) = .values(max)
124                 .values(max) = temp
126                 min = min + 1
128                 max = max - 1

                End If

            Loop

        End With
    
130     If First < max Then SortChildNodes Node, First, max
132     If min < Last Then SortChildNodes Node, min, Last

        
        Exit Sub

SortChildNodes_Err:
134     Call RegistrarError(Err.Number, Err.description, "clsIniReader.SortChildNodes", Erl)
136     Resume Next
        
End Sub

''
'Sorts all main nodes in the loaded INI file alphabetically by their names. Uses quicksort.
'
' @param    first The first index to consider when sorting.
' @param    last The last index to be considered when sorting.

Private Sub SortMainNodes(ByVal First As Integer, ByVal Last As Integer)
        
        On Error GoTo SortMainNodes_Err
        

        '**************************************************************
        'Author: Juan Martin Sotuyo Dodero
        'Last Modify Date: 5/01/2006
        'Sorts the MainNodes list using quicksort,
        'this allows the use of Binary Search for faster searches
        '**************************************************************
        Dim min  As Integer      'First item in the list

        Dim max  As Integer      'Last item in the list

        Dim comp As String      'Item used to compare

        Dim temp As MainNode
    
100     min = First
102     max = Last
    
104     comp = fileData((min + max) \ 2).name
    
106     Do While min <= max
108         Do While fileData(min).name < comp And min < Last
110             min = min + 1
            Loop

112         Do While fileData(max).name > comp And max > First
114             max = max - 1
            Loop

116         If min <= max Then
118             temp = fileData(min)
120             fileData(min) = fileData(max)
122             fileData(max) = temp
124             min = min + 1
126             max = max - 1

            End If

        Loop
    
128     If First < max Then SortMainNodes First, max
130     If min < Last Then SortMainNodes min, Last

        
        Exit Sub

SortMainNodes_Err:
132     Call RegistrarError(Err.Number, Err.description, "clsIniReader.SortMainNodes", Erl)
134     Resume Next
        
End Sub

''
'Searches for a given key within a given main section and if it exists retrieves it's value, otherwise a null string
'
' @param    Main The name of the main section in which we will be searching.
' @param    key The key of the value we are looking for.
' @returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.

Public Function GetValue(ByVal Main As String, ByVal Key As String) As String
        
        On Error GoTo GetValue_Err
        

        '**************************************************************
        'Author: Juan Martin Sotuyo Dodero
        'Last Modify Date: 5/01/2006
        'Returns a value if the key and main node exist, or a nullstring otherwise
        '**************************************************************
        Dim i As Long

        Dim j As Long
    
        'Search for the main node
100     i = FindMain(UCase$(Main))
    
102     If i >= 0 Then
            'If valid, binary search among keys
104         j = FindKey(fileData(i), UCase$(Key))
        
            'If we found it we return it
106         If j >= 0 Then GetValue = fileData(i).values(j).Value

        End If

        
        Exit Function

GetValue_Err:
108     Call RegistrarError(Err.Number, Err.description, "clsIniReader.GetValue", Erl)
110     Resume Next
        
End Function

''
'Searches for a given key within a given main section and if it exists retrieves it's value, otherwise a null string
'
' @param    Main The name of the main section in which we will be searching.
' @param    key The key of the value we are looking for.
' @returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.

Public Sub ChangeValue(ByVal Main As String, ByVal Key As String, ByVal Value As String)
        
        On Error GoTo ChangeValue_Err
        

        '**************************************************************
        'Author: ZaMa
        'Last Modify Date: 27/05/2009
        'If the key and main node exist, changes the value
        '**************************************************************
        Dim i As Long

        Dim j As Long
    
        'Search for the main node
100     i = FindMain(UCase$(Main))
    
102     If i < 0 Then i = AddNode(Main)
    
        'If valid, binary search among keys
104     j = FindKey(fileData(i), UCase$(Key))
    
106     If j < 0 Then
108         Call AddKey(Main, Key, Value)
        Else
            'If we found it we change it
110         fileData(i).values(j).Value = Value

        End If

        
        Exit Sub

ChangeValue_Err:
112     Call RegistrarError(Err.Number, Err.description, "clsIniReader.ChangeValue", Erl)
114     Resume Next
        
End Sub

''
'Searches for a given key within a given main node and returns the index in which it's stored or the negation of the index in which it should be if not found.
'
' @param    Node The MainNode among whose value entries we will be searching.
' @param    key The key of the value we are looking for.
' @returns  The index in which the value with the key we are looking for is stored or the negation of the index in which it should be if not found.

Private Function FindKey(ByRef Node As MainNode, ByVal Key As String) As Long
        
        On Error GoTo FindKey_Err
        

        '**************************************************************
        'Author: Juan Martin Sotuyo Dodero
        'Last Modify Date: 5/01/2006
        'Returns the index of the value which key matches the requested one,
        'or the negation of the position were it should be if not found
        '**************************************************************
        Dim min As Long

        Dim max As Long

        Dim mid As Long
    
100     min = 0
102     max = Node.numValues - 1
    
104     Do While min <= max
106         mid = (min + max) \ 2
        
108         If Node.values(mid).Key < Key Then
110             min = mid + 1
112         ElseIf Node.values(mid).Key > Key Then
114             max = mid - 1
            Else
                'We found it
116             FindKey = mid
                Exit Function

            End If

        Loop
    
        'Not found, return the negation of the position where it should be
        '(all higher values are to the right of the list and lower values are to the left)
118     FindKey = Not min

        
        Exit Function

FindKey_Err:
120     Call RegistrarError(Err.Number, Err.description, "clsIniReader.FindKey", Erl)
122     Resume Next
        
End Function

''
'Searches for a main section with the given name within the loaded INI file and returns the index in which it's stored or the negation of the index in which it should be if not found.
'
' @param    name The name of the MainNode we are looking for.
' @returns  The index in which the main section we are looking for is stored or the negation of the index in which it should be if not found.

Private Function FindMain(ByVal name As String) As Long
        
        On Error GoTo FindMain_Err
        

        '**************************************************************
        'Author: Juan Martin Sotuyo Dodero
        'Last Modify Date: 5/01/2006
        'Returns the index of the MainNode which name matches the requested one,
        'or the negation of the position were it should be if not found
        '**************************************************************
        Dim min As Long

        Dim max As Long

        Dim mid As Long
    
100     min = 0
102     max = MainNodes - 1
    
104     Do While min <= max
106         mid = (min + max) \ 2
        
108         If fileData(mid).name < name Then
110             min = mid + 1
112         ElseIf fileData(mid).name > name Then
114             max = mid - 1
            Else
                'We found it
116             FindMain = mid
                Exit Function

            End If

        Loop
    
        'Not found, return the negation of the position where it should be
        '(all higher values are to the right of the list and lower values are to the left)
118     FindMain = Not min

        
        Exit Function

FindMain_Err:
120     Call RegistrarError(Err.Number, Err.description, "clsIniReader.FindMain", Erl)
122     Resume Next
        
End Function

''
'Checks wether a given key exists or not.
'
' @param    name    The name of the element whose existance is being checked.
' @returns  True if the key exists, false otherwise.

Public Function KeyExists(ByVal name As String) As Boolean
        '**************************************************************
        'Author: Juan Martin Sotuyo Dodero
        'Last Modify Date: 04/01/2008
        'Returns true of the key exists, false otherwise.
        '**************************************************************
        
        On Error GoTo KeyExists_Err
        
100     KeyExists = FindMain(UCase$(name)) >= 0

        
        Exit Function

KeyExists_Err:
102     Call RegistrarError(Err.Number, Err.description, "clsIniReader.KeyExists", Erl)
104     Resume Next
        
End Function

Private Function AddNode(ByRef name As String) As Long
        
        On Error GoTo AddNode_Err
        

        '**************************************************************
        'Author: Torres Patricio (Pato)
        'Last Modify Date: 10/10/2010
        '
        '**************************************************************
        Dim i       As Long

        Dim MainPos As Long

100     name = UCase$(name)

        'Add a main node
102     ReDim Preserve fileData(MainNodes) As MainNode

104     If MainNodes Then
106         i = MainNodes - 1
108         MainPos = Not FindMain(name)
    
110         Do While i >= MainPos
112             fileData(i + 1) = fileData(i)
114             i = i - 1
            Loop
    
116         fileData(MainPos).numValues = 0
118         Erase fileData(MainPos).values

        End If

120     fileData(MainPos).name = name

122     MainNodes = MainNodes + 1

124     AddNode = MainPos

        
        Exit Function

AddNode_Err:
126     Call RegistrarError(Err.Number, Err.description, "clsIniReader.AddNode", Erl)
128     Resume Next
        
End Function

Private Function AddKey(ByRef Main As String, ByRef Key As String, Optional ByRef Value As String = vbNullString) As Long
        
        On Error GoTo AddKey_Err
        

        '**************************************************************
        'Author: Torres Patricio (Pato)
        'Last Modify Date: 10/10/2010
        '
        '**************************************************************
        Dim MainPos As Long

        Dim KeyPos  As Long

        Dim i       As Long

100     Main = UCase$(Main)
102     Key = UCase$(Key)

104     MainPos = FindMain(Main)

106     If MainPos < 0 Then MainPos = AddNode(Main)

108     With fileData(MainPos)
            'Add it to the main node's value
110         ReDim Preserve .values(.numValues) As ChildNode
    
112         If .numValues > 0 Then
114             i = .numValues - 1
116             KeyPos = Not FindKey(fileData(MainPos), Key)
        
118             Do While i >= KeyPos
120                 .values(i + 1) = .values(i)
122                 i = i - 1
                Loop

            End If
    
124         .values(KeyPos).Key = Key
126         .values(KeyPos).Value = Value
    
128         .numValues = .numValues + 1

        End With

130     AddKey = KeyPos

        
        Exit Function

AddKey_Err:
132     Call RegistrarError(Err.Number, Err.description, "clsIniReader.AddKey", Erl)
134     Resume Next
        
End Function

Public Sub DumpFile(ByRef File As String)
        
        On Error GoTo DumpFile_Err
        

        '**************************************************************
        'Author: Torres Patricio (Pato)
        'Last Modify Date: 10/10/2010
        '
        '**************************************************************
        Dim hFile As Integer

        Dim i     As Long

        Dim j     As Long

100     hFile = FreeFile()

102     Open File For Output As hFile

104     For i = 0 To MainNodes - 1

106         With fileData(i)
108             Print #hFile, "[" & .name & "]"
            
110             For j = 0 To .numValues - 1
112                 Print #hFile, .values(j).Key & "=" & .values(j).Value
114             Next j
            
116             Print #hFile, ""

            End With

118     Next i
    
120     Close hFile

        
        Exit Sub

DumpFile_Err:
122     Call RegistrarError(Err.Number, Err.description, "clsIniReader.DumpFile", Erl)
124     Resume Next
        
End Sub

