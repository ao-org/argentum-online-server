VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConsultasPopulares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' ConsultasPopulares.cls
'
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
'Modulo de consultas popluares
'En realidad el modulo inicial guardaba los datos de las votaciones
'en memoria, pero no le vi el punto, las votaciones son de vez en cuando nomas
'y el query a un .ini que registra todos los mails q ya votaron, es soportable
'si cuando uno vota y averiguo que el mail ya voto pero el PJ no, entonces seteo
'el flag de yavoto en el charfile ese tambien,
'por lo que la busqueda secuencial en el .dat que tiene todos los mails q ya votaron
'se hara .. 1 vez por PJ nomas.
'
'Hecha por el oso

Private Const ARCHIVOMAILS = "\logs\votaron.dat"

Private Const ARCHIVOCONFIG = "consultas.dat"

Private pEncuestaActualNum As Integer

Private pEncuestaActualTex As String

Private pNivelRequerido    As Integer

Private pOpciones()        As Integer

Public Property Let Numero(ByVal NumEncuesta As Integer)
        
        On Error GoTo Numero_Err
        
100     pEncuestaActualNum = NumEncuesta

        
        Exit Property

Numero_Err:
102     Call TraceError(Err.Number, Err.Description, "ConsultasPopulares.Numero", Erl)
104
        
End Property

Public Property Get Numero() As Integer
        
        On Error GoTo Numero_Err
        
100     Numero = pEncuestaActualNum

        
        Exit Property

Numero_Err:
102     Call TraceError(Err.Number, Err.Description, "ConsultasPopulares.Numero", Erl)
104
        
End Property

Public Property Let texto(ByVal Descripcion As String)
        
        On Error GoTo texto_Err
        
100     pEncuestaActualTex = Descripcion

        
        Exit Property

texto_Err:
102     Call TraceError(Err.Number, Err.Description, "ConsultasPopulares.texto", Erl)
104
        
End Property

Public Property Get texto() As String
        
        On Error GoTo texto_Err
        
100     texto = pEncuestaActualTex

        
        Exit Property

texto_Err:
102     Call TraceError(Err.Number, Err.Description, "ConsultasPopulares.texto", Erl)
104
        
End Property

Public Sub LoadData()
        
        On Error GoTo LoadData_Err
        

        Dim CantOpciones As Integer

        Dim i            As Integer

100     pEncuestaActualNum = val(GetVar(DatPath & ARCHIVOCONFIG, "INIT", "ConsultaActual"))
102     pEncuestaActualTex = GetVar(DatPath & ARCHIVOCONFIG, "INIT", "ConsultaActualTexto")
104     pNivelRequerido = GetVar(DatPath & ARCHIVOCONFIG, "INIT", "NivelRequerido")

106     If pEncuestaActualNum > 0 Then
            'cargo todas las opciones
108         CantOpciones = val(GetVar(DatPath & ARCHIVOCONFIG, "ENCUESTA" & pEncuestaActualNum, "CANTOPCIONES"))
110         ReDim pOpciones(1 To CantOpciones) As Integer

112         For i = 1 To CantOpciones
114             pOpciones(i) = val(GetVar(DatPath & ARCHIVOCONFIG, "ENCUESTA" & pEncuestaActualNum, "OPCION" & i))
116         Next i

        End If

        
        Exit Sub

LoadData_Err:
118     Call TraceError(Err.Number, Err.Description, "ConsultasPopulares.LoadData", Erl)
120
        
End Sub

Public Function doVotar(ByVal UserIndex As Integer, ByVal opcion As Integer) As String

        On Error GoTo errorh

        Dim YaVoto   As Boolean

        Dim CharFile As String

        Dim sufragio As Integer

        'revisar q no haya votado
        'grabar en el charfile el numero de encuesta
        'actualizar resultados encuesta
100     If pEncuestaActualNum = 0 Then
102         doVotar = "No hay consultas populares abiertas"
            Exit Function

        End If

104     CharFile = CharPath & UserList(UserIndex).Name & ".chr"
    
106     If (UserList(UserIndex).Stats.ELV >= pNivelRequerido) Then
108         If (OpcionValida(opcion)) Then
110             YaVoto = val(GetVar(CharFile, "CONSULTAS", "Voto")) >= pEncuestaActualNum

112             If Not YaVoto Then
114                 If Not MailYaVoto(UserList(UserIndex).Email) Then
                        'pj apto para votar
116                     sufragio = CLng(val(GetVar(DatPath & ARCHIVOCONFIG, "RESULTADOS" & pEncuestaActualNum, "V" & opcion)))
118                     sufragio = sufragio + 1
120                     Call WriteVar(DatPath & ARCHIVOCONFIG, "RESULTADOS" & pEncuestaActualNum, "V" & opcion, str(sufragio))
122                     doVotar = "Tu voto ha sido computado. Opcion: " & opcion
124                     Call MarcarPjComoQueYaVoto(UserIndex)
126                     Call MarcarMailComoQueYaVoto(UserList(UserIndex).Email)
                    Else
128                     Call MarcarPjComoQueYaVoto(UserIndex)
130                     doVotar = "Este email ya voto en la consulta: " & pEncuestaActualTex

                    End If

                Else
132                 doVotar = "Este personaje ya voto en la consulta: " & pEncuestaActualTex

                End If

            Else
134             doVotar = "Esa no es una opcion para votar"

            End If

        Else
136         doVotar = "Para votar en esta consulta debes ser nivel " & pNivelRequerido & " o superior"

        End If

        Exit Function
errorh:
138     Call LogError("Error en ConsultasPopularse.doVotar: " & Err.Description)

End Function

Public Function SendInfoEncuesta(ByVal UserIndex As Integer) As String
        
        On Error GoTo SendInfoEncuesta_Err
        

        Dim i As Integer

100     Call WriteConsoleMsg(UserIndex, "CONSULTA POPULAR NUMERO " & pEncuestaActualNum, e_FontTypeNames.FONTTYPE_GUILD)
102     Call WriteConsoleMsg(UserIndex, pEncuestaActualTex, e_FontTypeNames.FONTTYPE_GUILD)
104     Call WriteConsoleMsg(UserIndex, " Opciones de voto: ", e_FontTypeNames.FONTTYPE_GUILDMSG)

106     For i = 1 To UBound(pOpciones)
108         Call WriteConsoleMsg(UserIndex, "(Opcion " & i & "): " & GetVar(DatPath & ARCHIVOCONFIG, "ENCUESTA" & pEncuestaActualNum, "OPCION" & i), e_FontTypeNames.FONTTYPE_GUILDMSG)
110     Next i

112     Call WriteConsoleMsg(UserIndex, " Para votar una opcion, escribe /encuesta NUMERODEOPCION, por ejemplo para votar la opcion 1, escribe /encuesta 1. Tu voto no podra ser cambiado.", e_FontTypeNames.FONTTYPE_VENENO)

        
        Exit Function

SendInfoEncuesta_Err:
114     Call TraceError(Err.Number, Err.Description, "ConsultasPopulares.SendInfoEncuesta", Erl)
116
        
End Function

Private Sub MarcarPjComoQueYaVoto(ByVal UserIndex As Integer)
        
        On Error GoTo MarcarPjComoQueYaVoto_Err
        
102         Call SaveVotoDatabase(UserList(UserIndex).ID, pEncuestaActualNum)


        
        Exit Sub

MarcarPjComoQueYaVoto_Err:
106     Call TraceError(Err.Number, Err.Description, "ConsultasPopulares.MarcarPjComoQueYaVoto", Erl)
108
        
End Sub

Private Function MailYaVoto(ByVal Email As String) As Boolean
        
        On Error GoTo MailYaVoto_Err
        

        'abro el archivo, while not eof levnato 1 linea y comparo. Si da true, cierro
        Dim ArchN As Integer

        Dim Tmp   As String

100     MailYaVoto = False

102     ArchN = FreeFile()

104     Open DatPath & ARCHIVOMAILS For Input As #ArchN

106     Do While Not EOF(ArchN)
108         Line Input #ArchN, Tmp

110         If Email = Tmp Then
112             MailYaVoto = True
114             Close #ArchN
                Exit Function

            End If

        Loop

116     Close #ArchN

        
        Exit Function

MailYaVoto_Err:
118     Call TraceError(Err.Number, Err.Description, "ConsultasPopulares.MailYaVoto", Erl)
120
        
End Function

Private Sub MarcarMailComoQueYaVoto(ByVal Email As String)
        
        On Error GoTo MarcarMailComoQueYaVoto_Err
        

        Dim ArchN As Integer

100     ArchN = FreeFile()

102     Open DatPath & ARCHIVOMAILS For Append As #ArchN
104     Print #ArchN, Email

106     Close #ArchN

        
        Exit Sub

MarcarMailComoQueYaVoto_Err:
108     Call TraceError(Err.Number, Err.Description, "ConsultasPopulares.MarcarMailComoQueYaVoto", Erl)
110
        
End Sub

Private Function OpcionValida(ByVal opcion As Integer) As Boolean
        
        On Error GoTo OpcionValida_Err
        
100     OpcionValida = opcion > 0 And opcion <= UBound(pOpciones)

        
        Exit Function

OpcionValida_Err:
102     Call TraceError(Err.Number, Err.Description, "ConsultasPopulares.OpcionValida", Erl)
104
        
End Function

