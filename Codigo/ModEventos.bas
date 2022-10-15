Attribute VB_Name = "ModEventos"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Public HoraEvento           As Byte

Public TiempoRestanteEvento As Byte

Public EventoActivo         As Boolean

Public EventoAcutal         As t_EventoPropiedades

Public Evento(0 To 23)      As t_EventoPropiedades

Public Type t_EventoPropiedades

    Tipo As Byte
    Duracion As Byte
    multiplicacion As Byte

End Type

Public ExpMultOld         As Integer

Public OroMultOld         As Integer

Public DropMultOld        As Integer

Public RecoleccionMultOld As Double

Public PublicidadEvento   As String

Enum TipoEvento
    Invasion
End Enum

Public Sub CheckEvento(ByVal Hora As Byte)
        
        On Error GoTo CheckEvento_Err
        

100     If EventoActivo = True Then Exit Sub

        Dim aviso As String

102     aviso = "Eventos> Nuevo evento iniciado: "
104     PublicidadEvento = "Evento en curso>"

106     Select Case Evento(Hora).Tipo

            Case 1
108             OroMult = OroMult * Evento(Hora).multiplicacion
110             aviso = aviso & " Oro multiplicado por " & Evento(Hora).multiplicacion & " - Duración del evento: " & Evento(Hora).Duracion & " minutos."
112             PublicidadEvento = PublicidadEvento & " Oro multiplicado por " & Evento(Hora).multiplicacion

114         Case 2
116             ExpMult = ExpMult * Evento(Hora).multiplicacion
118             aviso = aviso & " Experiencia multiplicada por " & Evento(Hora).multiplicacion & " - Duración del evento: " & Evento(Hora).Duracion & " minutos."
120             PublicidadEvento = PublicidadEvento & " Experiencia multiplicada por " & Evento(Hora).multiplicacion

122         Case 3
124             RecoleccionMult = RecoleccionMult * Evento(Hora).multiplicacion
126             aviso = aviso & " Recoleccion multiplicada por " & Evento(Hora).multiplicacion & " - Duración del evento: " & Evento(Hora).Duracion & " minutos."
128             PublicidadEvento = PublicidadEvento & " Recoleccion multiplicada por " & Evento(Hora).multiplicacion

130         Case 4
132             DropMult = DropMult / Evento(Hora).multiplicacion
134             aviso = aviso & " Dropeo multiplicado por " & Evento(Hora).multiplicacion & " - Duración del evento: " & Evento(Hora).Duracion & " minutos."
136             PublicidadEvento = PublicidadEvento & " Dropeo multiplicado por " & Evento(Hora).multiplicacion

138         Case 5
140             ExpMult = ExpMult * Evento(Hora).multiplicacion
142             OroMult = OroMult * Evento(Hora).multiplicacion
144             aviso = aviso & " Oro y experiencia multiplicados por " & Evento(Hora).multiplicacion & " - Duración del evento: " & Evento(Hora).Duracion & " minutos."
146             PublicidadEvento = PublicidadEvento & " Oro y experiencia multiplicados por " & Evento(Hora).multiplicacion

148         Case 6
150             ExpMult = ExpMult * Evento(Hora).multiplicacion
152             OroMult = OroMult * Evento(Hora).multiplicacion
154             RecoleccionMult = RecoleccionMult * Evento(Hora).multiplicacion
156             aviso = aviso & " Oro, experiencia y recoleccion multiplicados por " & Evento(Hora).multiplicacion & " - Duración del evento: " & Evento(Hora).Duracion & " minutos."
158             PublicidadEvento = PublicidadEvento & " Oro, experiencia y recoleccion multiplicados por " & Evento(Hora).multiplicacion

160         Case 7
162             ExpMult = ExpMult * Evento(Hora).multiplicacion
164             OroMult = OroMult * Evento(Hora).multiplicacion
166             DropMult = DropMult / Evento(Hora).multiplicacion
168             RecoleccionMult = RecoleccionMult * Evento(Hora).multiplicacion
170             aviso = aviso & " Oro, experiencia, recoleccion y dropeo multiplicados por " & Evento(Hora).multiplicacion & " - Duración del evento: " & Evento(Hora).Duracion & " minutos."
172             PublicidadEvento = PublicidadEvento & " Oro, experiencia, recoleccion y dropeo multiplicados por " & Evento(Hora).multiplicacion

174         Case Else

176             EventoActivo = False
                Exit Sub
        
        End Select

178     Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(553, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno

180     Call AgregarAConsola(aviso)

182     EventoAcutal.Duracion = Evento(Hora).Duracion
184     EventoAcutal.multiplicacion = Evento(Hora).multiplicacion
186     EventoAcutal.Tipo = Evento(Hora).Tipo

188     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(aviso, e_FontTypeNames.FONTTYPE_New_Eventos))
190     TiempoRestanteEvento = Evento(Hora).Duracion
192     frmMain.Evento.Enabled = True
194     EventoActivo = True
        
        
        Exit Sub

CheckEvento_Err:
196     Call TraceError(Err.Number, Err.Description, "ModEventos.CheckEvento", Erl)

        
End Sub

Public Sub FinalizarEvento()
        
        On Error GoTo FinalizarEvento_Err
        
100     frmMain.Evento.Enabled = False
102     EventoActivo = False

104     Select Case EventoAcutal.Tipo

            Case 1
106             OroMult = OroMultOld
       
108         Case 2
110             ExpMult = ExpMultOld
       
112         Case 3
114             RecoleccionMult = RecoleccionMultOld
  
116         Case 4
118             DropMult = DropMultOld
        
120         Case 5
122             ExpMult = ExpMultOld
124             OroMult = OroMultOld

126         Case 6
128             ExpMult = ExpMultOld
130             OroMult = OroMultOld
132             RecoleccionMult = RecoleccionMultOld

134         Case 7
136             ExpMult = ExpMultOld
138             OroMult = OroMultOld
140             DropMult = DropMultOld
142             RecoleccionMult = RecoleccionMultOld

144         Case Else
                Exit Sub
        
        End Select

146     Call AgregarAConsola("Eventos > Evento finalizado.")
148     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos > Evento finalizado.", e_FontTypeNames.FONTTYPE_New_Eventos))
150     Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(551, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno

        
        Exit Sub

FinalizarEvento_Err:
152     Call TraceError(Err.Number, Err.Description, "ModEventos.FinalizarEvento", Erl)

        
End Sub

Public Function DescribirEvento(ByVal Hora As Byte) As String
        
        On Error GoTo DescribirEvento_Err
        

        Dim aviso As String

100     aviso = "("

102     Select Case Evento(Hora).Tipo

            Case 1

104             aviso = aviso & "Oro multiplicado por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).Duracion & " minutos"

106         Case 2
        
108             aviso = aviso & "Experiencia multiplicada por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).Duracion & " minutos"

110         Case 3
112             aviso = aviso & "Recoleccion multiplicada por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).Duracion & " minutos"

114         Case 4
116             aviso = aviso & "Dropeo multiplicado por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).Duracion & " minutos"
       
118         Case 5
120             aviso = aviso & "Oro y experiencia multiplicados por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).Duracion & " minutos"

122         Case 6

124             aviso = aviso & "Oro, experiencia y recoleccion multiplicados por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).Duracion & " minutos"

126         Case 7
128             aviso = aviso & "Oro, experiencia, recoleccion y dropeo multiplicados por " & Evento(Hora).multiplicacion & " - Duración: " & Evento(Hora).Duracion & " minutos"

130         Case Else
132             aviso = aviso & "sin información"
        
        End Select

134     aviso = aviso & ")"

136     DescribirEvento = aviso

        
        Exit Function

DescribirEvento_Err:
138     Call TraceError(Err.Number, Err.Description, "ModEventos.DescribirEvento", Erl)

        
End Function

Public Sub CargarEventos()
        
        On Error GoTo CargarEventos_Err
        

        Dim i          As Byte

        Dim EventoStrg As String

100     For i = 0 To 23
102         EventoStrg = GetVar(IniPath & "Configuracion.ini", "EVENTOS", i)
104         Evento(i).Tipo = val(ReadField(1, EventoStrg, Asc("-")))
106         Evento(i).Duracion = val(ReadField(2, EventoStrg, Asc("-")))
108         Evento(i).multiplicacion = val(ReadField(3, EventoStrg, Asc("-")))
110     Next i

112     ExpMultOld = ExpMult
114     OroMultOld = OroMult
116     DropMultOld = DropMult
118     RecoleccionMultOld = RecoleccionMult

        
        Exit Sub

CargarEventos_Err:
120     Call TraceError(Err.Number, Err.Description, "ModEventos.CargarEventos", Erl)

        
End Sub

Public Sub ForzarEvento(ByVal Tipo As Byte, ByVal Duracion As Byte, ByVal multi As Byte, ByVal Quien As String)
        
        On Error GoTo ForzarEvento_Err
        
        Dim tUser As t_UserReference
        tUser = NameIndex(Quien)
        If Not IsValidUserRef(tUser) Then
            Call LogError("Failed to force event, unknown user: " & Quien)
            Exit Sub
        End If
        
100     If Tipo > 3 Or Tipo < 1 Then
102         Call WriteConsoleMsg(tUser.ArrayIndex, "Tipo de evento invalido.", e_FontTypeNames.FONTTYPE_New_Eventos)
            Exit Sub

        End If
 
104     If Duracion > 59 Then
106         Call WriteConsoleMsg(tUser.ArrayIndex, "Duracion invalida, maxima 59 minutos.", e_FontTypeNames.FONTTYPE_New_Eventos)
            Exit Sub

        End If

108      If (Tipo = 1 And multi > 2) Then
110         Call WriteConsoleMsg(tUser.ArrayIndex, "Multiplicacion invalida, maxima x2.", e_FontTypeNames.FONTTYPE_New_Eventos)
            Exit Sub

        End If
        
112     If (Tipo = 2 And multi > 2) Then
114         Call WriteConsoleMsg(tUser.ArrayIndex, "Multiplicacion invalida, maxima x2.", e_FontTypeNames.FONTTYPE_New_Eventos)
            Exit Sub

        End If
        
116     If (Tipo = 3 And multi > 5) Then
118         Call WriteConsoleMsg(tUser.ArrayIndex, "Multiplicacion invalida, maxima x5.", e_FontTypeNames.FONTTYPE_New_Eventos)
            Exit Sub

        End If

        Dim aviso As String

120     aviso = "Eventos> " & Quien & " inicio un nuevo evento: "
122     PublicidadEvento = "Evento en curso>"

124     Select Case Tipo

            Case 1
126             OroMult = OroMult * multi
128             aviso = aviso & " Oro multiplicado por " & multi & " - Duración del evento: " & Duracion & " minutos."
130             PublicidadEvento = PublicidadEvento & " Oro multiplicado por " & multi

132         Case 2
134             ExpMult = ExpMult * multi
136             aviso = aviso & " Experiencia multiplicada por " & multi & " - Duración del evento: " & Duracion & " minutos."
138             PublicidadEvento = PublicidadEvento & " Experiencia multiplicada por " & multi

140         Case 3
142             RecoleccionMult = RecoleccionMult * multi
144             aviso = aviso & " Recoleccion multiplicada por " & multi & " - Duración del evento: " & Duracion & " minutos."
146             PublicidadEvento = PublicidadEvento & " Recoleccion multiplicada por " & multi

148         Case 4
150             DropMult = DropMult / multi
152             aviso = aviso & " Dropeo multiplicado por " & multi & " - Duración del evento: " & Duracion & " minutos."
154             PublicidadEvento = PublicidadEvento & " Dropeo multiplicado por " & multi

156         Case 5
158             ExpMult = ExpMult * multi
160             OroMult = OroMult * multi
162             aviso = aviso & " Oro y experiencia multiplicados por " & multi & " - Duración del evento: " & Duracion & " minutos."
164             PublicidadEvento = PublicidadEvento & " Oro y experiencia multiplicados por " & multi

166         Case 6
168             ExpMult = ExpMult * multi
170             OroMult = OroMult * multi
172             RecoleccionMult = RecoleccionMult * multi
174             aviso = aviso & " Oro, experiencia y recoleccion multiplicados por " & multi & " - Duración del evento: " & Duracion & " minutos."
176             PublicidadEvento = PublicidadEvento & " Oro, experiencia y recoleccion multiplicados por " & multi

178         Case 7
180             ExpMult = ExpMult * multi
182             OroMult = OroMult * multi
184             DropMult = DropMult / multi
186             RecoleccionMult = RecoleccionMult * multi
188             aviso = aviso & " Oro, experiencia, recoleccion y dropeo multiplicados por " & multi & " - Duración del evento: " & Duracion & " minutos."
190             PublicidadEvento = PublicidadEvento & " Oro, experiencia, recoleccion y dropeo multiplicados por " & multi

192         Case Else

194             EventoActivo = False
                Exit Sub
        
        End Select

196     Call AgregarAConsola(aviso)

198     EventoAcutal.Duracion = Duracion
200     EventoAcutal.multiplicacion = multi
202     EventoAcutal.Tipo = Tipo

204     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(aviso, e_FontTypeNames.FONTTYPE_New_Eventos))
206     Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(553, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno
208     TiempoRestanteEvento = Duracion
210     frmMain.Evento.Enabled = True
212     EventoActivo = True

        
        Exit Sub

ForzarEvento_Err:
214     Call TraceError(Err.Number, Err.Description, "ModEventos.ForzarEvento", Erl)

        
End Sub

Public Sub IniciarEvento(ByVal Tipo As TipoEvento, ByVal data As Variant)
100     Select Case Tipo
            Case TipoEvento.Invasion
102             Call IniciarInvasion(data)
        End Select
End Sub
