Attribute VB_Name = "ModLadder"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Declare Sub GetSystemTime Lib "kernel32.dll" (lpSystemTime As t_SYSTEMTIME)

Private theTime      As t_SYSTEMTIME

Public PaquetesCount As Long

Private Type t_SYSTEMTIME

    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer

End Type

Public Enum e_AccionBarra

    Runa = 1
    Resucitar = 2
    Intermundia = 3
    GoToPareja = 5
    Hogar = 6
    CancelarAccion = 99

End Enum

Public Function GetTickCount() As Long
        
        On Error GoTo GetTickCount_Err
    
        
    
100     GetTickCount = timeGetTime And &H7FFFFFFF
    
        
        Exit Function

GetTickCount_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLadder.GetTickCount", Erl)

        
End Function

Function GetTimeFormated() As String
        
        On Error GoTo GetTimeFormated_Err
        
        Dim Elapsed As Long
100     Elapsed = (GetTickCount() - HoraMundo) / DuracionDia
        
        Dim Mins As Long
102     Mins = (Elapsed - Fix(Elapsed)) * 1440

        Dim Horita    As Byte

        Dim Minutitos As Byte

104     Horita = Fix(Mins / 60)
106     Minutitos = Mins Mod 60

108     GetTimeFormated = Right$("00" & Horita, 2) & ":" & Right$("00" & Minutitos, 2)

        
        Exit Function

GetTimeFormated_Err:
110     Call TraceError(Err.Number, Err.Description, "ModLadder.GetTimeFormated - " + Erl, Erl)

        
End Function

Public Sub GetHoraActual()
        
        On Error GoTo GetHoraActual_Err
        
100     GetSystemTime theTime

102     HoraActual = (theTime.wHour - 3)

104     If HoraActual = -3 Then HoraActual = 21
106     If HoraActual = -2 Then HoraActual = 22
108     If HoraActual = -1 Then HoraActual = 23
110     frmMain.lblhora.Caption = HoraActual & ":" & Format(theTime.wMinute, "00") & ":" & Format(theTime.wSecond, "00")
112     HoraEvento = HoraActual

        
        Exit Sub

GetHoraActual_Err:
114     Call TraceError(Err.Number, Err.Description, "ModLadder.GetHoraActual", Erl)

        
End Sub

Public Function DarNameMapa(ByVal Map As Long) As String
        
        On Error GoTo DarNameMapa_Err
        
100     DarNameMapa = MapInfo(Map).map_name

        
        Exit Function

DarNameMapa_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLadder.DarNameMapa", Erl)

        
End Function





Public Function SumarTiempo(segundos As Integer) As String
        
        On Error GoTo SumarTiempo_Err
        

        Dim a As Variant, b As Variant

        Dim X As Integer

        Dim T As String

100     T = "00:00:00" 'Lo inicializamos en 0 horas, 0 minutos, 0 segundos
102     a = Format("00:00:01", "hh:mm:ss") 'guardamos en una variable el formato de 1 segundos
        
        If segundos > 0 Then
104         For X = 1 To segundos 'hacemos segundo a segundo
106             b = Format(T, "hh:mm:ss") 'En B guardamos un formato de hora:minuto:segundo segun lo que tenia T
108             T = Format(TimeValue(a) + TimeValue(b), "hh:mm:ss") 'asignamos a T la suma de A + B (osea, sumamos logicamente 1 segundo)
110         Next X
        End If

112     SumarTiempo = T 'a la funcion le damos el valor que hallamos en T

        
        Exit Function

SumarTiempo_Err:
114     Call TraceError(Err.Number, Err.Description, "ModLadder.SumarTiempo", Erl)

        
End Function

Public Sub AgregarAConsola(ByVal Text As String)
        
        On Error GoTo AgregarAConsola_Err
        

100     frmMain.List1.AddItem (Text)

        
        Exit Sub

AgregarAConsola_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLadder.AgregarAConsola", Erl)

        
End Sub

' TODO: Crear enum para la respuesta
Function PuedeUsarObjeto(UserIndex As Integer, ByVal ObjIndex As Integer, Optional ByVal writeInConsole As Boolean = False) As Byte
        On Error GoTo PuedeUsarObjeto_Err

        Dim Objeto As t_ObjData
        Dim Msg As String, i As Long
100     Objeto = ObjData(ObjIndex)
                
102     If EsGM(UserIndex) Then
104         PuedeUsarObjeto = 0
106         Msg = ""

108     ElseIf Objeto.Newbie = 1 And Not EsNewbie(UserIndex) Then
110         PuedeUsarObjeto = 7
112         Msg = "Solo los newbies pueden usar este objeto."
            
114     ElseIf UserList(UserIndex).Stats.ELV < Objeto.MinELV Then
116         PuedeUsarObjeto = 6
118         Msg = "Necesitas ser nivel " & Objeto.MinELV & " para usar este objeto."

120     ElseIf Not FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
122         PuedeUsarObjeto = 3
124         Msg = "Tu facción no te permite utilizarlo."

126     ElseIf Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
128         PuedeUsarObjeto = 2
130         Msg = "Tu clase no puede utilizar este objeto."

132     ElseIf Not SexoPuedeUsarItem(UserIndex, ObjIndex) Then
134         PuedeUsarObjeto = 1
136         Msg = "Tu sexo no puede utilizar este objeto."

138     ElseIf Not RazaPuedeUsarItem(UserIndex, ObjIndex) Then
140         PuedeUsarObjeto = 5
142         Msg = "Tu raza no puede utilizar este objeto."
144     ElseIf (Objeto.SkillIndex > 0) Then
146         If (UserList(UserIndex).Stats.UserSkills(Objeto.SkillIndex) < Objeto.SkillRequerido) Then
148             PuedeUsarObjeto = 4
150             Msg = "Necesitas " & Objeto.SkillRequerido & " puntos en " & SkillsNames(Objeto.SkillIndex) & " para usar este item."
            Else
152             PuedeUsarObjeto = 0
154             Msg = ""
            End If
        Else
156         PuedeUsarObjeto = 0
158         Msg = ""
        End If

160     If writeInConsole And Msg <> "" Then Call WriteConsoleMsg(UserIndex, Msg, e_FontTypeNames.FONTTYPE_INFO)

        Exit Function

PuedeUsarObjeto_Err:
162     Call TraceError(Err.Number, Err.Description, "ModLadder.PuedeUsarObjeto", Erl)
        'Resume Next ' WyroX: Si hay error que salga directamente

End Function

Public Function RequiereOxigeno(ByVal UserMap) As Boolean
        On Error GoTo RequiereOxigeno_Err
        
100     RequiereOxigeno = (UserMap = 265) Or _
                          (UserMap = 266) Or _
                          (UserMap = 267)
        
        Exit Function

RequiereOxigeno_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLadder.RequiereOxigeno", Erl)

        
End Function
