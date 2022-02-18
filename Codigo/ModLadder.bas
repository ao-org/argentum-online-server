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


Public Enum e_AccionBarra
    Runa = 1
    Resucitar = 2
    Intermundia = 3
    GoToPareja = 5
    Hogar = 6
    CancelarAccion = 99
End Enum



Public Function DarNameMapa(ByVal Map As Long) As String
        
        On Error GoTo DarNameMapa_Err
        
100     DarNameMapa = MapInfo(Map).map_name

        
        Exit Function

DarNameMapa_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLadder.DarNameMapa", Erl)

        
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
