Attribute VB_Name = "ModViajar"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Public Sub IniciarTransporte(ByVal UserIndex As Integer)
        On Error GoTo IniciarTransporte_Err
        Dim destinos As Byte
100     destinos = NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).NumDestinos
        Exit Sub
IniciarTransporte_Err:
102     Call TraceError(Err.Number, Err.Description, "ModViajar.IniciarTransporte", Erl)
End Sub
