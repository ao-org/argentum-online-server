Attribute VB_Name = "ModViajar"

Public Sub IniciarTransporte(ByVal UserIndex As Integer)
        
        On Error GoTo IniciarTransporte_Err
        

        Dim destinos As Byte

100     destinos = Npclist(UserList(UserIndex).flags.TargetNPC).NumDestinos

        
        Exit Sub

IniciarTransporte_Err:
        Call RegistrarError(Err.Number, Err.description, "ModViajar.IniciarTransporte", Erl)
        Resume Next
        
End Sub
