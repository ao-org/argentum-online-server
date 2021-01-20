Attribute VB_Name = "ModViajar"

Public Sub IniciarTransporte(ByVal UserIndex As Integer)
        
        On Error GoTo IniciarTransporte_Err
        

        Dim destinos As Byte

100     destinos = Npclist(UserList(UserIndex).flags.TargetNPC).NumDestinos

        
        Exit Sub

IniciarTransporte_Err:
102     Call RegistrarError(Err.Number, Err.Description, "ModViajar.IniciarTransporte", Erl)
104     Resume Next
        
End Sub
