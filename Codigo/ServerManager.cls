VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ServerManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum Packets
    LOGIN_ACCOUNT
    LOGIN_CHAR
    CREATE_CHAR
    SAVE_CHAR
End Enum

Private Enum ManagerPackets
    LOGIN_ACCOUNT_RESULT
    LOGIN_CHAR_RESULT
    CREATE_CHAR_RESULT
End Enum

Private WithEvents Socket As Winsock
Attribute Socket.VB_VarHelpID = -1

Public Sub Initialize(pSocket As Winsock, ByVal Port As Integer)
    Set Socket = pSocket
    
    Call Socket.Connect("localhost", Port)
End Sub

Public Sub Send(ByVal Packet As Packets, Data As String, Optional Instance As String)
    Dim message As String
    message = "{""id"":" & Packet & ",""data"":" & Data & IIf(LenB(Instance), ",""instance"":" & Instance, "") & "}" & Chr(0)
    
    Call Socket.SendData(message)
End Sub

Private Sub Class_Terminate()
    Set Socket = Nothing
End Sub

Private Sub Socket_Connect()
    Debug.Print "Conectado con el Manager"
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    ' TODO: Mejorar para evitar problemas de paquetes dobles o incompletos
    Dim message As String
    Call Socket.getData(message, vbString, bytesTotal)
    Debug.Print message
    
    Dim Data As Object
    Set Data = JSON.parse(message)
    
    Call HandlePacket(Data.Item("id"), Data.Item("data"), Data.Item("instance"))
End Sub

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Debug.Print Number, Description
End Sub

Private Sub HandlePacket(ByVal ID As Long, Data As Object, Instance As Object)
    Select Case ID
        Case ManagerPackets.LOGIN_ACCOUNT_RESULT
            Call HandleLoginAccount(ID, Data, Instance)
    End Select
End Sub

Private Function CheckPlayer(ByVal Slot As Integer, UUID As String) As Boolean
    ' TODO: Implementar comprobación de que el jugador sigue online
End Function

Private Sub HandleLoginAccount(ByVal ID As Long, Data As Object, Instance As Object)
    Debug.Print "HandleLoginAccount(" & ID & ", " & JSON.toString(Data) & ", " & JSON.toString(Instance) & ")" ' Borrame

    Dim UserIndex As Integer
    UserIndex = Instance.Item("slot")

    ' Si no está online el jugador, ignoramos la respuesta
    If Not CheckPlayer(UserIndex, Instance.Item("uuid")) Then Exit Sub

    ' Si el login fue exitoso
    If Data.Item("success") Then
    
        ' Seteamos variables de la cuenta
        With UserList(UserIndex)
            .AccountID = Data.Item("account_id")
            .Cuenta = .Email
        
            ' Verificación AoGuard
            Dim Verificar As Boolean
            
            Select Case AOGuard.AOG_STATUS
            
                ' Nunca enviara correo
                Case 0
                    Verificar = False
                
                ' Enviara correo solo si la IP no coinciden con lo que tenemos en la BD
                Case 1
                    Verificar = Not AOGuard.VerificarOrigen(.AccountID, .IP)
                
                ' Enviara correo cada vez que iniciemos sesion
                Case 2
                    Verificar = True
                
                Case Else
                    Verificar = True
            
            End Select
        End With
        
        If Verificar Then
            Call WriteGuardNotice(UserIndex)
            
            Call AOGuard.EnviarCodigo(UserIndex)
        Else
            ' TODO: Refactor: pedir directamente los personajes y hacer el chequeo de AoGuard directo en el Manager, para que con un sólo paquete se resuelva el login de la cuenta
            Call WritePersonajesDeCuenta(UserIndex)
            Call WriteMostrarCuenta(UserIndex)
        End If
    
    ' Si no pudo loguearse, le mostramos el mensaje de error y cerramos el socket
    Else
        Call WriteShowMessageBox(UserIndex, Data.Item("message"))
        Call CloseSocket(UserIndex)
    End If
End Sub
