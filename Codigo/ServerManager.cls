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
    GUARD_NOTICE_RESPONSE
End Enum

Private Enum ManagerPackets
    LOGIN_ACCOUNT_RESULT
    LOGIN_CHAR_RESULT
    CREATE_CHAR_RESULT
    GUARD_NOTICE_RESPONSE_RESULT
End Enum

Private Closing As Boolean

Private WithEvents Socket As Winsock
Attribute Socket.VB_VarHelpID = -1
Private Port As Integer

Private Buffer As cStringBuilder

Public Sub Initialize(pSocket As Winsock, ByVal pPort As Integer)
    Set Socket = pSocket
    Port = pPort
    Set Buffer = New cStringBuilder
    
    Call Connect
End Sub

Private Sub Connect()
    Call Socket.Connect("localhost", Port)
End Sub

Public Sub Send(ByVal Packet As Packets, Data As JS_Object, Optional Instance As JS_Object)
    Dim Msg As New JS_Object

    Msg.Item("id") = Packet
    Msg.Item("data") = Data
    
    If Not Instance Is Nothing Then
        Msg.Item("instance") = Instance
    End If

    Call Socket.SendData(Chr(1) & Msg.ToString() & Chr(0))
End Sub

Private Sub Class_Terminate()
    Closing = True
    Socket.Close
    Set Socket = Nothing
End Sub

Private Sub Socket_Close()
    If Not Closing Then
        Call Connect
    End If
End Sub

Private Sub Socket_Connect()
    Debug.Print "Conectado con el Manager"
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    Dim DataRecv As String
    
    Call Socket.getData(DataRecv, vbString, bytesTotal)
    Call Buffer.Append(DataRecv)

    Dim FirstSeparator As Long
    FirstSeparator = Buffer.Find(Chr(0))

    Dim message As String
    If FirstSeparator > 0 Then
        message = Buffer.SubStr(1, FirstSeparator - 2)

        Dim Data As Object
        Set Data = JSON.parse(message)

          Call HandlePacket(Data.Item("id"), Data.Item("data"), Data.Item("instance"))

        Call Buffer.Remove(0, FirstSeparator)
    End If
End Sub

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Not Closing Then
        Socket.Close
        DoEvents

        Call Connect
    End If
End Sub

Private Sub HandlePacket(ByVal ID As Long, Data As Object, Instance As Object)
    Select Case ID
    
        Case ManagerPackets.LOGIN_ACCOUNT_RESULT
            Call HandleLoginAccount(ID, Data, Instance)

        Case ManagerPackets.GUARD_NOTICE_RESPONSE_RESULT
            Call HandleGuardNotice(ID, Data, Instance)
            
    End Select
End Sub

Private Function CheckPlayer(ByVal Slot As Integer, UUID As String) As Boolean
    ' TODO: Implementar comprobación de que el jugador sigue online
    CheckPlayer = True

    'userlist(slot).

End Function

Private Sub HandleLoginAccount(ByVal ID As Long, Data As Object, Instance As Object)
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
        End With
        
        If Data.Item("verify") Then
            Call WriteGuardNotice(UserIndex)
            Call WriteShowMessageBox(UserIndex, Data.Item("message"))
        Else
            Call WritePersonajesDeCuenta(UserIndex, Data.Item("chars"))
            Call WriteMostrarCuenta(UserIndex)
        End If

    ' Si no pudo loguearse, le mostramos el mensaje de error y cerramos el socket
    Else
        Call WriteShowMessageBox(UserIndex, Data.Item("message"))
        Call CloseSocket(UserIndex)
    End If
End Sub

Private Sub HandleGuardNotice(ByVal ID As Long, Data As Object, Instance As Object)
    Dim UserIndex As Integer
    UserIndex = Instance.Item("slot")

    ' Si no está online el jugador, ignoramos la respuesta
    If Not CheckPlayer(UserIndex, Instance.Item("uuid")) Then Exit Sub

    ' Si el login fue exitoso
    If Data.Item("success") Then

        Call WritePersonajesDeCuenta(UserIndex, Data.Item("chars"))
        Call WriteMostrarCuenta(UserIndex)

    ' Si no pudo loguearse, le mostramos el mensaje de error y cerramos el socket
    Else
        Call WriteShowMessageBox(UserIndex, Data.Item("message"))
        Call CloseSocket(UserIndex)
    End If
End Sub

Private Sub HandleLoginChar(ByVal ID As Long, Data As Object, Instance As Object)
    Dim UserIndex As Integer
    UserIndex = Instance.Item("slot")

    ' Si no está online el jugador, ignoramos la respuesta
    If Not CheckPlayer(UserIndex, Instance.Item("uuid")) Then Exit Sub

    ' Si el login fue exitoso
    If Data.Item("success") Then
        ' Llenar el slot con los datos
        With UserList(UserIndex)

            Call ConnectUser_Complete(UserIndex, .Name, .Email)
        End With
    End If
End Sub
