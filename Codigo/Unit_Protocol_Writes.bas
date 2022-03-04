Attribute VB_Name = "Unit_Protocol_Writes"

Option Explicit
Private Writer As Network.Writer

Public Function writer_is_nothing() As Boolean
writer_is_nothing = Writer Is Nothing
End Function
Public Sub Initialize()
100     Set Writer = New Network.Writer
End Sub

Public Sub Clear()
100     Call Writer.Clear
End Sub

#If 0 Then

''
' Writes the "LoginExistingChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLoginExistingChar()
        '</EhHeader>
100     Call Writer.WriteInt(ClientPacketID.LoginExistingChar)
102     Call Writer.WriteString8(encrypted_session_token)


        Dim encrypted_username_b64 As String
        encrypted_username_b64 = AO20CryptoSysWrapper.ENCRYPT(cnvHexStrFromBytes(public_key), UserName)
        
104     Call Writer.WriteString8(encrypted_username_b64)
106     Call Writer.WriteInt8(App.Major)
108     Call Writer.WriteInt8(App.Minor)
110     Call Writer.WriteInt8(App.Revision)
     Call Writer.WriteString8(CheckMD5)
            
     Call modNetwork.Send(Writer)

End Sub

''
' Writes the "LoginNewChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLoginNewChar()
        
    Dim encrypted_username_b64 As String
        encrypted_username_b64 = AO20CryptoSysWrapper.ENCRYPT(cnvHexStrFromBytes(public_key), UserName)
        
     Call Writer.WriteInt(ClientPacketID.LoginNewChar)
     Call Writer.WriteString8(encrypted_session_token)
     Call Writer.WriteString8(encrypted_username_b64)
     Call Writer.WriteInt8(App.Major)
     Call Writer.WriteInt8(App.Minor)
     Call Writer.WriteInt8(App.Revision)
     Call Writer.WriteString8(CheckMD5)
     Call Writer.WriteInt8(UserRaza)
     Call Writer.WriteInt8(UserSexo)
     Call Writer.WriteInt8(UserClase)
     Call Writer.WriteInt16(MiCabeza)
     Call Writer.WriteInt8(UserHogar)
    
     Call modNetwork.Send(Writer)

End Sub
#End If

