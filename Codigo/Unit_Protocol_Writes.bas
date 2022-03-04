Attribute VB_Name = "Unit_Protocol_Writes"
Option Explicit
Private Writer As Network.Writer

Public Function writer_is_nothing() As Boolean
    writer_is_nothing = Writer Is Nothing
End Function
Public Sub Initialize()
    Set Writer = New Network.Writer
End Sub

Public Sub Clear()
    Call Writer.Clear
End Sub

#If 1 Then


'Public Sub WriteLoginExistingChar()
'     Call Writer.WriteInt(ClientPacketID.LoginExistingChar)
'     Call Writer.WriteString8(encrypted_session_token)
'
'
'        Dim encrypted_username_b64 As String
'        encrypted_username_b64 = AO20CryptoSysWrapper.ENCRYPT(cnvHexStrFromBytes(public_key), username)
'
'     Call Writer.WriteString8(encrypted_username_b64)
'     Call Writer.WriteInt8(App.Major)
'     Call Writer.WriteInt8(App.Minor)
'     Call Writer.WriteInt8(App.Revision)
'     Call Writer.WriteString8(CheckMD5)
'
'     Call modNetwork.Send(Writer)
'
'End Sub

''
' Writes the "LoginNewChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLoginNewChar(ByVal public_key As String, ByVal username As String)
        
    Dim encrypted_username_b64 As String
    
    encrypted_username_b64 = AO20CryptoSysWrapper.ENCRYPT(cnvHexStrFromString(public_key), username)
             
     Call Writer.WriteInt(ClientPacketID.LoginNewChar)
     Call Writer.WriteString8("encrypted_session_token")
     Call Writer.WriteString8(encrypted_username_b64)
     Call Writer.WriteInt8(App.Major)
     Call Writer.WriteInt8(App.Minor)
     Call Writer.WriteInt8(App.Revision)
     Call Writer.WriteString8("a944087c826163c4ed658b1ea00594be")
     Call Writer.WriteInt8(0)
     Call Writer.WriteInt8(0)
     Call Writer.WriteInt8(0)
     Call Writer.WriteInt16(0)
     Call Writer.WriteInt8(0)
    
     Call UnitClient.Send(Writer)

End Sub
#End If

