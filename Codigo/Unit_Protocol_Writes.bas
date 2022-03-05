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


Public Sub WriteLoginNewChar(ByVal public_key As String, ByVal username As String, _
    ByVal app_major As Byte, ByVal app_minor As Byte, ByVal app_revision As Byte, ByVal md5 As String, _
    ByVal race As Byte, ByVal gender As Byte, ByVal class As Byte, ByVal body As Byte, _
    ByVal head As Byte, ByVal home As Byte)
     
     Dim encrypted_username_b64 As String
     encrypted_username_b64 = AO20CryptoSysWrapper.ENCRYPT(cnvHexStrFromString(public_key), username)
     Call Writer.WriteInt(ClientPacketID.LoginNewChar)
     Call Writer.WriteString8(UnitTesting.encrypted_token)
     Call Writer.WriteString8(encrypted_username_b64)
     Call Writer.WriteInt8(App.Major)
     Call Writer.WriteInt8(App.Minor)
     Call Writer.WriteInt8(App.Revision)
     Call Writer.WriteString8(md5)
     Call Writer.WriteInt8(race)
     Call Writer.WriteInt8(gender)
     Call Writer.WriteInt8(class)
     Call Writer.WriteInt16(head)
     Call Writer.WriteInt8(home)
     Call UnitClient.Send(Writer)

End Sub
#End If

