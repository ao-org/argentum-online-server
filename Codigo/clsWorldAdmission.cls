VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsWorldAdmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ClaimId As String
Public ConnectionId As Long
Public UserIndex As Integer
Public UserVersion As Integer
Public PacketId As Long
Public Message As Network.reader
Public EncryptedToken As String
Public DecryptedToken As String
Public Stage As Long
Public Retired As Boolean
Public SaveFailed As Boolean
