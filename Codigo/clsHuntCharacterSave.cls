VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
  Persistable = 0
  DataBindingBehavior = 0
  DataSourceBehavior = 0
  MTSTransactionMode = 0
END
Attribute VB_Name = "clsHuntCharacterSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Argentum 20 Game Server. Copyright (C) 2026 Noland Studios LTD.
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit
Public CharacterId As Long
Public Sequence As Long
Public UserIndex As Integer
Public UserVersion As Integer
Public Commands As Collection
Private mPersist As t_UserPersistSnapshot
Private mCardQuantities(1 To 1024) As Byte
Private mCardsDirty As Boolean
Friend Sub Capture(ByVal source As Integer, ByVal revision As Long, ByVal logout As Boolean)
    Dim snapshot As t_User, i As Integer
    snapshot = UserList(source)
    CharacterId = snapshot.Id: Sequence = revision
    UserIndex = source: UserVersion = snapshot.VersionId
    Set Commands = CapturePartyQuestCharacterCommands(snapshot, logout)
    mPersist = snapshot.Persist
    mCardsDirty = snapshot.flags.DirtyCollectibleCardCollection
    For i = 1 To 1024: mCardQuantities(i) = snapshot.AccountCollectibleCardQuantities(i): Next i
End Sub
Friend Sub Acknowledge()
    If UserIndex < 1 Or UserIndex > UBound(UserList) Then Exit Sub
    If UserList(UserIndex).Id <> CharacterId Or UserList(UserIndex).VersionId <> UserVersion Then Exit Sub
    Call ApplyPartyQuestCharacterSave(UserIndex, Sequence, mPersist)
    If mCardsDirty Then
        Dim i As Integer, same As Boolean
        same = True
        For i = 1 To 1024
            If UserList(UserIndex).AccountCollectibleCardQuantities(i) <> mCardQuantities(i) Then same = False
        Next i
        If same Then UserList(UserIndex).flags.DirtyCollectibleCardCollection = False
    End If
End Sub
