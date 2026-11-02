Attribute VB_Name = "ModGrupos"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Public Grupo            As Tgrupo
Private UniqueIdCounter As Long

Private Function GetNextId() As Long
    UniqueIdCounter = (UniqueIdCounter + 1) And &H7FFFFFFF
    GetNextId = UniqueIdCounter
End Function

Public Sub InvitarMiembro(ByVal UserIndex As Integer, ByVal InvitadoIndex As Integer)
    Dim result As e_PartyResult
    result = InviteManagedParty(UserIndex, InvitadoIndex)
    If result = PartySuccess Then result = PartyInvitationSent
    Call WritePartyManagementState(UserIndex, 0, result)
End Sub
Public Sub EcharMiembro(ByVal UserIndex As Integer, ByVal Indice As Byte)
    Dim leader As Integer
    leader = PartyLeaderIndex(UserIndex)
    If leader <> UserIndex Or leader = 0 Then Exit Sub
    If Indice >= UserList(leader).Grupo.CantidadMiembros Then Exit Sub
    If Not IsValidUserRef(UserList(leader).Grupo.Miembros(Indice + 1)) Then Exit Sub
    Call RemoveManagedPartyMember(UserList(leader).Grupo.Miembros(Indice + 1).ArrayIndex)
End Sub

Public Sub SalirDeGrupo(ByVal UserIndex As Integer)
    Call RemoveManagedPartyMember(UserIndex)
End Sub

Public Sub SalirDeGrupoForzado(ByVal UserIndex As Integer)
    Call RemoveManagedPartyMember(UserIndex)
End Sub
Public Sub FinalizarGrupo(ByVal LiderIndex As Integer)
    Call DisbandManagedParty(LiderIndex)
End Sub

Public Sub CompartirUbicacion(ByVal UserIndex As Integer)
    On Error GoTo CompartirUbicacion_Err
    Dim i       As Byte
    Dim a       As Byte
    Dim indexpj As Byte
    Dim Lider   As t_User
    With UserList(UserIndex)
        Lider = UserList(.Grupo.Lider.ArrayIndex)
        For a = 1 To Lider.Grupo.CantidadMiembros
            If Lider.Grupo.Miembros(a).ArrayIndex = UserIndex Then
                indexpj = a
            End If
        Next a
        For i = 1 To Lider.Grupo.CantidadMiembros
            If Lider.Grupo.Miembros(i).ArrayIndex <> UserIndex Then
                If UserList(Lider.Grupo.Miembros(i).ArrayIndex).pos.Map = .pos.Map Then
                    Call WriteUbicacion(Lider.Grupo.Miembros(i).ArrayIndex, indexpj, UserIndex)
                    'Si va al mapa del compañero
                    Call WriteUbicacion(UserIndex, i, Lider.Grupo.Miembros(i).ArrayIndex)
                Else
                    ' Le borro la ubicacion a ellos
                    Call WriteUbicacion(Lider.Grupo.Miembros(i).ArrayIndex, indexpj, 0)
                    ' Le borro la ubicacion a mi
                    Call WriteUbicacion(UserIndex, i, 0)
                End If
            End If
        Next i
    End With
    Exit Sub
CompartirUbicacion_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.CompartirUbicacion", Erl)
End Sub

Public Sub GroupCreateSuccess(ByVal LiderIndex As Integer)
    On Error GoTo GroupCreateSuccess_Err
    If Not IsRewardUser(LiderIndex) Or EsGM(LiderIndex) Then Exit Sub
    If UserList(LiderIndex).Grupo.EnGrupo Then Exit Sub
    Call ClearPartyXpRotation(LiderIndex)
    With UserList(LiderIndex).Grupo
        .EnGrupo = True: .Id = GetNextId(): .RosterRevision = 1
        .CantidadMiembros = 1
        Call SetUserRef(.Lider, LiderIndex)
        Dim i As Integer
        For i = 1 To UBound(.Miembros)
            Call ClearUserRef(.Miembros(i))
        Next i
        Call SetUserRef(.Miembros(1), LiderIndex)
    End With
    Call PublishManagedParty(LiderIndex)
    Exit Sub
GroupCreateSuccess_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.GroupCreateSuccess", Erl)
End Sub

Public Sub AddUserToGRoup(ByVal UserIndex As Integer, ByVal GroupLiderIndex As Integer)
    If UserList(UserIndex).Grupo.PropuestaDe.ArrayIndex <> GroupLiderIndex Then Exit Sub
    Call AcceptManagedParty(UserIndex)
End Sub

' A snapshot, not the mutable roster, is used for both reward counting and payment.
' The leader is already in Miembros (normally slot 1). The credited player comes
' first for quest processing; XP separately sorts the same deduped snapshot.
Public Function IsRewardCredit(ByVal userIndex As Integer) As Boolean
    On Error GoTo IsRewardCredit_Err
    If userIndex < 1 Or userIndex > UBound(UserList) Then Exit Function
    With UserList(userIndex)
        IsRewardCredit = .flags.UserLogged And .ConnectionDetails.ConnIDValida
    End With
    Exit Function
IsRewardCredit_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.IsRewardCredit", Erl)
End Function

Public Function IsRewardUser(ByVal userIndex As Integer) As Boolean
    If Not IsRewardCredit(userIndex) Then Exit Function
    IsRewardUser = UserList(userIndex).flags.Muerto = 0
End Function

Public Sub CollectPartyRewardRecipients(ByVal credited As Integer, ByRef result As t_PartyRewardRecipients, Optional ByVal npcMap As Integer = 0)
    On Error GoTo CollectPartyRewardRecipients_Err
    Dim emptyResult As t_PartyRewardRecipients
    result = emptyResult
    If Not IsRewardCredit(credited) Then Exit Sub
    If npcMap <> 0 Then
        If UserList(credited).pos.Map <> npcMap Then Exit Sub
    End If
    If IsRewardUser(credited) Then
        result.Count = 1
        Call SetUserRef(result.Users(1), credited)
    End If
    Dim leader As Integer, groupId As Long, i As Integer, j As Integer, member As Integer
    Dim foundCredited As Boolean, duplicate As Boolean, rosterCount As Integer
    With UserList(credited).Grupo
        If Not .EnGrupo Or .Id <= 0 Then Exit Sub
        If Not IsValidUserRef(.Lider) Then Exit Sub
        leader = .Lider.ArrayIndex: groupId = .Id
    End With
    With UserList(leader).Grupo
        If Not .EnGrupo Or .Id <> groupId Then Exit Sub
        If Not IsValidUserRef(.Lider) Then Exit Sub
        If .Lider.ArrayIndex <> leader Then Exit Sub
        rosterCount = .CantidadMiembros
        If rosterCount < 1 Or rosterCount > UBound(.Miembros) Then Exit Sub
        For i = 1 To rosterCount
            If IsValidUserRef(.Miembros(i)) Then
                If .Miembros(i).ArrayIndex = credited Then foundCredited = True
            End If
        Next i
        If Not foundCredited Then Exit Sub
        result.IsParty = rosterCount > 1
        result.PartyId = groupId
        Call SetUserRef(result.Leader, leader)
        For i = 1 To rosterCount
            If IsValidUserRef(.Miembros(i)) Then
                member = .Miembros(i).ArrayIndex
                If IsRewardUser(member) Then
                    If UserList(member).Grupo.EnGrupo And UserList(member).Grupo.Id = groupId Then
                        If IsValidUserRef(UserList(member).Grupo.Lider) Then
                            If UserList(member).Grupo.Lider.ArrayIndex = leader Then
                                If UserList(member).pos.Map = UserList(credited).pos.Map Then
                                    If Distancia(UserList(credited).pos, UserList(member).pos) < 20 Then
                                        duplicate = False
                                        For j = 1 To result.Count
                                            If result.Users(j).ArrayIndex = member Then duplicate = True
                                        Next j
                                        If Not duplicate Then
                                            result.Count = result.Count + 1
                                            result.Users(result.Count) = .Miembros(i)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next i
    End With
    Exit Sub
CollectPartyRewardRecipients_Err:
    ' Never use a partially validated party projection. Keep the valid solo credit.
    result = emptyResult
    If IsRewardUser(credited) Then
        If npcMap = 0 Or UserList(credited).pos.Map = npcMap Then
            result.Count = 1: Call SetUserRef(result.Users(1), credited)
        End If
    End If
    Call TraceError(Err.Number, Err.Description, "ModGrupos.CollectPartyRewardRecipients", Erl)
End Sub

Public Sub ClearPartyXpRotation(ByVal userIndex As Integer)
    On Error GoTo ClearPartyXpRotation_Err
    If userIndex < 1 Or userIndex > UBound(UserList) Then Exit Sub
    Dim emptyRotation As t_PartyXpRotation
    UserList(userIndex).Grupo.XpRotation = emptyRotation
    Exit Sub
ClearPartyXpRotation_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.ClearPartyXpRotation", Erl)
End Sub

' Compute and commit one bounded rotation before the caller debits the NPC pool.
' No entitlement/XP is retained: only the next versioned identity on the ring.
Public Function AllocatePartyBaseExperience(ByRef recipients As t_PartyRewardRecipients, ByVal baseExp As Long, ByRef shares() As Long) As Boolean
    On Error GoTo AllocatePartyBaseExperience_Err
    If Not recipients.IsParty Or recipients.Count < 1 Or recipients.Count > UBound(recipients.Users) Or baseExp <= 0 Then Exit Function
    If Not IsValidUserRef(recipients.Leader) Then Exit Function
    Dim leader As Integer, i As Integer, j As Integer, entry As Integer
    Dim order(1 To 6) As Integer, quotient As Long, remainder As Integer
    Dim start As Integer, ordinal As Integer, current As t_UserReference
    leader = recipients.Leader.ArrayIndex
    If Not UserList(leader).Grupo.EnGrupo Or UserList(leader).Grupo.Id <> recipients.PartyId Then Exit Function
    If Not IsValidUserRef(UserList(leader).Grupo.Lider) Then Exit Function
    If UserList(leader).Grupo.Lider.ArrayIndex <> leader Then Exit Function
    quotient = baseExp \ recipients.Count: remainder = baseExp Mod recipients.Count
    ' Insertion sort of at most six snapshot indices; never mutate the roster or
    ' the quest recipient order. Slot order is stable during a character session.
    For i = 1 To recipients.Count
        shares(i) = quotient
        entry = i: j = i
        Do While j > 1
            If recipients.Users(order(j - 1)).ArrayIndex < recipients.Users(entry).ArrayIndex Then Exit Do
            order(j) = order(j - 1): j = j - 1
        Loop
        order(j) = entry
    Next i
    With UserList(leader).Grupo.XpRotation
        If .PartyId <> recipients.PartyId Or .Leader.ArrayIndex <> leader Or .Leader.VersionId <> recipients.Leader.VersionId Then
            Call ClearPartyXpRotation(leader)
            .PartyId = recipients.PartyId: .Leader = recipients.Leader
        End If
        If remainder > 0 Then
            start = 1
            If .NextRecipient.ArrayIndex > 0 Then
                ' Resume at the first eligible successor. A slot reused by a new
                ' version cannot inherit the old session's reserved next turn.
                For i = 1 To recipients.Count
                    current = recipients.Users(order(i))
                    If current.ArrayIndex > .NextRecipient.ArrayIndex Or (current.ArrayIndex = .NextRecipient.ArrayIndex And current.VersionId = .NextRecipient.VersionId) Then
                        start = i: Exit For
                    End If
                Next i
            End If
            For i = 0 To remainder - 1
                ordinal = ((start - 1 + i) Mod recipients.Count) + 1
                shares(order(ordinal)) = shares(order(ordinal)) + 1
            Next i
            ordinal = ((start - 1 + remainder) Mod recipients.Count) + 1
            .NextRecipient = recipients.Users(order(ordinal))
        End If
    End With
    AllocatePartyBaseExperience = True
    Exit Function
AllocatePartyBaseExperience_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.AllocatePartyBaseExperience", Erl)
End Function
