Attribute VB_Name = "modHooPriestForgiveness"
Option Explicit

Private Const OFFER_LIFETIME_MS As Long = 30000
Private mOfferActive() As Boolean
Private mOfferUserVersion() As Integer
Private mOfferRequestId() As Long
Private mOfferToken() As Long
Private mOfferPriestIndex() As Integer
Private mOfferPriestVersion() As Integer
Private mOfferPrice() As Long
Private mOfferCreatedAt() As Long
Private mOfferCapacity As Integer
Private mNextToken As Long

Private Function EnsureOfferStorage() As Boolean
    If MaxUsers <= 0 Then Exit Function
    If mOfferCapacity = MaxUsers Then
        EnsureOfferStorage = True
        Exit Function
    End If
    If mOfferCapacity = 0 Then
        ReDim mOfferActive(1 To MaxUsers)
        ReDim mOfferUserVersion(1 To MaxUsers)
        ReDim mOfferRequestId(1 To MaxUsers)
        ReDim mOfferToken(1 To MaxUsers)
        ReDim mOfferPriestIndex(1 To MaxUsers)
        ReDim mOfferPriestVersion(1 To MaxUsers)
        ReDim mOfferPrice(1 To MaxUsers)
        ReDim mOfferCreatedAt(1 To MaxUsers)
    Else
        ReDim Preserve mOfferActive(1 To MaxUsers)
        ReDim Preserve mOfferUserVersion(1 To MaxUsers)
        ReDim Preserve mOfferRequestId(1 To MaxUsers)
        ReDim Preserve mOfferToken(1 To MaxUsers)
        ReDim Preserve mOfferPriestIndex(1 To MaxUsers)
        ReDim Preserve mOfferPriestVersion(1 To MaxUsers)
        ReDim Preserve mOfferPrice(1 To MaxUsers)
        ReDim Preserve mOfferCreatedAt(1 To MaxUsers)
    End If
    mOfferCapacity = MaxUsers
    EnsureOfferStorage = True
End Function

Public Function CalculateHooForgivenessPrice(ByVal CitizensKilled As Long, ByRef Price As Long) As Boolean
    CalculateHooForgivenessPrice = CalculateHooForgivenessPriceForValues(CitizensKilled, _
        CDbl(SvrConfig.GetValue("GoldMult")), _
        CDbl(SvrConfig.GetValue("CostoPerdonPorCiudadano")), Price)
End Function

Public Function CalculateHooForgivenessPriceForValues(ByVal CitizensKilled As Long, _
        ByVal GoldMultiplier As Double, ByVal BaseCost As Double, ByRef Price As Long) As Boolean
    On Error GoTo InvalidPrice
    Dim amount As Double
    If CitizensKilled < 0 Then GoTo InvalidPrice
    amount = BaseCost
    If CitizensKilled > 0 Then
        amount = CDbl(CitizensKilled) * GoldMultiplier * amount
    Else
        amount = amount / 2
    End If
    If amount <= 0 Or amount > 2147483647# Then GoTo InvalidPrice
    Price = CLng(amount)
    If Price <= 0 Then GoTo InvalidPrice
    CalculateHooForgivenessPriceForValues = True
    Exit Function
InvalidPrice:
    Price = 0
    CalculateHooForgivenessPriceForValues = False
End Function

Public Function GetHooForgivenessEligibility(ByVal UserIndex As Integer, _
        ByRef PriestRef As t_NpcReference, ByRef Price As Long) As e_HooForgivenessResult
    On Error GoTo InvalidPrice
    GetHooForgivenessEligibility = eHooForgivenessInvalidPriest
    Price = 0
    If UserIndex < 1 Or UserIndex > UBound(UserList) Then Exit Function
    If Not UserList(UserIndex).InUse Or Not UserList(UserIndex).flags.UserLogged Then Exit Function
    If Not IsValidNpcRef(PriestRef) Then Exit Function
    Dim priest As t_Npc
    priest = NpcList(PriestRef.ArrayIndex)
    If priest.npcType <> e_NPCType.Revividor And _
            (priest.npcType <> e_NPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex)) Then Exit Function
    If UserList(UserIndex).flags.Muerto = 1 Then
        GetHooForgivenessEligibility = eHooForgivenessDead
        Exit Function
    End If
    If MapData(UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, UserList(UserIndex).pos.Map).trigger = 6 Then
        GetHooForgivenessEligibility = eHooForgivenessRestrictedLocation
        Exit Function
    End If
    If UserList(UserIndex).pos.Map <> priest.pos.Map Or Distancia(UserList(UserIndex).pos, priest.pos) > 3 Then
        GetHooForgivenessEligibility = eHooForgivenessTooFar
        Exit Function
    End If
    Select Case UserList(UserIndex).Faccion.Status
        Case e_Facciones.Ciudadano, e_Facciones.Armada, e_Facciones.consejo
            GetHooForgivenessEligibility = eHooForgivenessAlreadyCitizen
            Exit Function
        Case e_Facciones.Caos, e_Facciones.concilio
            GetHooForgivenessEligibility = eHooForgivenessFaction
            Exit Function
        Case e_Facciones.Criminal
            ' Criminal characters are the only faction eligible for forgiveness.
        Case Else
            GetHooForgivenessEligibility = eHooForgivenessFaction
            Exit Function
    End Select
    If UserList(UserIndex).GuildIndex > 0 Then
        Dim alignment As e_ALINEACION_GUILD
        alignment = modGuilds.Alineacion(UserList(UserIndex).GuildIndex)
        If alignment = e_ALINEACION_GUILD.ALINEACION_CRIMINAL Or _
                alignment = e_ALINEACION_GUILD.ALINEACION_CAOTICA Then
            GetHooForgivenessEligibility = eHooForgivenessGuild
            Exit Function
        End If
    End If
    If Not CalculateHooForgivenessPrice(UserList(UserIndex).Faccion.ciudadanosMatados, Price) Then GoTo InvalidPrice
    GetHooForgivenessEligibility = eHooForgivenessOffer
    Exit Function
InvalidPrice:
    Price = 0
    GetHooForgivenessEligibility = eHooForgivenessInvalidPrice
End Function

Public Sub ClearHooForgivenessOffer(ByVal UserIndex As Integer)
    If Not EnsureOfferStorage() Then Exit Sub
    If UserIndex < 1 Or UserIndex > mOfferCapacity Then Exit Sub
    mOfferActive(UserIndex) = False
    mOfferToken(UserIndex) = 0
    mOfferPriestIndex(UserIndex) = 0
    mOfferPriestVersion(UserIndex) = 0
End Sub

Public Function HooForgivenessOfferExpired(ByVal CreatedAt As Long, ByVal NowRaw As Long) As Boolean
    HooForgivenessOfferExpired = TicksElapsed(CreatedAt, NowRaw) > OFFER_LIFETIME_MS
End Function

Private Function NewOfferToken() As Long
    If mNextToken >= 2147483646 Then
        mNextToken = 0
        Dim i As Integer
        For i = LBound(mOfferActive) To mOfferCapacity
            mOfferActive(i) = False
        Next i
    End If
    mNextToken = mNextToken + 1
    NewOfferToken = mNextToken
End Function

Public Sub HandleHooPriestForgiveness(ByVal UserIndex As Integer)
    On Error GoTo HandleError
    Dim RequestId As Long, action As Byte, characterIndex As Integer, token As Long
    Dim paymentDeducted As Boolean
    RequestId = reader.ReadInt32
    action = reader.ReadInt8
    characterIndex = reader.ReadInt16
    token = reader.ReadInt32
    If RequestId <= 0 Or action > 1 Or characterIndex <= 0 Then Exit Sub
    If (action = 0 And token <> 0) Or (action = 1 And token <= 0) Then Exit Sub
    If Not UserSupportsPriestForgiveness(UserIndex) Then Exit Sub
    If Not EnsureOfferStorage() Then Exit Sub
    Dim targetRef As t_NpcReference
    targetRef = UserList(UserIndex).flags.TargetNPC
    If Not IsValidNpcRef(targetRef) Then
        Call WriteHooPriestForgivenessResult(UserIndex, RequestId, eHooForgivenessInvalidPriest, 0, 0)
        Exit Sub
    End If
    If NpcList(targetRef.ArrayIndex).Char.charindex <> characterIndex Then
        Call WriteHooPriestForgivenessResult(UserIndex, RequestId, eHooForgivenessInvalidPriest, 0, 0)
        Exit Sub
    End If
    If action = 0 Then
        Call ClearHooForgivenessOffer(UserIndex)
        Dim price As Long, eligibility As e_HooForgivenessResult
        eligibility = GetHooForgivenessEligibility(UserIndex, targetRef, price)
        If eligibility <> eHooForgivenessOffer Then
            Call WriteHooPriestForgivenessResult(UserIndex, RequestId, eligibility, price, 0)
            Exit Sub
        End If
        If UserList(UserIndex).Stats.GLD < price Then
            Call WriteHooPriestForgivenessResult(UserIndex, RequestId, eHooForgivenessInsufficientGold, price, 0)
            Exit Sub
        End If
        mOfferActive(UserIndex) = True
        mOfferUserVersion(UserIndex) = UserList(UserIndex).VersionId
        mOfferRequestId(UserIndex) = RequestId
        mOfferToken(UserIndex) = NewOfferToken()
        mOfferPriestIndex(UserIndex) = targetRef.ArrayIndex
        mOfferPriestVersion(UserIndex) = targetRef.VersionId
        mOfferPrice(UserIndex) = price
        mOfferCreatedAt(UserIndex) = GetTickCountRaw()
        Call WriteHooPriestForgivenessResult(UserIndex, RequestId, eHooForgivenessOffer, price, mOfferToken(UserIndex))
        Exit Sub
    End If
    If Not mOfferActive(UserIndex) Or _
            mOfferUserVersion(UserIndex) <> UserList(UserIndex).VersionId Or _
            mOfferRequestId(UserIndex) <> RequestId Or mOfferToken(UserIndex) <> token Or _
            mOfferPriestIndex(UserIndex) <> targetRef.ArrayIndex Or _
            mOfferPriestVersion(UserIndex) <> targetRef.VersionId Then
        Call WriteHooPriestForgivenessResult(UserIndex, RequestId, eHooForgivenessInvalidOffer, 0, 0)
        Exit Sub
    End If
    If HooForgivenessOfferExpired(mOfferCreatedAt(UserIndex), GetTickCountRaw()) Then
        Call ClearHooForgivenessOffer(UserIndex)
        Call WriteHooPriestForgivenessResult(UserIndex, RequestId, eHooForgivenessInvalidOffer, 0, 0)
        Exit Sub
    End If
    Dim currentPrice As Long, currentEligibility As e_HooForgivenessResult
    currentEligibility = GetHooForgivenessEligibility(UserIndex, targetRef, currentPrice)
    If currentEligibility <> eHooForgivenessOffer Then
        Call ClearHooForgivenessOffer(UserIndex)
        Call WriteHooPriestForgivenessResult(UserIndex, RequestId, currentEligibility, currentPrice, 0)
        Exit Sub
    End If
    If currentPrice <> mOfferPrice(UserIndex) Then
        mOfferPrice(UserIndex) = currentPrice
        mOfferToken(UserIndex) = NewOfferToken()
        mOfferCreatedAt(UserIndex) = GetTickCountRaw()
        Call WriteHooPriestForgivenessResult(UserIndex, RequestId, eHooForgivenessPriceChanged, currentPrice, mOfferToken(UserIndex))
        Exit Sub
    End If
    If UserList(UserIndex).Stats.GLD < currentPrice Then
        Call ClearHooForgivenessOffer(UserIndex)
        Call WriteHooPriestForgivenessResult(UserIndex, RequestId, eHooForgivenessInsufficientGold, currentPrice, 0)
        Exit Sub
    End If
    Call ClearHooForgivenessOffer(UserIndex)
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - currentPrice
    paymentDeducted = True
    Call WriteUpdateGold(UserIndex)
    Call VolverCiudadano(UserIndex)
    If UserList(UserIndex).Faccion.Status <> e_Facciones.Ciudadano Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + currentPrice
        paymentDeducted = False
        Call WriteUpdateGold(UserIndex)
        Call WriteHooPriestForgivenessResult(UserIndex, RequestId, eHooForgivenessTransitionFailed, currentPrice, 0)
        Exit Sub
    End If
    paymentDeducted = False
    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, "80", 100, False))
    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave("100", UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
    Call WriteHooPriestForgivenessResult(UserIndex, RequestId, eHooForgivenessSuccess, currentPrice, 0)
    Exit Sub
HandleError:
    If paymentDeducted Then
        If UserList(UserIndex).Faccion.Status <> e_Facciones.Ciudadano Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + currentPrice
            paymentDeducted = False
            On Error Resume Next
            Call WriteUpdateGold(UserIndex)
            On Error GoTo 0
        End If
    End If
    Call ClearHooForgivenessOffer(UserIndex)
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleHooPriestForgiveness", Erl)
End Sub
