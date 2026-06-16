Attribute VB_Name = "ModCollectibleCards"
Option Explicit

Public Enum e_CollectibleCardTags
    Hostile = 1
    Friendly = 2
    Fire = 4
    Water = 8
    Earth = 16
    Wind = 32
    Normal = 64
    Citizen = 128
    Draconic = 256
    Aberration = 512
    WildLife = 1024
    Forest = 2048
    Desert = 4096
    Dungeon = 8192
    Ocean = 16384
    Undead = 32768
    Mythologic = 65536
    Critter = 131072
    Humanoid = 262144
    NonHumanoid = 524288
    Boss = 1048576
End Enum

Public Enum e_CollectibleCardRarity
    Common = 1
    Uncommon = 2
    Rare = 4
    Epic = 8
    Legendary = 16
End Enum

Public Enum e_CollectibleCardAchievementSlot
    'simpletags are simple condition like have 10 cards from tag x and have bonus against that tag
    SimpleTags = 1
    
    'compoundtags are more robust unions of tags, have 10 and 20 from x and y tag and have an special bonus
    CompoundTags = 2
    
    'rarities validate if you have at least all rarities of a determined tag
    Rarities = 3
    
    'specific cards validate if you have all rarities of a specific card type
    SpecificCards = 4
End Enum

Public Enum e_CollectibleSpecificCard
    Wolf = 1
    Dragon = 2
    Skeleton = 4
    Goblin = 8
    Phoenix = 16
    ' Add more specific card types as needed
End Enum

Private Const UPSERT_NEW_COLLECTIBLE_CARD As String = _
    "INSERT INTO account_collectible_cards (account_id, card_id, last_updated, quantity) " & _
    "VALUES (?, ?, ?, 1) " & _
    "ON CONFLICT(account_id, card_id) " & _
    "DO UPDATE SET quantity = quantity + 1, last_updated = excluded.last_updated;"
    
Private Const SELECT_ALL_ACCOUNT_CARDS As String = "SELECT    "

Private Const GET_ACCOUNT_COLLECTIBLE_CARDS As String = _
    "SELECT acc.id, acc.account_id, acc.card_id, acc.quantity, " & _
    "acc.first_added, acc.last_updated, c.description, c.rarity, c.tags " & _
    "FROM account_collectible_cards acc " & _
    "INNER JOIN collectible_cards c ON acc.card_id = c.id " & _
    "WHERE acc.account_id = ? " & _
    "ORDER BY acc.first_added DESC;"

Public Function GetUserCollectibleCards(ByVal AccountId As Integer) As ADODB.Recordset
    On Error GoTo GetUserCollectibleCards_Err
    Set GetUserCollectibleCards = Query(GET_ACCOUNT_COLLECTIBLE_CARDS, AccountId)
    Exit Function
GetUserCollectibleCards_Err:
    Call TraceError(Err.Number, Err.Description, "ModCollectibleCards.GetUserCollectibleCards", Erl)
End Function

Public Sub AddCollectibleCardToUser(ByVal UserIndex As Integer, ByRef ObjCard As t_Obj)
    If ObjCard.ObjIndex = 0 Then Exit Sub
    With UserList(UserIndex)
        Call InsertCollectibleCardIntoDatabase(.AccountID, ObjCard)
    End With
End Sub

Public Sub InsertCollectibleCardIntoDatabase(ByVal Acount_Id As Integer, ByRef ObjCard As t_Obj)
    On Error GoTo InsertCollectibleCardIntoDatabase_Err
    Dim RS As ADODB.Recordset
    Set RS = Query(UPSERT_NEW_COLLECTIBLE_CARD, Acount_Id, ObjData(ObjCard.ObjIndex).CollectibleCardId, CStr(DateTime.Now))
    Exit Sub
InsertCollectibleCardIntoDatabase_Err:
    Call TraceError(Err.Number, Err.Description, "ModCollectibleCards.InsertCollectibleCardIntoDatabase", Erl)
End Sub

Public Function CountCardsByTag(ByRef Rs As ADODB.Recordset, ByVal Tag As e_CollectibleCardTags) As Integer
    Dim Count As Integer
    Dim CurrentPosition As Long
    Count = 0
    If Rs Is Nothing Or Rs.EOF Then
        CountCardsByTag = 0
        Exit Function
    End If
    ' Save current position
    If Rs.Supports(adBookmark) Then
        CurrentPosition = Rs.Bookmark
    End If
    ' Move to first record
    Rs.MoveFirst
    Do While Not Rs.EOF
        If (Rs!Tags And Tag) = Tag Then
            Count = Count + Rs!quantity
        End If
        Rs.MoveNext
    Loop
    ' Restore position
    If Rs.Supports(adBookmark) Then
        Rs.Bookmark = CurrentPosition
    End If
    CountCardsByTag = Count
End Function


'this function below controls business logic for card collections
Public Sub SetupUserAccountCollectibleCardAchievements(ByRef User As t_User)
    Dim Rs As ADODB.Recordset
    
    Set Rs = GetUserCollectibleCards(User.AccountId)
    If Rs Is Nothing Or Rs.EOF Then
        Exit Sub
    End If
    
    Dim HostileCount As Integer
    Dim FriendlyCount As Integer
    Dim FireCount As Integer
    Dim WaterCount As Integer
    Dim EarthCount As Integer
    Dim WindCount As Integer
    Dim NormalCount As Integer
    Dim CitizenCount As Integer
    Dim DraconicCount As Integer
    Dim AberrationCount As Integer
    Dim WildLifeCount As Integer
    Dim ForestCount As Integer
    Dim DesertCount As Integer
    Dim DungeonCount As Integer
    Dim OceanCount As Integer
    Dim UndeadCount As Integer
    Dim MythologicCount As Integer
    Dim CritterCount As Integer
    Dim HumanoidCount As Integer
    Dim NonHumanoidCount As Integer
    Dim BossCount As Integer
    
    HostileCount = CountCardsByTag(Rs, e_CollectibleCardTags.Hostile)
    FriendlyCount = CountCardsByTag(Rs, e_CollectibleCardTags.Friendly)
    FireCount = CountCardsByTag(Rs, e_CollectibleCardTags.Fire)
    WaterCount = CountCardsByTag(Rs, e_CollectibleCardTags.Water)
    EarthCount = CountCardsByTag(Rs, e_CollectibleCardTags.Earth)
    WindCount = CountCardsByTag(Rs, e_CollectibleCardTags.Wind)
    NormalCount = CountCardsByTag(Rs, e_CollectibleCardTags.Normal)
    CitizenCount = CountCardsByTag(Rs, e_CollectibleCardTags.Citizen)
    DraconicCount = CountCardsByTag(Rs, e_CollectibleCardTags.Draconic)
    AberrationCount = CountCardsByTag(Rs, e_CollectibleCardTags.Aberration)
    WildLifeCount = CountCardsByTag(Rs, e_CollectibleCardTags.WildLife)
    ForestCount = CountCardsByTag(Rs, e_CollectibleCardTags.Forest)
    DesertCount = CountCardsByTag(Rs, e_CollectibleCardTags.Desert)
    DungeonCount = CountCardsByTag(Rs, e_CollectibleCardTags.Dungeon)
    OceanCount = CountCardsByTag(Rs, e_CollectibleCardTags.Ocean)
    UndeadCount = CountCardsByTag(Rs, e_CollectibleCardTags.Undead)
    MythologicCount = CountCardsByTag(Rs, e_CollectibleCardTags.Mythologic)
    CritterCount = CountCardsByTag(Rs, e_CollectibleCardTags.Critter)
    HumanoidCount = CountCardsByTag(Rs, e_CollectibleCardTags.Humanoid)
    NonHumanoidCount = CountCardsByTag(Rs, e_CollectibleCardTags.NonHumanoid)
    BossCount = CountCardsByTag(Rs, e_CollectibleCardTags.Boss)
    
    ' ========================================
    ' SLOT 1: SimpleTags - Single tag achievements
    ' ========================================
    If HostileCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Hostile)
    End If
    
    If FriendlyCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Friendly)
    End If
    
    If FireCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Fire)
    End If
    
    If WaterCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Water)
    End If
    
    If EarthCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Earth)
    End If
    
    If WindCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Wind)
    End If
    
    If NormalCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Normal)
    End If
    
    If CitizenCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Citizen)
    End If
    
    If DraconicCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Draconic)
    End If
    
    If AberrationCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Aberration)
    End If
    
    If WildLifeCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.WildLife)
    End If
    
    If ForestCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Forest)
    End If
    
    If DesertCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Desert)
    End If
    
    If DungeonCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Dungeon)
    End If
    
    If OceanCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Ocean)
    End If
    
    If UndeadCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Undead)
    End If
    
    If MythologicCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Mythologic)
    End If
    
    If CritterCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Critter)
    End If
    
    If HumanoidCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Humanoid)
    End If
    
    If NonHumanoidCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.NonHumanoid)
    End If
    
    If BossCount >= 5 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SimpleTags), e_CollectibleCardTags.Boss)
    End If
    
    ' ========================================
    ' SLOT 2: CompoundTags - Multiple tag combinations
    ' ========================================
    
    ' Fire + Hostile (10 each) - Infernal Warrior
    If FireCount >= 10 And HostileCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.CompoundTags), e_CollectibleCardTags.Fire Or e_CollectibleCardTags.Hostile)
    End If
    
    ' Water + Friendly (10 each) - Aquatic Guardian
    If WaterCount >= 10 And FriendlyCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.CompoundTags), e_CollectibleCardTags.Water Or e_CollectibleCardTags.Friendly)
    End If
    
    ' Earth + Forest (10 each) - Nature's Protector
    If EarthCount >= 10 And ForestCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.CompoundTags), e_CollectibleCardTags.Earth Or e_CollectibleCardTags.Forest)
    End If
    
    ' Wind + WildLife (10 each) - Sky Hunter
    If WindCount >= 10 And WildLifeCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.CompoundTags), e_CollectibleCardTags.Wind Or e_CollectibleCardTags.WildLife)
    End If
    
    ' Fire + Draconic (10 each) - Dragon Master
    If FireCount >= 10 And DraconicCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.CompoundTags), e_CollectibleCardTags.Fire Or e_CollectibleCardTags.Draconic)
    End If
    
    ' Undead + Dungeon (10 each) - Necropolis Lord
    If UndeadCount >= 10 And DungeonCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.CompoundTags), e_CollectibleCardTags.Undead Or e_CollectibleCardTags.Dungeon)
    End If
    
    ' Water + Ocean (10 each) - Deep Sea Explorer
    If WaterCount >= 10 And OceanCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.CompoundTags), e_CollectibleCardTags.Water Or e_CollectibleCardTags.Ocean)
    End If
    
    ' Hostile + Boss (20 + 10) - Elite Slayer
    If HostileCount >= 20 And BossCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.CompoundTags), e_CollectibleCardTags.Hostile Or e_CollectibleCardTags.Boss)
    End If
    
    ' Mythologic + Draconic (10 each) - Legendary Tamer
    If MythologicCount >= 10 And DraconicCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.CompoundTags), e_CollectibleCardTags.Mythologic Or e_CollectibleCardTags.Draconic)
    End If
    
    ' Citizen + Humanoid (10 each) - Civilization Builder
    If CitizenCount >= 10 And HumanoidCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.CompoundTags), e_CollectibleCardTags.Citizen Or e_CollectibleCardTags.Humanoid)
    End If
    
    ' Aberration + Hostile (10 each) - Void Conqueror
    If AberrationCount >= 10 And HostileCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.CompoundTags), e_CollectibleCardTags.Aberration Or e_CollectibleCardTags.Hostile)
    End If
    
    ' Critter + Forest (10 each) - Woodland Keeper
    If CritterCount >= 10 And ForestCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.CompoundTags), e_CollectibleCardTags.Critter Or e_CollectibleCardTags.Forest)
    End If
    
    ' Desert + Hostile (10 each) - Sandstorm Raider
    If DesertCount >= 10 And HostileCount >= 10 Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.CompoundTags), e_CollectibleCardTags.Desert Or e_CollectibleCardTags.Hostile)
    End If
    
    ' ========================================
    ' SLOT 3: Rarities - Complete rarity sets per tag (PLACEHOLDER)
    ' ========================================
    ' TODO: Implement tag-based rarity collection achievements
    
    ' ========================================
    ' SLOT 4: SpecificCards - All rarities of specific card types
    ' ========================================
    
    ' Wolf Master - Has Common, Uncommon, Rare, Epic, and Legendary Wolf
    If HasAllRaritiesForSpecificCard(Rs, e_CollectibleSpecificCard.Wolf) Then
        Call SetMask(User.CollectibleCardAchievements(e_CollectibleCardAchievementSlot.SpecificCards), e_CollectibleSpecificCard.Wolf)
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub

' Helper function to check if user has all rarities for a specific card type
' Example: Wolf card exists in Common (1), Uncommon (2), Rare (4), Epic (8), Legendary (16)
Public Function HasAllRaritiesForSpecificCard(ByRef Rs As ADODB.Recordset, ByVal SpecificCard As e_CollectibleSpecificCard) As Boolean
    Dim RarityMask As Long
    Dim CurrentPosition As Long
    
    RarityMask = 0
    HasAllRaritiesForSpecificCard = False
    
    If Rs Is Nothing Or Rs.EOF Then
        Exit Function
    End If
    
    ' Save current position
    If Rs.Supports(adBookmark) Then
        CurrentPosition = Rs.Bookmark
    End If
    
    ' Move to first record
    Rs.MoveFirst
    
    ' Iterate and collect all rarities found for this specific card
    Do While Not Rs.EOF
        ' Check if this card matches the specific card type
        ' You need to add a field in your query or have a way to identify the card type
        ' For now, assuming you have a way to map card_id to specific card type
        If IsCardOfType(Rs!card_id, SpecificCard) Then
            ' Add this rarity to the mask
            RarityMask = RarityMask Or Rs!Rarity
        End If
        Rs.MoveNext
    Loop
    
    ' Restore position
    If Rs.Supports(adBookmark) Then
        Rs.Bookmark = CurrentPosition
    End If
    
    ' Check if all rarities are present
    ' Assuming rarities are: Common=1, Uncommon=2, Rare=4, Epic=8, Legendary=16
    Const ALL_RARITIES As Long = 31 ' 1 + 2 + 4 + 8 + 16
    
    HasAllRaritiesForSpecificCard = (RarityMask = ALL_RARITIES)
    
End Function

' Helper function to check if a card_id belongs to a specific card type
' You need to implement this based on your card data structure
Private Function IsCardOfType(ByVal CardId As Integer, ByVal SpecificCard As e_CollectibleSpecificCard) As Boolean
    ' TODO: Implement logic to map card_id to specific card types
    ' This could be done via:
    ' 1. A lookup table in the database
    ' 2. A mapping array in memory
    ' 3. A naming convention (e.g., card IDs 100-104 are all Wolf variants)
    ' 4. An additional field in collectible_cards table called "card_type"
    
    ' Placeholder example using card ID ranges:
    Select Case SpecificCard
        Case e_CollectibleSpecificCard.Wolf
            ' Assuming Wolf cards have IDs 1-5 (one for each rarity)
            IsCardOfType = (CardId >= 1 And CardId <= 5)
        Case e_CollectibleSpecificCard.Dragon
            IsCardOfType = (CardId >= 6 And CardId <= 10)
        Case e_CollectibleSpecificCard.Skeleton
            IsCardOfType = (CardId >= 11 And CardId <= 15)
        ' Add more cases as needed
        Case Else
            IsCardOfType = False
    End Select
End Function
