Attribute VB_Name = "LocaleDef"
Public Const MsgToFarToAttack = 8
Public Const MsgYouAreDeathAndCantAttack = 77
Public Const MsgToTired = 93
Public Const MsgRemoveSafeToAttack = 126
Public Const MsgInventoryIsFull = 328
Public Const MsgCantAttackYourself = 380
Public Const MSgNpcInmuneToEffect = 381
Public Const MsgInvalidGroupCount = 406
Public Const MsgCantChangeGroupSizeNow = 407
Public Const MsgInvalidUserState = 408
Public Const MsgTeamConfigSuccess = 409
Public Const MsgCantJoinPrivateLobby = 410
Public Const MsgDisconnectedPlayers = 411
Public Const MsgTeamRequiredToJoin = 412
Public Const MsgOnlyLeaderCanJoin = 413
Public Const MsgNotEnoughPlayersInGroup = 414
Public Const MsgNotEnoughPlayerForTeam = 415
Public Const MsgFactionForbidAction = 416
Public Const MsgClanForbidAction = 417
Public Const MsgDisableAttackGuardToContinue = 418
Public Const MsgInvalidTarget = 419
Public Const MsgTiredToPerformAction = 420
Public Const MsgRequiresMoreHealth = 421
Public Const MsgTargetAlreadyAffected = 422
Public Const MsgNotEnoughtStamina = 423
Public Const MsgToFar = 424
Public Const MsgInvalidTile = 425
Public Const MsgInvalidPass = 433
Public Const MsgPassForgat = 434
Public Const MsgPassNix = 435
Public Const MsgThanksForTravelNix = 436
Public Const MsgThanksForTravelForgat = 437
Public Const MsgStartingTrip = 438
Public Const MsgNotEnoughtAmunitions = 439
Public Const MsgEquipedArrowRequired = 440
Public Const MsgYouAreAlreadyInGroup = 441
Public Const MsgCantEquipYet = 442
Public Const MsgSkillAlreadyKnown = 443
Public Const MsgLandRequiredToUseSpell = 444
Public Const MsgWaterRequiredToUseSpell = 445
Public Const MsgCastOnlyOnSelf = 446
Public Const MsgExtraDamageDone = 447
Public Const MsgExtraDamageReceive = 448
Public Const MsgFacctionForbidAttack = 449
Public Const MsgYourTrapDidDamangeTo = 450
Public Const MsgTrapDidDamageToYou = 451
Public Const MsgTrapInmo = 452
Public Const MsgTrapPoison = 453
Public Const MsgFallIntoTrap = 454
Public Const MsgCaptainIsDeath = 455
Public Const MsgTeamNumberWin = 456
Public Const MsgCantPickFromYourStorage = 457
Public Const MsgCantCarryMoreThanOne = 458
Public Const MsgTeamGotAllCargo = 459
Public Const MsgBothTeamCargo = 460
Public Const MsgNavalConquestWinnerTeam = 461
Public Const MsgNavalConquestEvenMatch = 462
Public Const MsgCreateEventRoom = 463
Public Const MsgDeathMatchInstructions = 464
Public Const MsgHuntScenearioIntro = 465
Public Const MsgNavalConquestIntro = 466
Public Const MsgDeathMathInstructions = 467
Public Const MsgHuntScenarioInstructions = 468
Public Const MsgNavalConquestInstructions = 469
Public Const MsgCantDropCargoAtPos = 470
Public Const MsgSpellRequiresTransform = 471
Public Const MsgRequiredSpell = 472

Public Function GetRequiredWeaponLocaleId(ByVal WeaponType As e_WeaponType) As Integer
    Select Case WeaponType
        Case e_WeaponType.eAxe
            GetRequiredWeaponLocaleId = 426
        Case e_WeaponType.eBow
            GetRequiredWeaponLocaleId = 427
        Case e_WeaponType.eDagger
            GetRequiredWeaponLocaleId = 428
        Case e_WeaponType.eMace
            GetRequiredWeaponLocaleId = 429
        Case e_WeaponType.eStaff
            GetRequiredWeaponLocaleId = 430
        Case e_WeaponType.eSword
            GetRequiredWeaponLocaleId = 431
        Case e_WeaponType.eThrowableAxe
            GetRequiredWeaponLocaleId = 432
        Case Else
            Debug.Assert False
    End Select
End Function



