Attribute VB_Name = "Unit_GameStatus"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_gamestatus() As Boolean
    Call UnitTesting.RunTest("test_esnewbie_below", test_esnewbie_below())
    Call UnitTesting.RunTest("test_esnewbie_at_limit", test_esnewbie_at_limit())
    Call UnitTesting.RunTest("test_esnewbie_above", test_esnewbie_above())
    Call UnitTesting.RunTest("test_esnewbie_zero_index", test_esnewbie_zero_index())
    Call UnitTesting.RunTest("test_esciudadano_true", test_esciudadano_true())
    Call UnitTesting.RunTest("test_esciudadano_false", test_esciudadano_false())
    Call UnitTesting.RunTest("test_escriminal_true", test_escriminal_true())
    Call UnitTesting.RunTest("test_escriminal_false", test_escriminal_false())
    Call UnitTesting.RunTest("test_esarmada_true", test_esarmada_true())
    Call UnitTesting.RunTest("test_esarmada_consejo", test_esarmada_consejo())
    Call UnitTesting.RunTest("test_esarmada_false", test_esarmada_false())
    Call UnitTesting.RunTest("test_escaos_true", test_escaos_true())
    Call UnitTesting.RunTest("test_escaos_concilio", test_escaos_concilio())
    Call UnitTesting.RunTest("test_escaos_false", test_escaos_false())
    Call UnitTesting.RunTest("test_faction_zero_index", test_faction_zero_index())
    Call UnitTesting.RunTest("test_esgm_admin", test_esgm_admin())
    Call UnitTesting.RunTest("test_esgm_dios", test_esgm_dios())
    Call UnitTesting.RunTest("test_esgm_semidios", test_esgm_semidios())
    Call UnitTesting.RunTest("test_esgm_consejero", test_esgm_consejero())
    Call UnitTesting.RunTest("test_esgm_no_privs", test_esgm_no_privs())
    Call UnitTesting.RunTest("test_esgm_zero_index", test_esgm_zero_index())
    Call UnitTesting.RunTest("test_esnewbie_threshold_property", test_esnewbie_threshold_property())
    test_suite_gamestatus = True
End Function

' ============================================================
' EsNewbie Tests
' ============================================================

' EsNewbie returns True when ELV is below LimiteNewbie (12).
Private Function test_esnewbie_below() As Boolean
    On Error GoTo Err_Handler
    test_esnewbie_below = True

    Dim origELV As Byte
    origELV = UserList(1).Stats.ELV

    UserList(1).Stats.ELV = 5
    If Not EsNewbie(1) Then test_esnewbie_below = False

    UserList(1).Stats.ELV = origELV
    Exit Function
Err_Handler:
    UserList(1).Stats.ELV = origELV
    test_esnewbie_below = False
End Function

' EsNewbie returns True when ELV equals LimiteNewbie (12).
Private Function test_esnewbie_at_limit() As Boolean
    On Error GoTo Err_Handler
    test_esnewbie_at_limit = True

    Dim origELV As Byte
    origELV = UserList(1).Stats.ELV

    UserList(1).Stats.ELV = 12
    If Not EsNewbie(1) Then test_esnewbie_at_limit = False

    UserList(1).Stats.ELV = origELV
    Exit Function
Err_Handler:
    UserList(1).Stats.ELV = origELV
    test_esnewbie_at_limit = False
End Function

' EsNewbie returns False when ELV is above LimiteNewbie (12).
Private Function test_esnewbie_above() As Boolean
    On Error GoTo Err_Handler
    test_esnewbie_above = True

    Dim origELV As Byte
    origELV = UserList(1).Stats.ELV

    UserList(1).Stats.ELV = 13
    If EsNewbie(1) Then test_esnewbie_above = False

    UserList(1).Stats.ELV = origELV
    Exit Function
Err_Handler:
    UserList(1).Stats.ELV = origELV
    test_esnewbie_above = False
End Function

' EsNewbie returns False for UserIndex 0.
Private Function test_esnewbie_zero_index() As Boolean
    On Error GoTo Err_Handler
    test_esnewbie_zero_index = True

    If EsNewbie(0) Then test_esnewbie_zero_index = False

    Exit Function
Err_Handler:
    test_esnewbie_zero_index = False
End Function

' ============================================================
' Faction Status Helper Tests
' ============================================================

' esCiudadano returns True for Ciudadano status.
Private Function test_esciudadano_true() As Boolean
    On Error GoTo Err_Handler
    test_esciudadano_true = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.Ciudadano
    If Not esCiudadano(1) Then test_esciudadano_true = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_esciudadano_true = False
End Function

' esCiudadano returns False for Criminal status.
Private Function test_esciudadano_false() As Boolean
    On Error GoTo Err_Handler
    test_esciudadano_false = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.Criminal
    If esCiudadano(1) Then test_esciudadano_false = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_esciudadano_false = False
End Function

' esCriminal returns True for Criminal status.
Private Function test_escriminal_true() As Boolean
    On Error GoTo Err_Handler
    test_escriminal_true = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.Criminal
    If Not esCriminal(1) Then test_escriminal_true = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_escriminal_true = False
End Function

' esCriminal returns False for Ciudadano status.
Private Function test_escriminal_false() As Boolean
    On Error GoTo Err_Handler
    test_escriminal_false = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.Ciudadano
    If esCriminal(1) Then test_escriminal_false = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_escriminal_false = False
End Function

' esArmada returns True for Armada status.
Private Function test_esarmada_true() As Boolean
    On Error GoTo Err_Handler
    test_esarmada_true = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.Armada
    If Not esArmada(1) Then test_esarmada_true = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_esarmada_true = False
End Function

' esArmada returns True for consejo (allied with Armada).
Private Function test_esarmada_consejo() As Boolean
    On Error GoTo Err_Handler
    test_esarmada_consejo = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.consejo
    If Not esArmada(1) Then test_esarmada_consejo = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_esarmada_consejo = False
End Function

' esArmada returns False for Caos status.
Private Function test_esarmada_false() As Boolean
    On Error GoTo Err_Handler
    test_esarmada_false = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.Caos
    If esArmada(1) Then test_esarmada_false = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_esarmada_false = False
End Function

' esCaos returns True for Caos status.
Private Function test_escaos_true() As Boolean
    On Error GoTo Err_Handler
    test_escaos_true = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.Caos
    If Not esCaos(1) Then test_escaos_true = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_escaos_true = False
End Function

' esCaos returns True for concilio (allied with Caos).
Private Function test_escaos_concilio() As Boolean
    On Error GoTo Err_Handler
    test_escaos_concilio = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.concilio
    If Not esCaos(1) Then test_escaos_concilio = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_escaos_concilio = False
End Function

' esCaos returns False for Armada status.
Private Function test_escaos_false() As Boolean
    On Error GoTo Err_Handler
    test_escaos_false = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.Armada
    If esCaos(1) Then test_escaos_false = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_escaos_false = False
End Function

' All faction helpers return False for UserIndex 0.
Private Function test_faction_zero_index() As Boolean
    On Error GoTo Err_Handler
    test_faction_zero_index = True

    If esCiudadano(0) Then test_faction_zero_index = False: Exit Function
    If esCriminal(0) Then test_faction_zero_index = False: Exit Function
    If esArmada(0) Then test_faction_zero_index = False: Exit Function
    If esCaos(0) Then test_faction_zero_index = False: Exit Function

    Exit Function
Err_Handler:
    test_faction_zero_index = False
End Function

' ============================================================
' EsGM Privilege Tests
' ============================================================

' EsGM returns True for Admin privilege.
Private Function test_esgm_admin() As Boolean
    On Error GoTo Err_Handler
    test_esgm_admin = True

    Dim origPrivs As Long
    origPrivs = UserList(1).flags.Privilegios

    UserList(1).flags.Privilegios = e_PlayerType.Admin
    If Not EsGM(1) Then test_esgm_admin = False

    UserList(1).flags.Privilegios = origPrivs
    Exit Function
Err_Handler:
    UserList(1).flags.Privilegios = origPrivs
    test_esgm_admin = False
End Function

' EsGM returns True for Dios privilege.
Private Function test_esgm_dios() As Boolean
    On Error GoTo Err_Handler
    test_esgm_dios = True

    Dim origPrivs As Long
    origPrivs = UserList(1).flags.Privilegios

    UserList(1).flags.Privilegios = e_PlayerType.Dios
    If Not EsGM(1) Then test_esgm_dios = False

    UserList(1).flags.Privilegios = origPrivs
    Exit Function
Err_Handler:
    UserList(1).flags.Privilegios = origPrivs
    test_esgm_dios = False
End Function

' EsGM returns True for SemiDios privilege.
Private Function test_esgm_semidios() As Boolean
    On Error GoTo Err_Handler
    test_esgm_semidios = True

    Dim origPrivs As Long
    origPrivs = UserList(1).flags.Privilegios

    UserList(1).flags.Privilegios = e_PlayerType.SemiDios
    If Not EsGM(1) Then test_esgm_semidios = False

    UserList(1).flags.Privilegios = origPrivs
    Exit Function
Err_Handler:
    UserList(1).flags.Privilegios = origPrivs
    test_esgm_semidios = False
End Function

' EsGM returns True for Consejero privilege.
Private Function test_esgm_consejero() As Boolean
    On Error GoTo Err_Handler
    test_esgm_consejero = True

    Dim origPrivs As Long
    origPrivs = UserList(1).flags.Privilegios

    UserList(1).flags.Privilegios = e_PlayerType.Consejero
    If Not EsGM(1) Then test_esgm_consejero = False

    UserList(1).flags.Privilegios = origPrivs
    Exit Function
Err_Handler:
    UserList(1).flags.Privilegios = origPrivs
    test_esgm_consejero = False
End Function

' EsGM returns False when Privilegios is 0 (no privileges).
Private Function test_esgm_no_privs() As Boolean
    On Error GoTo Err_Handler
    test_esgm_no_privs = True

    Dim origPrivs As Long
    origPrivs = UserList(1).flags.Privilegios

    UserList(1).flags.Privilegios = 0
    If EsGM(1) Then test_esgm_no_privs = False

    UserList(1).flags.Privilegios = origPrivs
    Exit Function
Err_Handler:
    UserList(1).flags.Privilegios = origPrivs
    test_esgm_no_privs = False
End Function

' EsGM returns False for UserIndex 0.
Private Function test_esgm_zero_index() As Boolean
    On Error GoTo Err_Handler
    test_esgm_zero_index = True

    If EsGM(0) Then test_esgm_zero_index = False

    Exit Function
Err_Handler:
    test_esgm_zero_index = False
End Function

' ============================================================
' Property Test: EsNewbie Threshold
' ============================================================

' Property 5: EsNewbie threshold
' For any level 1 through 50, EsNewbie returns True iff level <= 12.
' Validates: Requirements 3.1, 3.2, 3.3
Private Function test_esnewbie_threshold_property() As Boolean
    On Error GoTo Err_Handler
    test_esnewbie_threshold_property = True

    Dim origELV As Byte
    origELV = UserList(1).Stats.ELV

    Dim lvl As Integer
    For lvl = 1 To 50
        UserList(1).Stats.ELV = CByte(lvl)

        If lvl <= 12 Then
            If Not EsNewbie(1) Then
                UserList(1).Stats.ELV = origELV
                test_esnewbie_threshold_property = False
                Exit Function
            End If
        Else
            If EsNewbie(1) Then
                UserList(1).Stats.ELV = origELV
                test_esnewbie_threshold_property = False
                Exit Function
            End If
        End If
    Next lvl

    UserList(1).Stats.ELV = origELV
    Exit Function
Err_Handler:
    UserList(1).Stats.ELV = origELV
    test_esnewbie_threshold_property = False
End Function

#End If
