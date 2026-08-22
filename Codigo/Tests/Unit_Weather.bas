Attribute VB_Name = "Unit_Weather"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_weather() As Boolean
    Call UnitTesting.RunTest("test_weather_fog_state", test_weather_fog_state())
    Call UnitTesting.RunTest("test_weather_reset_decisions", test_weather_reset_decisions())
    Call UnitTesting.RunTest("test_weather_initial_sync", test_weather_initial_sync())
    Call UnitTesting.RunTest("test_weather_canonical_reset", test_weather_canonical_reset())
    test_suite_weather = True
End Function

Private Function test_weather_fog_state() As Boolean
    ServidorNublado = True
    Nieblando = False
    If Not IsAtmosphericFogActive() Then Exit Function
    ServidorNublado = False
    Nieblando = True
    If Not IsAtmosphericFogActive() Then Exit Function
    ServidorNublado = False
    Nieblando = False
    test_weather_fog_state = Not IsAtmosphericFogActive()
End Function

Private Function test_weather_reset_decisions() As Boolean
    Dim NotifyRain As Boolean
    Dim NotifySnow As Boolean
    Dim NotifyFog As Boolean
    Call DetermineWeatherResetNotifications(True, True, True, NotifyRain, NotifySnow, NotifyFog)
    If Not NotifyRain Or Not NotifySnow Or Not NotifyFog Then Exit Function
    Call DetermineWeatherResetNotifications(False, False, False, NotifyRain, NotifySnow, NotifyFog)
    test_weather_reset_decisions = Not NotifyRain And Not NotifySnow And Not NotifyFog
End Function

Private Function test_weather_initial_sync() As Boolean
    Dim SendRain As Boolean
    Dim SendSnow As Boolean
    Dim ToggleFog As Boolean
    Lloviendo = False
    Nebando = True
    ServidorNublado = False
    Nieblando = False
    Call DetermineInitialWeatherSynchronization(SendRain, SendSnow, ToggleFog)
    If SendRain Or Not SendSnow Or ToggleFog Then Exit Function
    Lloviendo = True
    ServidorNublado = True
    Call DetermineInitialWeatherSynchronization(SendRain, SendSnow, ToggleFog)
    test_weather_initial_sync = SendRain And SendSnow And ToggleFog
End Function

Private Function test_weather_canonical_reset() As Boolean
    Lloviendo = True
    Nebando = True
    ServidorNublado = True
    Nieblando = True
    Call ResetMeteo(False)
    test_weather_canonical_reset = Not Lloviendo And Not Nebando And _
        Not ServidorNublado And Not Nieblando And _
        Not frmMain.Truenos.Enabled And frmMain.TimerMeteorologia.Enabled And _
        TimerMeteorologico = 30
End Function

#End If
