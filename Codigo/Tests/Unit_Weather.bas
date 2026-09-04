Attribute VB_Name = "Unit_Weather"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_weather() As Boolean
    Call UnitTesting.RunTest("test_weather_fog_state", test_weather_fog_state())
    Call UnitTesting.RunTest("test_weather_schedule_configuration", test_weather_schedule_configuration())
    Call UnitTesting.RunTest("test_weather_probability_decisions", test_weather_probability_decisions())
    Call UnitTesting.RunTest("test_weather_state_transitions", test_weather_state_transitions())
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

Private Function test_weather_schedule_configuration() As Boolean
    test_weather_schedule_configuration = METEOROLOGICAL_CYCLE_MINUTES = 20 And _
        METEOROLOGICAL_CLOUD_COUNTDOWN_MINUTES = 11 And _
        METEOROLOGICAL_PRECIPITATION_COUNTDOWN_MINUTES = 6
End Function

Private Function test_weather_probability_decisions() As Boolean
    If Not WeatherCloudRollSucceeds(1) Then Exit Function
    If WeatherCloudRollSucceeds(2) Then Exit Function
    If Not WeatherPrecipitationRollSucceeds(1, True) Then Exit Function
    If Not WeatherPrecipitationRollSucceeds(2, True) Then Exit Function
    If WeatherPrecipitationRollSucceeds(3, True) Then Exit Function
    If WeatherPrecipitationRollSucceeds(1, False) Then Exit Function
    test_weather_probability_decisions = True
End Function

Private Function test_weather_state_transitions() As Boolean
    Call ResetMeteo(False)
    Call SetAtmosphericFogState(True, 31)
    If Not ServidorNublado Or Not Nieblando Or IntensidadDeNubes <> 31 Then Exit Function
    Call SetPrecipitationState(True)
    If Not Lloviendo Or Not Nebando Then Exit Function
    Call SetAtmosphericFogState(False)
    If ServidorNublado Or Nieblando Or IntensidadDeNubes <> 0 Then Exit Function
    If Not Lloviendo Or Not Nebando Then Exit Function
    Call SetPrecipitationState(False)
    test_weather_state_transitions = Not Lloviendo And Not Nebando
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
        TimerMeteorologico = METEOROLOGICAL_CYCLE_MINUTES And _
        ProbabilidadNublar = 0 And ProbabilidadLLuvia = 0 And _
        IntensidadDeNubes = 0
End Function

#End If
