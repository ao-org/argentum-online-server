VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Servidor Argentum 20"
   ClientHeight    =   6225
   ClientLeft      =   1950
   ClientTop       =   1695
   ClientWidth     =   6930
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6225
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CerrarYForzarActualizar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cerrar y forzar actualización"
      Height          =   495
      Left            =   5160
      MaskColor       =   &H000040C0&
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Timer Invasion 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3120
      Top             =   3120
   End
   Begin VB.Timer TiempoRetos 
      Interval        =   10000
      Left            =   3120
      Top             =   4200
   End
   Begin VB.Timer TimerGuardarUsuarios 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   2640
      Top             =   3120
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Guardar y cerrar"
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Recargar intervalos.ini"
      Height          =   495
      Left            =   5160
      TabIndex        =   32
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Recargar Balance.dat"
      Height          =   495
      Left            =   5160
      TabIndex        =   31
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Recargar configuracion.ini"
      Height          =   495
      Left            =   5160
      TabIndex        =   30
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Recargar Server.ini"
      Height          =   495
      Left            =   5160
      TabIndex        =   29
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Guardar Usuarios"
      Height          =   495
      Left            =   5160
      TabIndex        =   28
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Recargar Npcs"
      Height          =   495
      Left            =   5160
      TabIndex        =   27
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Recargar Objetos"
      Height          =   495
      Left            =   5160
      TabIndex        =   26
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Recargar Administradores"
      Height          =   495
      Left            =   5160
      TabIndex        =   25
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eventos"
      Height          =   6015
      Left            =   7920
      TabIndex        =   22
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label cuentas 
         Caption         =   "0"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Timer TimerRespawn 
      Interval        =   1000
      Left            =   1680
      Top             =   4200
   End
   Begin VB.Timer EstadoTimer 
      Interval        =   1000
      Left            =   240
      Top             =   3480
   End
   Begin VB.Timer Evento 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   720
      Top             =   3480
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Logeo de eventos"
      Height          =   1455
      Left            =   120
      TabIndex        =   20
      Top             =   4680
      Width           =   4935
      Begin VB.ListBox List1 
         Height          =   1110
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Timer TimerMeteorologia 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   240
      Top             =   4200
   End
   Begin VB.Timer UptimeTimer 
      Interval        =   1000
      Left            =   3600
      Top             =   3060
   End
   Begin VB.Timer Truenos 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2160
      Top             =   4200
   End
   Begin VB.Timer HoraFantasia 
      Interval        =   1000
      Left            =   2640
      Top             =   4200
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Información general"
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4935
      Begin MSWinsockLib.Winsock auxSocket 
         Left            =   3960
         Top             =   960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nublado"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lloviendo"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label3"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hora Fantasia Servidor:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblhora 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label3"
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hora Actual Servidor:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label CantUsuarios 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de usuarios: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label lblLimpieza 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpieza de objetos cada: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   4215
      End
   End
   Begin VB.Timer LimpiezaTimer 
      Interval        =   60000
      Left            =   1200
      Top             =   4200
   End
   Begin VB.Timer SubastaTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   4200
   End
   Begin VB.Timer packetResend 
      Interval        =   5
      Left            =   240
      Top             =   3060
   End
   Begin VB.CommandButton CMDDUMP 
      Caption         =   "dump"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   3060
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1200
      Top             =   3060
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   2160
      Top             =   3120
   End
   Begin VB.Timer Minuto 
      Interval        =   60000
      Left            =   4080
      Top             =   3060
   End
   Begin VB.Timer KillLog 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1680
      Top             =   3120
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4560
      Top             =   3060
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mensaje a jugadores"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   4935
      Begin VB.CommandButton Command2 
         Caption         =   "Via consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Via ventana"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mensaje:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Estadisticas de Paquetes"
      Height          =   1695
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   4935
      Begin VB.ListBox listaDePaquetes 
         Height          =   1110
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label paquetesRecibidos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         Height          =   375
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Paquetes Recibidos:"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Basado en rao"
      Height          =   495
      Left            =   5280
      TabIndex        =   35
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   495
      Left            =   2880
      TabIndex        =   34
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   45
   End
   Begin VB.Menu mnuControles 
      Caption         =   "Opciones"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Systray Servidor"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar Servidor"
      End
   End
   Begin VB.Menu donador 
      Caption         =   "Donador"
      Begin VB.Menu addtimeDonador 
         Caption         =   "Cargar tiempo"
      End
      Begin VB.Menu loadcredit 
         Caption         =   "Cargar Creditos"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public ESCUCHADAS As Long

Private Type NOTIFYICONDATA

    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64

End Type
   
Const NIM_ADD = 0

Const NIM_DELETE = 2

Const NIF_MESSAGE = 1

Const NIF_ICON = 2

Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200

Const WM_LBUTTONDBLCLK = &H203

Const WM_RBUTTONUP = &H205

Private GuardarYCerrar As Boolean

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hwnd As Long, Id As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
        
        On Error GoTo setNOTIFYICONDATA_Err
        

        Dim nidTemp As NOTIFYICONDATA

100     nidTemp.cbSize = Len(nidTemp)
102     nidTemp.hwnd = hwnd
104     nidTemp.uID = Id
106     nidTemp.uFlags = flags
108     nidTemp.uCallbackMessage = CallbackMessage
110     nidTemp.hIcon = Icon
112     nidTemp.szTip = Tip & Chr$(0)

114     setNOTIFYICONDATA = nidTemp

        
        Exit Function

setNOTIFYICONDATA_Err:
116     Call RegistrarError(Err.Number, Err.Description, "frmMain.setNOTIFYICONDATA", Erl)
118     Resume Next
        
End Function

Sub CheckIdleUser()
        
        On Error GoTo CheckIdleUser_Err
        

        Dim iUserIndex As Long
    
100     For iUserIndex = 1 To MaxUsers

            'Conexion activa? y es un usuario loggeado?
102         If UserList(iUserIndex).ConnID <> -1 And UserList(iUserIndex).flags.UserLogged Then
                'Actualiza el contador de inactividad
104             UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1

106             If UserList(iUserIndex).Counters.IdleCount >= IdleLimit Then
108                 Call WriteShowMessageBox(iUserIndex, "Demasiado tiempo inactivo. Has sido desconectado..")

                    'mato los comercios seguros
110                 If UserList(iUserIndex).ComUsu.DestUsu > 0 Then
112                     If UserList(UserList(iUserIndex).ComUsu.DestUsu).flags.UserLogged Then
114                         If UserList(UserList(iUserIndex).ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
116                             Call WriteConsoleMsg(UserList(iUserIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
118                             Call FinComerciarUsu(UserList(iUserIndex).ComUsu.DestUsu)
                            

                            End If

                        End If

120                     Call FinComerciarUsu(iUserIndex)

                    End If

122                 Call Cerrar_Usuario(iUserIndex)

                End If

            End If

124     Next iUserIndex

        
        Exit Sub

CheckIdleUser_Err:
126     Call RegistrarError(Err.Number, Err.Description, "frmMain.CheckIdleUser", Erl)
128     Resume Next
        
End Sub

Private Sub addtimeDonador_Click()
        
        On Error GoTo addtimeDonador_Click_Err
        

        Dim Tmp  As String

        Dim tmp2 As String

100     Tmp = InputBox("Cuenta?", "Ingrese la cuenta")

102     If FileExist(CuentasPath & Tmp & ".act", vbNormal) Then
104         tmp2 = InputBox("¿Dias?", "Ingrese cantidad de días")

106         If IsNumeric(tmp2) Then
108             Call DonadorTiempo(Tmp, tmp2)
            Else
110             MsgBox ("Cantidad invalida")

            End If

        Else
112         MsgBox ("La cuenta no existe")

        End If

        
        Exit Sub

addtimeDonador_Click_Err:
114     Call RegistrarError(Err.Number, Err.Description, "frmMain.addtimeDonador_Click", Erl)
116     Resume Next
        
End Sub

Private Sub Auditoria_Timer()

        On Error GoTo errhand

        'Static centinelSecs As Byte

        'centinelSecs = centinelSecs + 1

        'If centinelSecs = 5 Then
        'Every 5 seconds, we try to call the player's attention so it will report the code.
        ' Call modCentinela.CallUserAttention
    
        ' centinelSecs = 0
        'End If

100     Call PasarSegundo 'sistema de desconexion de 10 segs
102     Call PurgarScroll

104     Call PurgarOxigeno

106     Call ActualizaStatsES

        Exit Sub

errhand:

108     Call LogError("Error en Timer Auditoria. Err: " & Err.Description & " - " & Err.Number)

110     Resume Next

End Sub

Private Sub auxSocket_Connect()

    'auxSocket.SendData "{""header"":{""action"":""LoadUser""},""data"":{""accountId"":1}}"
    
End Sub

Private Sub auxSocket_DataArrival(ByVal bytesTotal As Long)
        
        On Error GoTo auxSocket_DataArrival_Err
    
        
    
        Dim strData As String
    
        ' Recibimos la info.
100     Call auxSocket.GetData(strData)
    
        ' Si no llegó nada, nos vamos alv.
102     If Len(strData) = 0 Then Exit Sub
    
        ' Parseamos el JSON que recibimo.
        Dim response As Object
104     Set response = mod_JSON.parse(strData)
    
106     Select Case response.Item("header").Item("action")
    
            Case "LoadUser"
                'Call MsgBox(response!data)
            
        End Select
    
108     End
    
        
        Exit Sub

auxSocket_DataArrival_Err:
110     Call RegistrarError(Err.Number, Err.Description, "frmMain.auxSocket_DataArrival", Erl)

        
End Sub

Private Sub CerrarYForzarActualizar_Click()
    On Error GoTo Command4_Click_Err

        If MsgBox("¿Está seguro que desea guardar, forzar actualización a los usuarios y cerrar?", vbYesNo, "Confirmación") = vbNo Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor » Cerrando servidor y lanzando nuevo parche.", FontTypeNames.FONTTYPE_PROMEDIO_MENOR))

        Call ForzarActualizar
100     Call GuardarUsuarios
102     Call EcharPjsNoPrivilegiados

104     GuardarYCerrar = True
106     Unload frmMain
        
        Exit Sub

Command4_Click_Err:
108     Call RegistrarError(Err.Number, Err.Description, "frmMain.CerrarYForzarActualizar", Erl)
110     Resume Next
End Sub

Private Sub Invasion_Timer()


    On Error GoTo Handler
        Dim i As Integer

        ' **********************************
        ' **********  Invasiones  **********
        ' **********************************
100     For i = 1 To UBound(Invasiones)
102         With Invasiones(i)
                ' Aumentamos el contador para controlar cuando
                ' inicia la invasión o cuando debe terminar
104             .TimerInvasion = .TimerInvasion + 1
        
106             If .Activa Then
                    ' Chequeamos si el evento debe terminar
108                 If .TimerInvasion >= .Duracion Then
110                     Call FinalizarInvasion(i)
                
                    Else
                        ' Descripción del evento
112                     .TimerRepetirDesc = .TimerRepetirDesc + 1
    
114                     If .TimerRepetirDesc >= .RepetirDesc Then
116                         Call MensajeGlobal(.Desc, FontTypeNames.FONTTYPE_New_Eventos)
118                         .TimerRepetirDesc = 0
                        End If
                    End If
            
                ' Si no está activa, chequeamos si debemos iniciarla
120             ElseIf .Intervalo > 0 Then
122                 If .TimerInvasion >= .Intervalo Then
124                     Call IniciarInvasion(i)
    
                    ' Si no está activa ni hay que iniciar, chequeamos si hay que avisar que se acerca el evento
126                 ElseIf .TimerInvasion >= .Intervalo - .AvisarTiempo Then
128                     .TimerRepetirAviso = .TimerRepetirAviso - 1
    
130                     If .TimerRepetirAviso <= 0 Then
132                         Call MensajeGlobal(.aviso, FontTypeNames.FONTTYPE_New_Eventos)
134                         .TimerRepetirAviso = .RepetirAviso
                        End If
                    End If
                End If
        
            End With
        Next
        Exit Sub
    
Handler:
136     Call RegistrarError(Err.Number, Err.Description, "frmMain.Invasion_Timer")
138     Resume Next
    
        ' **********************************
End Sub

' WyroX: Comprobamos cada 10 segundos, porque no es necesaria tanta precisión
Private Sub TiempoRetos_Timer()

    On Error GoTo Handler
    
        Dim IntervaloTimerRetosEnSegundos As Integer
100     IntervaloTimerRetosEnSegundos = TiempoRetos.Interval * 0.001
    
        Dim Sala As Integer
102     For Sala = 1 To Retos.TotalSalas
        
104         With Retos.Salas(Sala)
106             If .EnUso Then
108                 .TiempoRestante = .TiempoRestante - IntervaloTimerRetosEnSegundos
                
110                 If .TiempoRestante <= 0 Then
112                     Call FinalizarReto(Sala, True)
                    End If
                
114                 If .TiempoItems > 0 Then
116                     .TiempoItems = .TiempoItems - IntervaloTimerRetosEnSegundos
118                     If .TiempoItems <= 0 Then Call TerminarTiempoAgarrarItems(Sala)
                    End If
                End If
            End With

        Next
        Exit Sub
    
Handler:
120     Call RegistrarError(Err.Number, Err.Description, "frmMain.TiempoRetos_Timer")
122     Resume Next
    
End Sub

Private Sub TimerGuardarUsuarios_Timer()

    On Error GoTo Handler
    
        ' Guardar usuarios (solo si pasó el tiempo mínimo para guardar)
        Dim UserIndex As Integer, UserGuardados As Integer

100     For UserIndex = 1 To LastUser
    
102         With UserList(UserIndex)

104             If .flags.UserLogged Then
106                 If GetTickCount - .Counters.LastSave > IntervaloGuardarUsuarios Then
                
108                     Call SaveUser(UserIndex)
                    
110                     UserGuardados = UserGuardados + 1
                    
112                     If UserGuardados >= LimiteGuardarUsuarios Then Exit For
    
                    End If
    
                End If
        
            End With

        Next
    
        Exit Sub
    
Handler:
114     Call RegistrarError(Err.Number, Err.Description, "frmMain.TimreGuardarUsuarios_Timer")
116     Resume Next
    
End Sub

Private Sub Minuto_Timer()

        On Error GoTo ErrHandler

        'fired every minute
        Static minutos          As Long

        Static MinutosLatsClean As Long

        Dim i                   As Integer

        Dim num                 As Long

100     MinsRunning = MinsRunning + 1

102     If MinsRunning = 60 Then
104         horas = horas + 1

106         If horas = 24 Then
108             Call SaveDayStats
110             DayStats.MaxUsuarios = 0
112             DayStats.segundos = 0
114             DayStats.Promedio = 0
        
116             horas = 0
        
            End If

118         MinsRunning = 0

        End If
    
120     minutos = minutos + 1

        '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
122     Call ModAreas.AreasOptimizacion
        '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

        'Actualizamos el centinela
124     Call modCentinela.PasarMinutoCentinela

126     If MinutosLatsClean >= 15 Then
128         MinutosLatsClean = 0
130         Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
        Else
132         MinutosLatsClean = MinutosLatsClean + 1
        End If

134     Call PurgarPenas

136     If IdleLimit > 0 Then
138         Call CheckIdleUser
        End If

        '<<<<<-------- Log the number of users online ------>>>
        Dim n As Integer

140     n = FreeFile()
142     Open App.Path & "\logs\numusers.log" For Output Shared As n
144     Print #n, NumUsers
146     Close #n
        '<<<<<-------- Log the number of users online ------>>>

        Exit Sub
ErrHandler:
148     Call LogError("Error en Timer Minuto " & Err.Number & ": " & Err.Description)

150     Resume Next

End Sub

Private Sub CMDDUMP_Click()
        
        On Error GoTo CMDDUMP_Click_Err
    
        

        

        Dim i As Integer

100     For i = 1 To MaxUsers
102         Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).name & " UserLogged: " & UserList(i).flags.UserLogged)
104     Next i

106     Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

        
        Exit Sub

CMDDUMP_Click_Err:
108     Call RegistrarError(Err.Number, Err.Description, "frmMain.CMDDUMP_Click", Erl)

        
End Sub

Private Sub Command1_Click()
        
        On Error GoTo Command1_Click_Err
        
100     Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(BroadMsg.Text))

        
        Exit Sub

Command1_Click_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmMain.Command1_Click", Erl)
104     Resume Next
        
End Sub

Public Sub InitMain(ByVal f As Byte)
        
        On Error GoTo InitMain_Err
        

100     If f = 1 Then
102         Call mnuSystray_Click
        Else
104         Call frmMain.Show
        End If

        
        Exit Sub

InitMain_Err:
106     Call RegistrarError(Err.Number, Err.Description, "frmMain.InitMain", Erl)
108     Resume Next
        
End Sub

Private Sub Command10_Click()
        
        On Error GoTo Command10_Click_Err
        
100     Call GuardarUsuarios

        
        Exit Sub

Command10_Click_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmMain.Command10_Click", Erl)
104     Resume Next
        
End Sub

Private Sub Command11_Click()
        
        On Error GoTo Command11_Click_Err
        
100     Call LoadSini

        
        Exit Sub

Command11_Click_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmMain.Command11_Click", Erl)
104     Resume Next
        
End Sub

Private Sub Command12_Click()
        
        On Error GoTo Command12_Click_Err
        
100     Call LoadConfiguraciones

        
        Exit Sub

Command12_Click_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmMain.Command12_Click", Erl)
104     Resume Next
        
End Sub

Private Sub Command13_Click()
        
        On Error GoTo Command13_Click_Err
        
100     Call LoadBalance

        
        Exit Sub

Command13_Click_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmMain.Command13_Click", Erl)
104     Resume Next
        
End Sub

Private Sub Command2_Click()
        
        On Error GoTo Command2_Click_Err
        
100     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor » " & BroadMsg.Text, FontTypeNames.FONTTYPE_SERVER))

        
        Exit Sub

Command2_Click_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmMain.Command2_Click", Erl)
104     Resume Next
        
End Sub

Private Sub Command4_Click()
        
        On Error GoTo Command4_Click_Err

        If MsgBox("¿Está seguro que desea guardar y cerrar?", vbYesNo, "Confirmación") = vbNo Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor » Cerrando servidor.", FontTypeNames.FONTTYPE_PROMEDIO_MENOR))

100     Call GuardarUsuarios
102     Call EcharPjsNoPrivilegiados

104     GuardarYCerrar = True
106     Unload frmMain

        
        Exit Sub

Command4_Click_Err:
108     Call RegistrarError(Err.Number, Err.Description, "frmMain.Command4_Click", Erl)
110     Resume Next
        
End Sub

Private Sub Command5_Click()
        'Dim tem As String
        'tem = InputBox("Ingreste clave")
        'MsgBox SDesencriptar(tem)
        
        On Error GoTo Command5_Click_Err
        

100     Call GuardarRanking

        
        Exit Sub

Command5_Click_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmMain.Command5_Click", Erl)
104     Resume Next
        
End Sub

Private Sub Command6_Click()
        
        On Error GoTo Command6_Click_Err
        
100     Call LoadIntervalos

        
        Exit Sub

Command6_Click_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmMain.Command6_Click", Erl)
104     Resume Next
        
End Sub

Private Sub Command7_Click()
        
        On Error GoTo Command7_Click_Err
        
100     Call loadAdministrativeUsers

        
        Exit Sub

Command7_Click_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmMain.Command7_Click", Erl)
104     Resume Next
        
End Sub

Private Sub Command8_Click()
        
        On Error GoTo Command8_Click_Err
        
100     Call LoadOBJData
102     Call LoadPesca
104     Call LoadRecursosEspeciales
106     Call LoadRangosFaccion
108     Call LoadRecompensasFaccion


        Exit Sub

Command8_Click_Err:
110     Call RegistrarError(Err.Number, Err.Description, "frmMain.Command8_Click", Erl)
112     Resume Next
        
End Sub

Private Sub Command9_Click()
        
        On Error GoTo Command9_Click_Err
        
100     Call CargaNpcsDat(True)

        
        Exit Sub

Command9_Click_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmMain.Command9_Click", Erl)
104     Resume Next
        
End Sub

Private Sub EstadoTimer_Timer()
        
        On Error GoTo EstadoTimer_Timer_Err
    
        

    

100     Call GetHoraActual
    
        Dim i As Long

102     For i = 1 To Baneos.Count

104         If Baneos(i).FechaLiberacion <= Now Then
106             Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor » Se ha concluido la sentencia de ban para " & Baneos(i).Name & ".", FontTypeNames.FONTTYPE_SERVER))
108             Call ChangeBan(Baneos(i).name, 0)
110             Call Baneos.Remove(i)
112             Call SaveBans

            End If

        Next

114     For i = 1 To Donadores.Count

116         If Donadores(i).FechaExpiracion <= Now Then
118             Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor » Se ha concluido el tiempo de donador para " & Donadores(i).Name & ".", FontTypeNames.FONTTYPE_SERVER))
120             Call ChangeDonador(Donadores(i).name, 0)
122             Call Donadores.Remove(i)
124             Call SaveDonadores

            End If

        Next

126     Select Case frmMain.lblhora.Caption

            Case "0:00:00"
128             HoraEvento = 0

130         Case "1:00:00"
132             HoraEvento = 1

134         Case "2:00:00"
136             HoraEvento = 2

138         Case "3:00:00"
140             HoraEvento = 3

142         Case "4:00:00"
144             HoraEvento = 4

146         Case "5:00:00"
148             HoraEvento = 5

150         Case "6:00:00"
152             HoraEvento = 6

154         Case "7:00:00"
156             HoraEvento = 7

158         Case "8:00:00"
160             HoraEvento = 8

162         Case "9:00:00"
164             HoraEvento = 9

166         Case "10:00:00"
168             HoraEvento = 10

170         Case "11:00:00"
172             HoraEvento = 11

174         Case "12:00:00"
176             HoraEvento = 12

178         Case "13:00:00"
180             HoraEvento = 13

182         Case "14:00:00"
184             HoraEvento = 14

186         Case "15:00:00"
188             HoraEvento = 15

190         Case "16:00:00"
192             HoraEvento = 16

194         Case "17:00:00"
196             HoraEvento = 17

198         Case "18:00:00"
200             HoraEvento = 18

202         Case "19:00:00"
204             HoraEvento = 19

206         Case "20:00:00"
208             HoraEvento = 20

210         Case "21:00:00"
212             HoraEvento = 21

214         Case "22:00:00"
216             HoraEvento = 22

218         Case "23:00:00"
220             HoraEvento = 23

222         Case Else
                Exit Sub

        End Select

224     Call CheckEvento(HoraEvento)

        
        Exit Sub

EstadoTimer_Timer_Err:
226     Call RegistrarError(Err.Number, Err.Description, "frmMain.EstadoTimer_Timer", Erl)

        
End Sub

Private Sub Evento_Timer()
        
        On Error GoTo Evento_Timer_Err
        
100     TiempoRestanteEvento = TiempoRestanteEvento - 1

102     If TiempoRestanteEvento = 0 Then
104         Call FinalizarEvento

        End If

        
        Exit Sub

Evento_Timer_Err:
106     Call RegistrarError(Err.Number, Err.Description, "frmMain.Evento_Timer", Erl)
108     Resume Next
        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
        On Error GoTo Form_MouseMove_Err
    
        

        
   
100     If Not Visible Then

102         Select Case X \ Screen.TwipsPerPixelX
                
                Case WM_LBUTTONDBLCLK
104                 WindowState = vbNormal
106                 Visible = True

                    Dim hProcess As Long

108                 GetWindowThreadProcessId hwnd, hProcess
110                 AppActivate hProcess

112             Case WM_RBUTTONUP
114                 hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
116                 PopupMenu mnuPopUp

118                 If hHook Then
120                     UnhookWindowsHookEx hHook
122                     hHook = 0
                    End If


            End Select

        End If
   
        
        Exit Sub

Form_MouseMove_Err:
124     Call RegistrarError(Err.Number, Err.Description, "frmMain.Form_MouseMove", Erl)

        
End Sub

Public Sub QuitarIconoSystray()
        
        On Error GoTo QuitarIconoSystray_Err
    
        

        

        'Borramos el icono del systray
        Dim i   As Integer
        Dim nid As NOTIFYICONDATA

100     nid = setNOTIFYICONDATA(frmMain.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

102     i = Shell_NotifyIconA(NIM_DELETE, nid)

        
        Exit Sub

QuitarIconoSystray_Err:
104     Call RegistrarError(Err.Number, Err.Description, "frmMain.QuitarIconoSystray", Erl)

        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        
        On Error GoTo Form_QueryUnload_Err
    
        
100     If GuardarYCerrar Then Exit Sub
102     If MsgBox("¿Deseas FORZAR el CIERRE del servidor?" & vbNewLine & vbNewLine & "Ten en cuenta que ES POSIBLE PIERDAS DATOS!", vbYesNo, "¡FORZAR CIERRE!") = vbNo Then
104         Cancel = True
        End If
    
        
        Exit Sub

Form_QueryUnload_Err:
106     Call RegistrarError(Err.Number, Err.Description, "frmMain.Form_QueryUnload", Erl)

        
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
        On Error GoTo Form_Unload_Err
    
        

100     Call CerrarServidor

        
        Exit Sub

Form_Unload_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmMain.Form_Unload", Erl)

        
End Sub

Private Sub GameTimer_Timer()

        Dim iUserIndex   As Long
        Dim bEnviarAyS   As Boolean
    
        On Error GoTo HayError
    
        '<<<<<< Procesa eventos de los usuarios >>>>>>
100     For iUserIndex = 1 To MaxUsers 'LastUser

102         With UserList(iUserIndex)

                'Conexion activa?
104             If .ConnID <> -1 Then
                    '¿User valido?
                    
106                 If .ConnIDValida And .flags.UserLogged Then
                    
                        '[Alejo-18-5]
110                     bEnviarAyS = False
                    
112                     .NumeroPaquetesPorMiliSec = 0
                    
114                     Call DoTileEvents(iUserIndex, .Pos.Map, .Pos.X, .Pos.Y)

116                     If .flags.Muerto = 0 Then
                        
                            'Efectos en mapas
118                         If (.flags.Privilegios And PlayerType.user) <> 0 Then
120                             Call EfectoLava(iUserIndex)
122                             Call EfectoFrio(iUserIndex)
                            End If

124                         If .flags.Meditando Then Call DoMeditar(iUserIndex)
126                         If .flags.Envenenado <> 0 Then Call EfectoVeneno(iUserIndex)
128                         If .flags.Ahogandose <> 0 Then Call EfectoAhogo(iUserIndex)
130                         If .flags.Incinerado <> 0 Then Call EfectoIncineramiento(iUserIndex)
132                         If .flags.Mimetizado <> 0 Then Call EfectoMimetismo(iUserIndex)
                        
134                         If .flags.AdminInvisible <> 1 Then
136                             If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)
                            End If
                        
138                         If .NroMascotas > 0 Then Call TiempoInvocacion(iUserIndex)
                        
140                         Call HambreYSed(iUserIndex, bEnviarAyS)
                            Call EfectoStamina(iUserIndex)
                                    
256                         If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)
                        
                        Else
258                         If .flags.Traveling <> 0 Then Call TravelingEffect(iUserIndex)
                                                
                        End If 'Muerto

                    Else 'no esta logeado?
                        'Inactive players will be removed!
260                     .Counters.IdleCount = .Counters.IdleCount + 1
                    
                        'El intervalo cambia según si envió el primer paquete
262                     If .Counters.IdleCount > IIf(.flags.FirstPacket, TimeoutEsperandoLoggear, TimeoutPrimerPaquete) Then
264                         Call CloseSocket(iUserIndex)
                        End If

                    End If 'UserLogged

                End If

            End With

266     Next iUserIndex

        Exit Sub

HayError:
268     LogError ("Error en GameTimer: " & Err.Description & " UserIndex = " & iUserIndex)

End Sub

Private Sub HoraFantasia_Timer()
        
        On Error GoTo HoraFantasia_Timer_Err
        

100     If Lloviendo Then
102         Label6.Caption = "Lloviendo"
        Else
104         Label6.Caption = "No llueve"

        End If

106     If ServidorNublado Then
108         Label7.Caption = "Nublado"
        Else
110         Label7.Caption = "Sin nubes"

        End If

112     frmMain.Label4.Caption = GetTimeFormated
        
        Exit Sub

HoraFantasia_Timer_Err:
114     Call RegistrarError(Err.Number, Err.Description, "frmMain.HoraFantasia_Timer", Erl)
116     Resume Next
        
End Sub

Private Sub LimpiezaTimer_Timer()
        
        On Error GoTo LimpiezaTimer_Timer_Err
    
        

100     Call LimpiarItemsViejos
        
        
        Exit Sub

LimpiezaTimer_Timer_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmMain.LimpiezaTimer_Timer", Erl)

        
End Sub

Private Sub loadcredit_Click()
        
        On Error GoTo loadcredit_Click_Err
        

        Dim Tmp  As String

        Dim tmp2 As String

100     Tmp = InputBox("¿Cuenta?", "Ingrese la cuenta")

102     If FileExist(CuentasPath & Tmp & ".act", vbNormal) Then
104         tmp2 = InputBox("¿Cantidad?", "Ingrese cantidad de creditos a agregar")

106         If IsNumeric(tmp2) Then
108             Call AgregarCreditosDonador(Tmp, CLng(tmp2))
            Else
110             MsgBox ("Cantidad invalida")

            End If

        Else
112         MsgBox ("La cuenta no existe")

        End If

        
        Exit Sub

loadcredit_Click_Err:
114     Call RegistrarError(Err.Number, Err.Description, "frmMain.loadcredit_Click", Erl)
116     Resume Next
        
End Sub

Private Sub mnuCerrar_Click()
        
        On Error GoTo mnuCerrar_Click_Err
        

100     If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then

            Dim f
102         For Each f In Forms
104             Unload f
            Next

        End If

        
        Exit Sub

mnuCerrar_Click_Err:
106     Call RegistrarError(Err.Number, Err.Description, "frmMain.mnuCerrar_Click", Erl)
108     Resume Next
        
End Sub

Private Sub mnusalir_Click()
        
        On Error GoTo mnusalir_Click_Err
        
100     Call mnuCerrar_Click

        
        Exit Sub

mnusalir_Click_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmMain.mnusalir_Click", Erl)
104     Resume Next
        
End Sub

Public Sub mnuMostrar_Click()
        
        On Error GoTo mnuMostrar_Click_Err
    
        

        

100     WindowState = vbNormal
102     Form_MouseMove 0, 0, 7725, 0

        
        Exit Sub

mnuMostrar_Click_Err:
104     Call RegistrarError(Err.Number, Err.Description, "frmMain.mnuMostrar_Click", Erl)

        
End Sub

Private Sub KillLog_Timer()
        
        On Error GoTo KillLog_Timer_Err
    
        

    

100     If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\connect.log"
102     If FileExist(App.Path & "\logs\haciendo.log", vbNormal) Then Kill App.Path & "\logs\haciendo.log"
104     If FileExist(App.Path & "\logs\stats.log", vbNormal) Then Kill App.Path & "\logs\stats.log"
106     If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
108     If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
110     If Not FileExist(App.Path & "\logs\nokillwsapi.txt") Then
112         If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then Kill App.Path & "\logs\wsapi.log"
        End If

        
        Exit Sub

KillLog_Timer_Err:
114     Call RegistrarError(Err.Number, Err.Description, "frmMain.KillLog_Timer", Erl)

        
End Sub

Private Sub mnuServidor_Click()
        
        On Error GoTo mnuServidor_Click_Err
        
100     frmServidor.Visible = True

        
        Exit Sub

mnuServidor_Click_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmMain.mnuServidor_Click", Erl)
104     Resume Next
        
End Sub

Private Sub mnuSystray_Click()
        
        On Error GoTo mnuSystray_Click_Err
        

        Dim i   As Integer
        Dim S   As String
        Dim nid As NOTIFYICONDATA

100     S = "ARGENTUM-ONLINE"
102     nid = setNOTIFYICONDATA(frmMain.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
104     i = Shell_NotifyIconA(NIM_ADD, nid)
    
106     If WindowState <> vbMinimized Then WindowState = vbMinimized

108     Visible = False

        
        Exit Sub

mnuSystray_Click_Err:
110     Call RegistrarError(Err.Number, Err.Description, "frmMain.mnuSystray_Click", Erl)
112     Resume Next
        
End Sub

Private Sub packetResend_Timer()

    On Error GoTo Handler

        'If there is anything to be sent, we send it
        Dim i As Long
100     For i = 1 To LastUser
102         If UserList(i).ConnIDValida Then
104             Call FlushBuffer(i)
            End If
        Next
    
        Exit Sub
    
Handler:
106     Call RegistrarError(Err.Number, Err.Description, "frmMain.packetResend_Timer")
108     Resume Next
    
End Sub

Private Sub SubastaTimer_Timer()
        
        On Error GoTo SubastaTimer_Timer_Err
        

        'Si ya paso un minuto y todavia no hubo oferta, avisamos que se cancela en un minuto
100     If Subasta.TiempoRestanteSubasta = 240 And Subasta.HuboOferta = False Then
102         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Quedan 4 minuto(s) para finalizar la subasta! Escribe /SUBASTA para mas información. La subasta será cancelada si no hay ofertas en el próximo minuto.", FontTypeNames.FONTTYPE_SUBASTA))
104         Subasta.MinutosDeSubasta = 4
106         Subasta.PosibleCancelo = True

        End If
    
        'Si ya pasaron dos minutos y no hubo ofertas, cancelamos la subasta
108     If Subasta.TiempoRestanteSubasta = 180 And Subasta.HuboOferta = False Then
110         Subasta.HaySubastaActiva = False
112         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Subasta cancelada por falta de ofertas.", FontTypeNames.FONTTYPE_SUBASTA))
            'Devolver item antes de resetear datos
114         Call DevolverItem
            Exit Sub

        End If

116     If Subasta.PosibleCancelo = True Then
118         Subasta.TiempoRestanteSubasta = Subasta.TiempoRestanteSubasta - 1

        End If
    
120     If Subasta.TiempoRestanteSubasta > 0 And Subasta.PosibleCancelo = False Then
122         If Subasta.TiempoRestanteSubasta = 240 Then
124             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Quedan 4 minuto(s) para finalizar la subasta! Escribe /SUBASTA para mas información.", FontTypeNames.FONTTYPE_SUBASTA))
126             Subasta.MinutosDeSubasta = "4"

            End If
        
128         If Subasta.TiempoRestanteSubasta = 180 Then
130             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Quedan 3 minuto(s) para finalizar la subasta! Escribe /SUBASTA para mas información.", FontTypeNames.FONTTYPE_SUBASTA))
132             Subasta.MinutosDeSubasta = "3"

            End If

134         If Subasta.TiempoRestanteSubasta = 120 Then
136             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Quedan 2 minuto(s) para finalizar la subasta! Escribe /SUBASTA para mas información.", FontTypeNames.FONTTYPE_SUBASTA))
138             Subasta.MinutosDeSubasta = "2"

            End If

140         If Subasta.TiempoRestanteSubasta = 60 Then
142             Subasta.MinutosDeSubasta = "1"
144             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Quedan 1 minuto(s) para finalizar la subasta! Escribe /SUBASTA para mas información.", FontTypeNames.FONTTYPE_SUBASTA))

            End If

146         Subasta.TiempoRestanteSubasta = Subasta.TiempoRestanteSubasta - 1

        End If
    
148     If Subasta.TiempoRestanteSubasta = 1 Then
150         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡La subasta a terminado! El ganador fue: " & Subasta.Comprador, FontTypeNames.FONTTYPE_SUBASTA))
152         Call FinalizarSubasta

        End If

        
        Exit Sub

SubastaTimer_Timer_Err:
154     Call RegistrarError(Err.Number, Err.Description, "frmMain.SubastaTimer_Timer", Erl)
156     Resume Next
        
End Sub

Private Sub TIMER_AI_Timer()

        On Error GoTo ErrorHandler

        Dim NpcIndex As Long
        Dim Mapa     As Integer
    
        Dim X        As Integer
        Dim Y        As Integer

        'Barrin 29/9/03
100     If Not haciendoBK And Not EnPausa Then

            'Update NPCs
102         For NpcIndex = 1 To LastNPC
            
104             With NpcList(NpcIndex)
            
106                 If .flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
                
108                     If .NPCtype = DummyTarget Then
                            ' Regenera vida después de X tiempo sin atacarlo
110                         If .Stats.MinHp < .Stats.MaxHp Then
112                             .Contadores.UltimoAtaque = .Contadores.UltimoAtaque - 1
                            
114                             If .Contadores.UltimoAtaque <= 0 Then
116                                 Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageTextOverChar(.Stats.MaxHp - .Stats.MinHp, .Char.CharIndex, vbGreen))
118                                 .Stats.MinHp = .Stats.MaxHp
                                End If
                            End If

                        Else
                            'Usamos AI si hay algun user en el mapa
120                         Mapa = .Pos.Map
                        
122                         If .flags.Paralizado > 0 Then Call EfectoParalisisNpc(NpcIndex)
124                         If .flags.Inmovilizado > 0 Then Call EfectoInmovilizadoNpc(NpcIndex)
                        
126                         If Mapa > 0 Then
                                'Emancu: Vamos a probar si el server se degrada moviendo TODOS los npc, con o sin users.
128                             'If MapInfo(Mapa).NumUsers > 0 Or NpcList(NpcIndex).NPCtype = eNPCType.GuardiaNpc Then
    
130                                 If IntervaloPermiteMoverse(NpcIndex) Then
                                        
                                        'Si NO es pretoriano...
                                        If .NPCtype = eNPCType.Pretoriano Then
132                                         Call ClanPretoriano(.ClanIndex).PerformPretorianAI(NpcIndex)
                                    
                                        Else '... si es pretoriano.
136                                         Call NpcAI(NpcIndex)
                                        
                                        End If
                                        
                                        
                                    End If
    
                                'End If
    
                            End If

                        End If

                    End If
            
                End With

138         Next NpcIndex

        End If

        Exit Sub

ErrorHandler:
140     Call LogError("Error en TIMER_AI_Timer " & NpcList(NpcIndex).name & " mapa:" & NpcList(NpcIndex).Pos.Map)
142     Call MuereNpc(NpcIndex, 0)

End Sub

Private Sub TimerMeteorologia_Timer()
        'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor > Timer de lluvia en :" & TimerMeteorologico, FontTypeNames.FONTTYPE_SERVER))
        
        On Error GoTo TimerMeteorologia_Timer_Err
        

100     If TimerMeteorologico > 7 Then
102         TimerMeteorologico = TimerMeteorologico - 1
            Exit Sub

        End If

104     If TimerMeteorologico = 7 Then
106         ProbabilidadNublar = RandomNumber(1, 3)

108         If ProbabilidadNublar = 1 Then
110             IntensidadDeNubes = RandomNumber(10, 45)
112             ServidorNublado = True
                'Enviar Nubes a todos
114             Nieblando = True
116             ServidorNublado = True
118             Call SendData(SendTarget.ToAll, 0, PrepareMessageNieblandoToggle(IntensidadDeNubes))
                ' Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor > Empezaron las nubes con intensidad: " & IntensidadDeNubes & "%.", FontTypeNames.FONTTYPE_SERVER))
120             Call AgregarAConsola("Servidor > Empezaron las nubes")
            
122             TimerMeteorologico = TimerMeteorologico - 1
            Else
124             ServidorNublado = False
                ' Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor > Tranquilo, no hay nubes ni va a llover.", FontTypeNames.FONTTYPE_SERVER))
126             Call AgregarAConsola("Servidor >Tranquilo, no hay nubes ni va a llover.")
128             Call ResetMeteo
                Exit Sub

            End If

        End If

130     If TimerMeteorologico < 7 And TimerMeteorologico > 3 Then
132         TimerMeteorologico = TimerMeteorologico - 1
            'Enviar Truenos y rayos
134         Truenos.Enabled = True
            'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor > Envio un truenito para que te asustes.", FontTypeNames.FONTTYPE_SERVER))
136         Call AgregarAConsola("Servidor >Truenos y nubes activados.")
            Exit Sub

        End If

138     If TimerMeteorologico = 3 Then
140         ProbabilidadLLuvia = RandomNumber(1, 5)

142         If ProbabilidadLLuvia = 1 Then
                'Envia Lluvia
144             Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(404, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno
146             Call SendData(SendTarget.ToAll, 0, PrepareMessageFlashScreen(&HD254D6, 250)) 'Rayo
148             Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
150             Nebando = True
        
152             Call SendData(SendTarget.ToAll, 0, PrepareMessageNevarToggle())
                '  Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor > LLuvia lluvia y mas lluvia!", FontTypeNames.FONTTYPE_SERVER))
154             Call AgregarAConsola("Servidor >Lloviendo.")
156             Lloviendo = True
158             TimerMeteorologico = TimerMeteorologico - 1
            Else
160             Nieblando = False
162             Call SendData(SendTarget.ToAll, 0, PrepareMessageNieblandoToggle(IntensidadDeNubes))
164             Call AgregarAConsola("Servidor >Truenos y nubes desactivados.")
                ' Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor > Tranquilo, las nubes se fueron.", FontTypeNames.FONTTYPE_SERVER))
166             Lloviendo = False
168             ServidorNublado = False
170             Truenos.Enabled = False
172             Call ResetMeteo
                Exit Sub

            End If

        End If

174     If TimerMeteorologico < 3 And TimerMeteorologico > 0 Then

176         TimerMeteorologico = TimerMeteorologico - 1
            Exit Sub

        End If

178     If TimerMeteorologico = 0 Then
            'dejar de llover y sacar nubes
180         Nieblando = False
182         Call SendData(SendTarget.ToAll, 0, PrepareMessageNieblandoToggle(IntensidadDeNubes))
184         Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        
186         Call SendData(SendTarget.ToAll, 0, PrepareMessageNevarToggle())
            ' Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor > Se acabo la lluvia señores.", FontTypeNames.FONTTYPE_SERVER))
188         Call AgregarAConsola("Servidor >Lluvia desactivada.")
190         Lloviendo = False
192         Truenos.Enabled = False
194         Nebando = False
196         Call ResetMeteo
            Exit Sub

        End If

        Exit Sub

        
        Exit Sub

TimerMeteorologia_Timer_Err:
198     Call RegistrarError(Err.Number, Err.Description, "frmMain.TimerMeteorologia_Timer", Erl)
200     Resume Next
        
End Sub

Private Sub TimerRespawn_Timer()

        On Error GoTo ErrorHandler

        Dim NpcIndex As Long

        'Update NPCs
100     For NpcIndex = 1 To MaxRespawn
            'Debug.Print RespawnList(NpcIndex).name
102         If RespawnList(NpcIndex).flags.NPCActive Then  'Nos aseguramos que este muerto
104             If RespawnList(NpcIndex).Contadores.IntervaloRespawn > 0 Then
106                 RespawnList(NpcIndex).Contadores.IntervaloRespawn = RespawnList(NpcIndex).Contadores.IntervaloRespawn - 1
                Else
108                 RespawnList(NpcIndex).flags.NPCActive = False

110                 If RespawnList(NpcIndex).InformarRespawn = 1 Then
112                     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(RespawnList(NpcIndex).name & " ha vuelto a este mundo.", FontTypeNames.FONTTYPE_EXP))
114                     Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(257, NO_3D_SOUND, NO_3D_SOUND)) 'Para evento de respwan
                        
                        'Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(246, NO_3D_SOUND, NO_3D_SOUND)) 'Para evento de respwan
                    End If

116                 Call ReSpawnNpc(RespawnList(NpcIndex))

                End If

            End If

118     Next NpcIndex

        Exit Sub

ErrorHandler:
120     Call LogError("Error en TIMER_RESPAWN " & NpcList(NpcIndex).name & " mapa:" & NpcList(NpcIndex).Pos.Map)
122     Call MuereNpc(NpcIndex, 0)

End Sub

Private Sub tPiqueteC_Timer()

        On Error GoTo ErrHandler

        Static segundos As Integer

        Dim NuevaA      As Boolean

        Dim NuevoL      As Boolean

        Dim GI          As Integer

100     segundos = segundos + 6

        Dim i As Long

102     For i = 1 To LastUser

104         If UserList(i).flags.UserLogged Then
106             If MapData(UserList(i).Pos.Map, UserList(i).Pos.X, UserList(i).Pos.Y).trigger = eTrigger.ANTIPIQUETE Then
108                 UserList(i).Counters.PiqueteC = UserList(i).Counters.PiqueteC + 1
                    'Call WriteConsoleMsg(i, "Estás obstruyendo la via pública, muévete o serás encarcelado!!!", FontTypeNames.FONTTYPE_INFO)
                
                    'WyroX: Le empiezo a avisar a partir de los 18 segundos, para no spamear
110                 If UserList(i).Counters.PiqueteC > 3 Then
112                     Call WriteLocaleMsg(i, "70", FontTypeNames.FONTTYPE_INFO)
                    End If
            
114                 If UserList(i).Counters.PiqueteC > 10 Then
116                     UserList(i).Counters.PiqueteC = 0
                        'Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
                        'WyroX: En vez de encarcelarlo, lo sacamos del juego.
                        'Ojo! No sé si se puede abusar de esto para evitar los 10 segundos al salir
118                     Call WriteDisconnect(i)
120                     Call CloseSocket(i)
                    End If

                Else

122                 If UserList(i).Counters.PiqueteC > 0 Then UserList(i).Counters.PiqueteC = 0

                End If

124             If segundos >= 18 Then
126                 If segundos >= 18 Then UserList(i).Counters.Pasos = 0

                End If

            End If
    
128     Next i

130     If segundos >= 18 Then segundos = 0

        Exit Sub

ErrHandler:
132     Call LogError("Error en tPiqueteC_Timer " & Err.Number & ": " & Err.Description)

End Sub

Private Sub Truenos_Timer()
        
        On Error GoTo Truenos_Timer_Err
        

        Dim Enviar    As Byte

        Dim TruenoWav As Integer

100     Enviar = RandomNumber(1, 15)

        Dim Duracion As Long

102     If Enviar < 8 Then
104         TruenoWav = 399 + Enviar

106         If TruenoWav = 404 Then TruenoWav = 406
108         Duracion = RandomNumber(80, 250)
110         Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(TruenoWav, NO_3D_SOUND, NO_3D_SOUND))
112         Call SendData(SendTarget.ToAll, 0, PrepareMessageFlashScreen(&HEFEECB, Duracion))
        
        End If

        
        Exit Sub

Truenos_Timer_Err:
114     Call RegistrarError(Err.Number, Err.Description, "frmMain.Truenos_Timer", Erl)
116     Resume Next
        
End Sub

Private Sub UptimeTimer_Timer()
        
        On Error GoTo UptimeTimer_Timer_Err
        
100     SERVER_UPTIME = SERVER_UPTIME + 1

        
        Exit Sub

UptimeTimer_Timer_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmMain.UptimeTimer_Timer", Erl)
104     Resume Next
        
End Sub
