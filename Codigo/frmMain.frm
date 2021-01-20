VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Servidor Argentum 20 - Basado en RAO de Pablo Mercavides"
   ClientHeight    =   6315
   ClientLeft      =   1950
   ClientTop       =   1695
   ClientWidth     =   6915
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
   ScaleHeight     =   6315
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerGuardarUsuarios 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   2640
      Top             =   3120
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Guardar y cerrar"
      Height          =   975
      Left            =   5160
      TabIndex        =   33
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Recargar intervalos.ini"
      Height          =   495
      Left            =   5160
      TabIndex        =   32
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Recargar Balance.dat"
      Height          =   495
      Left            =   5160
      TabIndex        =   31
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Recargar configuracion.ini"
      Height          =   495
      Left            =   5160
      TabIndex        =   30
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Recargar Server.ini"
      Height          =   495
      Left            =   5160
      TabIndex        =   29
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Guardar Usuarios"
      Height          =   495
      Left            =   5160
      TabIndex        =   28
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Recargar Npcs"
      Height          =   495
      Left            =   5160
      TabIndex        =   27
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Recargar Objetos"
      Height          =   495
      Left            =   5160
      TabIndex        =   26
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Recargar Administradores"
      Height          =   495
      Left            =   5160
      TabIndex        =   25
      Top             =   240
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
   Begin VB.Timer npcataca 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   3120
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
      Begin VB.Timer t_ColaAPI 
         Interval        =   50
         Left            =   3120
         Top             =   1200
      End
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
   Begin VB.Menu mnuConsolaAPI 
      Caption         =   "Consola API"
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuConsolaAPI_Popup 
         Caption         =   "&Consola API"
      End
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

    Call PasarSegundo 'sistema de desconexion de 10 segs
    Call PurgarScroll

    Call PurgarOxigeno

    Call ActualizaStatsES

    Exit Sub

errhand:

    Call LogError("Error en Timer Auditoria. Err: " & Err.Description & " - " & Err.Number)

    Resume Next

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

Private Sub mnuConsolaAPI_Click()
    frmAPISocket.Show vbModeless
End Sub

Private Sub mnuConsolaAPI_Popup_Click()
    Call mnuConsolaAPI_Click
End Sub

Private Sub t_ColaAPI_Timer()

    If API.packetResend.Count = 0 Then Exit Sub
    
    With frmAPISocket
    
        Select Case .Socket.State
        
            Case sckConnected
            
                'Iteramos la cola y mandamos todo.
                Do While (Not API.packetResend.IsEmpty)
                    Call .Socket.SendData(API.packetResend.Pop)
                Loop
                
                Debug.Print "API: Enviado!"
                
                Exit Sub
            
            Case sckClosed
                Call .Connect
                Debug.Print "API: El socket estaba cerrado! Reconectando..."
                
            Case sckError
                Call .Socket.CloseSck
                Call .Connect
                Debug.Print "API: Error en el socket! Reconectando..."
                
        End Select
    
    End With
    
End Sub

Private Sub TimerGuardarUsuarios_Timer()

On Error GoTo Handler
    
    ' Guardar usuarios (solo si pasó el tiempo mínimo para guardar)
    Dim UserIndex As Integer, UserGuardados As Integer

    For UserIndex = 1 To LastUser
    
        With UserList(UserIndex)

            If .flags.UserLogged Then
                If GetTickCount - .Counters.LastSave > IntervaloGuardarUsuarios Then
                
                    Call SaveUser(UserIndex)
                    
                    UserGuardados = UserGuardados + 1
                    
                    If UserGuardados >= LimiteGuardarUsuarios Then Exit For
    
                End If
    
            End If
        
        End With

    Next
    
    Exit Sub
    
Handler:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.TimreGuardarUsuarios_Timer")
    Resume Next
    
End Sub

Private Sub Minuto_Timer()

    On Error GoTo ErrHandler

    'fired every minute
    Static minutos          As Long

    Static MinutosLatsClean As Long

    Dim i                   As Integer

    Dim num                 As Long

    MinsRunning = MinsRunning + 1

    If MinsRunning = 60 Then
        horas = horas + 1

        If horas = 24 Then
            Call SaveDayStats
            DayStats.MaxUsuarios = 0
            DayStats.segundos = 0
            DayStats.Promedio = 0
        
            horas = 0
        
        End If

        MinsRunning = 0

    End If
    
    minutos = minutos + 1

    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    Call ModAreas.AreasOptimizacion
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

    'Actualizamos el centinela
    Call modCentinela.PasarMinutoCentinela

    If MinutosLatsClean >= 15 Then
        MinutosLatsClean = 0
        Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
    Else
        MinutosLatsClean = MinutosLatsClean + 1

    End If

    Call PurgarPenas

    If IdleLimit > 0 Then
        Call CheckIdleUser

    End If

    '<<<<<-------- Log the number of users online ------>>>
    Dim n As Integer

    n = FreeFile()
    Open App.Path & "\logs\numusers.log" For Output Shared As n
    Print #n, NumUsers
    Close #n
    '<<<<<-------- Log the number of users online ------>>>

    Exit Sub
ErrHandler:
    Call LogError("Error en Timer Minuto " & Err.Number & ": " & Err.Description)

    Resume Next

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
        
100     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & BroadMsg.Text, FontTypeNames.FONTTYPE_SERVER))

        
        Exit Sub

Command2_Click_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmMain.Command2_Click", Erl)
104     Resume Next
        
End Sub

Private Sub Command4_Click()
        
        On Error GoTo Command4_Click_Err
        
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

        
        Exit Sub

Command8_Click_Err:
106     Call RegistrarError(Err.Number, Err.Description, "frmMain.Command8_Click", Erl)
108     Resume Next
        
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
    
        

    

    Call GetHoraActual
    
    Dim i As Long

    For i = 1 To Baneos.Count

        If Baneos(i).FechaLiberacion <= Now Then
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> Se ha concluido la sentencia de ban para " & Baneos(i).name & ".", FontTypeNames.FONTTYPE_SERVER))
            Call ChangeBan(Baneos(i).name, 0)
            Call Baneos.Remove(i)
            Call SaveBans

        End If

    Next

    For i = 1 To Donadores.Count

        If Donadores(i).FechaExpiracion <= Now Then
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> Se ha concluido el tiempo de donador para " & Donadores(i).name & ".", FontTypeNames.FONTTYPE_SERVER))
            Call ChangeDonador(Donadores(i).name, 0)
            Call Donadores.Remove(i)
            Call SaveDonadores

        End If

    Next

    Select Case frmMain.lblhora.Caption

        Case "0:00:00"
            HoraEvento = 0

        Case "1:00:00"
            HoraEvento = 1

        Case "2:00:00"
            HoraEvento = 2

        Case "3:00:00"
            HoraEvento = 3

        Case "4:00:00"
            HoraEvento = 4

        Case "5:00:00"
            HoraEvento = 5

        Case "6:00:00"
            HoraEvento = 6

        Case "7:00:00"
            HoraEvento = 7

        Case "8:00:00"
            HoraEvento = 8

        Case "9:00:00"
            HoraEvento = 9

        Case "10:00:00"
            HoraEvento = 10

        Case "11:00:00"
            HoraEvento = 11

        Case "12:00:00"
            HoraEvento = 12

        Case "13:00:00"
            HoraEvento = 13

        Case "14:00:00"
            HoraEvento = 14

        Case "15:00:00"
            HoraEvento = 15

        Case "16:00:00"
            HoraEvento = 16

        Case "17:00:00"
            HoraEvento = 17

        Case "18:00:00"
            HoraEvento = 18

        Case "19:00:00"
            HoraEvento = 19

        Case "20:00:00"
            HoraEvento = 20

        Case "21:00:00"
            HoraEvento = 21

        Case "22:00:00"
            HoraEvento = 22

        Case "23:00:00"
            HoraEvento = 23

        Case Else
            Exit Sub

    End Select

    Call CheckEvento(HoraEvento)

        
    Exit Sub

EstadoTimer_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.EstadoTimer_Timer", Erl)

        
End Sub

Private Sub Evento_Timer()
        
    On Error GoTo Evento_Timer_Err
        
    TiempoRestanteEvento = TiempoRestanteEvento - 1

    If TiempoRestanteEvento = 0 Then
        Call FinalizarEvento

    End If

        
    Exit Sub

Evento_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Evento_Timer", Erl)
    Resume Next
        
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
    Dim bEnviarStats As Boolean
    Dim bEnviarAyS   As Boolean
    
    On Error GoTo HayError
    
    '<<<<<< Procesa eventos de los usuarios >>>>>>
    For iUserIndex = 1 To MaxUsers 'LastUser

        With UserList(iUserIndex)

            'Conexion activa?
            If .ConnID <> -1 Then
                '¿User valido?
                
                If .ConnIDValida And .flags.UserLogged Then
                    
                    '[Alejo-18-5]
                    bEnviarStats = False
                    bEnviarAyS = False
                    
                    .NumeroPaquetesPorMiliSec = 0
                    
                    Call DoTileEvents(iUserIndex, .Pos.Map, .Pos.X, .Pos.Y)

                    If .flags.Muerto = 0 Then
                        
                        'Efectos en mapas
                        If (.flags.Privilegios And PlayerType.user) <> 0 Then
                            Call EfectoLava(iUserIndex)
                            Call EfectoFrio(iUserIndex)
                        End If

                        If .flags.Meditando Then Call DoMeditar(iUserIndex)
                        If .flags.Envenenado <> 0 Then Call EfectoVeneno(iUserIndex)
                        If .flags.Ahogandose <> 0 Then Call EfectoAhogo(iUserIndex)
                        If .flags.Incinerado <> 0 Then Call EfectoIncineramiento(iUserIndex, False)
                        If .flags.Mimetizado <> 0 Then Call EfectoMimetismo(iUserIndex)
                        If .flags.AdminInvisible <> 1 Then
                            If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)
                        End If
                        
                        If .NroMascotas > 0 Then Call TiempoInvocacion(iUserIndex)
                        
                        Call HambreYSed(iUserIndex, bEnviarAyS)
                        
                        If .flags.Hambre = 0 And .flags.Sed = 0 Then
                            
                            If Lloviendo Then
                            
                                If Not Intemperie(iUserIndex) Then
                                    
                                    'No esta descansando
                                    If Not .flags.Descansar Then

                                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)

                                        If bEnviarStats Then
                                            Call WriteUpdateHP(iUserIndex)
                                            bEnviarStats = False
                                        End If
    
                                        If .flags.Desnudo = 0 Then
                                            Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                                        Else
                                            Call PierdeEnergia(iUserIndex, bEnviarStats, IntervaloPerderStamina)
                                        End If

                                        If bEnviarStats Then
                                            Call WriteUpdateSta(iUserIndex)
                                            bEnviarStats = False
                                        End If

                                    Else
                                        'esta descansando
                                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)

                                        If bEnviarStats Then
                                            Call WriteUpdateHP(iUserIndex)
                                            bEnviarStats = False
                                        End If

                                        If .flags.Desnudo = 0 Then
                                            Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                                        Else
                                            Call PierdeEnergia(iUserIndex, bEnviarStats, IntervaloPerderStamina * 2)
                                        End If

                                        If bEnviarStats Then
                                            Call WriteUpdateSta(iUserIndex)
                                            bEnviarStats = False
                                        End If

                                        'termina de descansar automaticamente
                                        If .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
                                            Call WriteRestOK(iUserIndex)
                                            Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                            .flags.Descansar = False
                                        End If
                                        
                                    End If

                                Else
                                    If .flags.Desnudo = 0 Then
                                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar * 4)
                                    Else
                                        Call PierdeEnergia(iUserIndex, bEnviarStats, IntervaloPerderStamina * 0.5)
                                    End If

                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False
                                    End If

                                End If
                                
                            Else

                                If Not .flags.Descansar Then
                                    'No esta descansando
                                    
                                    Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)

                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False
                                    End If

                                    If .flags.Desnudo = 0 Then
                                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                                    Else
                                        Call PierdeEnergia(iUserIndex, bEnviarStats, IntervaloPerderStamina)
                                    End If

                                    If bEnviarStats Then
                                        'borrar este
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    
                                Else
                                    'esta descansando
                                    
                                    Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)

                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False
                                    End If

                                    If .flags.Desnudo = 0 Then
                                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                                    Else
                                        Call PierdeEnergia(iUserIndex, bEnviarStats, IntervaloPerderStamina * 2)
                                    End If

                                    If bEnviarStats Then
                                        '  Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False
                                    End If

                                    'termina de descansar automaticamente
                                    If .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
                                        Call WriteRestOK(iUserIndex)
                                        Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                        .flags.Descansar = False
                                    End If
                                    
                                End If

                            End If

                        End If
                        
                        If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)
                        
                    Else
                        If .flags.Traveling <> 0 Then Call TravelingEffect(iUserIndex)
                                                
                    End If 'Muerto

                Else 'no esta logeado?
                    'Inactive players will be removed!
                    .Counters.IdleCount = .Counters.IdleCount + 1
                    
                    'El intervalo cambia según si envió el primer paquete
                    If .Counters.IdleCount > IIf(.flags.FirstPacket, TimeoutEsperandoLoggear, TimeoutPrimerPaquete) Then
                        Call CloseSocket(iUserIndex)
                    End If

                End If 'UserLogged

            End If

        End With

    Next iUserIndex

    Exit Sub

HayError:
    LogError ("Error en GameTimer: " & Err.Description & " UserIndex = " & iUserIndex)

End Sub

Private Sub HoraFantasia_Timer()
        
    On Error GoTo HoraFantasia_Timer_Err
        

    If Lloviendo Then
        Label6.Caption = "Lloviendo"
    Else
        Label6.Caption = "No llueve"

    End If

    If ServidorNublado Then
        Label7.Caption = "Nublado"
    Else
        Label7.Caption = "Sin nubes"

    End If

    frmMain.Label4.Caption = GetTimeFormated
        
    Exit Sub

HoraFantasia_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.HoraFantasia_Timer", Erl)
    Resume Next
        
End Sub

Private Sub LimpiezaTimer_Timer()
        
    On Error GoTo LimpiezaTimer_Timer_Err
    
        

    Call LimpiarItemsViejos
        
        
    Exit Sub

LimpiezaTimer_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.LimpiezaTimer_Timer", Erl)

        
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
    
        

    

    If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\connect.log"
    If FileExist(App.Path & "\logs\haciendo.log", vbNormal) Then Kill App.Path & "\logs\haciendo.log"
    If FileExist(App.Path & "\logs\stats.log", vbNormal) Then Kill App.Path & "\logs\stats.log"
    If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
    If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
    If Not FileExist(App.Path & "\logs\nokillwsapi.txt") Then
        If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then Kill App.Path & "\logs\wsapi.log"
    End If

        
    Exit Sub

KillLog_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.KillLog_Timer", Erl)

        
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
        Dim s   As String
        Dim nid As NOTIFYICONDATA

100     s = "ARGENTUM-ONLINE"
102     nid = setNOTIFYICONDATA(frmMain.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, s)
104     i = Shell_NotifyIconA(NIM_ADD, nid)
    
106     If WindowState <> vbMinimized Then WindowState = vbMinimized

108     Visible = False

        
        Exit Sub

mnuSystray_Click_Err:
110     Call RegistrarError(Err.Number, Err.Description, "frmMain.mnuSystray_Click", Erl)
112     Resume Next
        
End Sub

Private Sub npcataca_Timer()
    
    On Error GoTo npcataca_Timer_Err
    
    

    

    Dim npc As Integer

    'For npc = 1 To LastNPC
    '  NpcList(npc).CanAttack = 1
    'Next npc

    
    Exit Sub

npcataca_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.npcataca_Timer", Erl)

    
End Sub

Private Sub packetResend_Timer()

On Error GoTo Handler

    'If there is anything to be sent, we send it
    Dim i As Long
    For i = 1 To LastUser
        If UserList(i).ConnIDValida Then
            Call FlushBuffer(i)
        End If
    Next
    
    Exit Sub
    
Handler:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.packetResend_Timer")
    Resume Next
    
End Sub

Private Sub SubastaTimer_Timer()
        
    On Error GoTo SubastaTimer_Timer_Err
        

    'Si ya paso un minuto y todavia no hubo oferta, avisamos que se cancela en un minuto
    If Subasta.TiempoRestanteSubasta = 240 And Subasta.HuboOferta = False Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Quedan 4 minuto(s) para finalizar la subasta! Escribe /SUBASTA para mas información. La subasta será cancelada si no hay ofertas en el próximo minuto.", FontTypeNames.FONTTYPE_SUBASTA))
        Subasta.MinutosDeSubasta = 4
        Subasta.PosibleCancelo = True

    End If
    
    'Si ya pasaron dos minutos y no hubo ofertas, cancelamos la subasta
    If Subasta.TiempoRestanteSubasta = 180 And Subasta.HuboOferta = False Then
        Subasta.HaySubastaActiva = False
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Subasta cancelada por falta de ofertas.", FontTypeNames.FONTTYPE_SUBASTA))
        'Devolver item antes de resetear datos
        Call DevolverItem
        Exit Sub

    End If

    If Subasta.PosibleCancelo = True Then
        Subasta.TiempoRestanteSubasta = Subasta.TiempoRestanteSubasta - 1

    End If
    
    If Subasta.TiempoRestanteSubasta > 0 And Subasta.PosibleCancelo = False Then
        If Subasta.TiempoRestanteSubasta = 240 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Quedan 4 minuto(s) para finalizar la subasta! Escribe /SUBASTA para mas información.", FontTypeNames.FONTTYPE_SUBASTA))
            Subasta.MinutosDeSubasta = "4"

        End If
        
        If Subasta.TiempoRestanteSubasta = 180 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Quedan 3 minuto(s) para finalizar la subasta! Escribe /SUBASTA para mas información.", FontTypeNames.FONTTYPE_SUBASTA))
            Subasta.MinutosDeSubasta = "3"

        End If

        If Subasta.TiempoRestanteSubasta = 120 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Quedan 2 minuto(s) para finalizar la subasta! Escribe /SUBASTA para mas información.", FontTypeNames.FONTTYPE_SUBASTA))
            Subasta.MinutosDeSubasta = "2"

        End If

        If Subasta.TiempoRestanteSubasta = 60 Then
            Subasta.MinutosDeSubasta = "1"
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Quedan 1 minuto(s) para finalizar la subasta! Escribe /SUBASTA para mas información.", FontTypeNames.FONTTYPE_SUBASTA))

        End If

        Subasta.TiempoRestanteSubasta = Subasta.TiempoRestanteSubasta - 1

    End If
    
    If Subasta.TiempoRestanteSubasta = 1 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡La subasta a terminado! El ganador fue: " & Subasta.Comprador, FontTypeNames.FONTTYPE_SUBASTA))
        Call FinalizarSubasta

    End If

        
    Exit Sub

SubastaTimer_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.SubastaTimer_Timer", Erl)
    Resume Next
        
End Sub

Private Sub TIMER_AI_Timer()

    On Error GoTo ErrorHandler

    Dim NpcIndex As Long
    Dim Mapa     As Integer
    
    Dim X        As Integer
    Dim Y        As Integer

    'Barrin 29/9/03
    If Not haciendoBK And Not EnPausa Then

        'Update NPCs
        For NpcIndex = 1 To LastNPC
            
            With Npclist(NpcIndex)
            
                If .flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
                
                    If .NPCtype = DummyTarget Then
                        ' Regenera vida después de X tiempo sin atacarlo
                        If .Stats.MinHp < .Stats.MaxHp Then
                            .Contadores.UltimoAtaque = .Contadores.UltimoAtaque - 1
                            
                            If .Contadores.UltimoAtaque <= 0 Then
                                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageTextOverChar(.Stats.MaxHp - .Stats.MinHp, .Char.CharIndex, vbGreen))
                                .Stats.MinHp = .Stats.MaxHp
                            End If
                        End If
                
                    ElseIf .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
                        Call EfectoParalisisNpc(NpcIndex)

                    Else
                        'Usamos AI si hay algun user en el mapa
                        Mapa = .Pos.Map
                        
                        If Mapa > 0 Then
                            
                            If MapInfo(Mapa).NumUsers > 0 Then
    
                                If IntervaloPermiteMoverse(NpcIndex) Then
                                        
                                    'Si NO es pretoriano...
                                    If .NPCtype <> eNPCType.Pretoriano Then
                                        Call NPCAI(NpcIndex)
                                    
                                    Else '... si es pretoriano.
                                        Call ClanPretoriano(.ClanIndex).PerformPretorianAI(NpcIndex)
                                        
                                    End If
                                        
                                        
                                End If
    
                            End If
    
                        End If

                    End If

                End If
            
            End With

        Next NpcIndex

    End If

    Exit Sub

ErrorHandler:
    Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).name & " mapa:" & Npclist(NpcIndex).Pos.Map)
    Call MuereNpc(NpcIndex, 0)

End Sub

Private Sub TimerMeteorologia_Timer()
    'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor > Timer de lluvia en :" & TimerMeteorologico, FontTypeNames.FONTTYPE_SERVER))
        
    On Error GoTo TimerMeteorologia_Timer_Err
        

    If TimerMeteorologico > 7 Then
        TimerMeteorologico = TimerMeteorologico - 1
        Exit Sub

    End If

    If TimerMeteorologico = 7 Then
        ProbabilidadNublar = RandomNumber(1, 3)

        If ProbabilidadNublar = 1 Then
            IntensidadDeNubes = RandomNumber(10, 45)
            ServidorNublado = True
            'Enviar Nubes a todos
            Nieblando = True
            ServidorNublado = True
            Call SendData(SendTarget.ToAll, 0, PrepareMessageNieblandoToggle(IntensidadDeNubes))
            ' Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor > Empezaron las nubes con intensidad: " & IntensidadDeNubes & "%.", FontTypeNames.FONTTYPE_SERVER))
            Call AgregarAConsola("Servidor > Empezaron las nubes")
            
            TimerMeteorologico = TimerMeteorologico - 1
        Else
            ServidorNublado = False
            ' Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor > Tranquilo, no hay nubes ni va a llover.", FontTypeNames.FONTTYPE_SERVER))
            Call AgregarAConsola("Servidor >Tranquilo, no hay nubes ni va a llover.")
            Call ResetMeteo
            Exit Sub

        End If

    End If

    If TimerMeteorologico < 7 And TimerMeteorologico > 3 Then
        TimerMeteorologico = TimerMeteorologico - 1
        'Enviar Truenos y rayos
        Truenos.Enabled = True
        'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor > Envio un truenito para que te asustes.", FontTypeNames.FONTTYPE_SERVER))
        Call AgregarAConsola("Servidor >Truenos y nubes activados.")
        Exit Sub

    End If

    If TimerMeteorologico = 3 Then
        ProbabilidadLLuvia = RandomNumber(1, 5)

        If ProbabilidadLLuvia = 1 Then
            'Envia Lluvia
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(404, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno
            Call SendData(SendTarget.ToAll, 0, PrepareMessageFlashScreen(&HD254D6, 250)) 'Rayo
            Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
            Nebando = True
        
            Call SendData(SendTarget.ToAll, 0, PrepareMessageNevarToggle())
            '  Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor > LLuvia lluvia y mas lluvia!", FontTypeNames.FONTTYPE_SERVER))
            Call AgregarAConsola("Servidor >Lloviendo.")
            Lloviendo = True
            TimerMeteorologico = TimerMeteorologico - 1
        Else
            Nieblando = False
            Call SendData(SendTarget.ToAll, 0, PrepareMessageNieblandoToggle(IntensidadDeNubes))
            Call AgregarAConsola("Servidor >Truenos y nubes desactivados.")
            ' Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor > Tranquilo, las nubes se fueron.", FontTypeNames.FONTTYPE_SERVER))
            Lloviendo = False
            ServidorNublado = False
            Truenos.Enabled = False
            Call ResetMeteo
            Exit Sub

        End If

    End If

    If TimerMeteorologico < 3 And TimerMeteorologico > 0 Then

        TimerMeteorologico = TimerMeteorologico - 1
        Exit Sub

    End If

    If TimerMeteorologico = 0 Then
        'dejar de llover y sacar nubes
        Nieblando = False
        Call SendData(SendTarget.ToAll, 0, PrepareMessageNieblandoToggle(IntensidadDeNubes))
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageNevarToggle())
        ' Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor > Se acabo la lluvia señores.", FontTypeNames.FONTTYPE_SERVER))
        Call AgregarAConsola("Servidor >Lluvia desactivada.")
        Lloviendo = False
        Truenos.Enabled = False
        Nebando = False
        Call ResetMeteo
        Exit Sub

    End If

    Exit Sub

        
    Exit Sub

TimerMeteorologia_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.TimerMeteorologia_Timer", Erl)
    Resume Next
        
End Sub

Private Sub TimerRespawn_Timer()

    On Error GoTo ErrorHandler

    Dim NpcIndex As Long

    'Update NPCs
    For NpcIndex = 1 To MaxRespawn

        If RespawnList(NpcIndex).flags.NPCActive Then  'Nos aseguramos que este muerto
            If RespawnList(NpcIndex).Contadores.InvervaloRespawn <> 0 Then
                RespawnList(NpcIndex).Contadores.InvervaloRespawn = RespawnList(NpcIndex).Contadores.InvervaloRespawn - 1

                If RespawnList(NpcIndex).Contadores.InvervaloRespawn = 0 Then
                    RespawnList(NpcIndex).flags.NPCActive = False

                    If RespawnList(NpcIndex).InformarRespawn = 1 Then
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(RespawnList(NpcIndex).name & " ha regresado y está listo para enfrentarte.", FontTypeNames.FONTTYPE_EXP))
                        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(257, NO_3D_SOUND, NO_3D_SOUND)) 'Para evento de respwan
                        
                        'Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(246, NO_3D_SOUND, NO_3D_SOUND)) 'Para evento de respwan
                    End If

                    Call ReSpawnNpc(RespawnList(NpcIndex))

                End If

            End If

        End If

    Next NpcIndex

    Exit Sub

ErrorHandler:
    Call LogError("Error en TIMER_RESPAWN " & Npclist(NpcIndex).name & " mapa:" & Npclist(NpcIndex).Pos.Map)
    Call MuereNpc(NpcIndex, 0)

End Sub

Private Sub tPiqueteC_Timer()

    On Error GoTo ErrHandler

    Static segundos As Integer

    Dim NuevaA      As Boolean

    Dim NuevoL      As Boolean

    Dim GI          As Integer

    segundos = segundos + 6

    Dim i As Long

    For i = 1 To LastUser

        If UserList(i).flags.UserLogged Then
            If MapData(UserList(i).Pos.Map, UserList(i).Pos.X, UserList(i).Pos.Y).trigger = eTrigger.ANTIPIQUETE Then
                UserList(i).Counters.PiqueteC = UserList(i).Counters.PiqueteC + 1
                'Call WriteConsoleMsg(i, "Estás obstruyendo la via pública, muévete o serás encarcelado!!!", FontTypeNames.FONTTYPE_INFO)
                
                'WyroX: Le empiezo a avisar a partir de los 18 segundos, para no spamear
                If UserList(i).Counters.PiqueteC > 3 Then
                    Call WriteLocaleMsg(i, "70", FontTypeNames.FONTTYPE_INFO)
                End If
            
                If UserList(i).Counters.PiqueteC > 10 Then
                    UserList(i).Counters.PiqueteC = 0
                    'Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
                    'WyroX: En vez de encarcelarlo, lo sacamos del juego.
                    'Ojo! No sé si se puede abusar de esto para evitar los 10 segundos al salir
                    Call WriteDisconnect(i)
                    Call CloseSocket(i)
                End If

            Else

                If UserList(i).Counters.PiqueteC > 0 Then UserList(i).Counters.PiqueteC = 0

            End If

            'ustedes se preguntaran que hace esto aca?
            'bueno la respuesta es simple: el codigo de AO es una mierda y encontrar
            'todos los puntos en los cuales la alineacion puede cambiar es un dolor de
            'huevos, asi que lo controlo aca, cada 6 segundos, lo cual es razonable

            'GI = UserList(i).guildIndex
            ' If GI > 0 Then
            '  NuevaA = False
            ' NuevoL = False
            ' If Not modGuilds.m_ValidarPermanencia(i, True, NuevaA, NuevoL) Then
            '  Call WriteConsoleMsg(i, "Has sido expulsado del clan. ¡El clan ha sumado un punto de antifacción!", FontTypeNames.FONTTYPE_GUILD)
            ' End If
            'If NuevaA Then
            '   Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg("¡El clan ha pasado a tener alineación neutral!", FontTypeNames.FONTTYPE_GUILD))
            '   Call LogClanes("El clan cambio de alineacion!")
            'End If
            '  If NuevoL Then
            '   Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg("¡El clan tiene un nuevo líder!", FontTypeNames.FONTTYPE_GUILD))
            '  Call LogClanes("El clan tiene nuevo lider!")
            ' End If
            ' End If

            If segundos >= 18 Then
                If segundos >= 18 Then UserList(i).Counters.Pasos = 0

            End If

            

        End If
    
    Next i

    If segundos >= 18 Then segundos = 0

    Exit Sub

ErrHandler:
    Call LogError("Error en tPiqueteC_Timer " & Err.Number & ": " & Err.Description)

End Sub

Private Sub Truenos_Timer()
        
    On Error GoTo Truenos_Timer_Err
        

    Dim Enviar    As Byte

    Dim TruenoWav As Integer

    Enviar = RandomNumber(1, 15)

    Dim duracion As Long

    If Enviar < 8 Then
        TruenoWav = 399 + Enviar

        If TruenoWav = 404 Then TruenoWav = 406
        duracion = RandomNumber(80, 250)
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(TruenoWav, NO_3D_SOUND, NO_3D_SOUND))
        Call SendData(SendTarget.ToAll, 0, PrepareMessageFlashScreen(&HEFEECB, duracion))
        
    End If

        
    Exit Sub

Truenos_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.Truenos_Timer", Erl)
    Resume Next
        
End Sub

Private Sub UptimeTimer_Timer()
        
    On Error GoTo UptimeTimer_Timer_Err
        
    SERVER_UPTIME = SERVER_UPTIME + 1

        
    Exit Sub

UptimeTimer_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMain.UptimeTimer_Timer", Erl)
    Resume Next
        
End Sub
