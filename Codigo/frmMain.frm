VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RevolucionAo Server by Ladder -  Pablo Mercavides 2008-2017"
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
   WindowState     =   1  'Minimized
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
      Begin VB.Timer grabado 
         Interval        =   60000
         Left            =   1080
         Top             =   3360
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Timer TimerRespawn 
         Interval        =   1000
         Left            =   240
         Top             =   4800
      End
      Begin VB.Timer EstadoTimer 
         Interval        =   1000
         Left            =   240
         Top             =   3480
      End
      Begin VB.Timer Evento 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   240
         Top             =   3960
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
   Begin VB.Timer AutoSave 
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
    hWnd As Long
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

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, Id As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
        
        On Error GoTo setNOTIFYICONDATA_Err
        

        Dim nidTemp As NOTIFYICONDATA

100     nidTemp.cbSize = Len(nidTemp)
102     nidTemp.hWnd = hWnd
104     nidTemp.uID = Id
106     nidTemp.uFlags = flags
108     nidTemp.uCallbackMessage = CallbackMessage
110     nidTemp.hIcon = Icon
112     nidTemp.szTip = Tip & Chr$(0)

114     setNOTIFYICONDATA = nidTemp

        
        Exit Function

setNOTIFYICONDATA_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.setNOTIFYICONDATA", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "frmMain.CheckIdleUser", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "frmMain.addtimeDonador_Click", Erl)
        Resume Next
        
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

    Call LogError("Error en Timer Auditoria. Err: " & Err.description & " - " & Err.Number)

    Resume Next

End Sub

Private Sub AutoSave_Timer()

    On Error GoTo ErrHandler

    'fired every minute
    Static minutos          As Long

    Static MinutosLatsClean As Long

    Static MinsPjesSave     As Long

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
    
    MinsPjesSave = MinsPjesSave + 1
    
    If MinsPjesSave >= 10 Then
        Call GuardarUsuarios
        MinsPjesSave = 0
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
    Call LogError("Error en TimerAutoSave " & Err.Number & ": " & Err.description)

    Resume Next

End Sub

Private Sub CMDDUMP_Click()

    On Error Resume Next

    Dim i As Integer

    For i = 1 To MaxUsers
        Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).name & " UserLogged: " & UserList(i).flags.UserLogged)
    Next i

    Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub

Private Sub Command1_Click()
        
        On Error GoTo Command1_Click_Err
        
100     Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(BroadMsg.Text))

        
        Exit Sub

Command1_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.Command1_Click", Erl)
        Resume Next
        
End Sub

Public Sub InitMain(ByVal f As Byte)
        
        On Error GoTo InitMain_Err
        

100     If f = 1 Then
102         Call mnuSystray_Click

        Else
104         frmMain.Show

        End If

        
        Exit Sub

InitMain_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.InitMain", Erl)
        Resume Next
        
End Sub

Private Sub Command10_Click()
        
        On Error GoTo Command10_Click_Err
        
100     Call GuardarUsuarios

        
        Exit Sub

Command10_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.Command10_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command11_Click()
        
        On Error GoTo Command11_Click_Err
        
100     Call LoadSini

        
        Exit Sub

Command11_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.Command11_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command12_Click()
        
        On Error GoTo Command12_Click_Err
        
100     Call LoadConfiguraciones

        
        Exit Sub

Command12_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.Command12_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command13_Click()
        
        On Error GoTo Command13_Click_Err
        
100     Call LoadBalance

        
        Exit Sub

Command13_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.Command13_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command2_Click()
        
        On Error GoTo Command2_Click_Err
        
100     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & BroadMsg.Text, FontTypeNames.FONTTYPE_SERVER))

        
        Exit Sub

Command2_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.Command2_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command4_Click()
        
        On Error GoTo Command4_Click_Err
        
100     Call GuardarUsuarios
102     Call EcharPjsNoPrivilegiados

        GuardarYCerrar = True
104     Unload frmMain

        
        Exit Sub

Command4_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.Command4_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command5_Click()
        'Dim tem As String
        'tem = InputBox("Ingreste clave")
        'MsgBox SDesencriptar(tem)
        
        On Error GoTo Command5_Click_Err
        

100     Call GuardarRanking

        
        Exit Sub

Command5_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.Command5_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command6_Click()
        
        On Error GoTo Command6_Click_Err
        
100     Call LoadIntervalos

        
        Exit Sub

Command6_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.Command6_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command7_Click()
        
        On Error GoTo Command7_Click_Err
        
100     Call loadAdministrativeUsers

        
        Exit Sub

Command7_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.Command7_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command8_Click()
        
        On Error GoTo Command8_Click_Err
        
100     Call LoadOBJData
102     Call LoadPesca
104     Call LoadRecursosEspeciales

        
        Exit Sub

Command8_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.Command8_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command9_Click()
        
        On Error GoTo Command9_Click_Err
        
100     Call CargaNpcsDat(True)

        
        Exit Sub

Command9_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.Command9_Click", Erl)
        Resume Next
        
End Sub

Private Sub EstadoTimer_Timer()

    On Error Resume Next

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

End Sub

Private Sub Evento_Timer()
        
        On Error GoTo Evento_Timer_Err
        
100     TiempoRestanteEvento = TiempoRestanteEvento - 1

102     If TiempoRestanteEvento = 0 Then
104         Call FinalizarEvento

        End If

        
        Exit Sub

Evento_Timer_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.Evento_Timer", Erl)
        Resume Next
        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
   
    If Not Visible Then

        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True

                Dim hProcess As Long

                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess

            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp

                If hHook Then
                    UnhookWindowsHookEx hHook
                    hHook = 0
                End If


        End Select

    End If
   
End Sub

Public Sub QuitarIconoSystray()

    On Error Resume Next

    'Borramos el icono del systray
    Dim i   As Integer
    Dim nid As NOTIFYICONDATA

    nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

    i = Shell_NotifyIconA(NIM_DELETE, nid)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If GuardarYCerrar Then Exit Sub
    If MsgBox("¿Deseas FORZAR el CIERRE del servidor?" & vbNewLine & vbNewLine & "Ten en cuenta que ES POSIBLE PIERDAS DATOS!", vbYesNo, "¡FORZAR CIERRE!") = vbNo Then
        Cancel = True
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call CerrarServidor

End Sub

Private Sub GameTimer_Timer()

    Dim iUserIndex   As Long
    Dim bEnviarStats As Boolean
    Dim bEnviarAyS   As Boolean
    
    On Error GoTo hayerror
    
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
                                    
                                    If Not .flags.Descansar Then
                                        
                                        'No esta descansando
                                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)

                                        If bEnviarStats Then
                                            Call WriteUpdateHP(iUserIndex)
                                            bEnviarStats = False

                                        End If

                                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)

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

                                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)

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
                                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar * 4)
                                    Call WriteUpdateSta(iUserIndex)

                                End If
                                
                            Else

                                If Not .flags.Descansar Then
                                    'No esta descansando
                                    
                                    Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)

                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False

                                    End If

                                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)

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

                                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)

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

hayerror:
    LogError ("Error en GameTimer: " & Err.description & " UserIndex = " & iUserIndex)

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

114     frmMain.Label4.Caption = GetTimeFormated
        
        Exit Sub

HoraFantasia_Timer_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.HoraFantasia_Timer", Erl)
        Resume Next
        
End Sub

Private Sub LimpiezaTimer_Timer()

    Call LimpiarItemsViejos
        
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
        Call RegistrarError(Err.Number, Err.description, "frmMain.loadcredit_Click", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "frmMain.mnuCerrar_Click", Erl)
        Resume Next
        
End Sub

Private Sub mnusalir_Click()
        
        On Error GoTo mnusalir_Click_Err
        
100     Call mnuCerrar_Click

        
        Exit Sub

mnusalir_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.mnusalir_Click", Erl)
        Resume Next
        
End Sub

Public Sub mnuMostrar_Click()

    On Error Resume Next

    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0

End Sub

Private Sub KillLog_Timer()

    On Error Resume Next

    If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\connect.log"
    If FileExist(App.Path & "\logs\haciendo.log", vbNormal) Then Kill App.Path & "\logs\haciendo.log"
    If FileExist(App.Path & "\logs\stats.log", vbNormal) Then Kill App.Path & "\logs\stats.log"
    If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
    If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
    If Not FileExist(App.Path & "\logs\nokillwsapi.txt") Then
        If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then Kill App.Path & "\logs\wsapi.log"
    End If

End Sub

Private Sub mnuServidor_Click()
        
        On Error GoTo mnuServidor_Click_Err
        
100     frmServidor.Visible = True

        
        Exit Sub

mnuServidor_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.mnuServidor_Click", Erl)
        Resume Next
        
End Sub

Private Sub mnuSystray_Click()
        
        On Error GoTo mnuSystray_Click_Err
        

        Dim i   As Integer
        Dim S   As String
        Dim nid As NOTIFYICONDATA

100     S = "ARGENTUM-ONLINE"
102     nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
104     i = Shell_NotifyIconA(NIM_ADD, nid)
    
106     If WindowState <> vbMinimized Then WindowState = vbMinimized

108     Visible = False

        
        Exit Sub

mnuSystray_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.mnuSystray_Click", Erl)
        Resume Next
        
End Sub

Private Sub npcataca_Timer()

    On Error Resume Next

    Dim npc As Integer

    'For npc = 1 To LastNPC
    '  Npclist(npc).CanAttack = 1
    'Next npc

End Sub

Private Sub packetResend_Timer()

    'If there is anything to be sent, we send it
    Dim i As Long
    For i = 1 To LastUser
        If UserList(i).ConnIDValida Then
            Call FlushBuffer(i)
        End If
    Next
    
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
        Call RegistrarError(Err.Number, Err.description, "frmMain.SubastaTimer_Timer", Erl)
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
                
                    If .flags.Paralizado = 1 Then
                        Call EfectoParalisisNpc(NpcIndex)

                    Else
                        
                        'Usamos AI si hay algun user en el mapa
                        If .flags.Inmovilizado = 1 Then
                            Call EfectoParalisisNpc(NpcIndex)
                        End If
                        
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
146             Call SendData(SendTarget.ToAll, 0, PrepareMessageEfectToScreen(&HD254D6, 250)) 'Rayo
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
        Call RegistrarError(Err.Number, Err.description, "frmMain.TimerMeteorologia_Timer", Erl)
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
                Call WriteLocaleMsg(i, "70", FontTypeNames.FONTTYPE_INFO)
            
                If UserList(i).Counters.PiqueteC > 15 Then
                    UserList(i).Counters.PiqueteC = 0
                    Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)

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
    Call LogError("Error en tPiqueteC_Timer " & Err.Number & ": " & Err.description)

End Sub





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''USO DEL CONTROL TCPSERV'''''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#If UsarQueSocket = 3 Then

Private Sub TCPServ_Eror(ByVal Numero As Long, ByVal Descripcion As String)
    Call LogError("TCPSERVER SOCKET ERROR: " & Numero & "/" & Descripcion)
End Sub

Private Sub TCPServ_NuevaConn(ByVal Id As Long)
On Error GoTo errorHandlerNC

    ESCUCHADAS = ESCUCHADAS + 1

    
    Dim i As Integer
    
    Dim NewIndex As Integer
    NewIndex = NextOpenUser
    
    If NewIndex <= MaxUsers Then
        'call logindex(NewIndex, "******> Accept. ConnId: " & ID)
        
        TCPServ.SetDato Id, NewIndex
        
        If aDos.MaxConexiones(TCPServ.GetIP(Id)) Then
            Call aDos.RestarConexion(TCPServ.GetIP(Id))
            Call ResetUserSlot(NewIndex)
            Exit Sub
        End If

        If NewIndex > LastUser Then LastUser = NewIndex

        UserList(NewIndex).ConnID = Id
        UserList(NewIndex).ip = TCPServ.GetIP(Id)
        UserList(NewIndex).ConnIDValida = True
        Set UserList(NewIndex).CommandsBuffer = New CColaArray
        
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = TCPServ.GetIP(Id) Then
                Call ResetUserSlot(NewIndex)
                Exit Sub
            End If
        Next i

    Else
        Call CloseSocket(NewIndex, True)
        LogCriticEvent ("NEWINDEX > MAXUSERS. IMPOSIBLE ALOCATEAR SOCKETS")
    End If

Exit Sub

errorHandlerNC:
Call LogError("TCPServer::NuevaConexion " & Err.description)
End Sub

Private Sub TCPServ_Close(ByVal Id As Long, ByVal MiDato As Long)
    On Error GoTo eh
    '' No cierro yo el socket. El on_close lo cierra por mi.
    'call logindex(MiDato, "******> Remote Close. ConnId: " & ID & " Midato: " & MiDato)
    Call CloseSocket(MiDato, False)
Exit Sub
eh:
    Call LogError("Ocurrio un error en el evento TCPServ_Close. ID/miDato:" & Id & "/" & MiDato)
End Sub

Private Sub TCPServ_Read(ByVal Id As Long, Datos As Variant, ByVal Cantidad As Long, ByVal MiDato As Long)
On Error GoTo errorh

With UserList(MiDato)
    Datos = StrConv(StrConv(Datos, vbUnicode), vbFromUnicode)
    
    Call .incomingData.WriteASCIIStringFixed(Datos)
    
    If .ConnID <> -1 Then
        Call HandleIncomingData(MiDato)
    Else
        Exit Sub
    End If
End With

Exit Sub

errorh:
Call LogError("Error socket read: " & MiDato & " dato:" & RD & " userlogged: " & UserList(MiDato).flags.UserLogged & " connid:" & UserList(MiDato).ConnID & " ID Parametro" & Id & " error:" & Err.description)

End Sub

#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FIN  USO DEL CONTROL TCPSERV'''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Truenos_Timer()
        
        On Error GoTo Truenos_Timer_Err
        

        Dim Enviar    As Byte

        Dim TruenoWav As Integer

100     Enviar = RandomNumber(1, 15)

        Dim duracion As Long

102     If Enviar < 8 Then
104         TruenoWav = 399 + Enviar

106         If TruenoWav = 404 Then TruenoWav = 406
108         duracion = RandomNumber(80, 250)
110         Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(TruenoWav, NO_3D_SOUND, NO_3D_SOUND))
112         Call SendData(SendTarget.ToAll, 0, PrepareMessageEfectToScreen(&HEFEECB, duracion))
        
        End If

        
        Exit Sub

Truenos_Timer_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.Truenos_Timer", Erl)
        Resume Next
        
End Sub

Private Sub UptimeTimer_Timer()
        
        On Error GoTo UptimeTimer_Timer_Err
        
100     SERVER_UPTIME = SERVER_UPTIME + 1

        
        Exit Sub

UptimeTimer_Timer_Err:
        Call RegistrarError(Err.Number, Err.description, "frmMain.UptimeTimer_Timer", Erl)
        Resume Next
        
End Sub
