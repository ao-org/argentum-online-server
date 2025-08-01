VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Servidor Argentum 20"
   ClientHeight    =   6255
   ClientLeft      =   1950
   ClientTop       =   1695
   ClientWidth     =   8595
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
   ScaleHeight     =   6255
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerBarco 
      Enabled         =   0   'False
      Interval        =   12000
      Left            =   7200
      Top             =   4560
   End
   Begin VB.Timer tControlHechizos 
      Left            =   4440
      Top             =   4800
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Recargar baneos.dat"
      Height          =   495
      Left            =   6840
      TabIndex        =   40
      Top             =   120
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   4440
      Top             =   1920
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Recargar Donadore"
      Height          =   495
      Left            =   1320
      TabIndex        =   39
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdDbControl 
      Caption         =   "DbControl"
      Height          =   495
      Left            =   3720
      TabIndex        =   38
      Top             =   720
      Width           =   1215
   End
   Begin VB.Timer t_Extraer 
      Left            =   4440
      Top             =   4200
   End
   Begin VB.Timer T_UsersOnline 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3600
      Top             =   4200
   End
   Begin VB.CommandButton CerrarYForzarActualizar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cerrar y forzar actualizacion"
      Height          =   495
      Left            =   5160
      MaskColor       =   &H000040C0&
      Style           =   1  'Graphical
      TabIndex        =   34
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
      TabIndex        =   31
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Recargar intervalos y rates"
      Height          =   495
      Left            =   5160
      TabIndex        =   30
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Recargar Balance.dat"
      Height          =   495
      Left            =   5160
      TabIndex        =   29
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Recargar configuracion.ini"
      Height          =   495
      Left            =   5160
      TabIndex        =   28
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Recargar Server.ini"
      Height          =   495
      Left            =   5160
      TabIndex        =   27
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Guardar Usuarios"
      Height          =   495
      Left            =   5160
      TabIndex        =   26
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Recargar Npcs"
      Height          =   495
      Left            =   5160
      TabIndex        =   25
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Recargar Objetos"
      Height          =   495
      Left            =   5160
      TabIndex        =   24
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Recargar Administradores"
      Height          =   495
      Left            =   5160
      TabIndex        =   23
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eventos"
      Height          =   6015
      Left            =   8880
      TabIndex        =   21
      Top             =   120
      Width           =   1575
      Begin VB.Label cuentas 
         Caption         =   "0"
         Height          =   375
         Left            =   240
         TabIndex        =   22
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
      TabIndex        =   19
      Top             =   4680
      Width           =   4935
      Begin VB.ListBox List1 
         Height          =   900
         Left            =   120
         TabIndex        =   20
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
      Caption         =   "Informacion general"
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4935
      Begin VB.CheckBox chkLogDbPerfomance 
         BackColor       =   &H80000016&
         Caption         =   "Log DB perfomance"
         Height          =   375
         Left            =   2640
         TabIndex        =   37
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Estabilidad:"
         Height          =   255
         Left            =   3360
         TabIndex        =   36
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Estabilidad 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         Height          =   210
         Left            =   4560
         TabIndex        =   35
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nublado"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lloviendo"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label3"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hora Fantasia Servidor:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblhora 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label3"
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hora Actual Servidor:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   240
         Width           =   1875
      End
   End
   Begin VB.Timer SubastaTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   4200
   End
   Begin VB.Timer Segundo 
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
      Top             =   3000
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
      TabIndex        =   13
      Top             =   3000
      Width           =   4935
      Begin VB.ListBox listaDePaquetes 
         Height          =   900
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label paquetesRecibidos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Paquetes Recibidos:"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
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
   Begin VB.Label Label9 
      Caption         =   "Basado en rao"
      Height          =   495
      Left            =   5280
      TabIndex        =   33
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   495
      Left            =   2880
      TabIndex        =   32
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
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
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

Public GuardarYCerrar As Boolean
Private tHechizosMinutesCounter As Byte

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private SERVER_UPTIME As Long

#If DIRECT_PLAY = 1 Then
Implements DirectPlay8Event
Private mfExit As Boolean
Private Enum MsgTypes
    Msg_NoOtherPlayers
    Msg_NumPlayers
    Msg_SendWave
End Enum

Private Sub Form_Load()
    Dim lCount As Long
    Dim dpn As DPN_SERVICE_PROVIDER_INFO
    dps.RegisterMessageHandler Me
    'First load our list of Service Providers into our box
    For lCount = 1 To dps.GetCountServiceProviders
        dpn = dps.GetServiceProvider(lCount)
        Debug.Print dpn.Name; " " & dpn.Guid
    Next
End Sub

Private Sub DirectPlay8Event_AddRemovePlayerGroup(ByVal lMsgID As Long, ByVal lPlayerID As Long, ByVal lGroupID As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_AppDesc(fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
    Debug.Print "DirectPlay8Event_AppDesc"
End Sub

Private Sub DirectPlay8Event_AsyncOpComplete(dpnotify As DxVBLibA.DPNMSG_ASYNC_OP_COMPLETE, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
    Debug.Print "DirectPlay8Event_AsyncOpComplete"
End Sub

Private Sub DirectPlay8Event_ConnectComplete(dpnotify As DxVBLibA.DPNMSG_CONNECT_COMPLETE, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
    Debug.Print "DirectPlay8Event_ConnectComplete"
End Sub

Private Sub DirectPlay8Event_CreateGroup(ByVal lGroupID As Long, ByVal lOwnerID As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_CreatePlayer(ByVal lPlayerID As Long, fRejectMsg As Boolean)
   Call modNetwork.CreatePlayer(lPlayerID, fRejectMsg)
End Sub

Private Sub DirectPlay8Event_DestroyGroup(ByVal lGroupID As Long, ByVal lReason As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_DestroyPlayer(ByVal lPlayerID As Long, ByVal lReason As Long, fRejectMsg As Boolean)
    Debug.Print "DirectPlay8Event_DestroyPlayer"
    Call modNetwork.DestroyPlayer(lPlayerID, lReason, fRejectMsg)
    
End Sub

Private Sub DirectPlay8Event_EnumHostsQuery(dpnotify As DxVBLibA.DPNMSG_ENUM_HOSTS_QUERY, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
     Debug.Print "DirectPlay8Event_EnumHostsQuery"
End Sub

Private Sub DirectPlay8Event_EnumHostsResponse(dpnotify As DxVBLibA.DPNMSG_ENUM_HOSTS_RESPONSE, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
     Debug.Print "DirectPlay8Event_EnumHostsResponse"
End Sub

Private Sub DirectPlay8Event_HostMigrate(ByVal lNewHostID As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
     Debug.Print "DirectPlay8Event_HostMigrate"
End Sub

Private Sub DirectPlay8Event_IndicateConnect(dpnotify As DxVBLibA.DPNMSG_INDICATE_CONNECT, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
    Debug.Print "DirectPlay8Event_IndicateConnect"
End Sub

Private Sub DirectPlay8Event_IndicatedConnectAborted(fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
     Debug.Print "DirectPlay8Event_IndicatedConnectAborted"
End Sub

Private Sub DirectPlay8Event_InfoNotify(ByVal lMsgID As Long, ByVal lNotifyID As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
    Debug.Print "DirectPlay8Event_InfoNotify"
End Sub

Private Sub DirectPlay8Event_Receive(dpnotify As DxVBLibA.DPNMSG_RECEIVE, fRejectMsg As Boolean)
    Call modNetwork.Receive(dpnotify, fRejectMsg)
End Sub

Private Sub DirectPlay8Event_SendComplete(dpnotify As DxVBLibA.DPNMSG_SEND_COMPLETE, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
  
End Sub

Private Sub DirectPlay8Event_TerminateSession(dpnotify As DxVBLibA.DPNMSG_TERMINATE_SESSION, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
    Debug.Print "DirectPlay8Event_TerminateSession"
End Sub
#End If



Private Function setNOTIFYICONDATA(hwnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
        
        On Error GoTo setNOTIFYICONDATA_Err
        

        Dim nidTemp As NOTIFYICONDATA

100     nidTemp.cbSize = Len(nidTemp)
102     nidTemp.hwnd = hwnd
104     nidTemp.uID = ID
106     nidTemp.uFlags = flags
108     nidTemp.uCallbackMessage = CallbackMessage
110     nidTemp.hIcon = Icon
112     nidTemp.szTip = Tip & Chr$(0)

114     setNOTIFYICONDATA = nidTemp

        
        Exit Function

setNOTIFYICONDATA_Err:
116     Call TraceError(Err.Number, Err.Description, "frmMain.setNOTIFYICONDATA", Erl)

        
End Function

Sub CheckIdleUser()
        
        On Error GoTo CheckIdleUser_Err
        

        Dim iUserIndex As Long
    
100     For iUserIndex = 1 To MaxUsers

            'Conexion activa? y es un usuario loggeado?
102         If UserList(iUserIndex).ConnectionDetails.ConnIDValida And UserList(iUserIndex).flags.UserLogged Then
                'Actualiza el contador de inactividad
104             UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1

106             If UserList(iUserIndex).Counters.IdleCount >= IdleLimit Then
108                 Call WriteShowMessageBox(iUserIndex, 1775, vbNullString) 'Msg1775=Demasiado tiempo inactivo. Has sido desconectado...

                    'mato los comercios seguros
110                 If IsValidUserRef(UserList(iUserIndex).ComUsu.DestUsu) Then
112                     If UserList(UserList(iUserIndex).ComUsu.DestUsu.ArrayIndex).flags.UserLogged Then
114                         If UserList(UserList(iUserIndex).ComUsu.DestUsu.ArrayIndex).ComUsu.DestUsu.ArrayIndex = iUserIndex Then
116                             Call WriteConsoleMsg(UserList(iUserIndex).ComUsu.DestUsu.ArrayIndex, PrepareMessageLocaleMsg(1844, vbNullString, e_FontTypeNames.FONTTYPE_TALK)) ' Msg1844=Comercio cancelado por el otro usuario.
118                             Call FinComerciarUsu(UserList(iUserIndex).ComUsu.DestUsu.ArrayIndex)
                            

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
126     Call TraceError(Err.Number, Err.Description, "frmMain.CheckIdleUser", Erl)

        
End Sub

Private Sub cmdDbControl_Click()
    frmDbControl.Show
End Sub

Private Sub Command3_Click()
    Call CargarDonadores
End Sub




Private Sub Segundo_Timer()

    On Error GoTo errhand
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    ' WyroX - Control de estabilidad del servidor
    Static LastTime As Currency
    Static Frequency As Currency
    Dim CurTime As Currency
    
    'Get the timer frequency
    If Frequency = 0 Then
        Call QueryPerformanceFrequency(Frequency)
    End If

    Call QueryPerformanceCounter(CurTime)

    If LastTime <> 0 Then
        Estabilidad.Caption = Round(Clamp(200 + (LastTime - CurTime) * 100 / Frequency, 0, 100), 1) & "%"
    End If

    LastTime = CurTime
    ' -----------------------------------

    Call PasarSegundo 'sistema de desconexion de 10 segs
    Call CheckDisconnectedUsers
    Call PerformTimeLimitCheck(PerformanceTimer, "Segundo_Timer", 100)
    Exit Sub

errhand:
    Call TraceError(Err.Number, Err.Description, "frmMain.Auditoria", Erl)
        
End Sub

Private Sub CerrarYForzarActualizar_Click()
    On Error GoTo Command4_Click_Err

100     If MsgBox("¿Está seguro que desea guardar, forzar actualización a los usuarios y cerrar?", vbYesNo, "Confirmación") = vbNo Then Exit Sub
        
102     Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1659, vbNullString, e_FontTypeNames.FONTTYPE_PROMEDIO_MENOR)) 'Msg1659=Servidor » Cerrando servidor y lanzando nuevo parche.

106     Call GuardarUsuarios
108     Call EcharPjsNoPrivilegiados

110     GuardarYCerrar = True
112     Unload frmMain
        
        Exit Sub

Command4_Click_Err:
114     Call TraceError(Err.Number, Err.Description, "frmMain.CerrarYForzarActualizar", Erl)

End Sub

Private Sub Invasion_Timer()


On Error GoTo Handler
    Dim i As Integer

    ' **********************************
    ' **********  Invasiones  **********
    ' **********************************
    For i = 1 To UBound(Invasiones)
        With Invasiones(i)
            ' Aumentamos el contador para controlar cuando
            ' inicia la invasión o cuando debe terminar
            .TimerInvasion = .TimerInvasion + 1

            If .Activa Then
                ' Chequeamos si el evento debe terminar
                If .TimerInvasion >= .Duracion Then
                    Call FinalizarInvasion(i)
                
                Else
                    ' Descripción del evento
                    .TimerRepetirDesc = .TimerRepetirDesc + 1
    
                    If .TimerRepetirDesc >= .RepetirDesc Then
                        Call MensajeGlobal(.Desc, e_FontTypeNames.FONTTYPE_New_Eventos)
                        .TimerRepetirDesc = 0
                    End If
                End If
            
            ' Si no está activa, chequeamos si debemos iniciarla
            ElseIf .Intervalo > 0 Then
                If .TimerInvasion >= .Intervalo Then
                    Call IniciarInvasion(i)
    
                ' Si no está activa ni hay que iniciar, chequeamos si hay que avisar que se acerca el evento
                ElseIf .TimerInvasion >= .Intervalo - .AvisarTiempo Then
                    .TimerRepetirAviso = .TimerRepetirAviso - 1
    
                    If .TimerRepetirAviso <= 0 Then
                        Call MensajeGlobal(.aviso, e_FontTypeNames.FONTTYPE_New_Eventos)
                        .TimerRepetirAviso = .RepetirAviso
                    End If
                End If
            End If
        
        End With
    Next
    Exit Sub
    
Handler:
    Call TraceError(Err.Number, Err.Description, "frmMain.Invasion_Timer")

    
    ' **********************************
End Sub

Private Sub t_Extraer_Timer()
    Dim i As Long
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    For i = 1 To LastUser
        If UserList(i).Counters.Trabajando > 0 Then
            Call Trabajar(i, UserList(i).Trabajo.TargetSkill)
        End If
    Next i
    Call PerformTimeLimitCheck(PerformanceTimer, "t_Extraer_Timer", 100)
End Sub

Private Sub T_UsersOnline_Timer()

On Error GoTo T_UsersOnline_Err

    Call MostrarNumUsers
    
    Exit Sub

T_UsersOnline_Err:
106     Call TraceError(Err.Number, Err.Description, "General.T_UsersOnline", Erl)

End Sub

Private Sub tControlHechizos_Timer()
    Dim UserIndex As Integer
    'Reseteo control de hechizos
    tHechizosMinutesCounter = tHechizosMinutesCounter + 1
    
    If tHechizosMinutesCounter = 2 Then
        For UserIndex = 1 To LastUser
            With UserList(UserIndex)
                UserList(UserIndex).Counters.controlHechizos.HechizosTotales = 0
                UserList(UserIndex).Counters.controlHechizos.HechizosCasteados = 0
            End With
        Next UserIndex
        tHechizosMinutesCounter = 0
    End If
        
End Sub

'  Comprobamos cada 10 segundos, porque no es necesaria tanta precisión
Private Sub TiempoRetos_Timer()

On Error GoTo Handler
    
    Dim IntervaloTimerRetosEnSegundos As Integer
    IntervaloTimerRetosEnSegundos = TiempoRetos.Interval * 0.001
    
    Dim Sala As Integer
    For Sala = 1 To Retos.TotalSalas
        
        With Retos.Salas(Sala)
            If .EnUso Then
                .TiempoRestante = .TiempoRestante - IntervaloTimerRetosEnSegundos
                
                If .TiempoItems > 0 Then
                    .TiempoItems = .TiempoItems - IntervaloTimerRetosEnSegundos
                    If .TiempoItems <= 0 Then Call TerminarTiempoAgarrarItems(Sala)
                ElseIf .TiempoRestante <= 0 Then
                    Call FinalizarReto(Sala, True)
                End If
            End If
        End With

    Next
    Exit Sub
    
Handler:
    Call TraceError(Err.Number, Err.Description, "frmMain.TiempoRetos_Timer")

    
End Sub

Private Sub TimerBarco_Timer()
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    Call UpdateBarcoForgatNix
    Call UpdateBarcoNixArghal
    Call UpdateBarcoArghalForgat
    Call MsnEnbarque(ForgatDock)
    Call MsnEnbarque(ArghalDock)
    Call MsnEnbarque(NixDock)
    Call PerformTimeLimitCheck(PerformanceTimer, "TimerBarco_Timer", 100)
End Sub

Private Function GetPassSlot(ByVal UserIndex As Integer) As Integer
Dim i As Integer
    With UserList(UserIndex)
        For i = 1 To UBound(.invent.Object)
            
            ' Le saco el item requerido de Forgat a Nix
            ' Es el mismo item que de Nix a Arghal y de Arghal a Forgat
            If .invent.Object(i).ObjIndex = BarcoNavegandoForgatNix.RequiredPassID Then
                GetPassSlot = i
                Exit Function
            End If
        Next
    End With
    GetPassSlot = -1
End Function

Private Sub MsnEnbarque(ByRef ShipInfo As t_Transport)
    On Error GoTo SendToMap_Err
    Dim LoopC     As Long
    Dim tempIndex As Integer

100     If Not MapaValido(ShipInfo.map) Then Exit Sub

102     For LoopC = 1 To ConnGroups(ShipInfo.map).CountEntrys
104         tempIndex = ConnGroups(ShipInfo.map).UserEntrys(LoopC)
106         If UserList(tempIndex).ConnectionDetails.ConnIDValida And UserList(tempIndex).pos.x >= ShipInfo.startX And UserList(tempIndex).pos.x <= ShipInfo.EndX And UserList(tempIndex).pos.y >= ShipInfo.startY And UserList(tempIndex).pos.y <= ShipInfo.EndY Then
                If Not GetPassSlot(tempIndex) > 0 Then
                    Call WriteLocaleMsg(tempIndex, MsgInvalidPass, e_FontTypeNames.FONTTYPE_GUILD)
                Else
                    Call WriteLocaleMsg(tempIndex, MsgStartingTrip, e_FontTypeNames.FONTTYPE_GUILD)
                End If
            End If
110     Next LoopC
        Exit Sub

SendToMap_Err:
112     Call TraceError(Err.Number, Err.Description, "modSendData.SendToMap", Erl)
End Sub
Private Sub UpdateBarcoForgatNix()

    Dim TileX, TileY As Integer
    Dim user As Integer
    Dim PassFound As Boolean
    Dim PassSlot As Integer
    
    ' Modificado por Shugar 5/6/24
    ' Viaje de Forgat a Nix
    
    ' Verificar si el barco está en un muelle:
    ' Para ver si está en el muelle o no, miramos hay un NpcIndex en Map DockX DockY del mapa BarcoNavegando.
    ' Ese Npc solía ser un muelle, y se usa de referencia para saber si el barco partió o sigue quieto.
    ' Si no hay NPC ahí es que el barco está navegando, por lo tanto no hay movimiento de pasajeros.
    
    If MapData(BarcoNavegandoForgatNix.Map, BarcoNavegandoForgatNix.DockX, BarcoNavegandoForgatNix.DockY).NpcIndex = 0 Then
        Exit Sub
    End If
          
    
    ' Desembarcar: bajamos del barco a los usuarios que llegan a Nix
    
    ' Para cada tile en el área del barco
    For TileX = BarcoNavegandoForgatNix.startX To BarcoNavegandoForgatNix.EndX
        For TileY = BarcoNavegandoForgatNix.startY To BarcoNavegandoForgatNix.EndY
        
            ' Si hay un usuario en el tile del barco
            user = MapData(BarcoNavegandoForgatNix.Map, TileX, TileY).UserIndex
            If user > 0 Then
            
                ' Enviar usuario a Nix
                Call WarpToLegalPos(user, NixDock.Map, NixDock.DestX, NixDock.DestY, True)
                Call WriteLocaleMsg(user, MsgThanksForTravelNix, e_FontTypeNames.FONTTYPE_GUILD)
                
            End If
        Next TileY
    Next TileX
        
    ' Embarcacar: subimos al barco a los usuarios que salen de Forgat
    
    ' Para cada tile en el área del muelle de Forgat
    For TileX = ForgatDock.startX To ForgatDock.EndX
        For TileY = ForgatDock.startY To ForgatDock.EndY
        
            ' Si hay un usuario en el tile del muelle
            user = MapData(ForgatDock.Map, TileX, TileY).UserIndex
            If user > 0 Then
            
                ' Sacarle el pasaje y moverlo al barco navegando
                PassSlot = GetPassSlot(user)
                If PassSlot > 0 Then
                   Call WriteLocaleMsg(user, MsgPassForgat, e_FontTypeNames.FONTTYPE_GUILD)
                   Call QuitarUserInvItem(user, PassSlot, 1)
                   Call UpdateUserInv(False, user, PassSlot)
                   Call WarpToLegalPos(user, BarcoNavegandoForgatNix.Map, BarcoNavegandoForgatNix.DestX, BarcoNavegandoForgatNix.DestY, True)
                Else
                   Call WriteLocaleMsg(user, MsgInvalidPass, e_FontTypeNames.FONTTYPE_GUILD)
                End If
            End If
        Next TileY
    Next TileX

End Sub


Private Sub UpdateBarcoNixArghal()

    Dim TileX, TileY As Integer
    Dim user As Integer
    Dim PassFound As Boolean
    Dim PassSlot As Integer
    
    ' Modificado por Shugar 5/6/24
    ' Viaje de Nix a Arghal
    
    ' Verificar si el barco está en un muelle:
    ' Para ver si está en el muelle o no, miramos hay un NpcIndex en Map DockX DockY del mapa BarcoNavegando.
    ' Ese Npc solía ser un muelle, y se usa de referencia para saber si el barco partió o sigue quieto.
    ' Si no hay NPC ahí es que el barco está navegando, por lo tanto no hay movimiento de pasajeros.
    
    If MapData(BarcoNavegandoNixArghal.Map, BarcoNavegandoNixArghal.DockX, BarcoNavegandoNixArghal.DockY).NpcIndex = 0 Then
        Exit Sub
    End If
    
    ' Desembarcar: bajamos del barco a los usuarios que llegan a Arghal
    
    ' Para cada tile en el área del barco
    For TileX = BarcoNavegandoNixArghal.startX To BarcoNavegandoNixArghal.EndX
        For TileY = BarcoNavegandoNixArghal.startY To BarcoNavegandoNixArghal.EndY
        
            ' Si hay un usuario en el tile del barco
            user = MapData(BarcoNavegandoNixArghal.Map, TileX, TileY).UserIndex
            If user > 0 Then
            
                ' Enviar al usuario a Arghal
                Call WarpToLegalPos(user, ArghalDock.Map, ArghalDock.DestX, ArghalDock.DestY, True)
                Call WriteLocaleMsg(user, MsgThanksForTravelArghal, e_FontTypeNames.FONTTYPE_GUILD)
                
            End If
        Next TileY
    Next TileX
  
    ' Embarcacar: subimos al barco a los usuarios que salen de Nix
    
    ' Para cada tile en el área del muelle de Nix
    For TileX = NixDock.startX To NixDock.EndX
        For TileY = NixDock.startY To NixDock.EndY
        
            ' Si hay un usuario en el tile del muelle
            user = MapData(NixDock.Map, TileX, TileY).UserIndex
            If user > 0 Then
            
                ' Sacarle el pasaje y moverlo al barco navegando
                PassSlot = GetPassSlot(user)
                If PassSlot > 0 Then
                   Call WriteLocaleMsg(user, MsgPassNix, e_FontTypeNames.FONTTYPE_GUILD)
                   Call QuitarUserInvItem(user, PassSlot, 1)
                   Call UpdateUserInv(False, user, PassSlot)
                   Call WarpToLegalPos(user, BarcoNavegandoNixArghal.Map, BarcoNavegandoNixArghal.DestX, BarcoNavegandoNixArghal.DestY, True)
                Else
                   Call WriteLocaleMsg(user, MsgInvalidPass, e_FontTypeNames.FONTTYPE_GUILD)
                End If
            End If
        Next TileY
    Next TileX

End Sub

Private Sub UpdateBarcoArghalForgat()

    Dim TileX, TileY As Integer
    Dim user As Integer
    Dim PassFound As Boolean
    Dim PassSlot As Integer
    
    ' Modificado por Shugar 5/6/24
    ' Viaje de Arghal a Forgat
    
    ' Verificar si el barco está en un muelle:
    ' Para ver si está en el muelle o no, miramos hay un NpcIndex en Map DockX DockY del mapa BarcoNavegando.
    ' Ese Npc solía ser un muelle, y se usa de referencia para saber si el barco partió o sigue quieto.
    ' Si no hay NPC ahí es que el barco está navegando, por lo tanto no hay movimiento de pasajeros.
    
    If MapData(BarcoNavegandoArghalForgat.Map, BarcoNavegandoArghalForgat.DockX, BarcoNavegandoArghalForgat.DockY).NpcIndex = 0 Then
        Exit Sub
    End If
    
    ' Desembarcar: bajamos del barco a los usuarios que llegan a Forgat
    
    ' Para cada tile en el área del barco
    For TileX = BarcoNavegandoArghalForgat.startX To BarcoNavegandoArghalForgat.EndX
        For TileY = BarcoNavegandoArghalForgat.startY To BarcoNavegandoArghalForgat.EndY
        
            ' Si hay un usuario en el tile del barco
            user = MapData(BarcoNavegandoArghalForgat.Map, TileX, TileY).UserIndex
            If user > 0 Then
            
                ' Enviar al usuario a Forgat
                Call WarpToLegalPos(user, ForgatDock.Map, ForgatDock.DestX, ForgatDock.DestY, True)
                Call WriteLocaleMsg(user, MsgThanksForTravelForgat, e_FontTypeNames.FONTTYPE_GUILD)
                
            End If
        Next TileY
    Next TileX
            
    
    ' Embarcacar: subimos al barco a los usuarios que salen de Arghal
    
    ' Para cada tile en el área del muelle de Arghal
    For TileX = ArghalDock.startX To ArghalDock.EndX
        For TileY = ArghalDock.startY To ArghalDock.EndY
        
            ' Si hay un usuario en el tile del muelle
            user = MapData(ArghalDock.Map, TileX, TileY).UserIndex
            If user > 0 Then
            
                ' Sacarle el pasaje y moverlo al barco navegando
                PassSlot = GetPassSlot(user)
                If PassSlot > 0 Then
                   Call WriteLocaleMsg(user, MsgPassArghal, e_FontTypeNames.FONTTYPE_GUILD)
                   Call QuitarUserInvItem(user, PassSlot, 1)
                   Call UpdateUserInv(False, user, PassSlot)
                   Call WarpToLegalPos(user, BarcoNavegandoArghalForgat.Map, BarcoNavegandoArghalForgat.DestX, BarcoNavegandoArghalForgat.DestY, True)
                Else
                   Call WriteLocaleMsg(user, MsgInvalidPass, e_FontTypeNames.FONTTYPE_GUILD)
                End If
            End If
        Next TileY
    Next TileX

End Sub



Private Sub TimerGuardarUsuarios_Timer()

On Error GoTo Handler
    If IsFeatureEnabled("auto_save_chars") Then
        ' Guardar usuarios (solo si pasó el tiempo mínimo para guardar)
        Dim UserIndex As Integer, UserGuardados As Integer
        Dim PerformanceTimer As Long
        Call PerformanceTestStart(PerformanceTimer)
        For UserIndex = 1 To LastUser
            With UserList(UserIndex)
                If .flags.UserLogged Then
                    If GetTickCount - .Counters.LastSave > IntervaloGuardarUsuarios Then
                        Call SaveUser(UserIndex)
                        UserGuardados = UserGuardados + 1
                        If UserGuardados > NumUsers Then Exit For
                        'limit the amount of time we block the only thread we have here, lets save some user on the next loop
                        If (GetTickCount - PerformanceTimer) > 100 Then Exit For
                    End If
                End If
            End With
        Next
        Call PerformTimeLimitCheck(PerformanceTimer, "TimerGuardarUsuarios_Timer", 100)
    End If
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "frmMain.TimreGuardarUsuarios_Timer")
End Sub

Private Sub Minuto_Timer()

    On Error GoTo ErrHandler

    'fired every minute
    Static minutos          As Long

    Static MinutosLatsClean As Long

    Dim i                   As Integer

    Dim Num                 As Long
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    MinsRunning = MinsRunning + 1

    If MinsRunning = 60 Then
        horas = horas + 1

        If horas = 24 Then
            DayStats.MaxUsuarios = 0
            DayStats.segundos = 0
            DayStats.Promedio = 0
        
            horas = 0
        
        End If

        MinsRunning = 0

    End If
    
    minutos = minutos + 1
    
    If MinutosLatsClean >= 15 Then
        MinutosLatsClean = 0
        Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
    Else
        MinutosLatsClean = MinutosLatsClean + 1
    End If

    Call PurgarPenas

   ' If IdleLimit > 0 Then
   '     Call CheckIdleUser
   ' End If

    If IsFeatureEnabled("automatic_events") Then
        Call Automatic_Event_Timer
    End If
    Call dump_stats
    Call PerformTimeLimitCheck(PerformanceTimer, "Minuto_Timer", 500)
    Exit Sub
        
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "General.Minuto_Timer", Erl)
        
End Sub

Private Sub CMDDUMP_Click()
On Error GoTo CMDDUMP_Click_Err
        Dim i As Integer

100     For i = 1 To MaxUsers
102         Call LogCriticEvent(i & ") ConnIDValida: " & UserList(i).ConnectionDetails.ConnIDValida & " Name: " & UserList(i).name & " UserLogged: " & UserList(i).flags.UserLogged)
104     Next i

106     Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)
        Exit Sub

CMDDUMP_Click_Err:
108     Call TraceError(Err.Number, Err.Description, "frmMain.CMDDUMP_Click", Erl)
End Sub

Private Sub Command1_Click()
        
        On Error GoTo Command1_Click_Err
        
100     Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(BroadMsg.Text))

        
        Exit Sub

Command1_Click_Err:
102     Call TraceError(Err.Number, Err.Description, "frmMain.Command1_Click", Erl)

        
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
106     Call TraceError(Err.Number, Err.Description, "frmMain.InitMain", Erl)

        
End Sub

Private Sub Command10_Click()
        
        On Error GoTo Command10_Click_Err
        
100     Call GuardarUsuarios

        
        Exit Sub

Command10_Click_Err:
102     Call TraceError(Err.Number, Err.Description, "frmMain.Command10_Click", Erl)

        
End Sub

Private Sub Command11_Click()
        
        On Error GoTo Command11_Click_Err
        
100     Call LoadSini
        Call LoadMD5
133     Call LoadPrivateKey
        
        Exit Sub

Command11_Click_Err:
102     Call TraceError(Err.Number, Err.Description, "frmMain.Command11_Click", Erl)

        
End Sub

Private Sub Command12_Click()
        
        On Error GoTo Command12_Click_Err
        
100     Call LoadMainConfigFile
        
        Exit Sub

Command12_Click_Err:
104     Call TraceError(Err.Number, Err.Description, "frmMain.Command12_Click", Erl)

        
End Sub

Private Sub Command13_Click()
        
        On Error GoTo Command13_Click_Err
        
100     Call LoadBalance

        
        Exit Sub

Command13_Click_Err:
102     Call TraceError(Err.Number, Err.Description, "frmMain.Command13_Click", Erl)

        
End Sub

Private Sub Command2_Click()
        
        On Error GoTo Command2_Click_Err
        
100     Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1660, BroadMsg.Text, e_FontTypeNames.FONTTYPE_SERVER)) 'Msg1660=Servidor » ¬1
        
        Exit Sub

Command2_Click_Err:
102     Call TraceError(Err.Number, Err.Description, "frmMain.Command2_Click", Erl)

        
End Sub

Private Sub Command4_Click()
        
        On Error GoTo Command4_Click_Err

100     If MsgBox("¿Está seguro que desea guardar y cerrar?", vbYesNo, "Confirmación") = vbNo Then Exit Sub
        
102     Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1661, vbNullString, e_FontTypeNames.FONTTYPE_PROMEDIO_MENOR)) 'Msg1661=Servidor » Cerrando servidor.

104     Call GuardarUsuarios
106     Call EcharPjsNoPrivilegiados

108     GuardarYCerrar = True
110     Unload frmMain

        
        Exit Sub

Command4_Click_Err:
112     Call TraceError(Err.Number, Err.Description, "frmMain.Command4_Click", Erl)

        
End Sub


Private Sub Command6_Click()
        
        On Error GoTo Command6_Click_Err
        
100     Call LoadIntervalos
101     Call LoadPacketRatePolicy
        
        Exit Sub

Command6_Click_Err:
102     Call TraceError(Err.Number, Err.Description, "frmMain.Command6_Click", Erl)

        
End Sub

Private Sub Command7_Click()
        
        On Error GoTo Command7_Click_Err
        
100     Call loadAdministrativeUsers

        
        Exit Sub

Command7_Click_Err:
102     Call TraceError(Err.Number, Err.Description, "frmMain.Command7_Click", Erl)

        
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
110     Call TraceError(Err.Number, Err.Description, "frmMain.Command8_Click", Erl)

        
End Sub

Private Sub Command9_Click()
        
        On Error GoTo Command9_Click_Err
        
100     Call CargaNpcsDat(True)

        
        Exit Sub

Command9_Click_Err:
102     Call TraceError(Err.Number, Err.Description, "frmMain.Command9_Click", Erl)

        
End Sub

Private Sub EstadoTimer_Timer()
    On Error GoTo EstadoTimer_Timer_Err
    Call GetHoraActual
    Dim i As Long
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    For i = 1 To Baneos.Count
        If Baneos(i).FechaLiberacion <= Now Then
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageLocaleMsg(1787, Baneos(i).name, e_FontTypeNames.FONTTYPE_SERVER)) ' Msg1787=Servidor » Se ha concluido la sentencia de ban para ¬1.
            Call UnBan(Baneos(i).Name)
            Call Baneos.Remove(i)
            Call SaveBans
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
    Call PerformTimeLimitCheck(PerformanceTimer, "FrmMain EstadoTimer_Timer", 100)
    Exit Sub
EstadoTimer_Timer_Err:
    Call TraceError(Err.Number, Err.Description, "frmMain.EstadoTimer_Timer", Erl)
End Sub

Private Sub Automatic_Event_Timer()
        
    On Error GoTo Evento_Timer_Err

    'validacion para captura la bandera que es unico
    If EventoActivo Then Exit Sub
    
    Dim CurrentDay As Byte
    Dim CurrentHour As Byte

    Dim EventOfTheDay As Integer

    CurrentDay = Weekday(Date) 'domingo=1 , lunes=2, martes=3, miercoles=4, jueves=5, viernes=6, sabado=7
    CurrentHour = Hour(Time) 'number between 0 and 23

    'si no es la hora de activar el evento salimos
    If SvrConfig.GetValue("AUTOEVENTHOURCALENDAR_Horarios") <> CurrentHour Then
        Exit Sub
    End If
    
    
    If AlreadyDidAutoEventToday Then
        '3600000 = 1 hora en milisegundos
        If GetTickCount - LastAutoEventAttempt > 3600000 Then
            AlreadyDidAutoEventToday = False
        End If
        Exit Sub
    End If
    
    Select Case CurrentDay
        Case 1 'domingo
            EventOfTheDay = SvrConfig.GetValue("AUTOEVENTDAYCALENDAR_Domingo")
        Case 2 'lunes
            EventOfTheDay = SvrConfig.GetValue("AUTOEVENTDAYCALENDAR_Lunes")
        Case 3 'martes
            EventOfTheDay = SvrConfig.GetValue("AUTOEVENTDAYCALENDAR_Martes")
        Case 4 'miÃ©rcoles
            EventOfTheDay = SvrConfig.GetValue("AUTOEVENTDAYCALENDAR_Miercoles")
        Case 5 'jueves
            EventOfTheDay = SvrConfig.GetValue("AUTOEVENTDAYCALENDAR_Jueves")
        Case 6 'viernes
            EventOfTheDay = SvrConfig.GetValue("AUTOEVENTDAYCALENDAR_Viernes")
        Case 7 'sÃ¡bado
            EventOfTheDay = SvrConfig.GetValue("AUTOEVENTDAYCALENDAR_Sabado")
    End Select

    'no hay evento programado para hoy
    If EventOfTheDay = 0 Then
        Exit Sub
    End If

    Dim LobbySettings As t_NewScenearioSettings

    
    Select Case EventOfTheDay

        Case e_EventType.CaptureTheFlag
            LobbySettings.ScenearioType = e_EventType.CaptureTheFlag
            LobbySettings.MinLevel = SvrConfig.GetValue("AUTOCAPTURETHEFLAG_MinLevel")
            LobbySettings.MaxLevel = SvrConfig.GetValue("AUTOCAPTURETHEFLAG_MaxLevel")
            LobbySettings.MinPlayers = SvrConfig.GetValue("AUTOCAPTURETHEFLAG_MinPlayers")
            LobbySettings.MaxPlayers = SvrConfig.GetValue("AUTOCAPTURETHEFLAG_MaxPlayers")
            LobbySettings.TeamSize = SvrConfig.GetValue("AUTOCAPTURETHEFLAG_TeamSize")
            LobbySettings.TeamType = SvrConfig.GetValue("AUTOCAPTURETHEFLAG_TeamType")
            LobbySettings.RoundNumber = SvrConfig.GetValue("AUTOCAPTURETHEFLAG_RoundNumber")
            LobbySettings.InscriptionFee = SvrConfig.GetValue("AUTOCAPTURETHEFLAG_InscriptionFee")
            LobbySettings.Description = SvrConfig.GetValue("AUTOCAPTURETHEFLAG_Description")

        Case e_EventType.NpcHunt
            LobbySettings.ScenearioType = e_EventType.NpcHunt
            LobbySettings.MinLevel = SvrConfig.GetValue("AUTONPCHUNT_MinLevel")
            LobbySettings.MaxLevel = SvrConfig.GetValue("AUTONPCHUNT_MaxLevel")
            LobbySettings.MinPlayers = SvrConfig.GetValue("AUTONPCHUNT_MinPlayers")
            LobbySettings.MaxPlayers = SvrConfig.GetValue("AUTONPCHUNT_MaxPlayers")
            LobbySettings.TeamSize = SvrConfig.GetValue("AUTONPCHUNT_TeamSize")
            LobbySettings.TeamType = SvrConfig.GetValue("AUTONPCHUNT_TeamType")
            LobbySettings.RoundNumber = SvrConfig.GetValue("AUTONPCHUNT_RoundNumber")
            LobbySettings.InscriptionFee = SvrConfig.GetValue("AUTONPCHUNT_InscriptionFee")
            LobbySettings.Description = SvrConfig.GetValue("AUTONPCHUNT_Description")

        Case e_EventType.DeathMatch
            LobbySettings.ScenearioType = e_EventType.DeathMatch
            LobbySettings.MinLevel = SvrConfig.GetValue("AUTODEATHMATCH_MinLevel")
            LobbySettings.MaxLevel = SvrConfig.GetValue("AUTODEATHMATCH_MaxLevel")
            LobbySettings.MinPlayers = SvrConfig.GetValue("AUTODEATHMATCH_MinPlayers")
            LobbySettings.MaxPlayers = SvrConfig.GetValue("AUTODEATHMATCH_MaxPlayers")
            LobbySettings.TeamSize = SvrConfig.GetValue("AUTODEATHMATCH_TeamSize")
            LobbySettings.TeamType = SvrConfig.GetValue("AUTODEATHMATCH_TeamType")
            LobbySettings.RoundNumber = SvrConfig.GetValue("AUTODEATHMATCH_RoundNumber")
            LobbySettings.InscriptionFee = SvrConfig.GetValue("AUTODEATHMATCH_InscriptionFee")
            LobbySettings.Description = SvrConfig.GetValue("AUTODEATHMATCH_Description")

        Case e_EventType.NavalBattle
            LobbySettings.ScenearioType = e_EventType.NavalBattle
            LobbySettings.MinLevel = SvrConfig.GetValue("AUTONAVALBATTLE_MinLevel")
            LobbySettings.MaxLevel = SvrConfig.GetValue("AUTONAVALBATTLE_MaxLevel")
            LobbySettings.MinPlayers = SvrConfig.GetValue("AUTONAVALBATTLE_MinPlayers")
            LobbySettings.MaxPlayers = SvrConfig.GetValue("AUTONAVALBATTLE_MaxPlayers")
            LobbySettings.TeamSize = SvrConfig.GetValue("AUTONAVALBATTLE_TeamSize")
            LobbySettings.TeamType = SvrConfig.GetValue("AUTONAVALBATTLE_TeamType")
            LobbySettings.RoundNumber = SvrConfig.GetValue("AUTONAVALBATTLE_RoundNumber")
            LobbySettings.InscriptionFee = SvrConfig.GetValue("AUTONAVALBATTLE_InscriptionFee")
            LobbySettings.Description = SvrConfig.GetValue("AUTONAVALBATTLE_Description")
    End Select

    LobbySettings.Password = ""

    CurrentActiveEventType = LobbySettings.ScenearioType
    
    If LobbySettings.ScenearioType = e_EventType.CaptureTheFlag Then
        Call HandleIniciarCaptura(LobbySettings)
    Else
        Call CreatePublicEvent(LobbySettings)
    End If
    
    AlreadyDidAutoEventToday = True
    LastAutoEventAttempt = GetTickCount
    
    Exit Sub
Evento_Timer_Err:
    Call TraceError(Err.Number, Err.Description, "frmMain.Evento_Timer", Erl)
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
124     Call TraceError(Err.Number, Err.Description, "frmMain.Form_MouseMove", Erl)

        
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
104     Call TraceError(Err.Number, Err.Description, "frmMain.QuitarIconoSystray", Erl)

        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        
        On Error GoTo Form_QueryUnload_Err
    
        
100     If GuardarYCerrar Then Exit Sub
102     If MsgBox("¿Deseas FORZAR el CIERRE del servidor?" & vbNewLine & vbNewLine & "Ten en cuenta que ES POSIBLE PIERDAS DATOS!", vbYesNo, "¡FORZAR CIERRE!") = vbNo Then
104         Cancel = True
        End If
    
        
        Exit Sub

Form_QueryUnload_Err:
106     Call TraceError(Err.Number, Err.Description, "frmMain.Form_QueryUnload", Erl)

        
End Sub

Private Sub Form_Unload(Cancel As Integer)
        

  Call CerrarServidor

        
End Sub

Private Sub GameTimer_Timer()
On Error GoTo HayError
    Dim iUserIndex   As Long
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    '<<<<<< Procesa eventos de los usuarios >>>>>>
    For iUserIndex = 1 To LastUser
        With UserList(iUserIndex)
            If .flags.UserLogged Then
                Call DoTileEvents(iUserIndex, .Pos.Map, .Pos.X, .Pos.Y)
                If .flags.Muerto = 0 Then
                    'Efectos en mapas
                    If (.flags.Privilegios And e_PlayerType.user) <> 0 Then
                        Call EfectoLava(iUserIndex)
                        Call EfectoFrio(iUserIndex)
                        If .flags.Envenenado <> 0 Then Call EfectoVeneno(iUserIndex)
                        If .flags.Incinerado <> 0 Then Call EfectoIncineramiento(iUserIndex)
                    End If
                    If .flags.Meditando Then Call DoMeditar(iUserIndex)
                    If .flags.Mimetizado <> 0 Then Call EfectoMimetismo(iUserIndex)
                    If .flags.AdminInvisible <> 1 Then
                        If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)
                    End If
                    If .NroMascotas > 0 Then Call TiempoInvocacion(iUserIndex)
                    Call EfectoStamina(iUserIndex)
                End If 'Muerto
            End If 'UserLogged
        End With
    Next iUserIndex
    Call PerformTimeLimitCheck(PerformanceTimer, "GameTimer_Timer User loop", 400)
    Call CustomScenarios.UpdateAll
    Call PerformTimeLimitCheck(PerformanceTimer, "GameTimer_Timer customScenarios", 100)
    Exit Sub
HayError:
    Call TraceError(Err.Number, Err.Description & vbNewLine & "UserIndex:" & iUserIndex, "frmMain.GameTimer", Erl)
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
    Call TraceError(Err.Number, Err.Description, "frmMain.HoraFantasia_Timer", Erl)
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
106     Call TraceError(Err.Number, Err.Description, "frmMain.mnuCerrar_Click", Erl)

        
End Sub

Private Sub mnusalir_Click()
        
        On Error GoTo mnusalir_Click_Err
        
100     Call mnuCerrar_Click

        
        Exit Sub

mnusalir_Click_Err:
102     Call TraceError(Err.Number, Err.Description, "frmMain.mnusalir_Click", Erl)

        
End Sub

Public Sub mnuMostrar_Click()
        
        On Error GoTo mnuMostrar_Click_Err
    
        

        

100     WindowState = vbNormal
102     Form_MouseMove 0, 0, 7725, 0

        
        Exit Sub

mnuMostrar_Click_Err:
104     Call TraceError(Err.Number, Err.Description, "frmMain.mnuMostrar_Click", Erl)

        
End Sub

Private Sub KillLog_Timer()
    On Error GoTo KillLog_Timer_Err
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\connect.log"
    If FileExist(App.Path & "\logs\haciendo.log", vbNormal) Then Kill App.Path & "\logs\haciendo.log"
    If FileExist(App.Path & "\logs\stats.log", vbNormal) Then Kill App.Path & "\logs\stats.log"
    If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
    If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
    If Not FileExist(App.Path & "\logs\nokillwsapi.txt") Then
        If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then Kill App.Path & "\logs\wsapi.log"
    End If
    Call PerformTimeLimitCheck(PerformanceTimer, "KillLog_Timer", 100)
    Exit Sub
KillLog_Timer_Err:
    Call TraceError(Err.Number, Err.Description, "frmMain.KillLog_Timer", Erl)
End Sub

Private Sub mnuServidor_Click()
        
        On Error GoTo mnuServidor_Click_Err
        
100     frmServidor.Visible = True

        
        Exit Sub

mnuServidor_Click_Err:
102     Call TraceError(Err.Number, Err.Description, "frmMain.mnuServidor_Click", Erl)

        
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
110     Call TraceError(Err.Number, Err.Description, "frmMain.mnuSystray_Click", Erl)

        
End Sub

Private Sub SubastaTimer_Timer()
        
    On Error GoTo SubastaTimer_Timer_Err
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    'Si ya paso un minuto y todavia no hubo oferta, avisamos que se cancela en un minuto
    If Subasta.TiempoRestanteSubasta = 240 And Subasta.HuboOferta = False Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1662, vbNullString, e_FontTypeNames.FONTTYPE_SUBASTA)) 'Msg1662=¡Quedan 4 minuto(s) para finalizar la subasta! Escribe /SUBASTA para mas información. La subasta será cancelada si no hay ofertas en el próximo minuto.
        Subasta.MinutosDeSubasta = 4
        Subasta.PosibleCancelo = True
    End If
    
    'Si ya pasaron dos minutos y no hubo ofertas, cancelamos la subasta
    If Subasta.TiempoRestanteSubasta = 180 And Subasta.HuboOferta = False Then
        Subasta.HaySubastaActiva = False
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1663, vbNullString, e_FontTypeNames.FONTTYPE_SUBASTA)) 'Msg1663=Subasta cancelada por falta de ofertas.
        'Devolver item antes de resetear datos
        Call DevolverItem
        Exit Sub
    End If

    If Subasta.PosibleCancelo = True Then
        Subasta.TiempoRestanteSubasta = Subasta.TiempoRestanteSubasta - 1
    End If
    
    If Subasta.TiempoRestanteSubasta > 0 And Subasta.PosibleCancelo = False Then
        If Subasta.TiempoRestanteSubasta = 240 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1664, vbNullString, e_FontTypeNames.FONTTYPE_SUBASTA)) 'Msg1664=¡Quedan 4 minuto(s) para finalizar la subasta! Escribe /SUBASTA para mas información.
            Subasta.MinutosDeSubasta = "4"
        End If
        If Subasta.TiempoRestanteSubasta = 180 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1665, vbNullString, e_FontTypeNames.FONTTYPE_SUBASTA)) 'Msg1665=¡Quedan 3 minuto(s) para finalizar la subasta! Escribe /SUBASTA para mas información.
            Subasta.MinutosDeSubasta = "3"
        End If

        If Subasta.TiempoRestanteSubasta = 120 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1666, vbNullString, e_FontTypeNames.FONTTYPE_SUBASTA)) 'Msg1666=¡Quedan 2 minuto(s) para finalizar la subasta! Escribe /SUBASTA para mas información.
            Subasta.MinutosDeSubasta = "2"
        End If

        If Subasta.TiempoRestanteSubasta = 60 Then
            Subasta.MinutosDeSubasta = "1"
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1667, vbNullString, e_FontTypeNames.FONTTYPE_SUBASTA)) 'Msg1667=¡Quedan 1 minuto(s) para finalizar la subasta! Escribe /SUBASTA para mas información.
        End If
        Subasta.TiempoRestanteSubasta = Subasta.TiempoRestanteSubasta - 1
    End If
    
    If Subasta.TiempoRestanteSubasta = 1 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1668, Subasta.Comprador, e_FontTypeNames.FONTTYPE_SUBASTA)) 'Msg1668=¡La subasta ha terminado! El ganador fue: ¬1

        Call FinalizarSubasta
    End If
    Call PerformTimeLimitCheck(PerformanceTimer, "SubastaTimer_Timer")
        
    Exit Sub

SubastaTimer_Timer_Err:
    Call TraceError(Err.Number, Err.Description, "frmMain.SubastaTimer_Timer", Erl)

        
End Sub






Private Sub TIMER_AI_Timer()
    On Error GoTo ErrorHandler
    Dim NpcIndex As Long
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    If Not haciendoBK And Not EnPausa Then
        For NpcIndex = 1 To LastNPC
            With NpcList(NpcIndex)
                If .Pos.Map > 0 Then
                    If MapInfo(.Pos.Map).NumUsers > 0 Or MapInfo(.Pos.Map).ForceUpdate Then
                                If .flags.NPCActive Then
                                    If .npcType = DummyTarget Then
                                        Call NpcDummyUpdate(NpcIndex)
                                     Else
                                        If .flags.Paralizado > 0 Then Call EfectoParalisisNpc(NpcIndex)
                                        If .flags.Inmovilizado > 0 Then Call EfectoInmovilizadoNpc(NpcIndex)
                                        If IntervaloPermiteMoverse(NpcIndex) Then Call NpcAI(NpcIndex)
                                    End If 'If .npcType = DummyTarget Then
                                End If 'If .flags.NPCActive Then
                    End If 'If MapInfo(.Pos.Map).NumUsers > 0 Then
                End If 'If .Pos.Map > 0 Then
            End With
        Next NpcIndex
    End If
    Call PerformTimeLimitCheck(PerformanceTimer, "TIMER_AI_Timer", 600)
    Exit Sub

ErrorHandler:
    Call TraceError(Err.Number, Err.Description & vbNewLine & _
                                    "NPC: " & NpcList(NpcIndex).Name & _
                                    " en la posicion: " & NpcList(NpcIndex).Pos.Map & "-" & NpcList(NpcIndex).Pos.X & "-" & NpcList(NpcIndex).Pos.Y, "frmMain.Timer_AI", Erl)
    Call MuereNpc(NpcIndex, 0)

End Sub

Private Sub TimerMeteorologia_Timer()
    'Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1741, TimerMeteorologico, e_FontTypeNames.FONTTYPE_SERVER)) 'Msg1741=Servidor > Timer de lluvia en : ¬1
        
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
            ' Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor > Empezaron las nubes con intensidad: " & IntensidadDeNubes & "%.", e_FontTypeNames.FONTTYPE_SERVER))
            Call AgregarAConsola("Servidor » Empezaron las nubes")
            TimerMeteorologico = TimerMeteorologico - 1
        Else
            ServidorNublado = False
            ' Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor > Tranquilo, no hay nubes ni va a llover.", e_FontTypeNames.FONTTYPE_SERVER))
            Call AgregarAConsola("Servidor » Tranquilo, no hay nubes ni va a llover.")
            Call ResetMeteo
            Exit Sub
        End If
    End If

    If TimerMeteorologico < 7 And TimerMeteorologico > 3 Then
        TimerMeteorologico = TimerMeteorologico - 1
        'Enviar Truenos y rayos
        Truenos.Enabled = True
        'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor > Envio un truenito para que te asustes.", e_FontTypeNames.FONTTYPE_SERVER))
        Call AgregarAConsola("Servidor » Truenos y nubes activados.")
        Exit Sub
    End If

    If TimerMeteorologico = 3 Then
        ProbabilidadLLuvia = RandomNumber(1, 5)
        If ProbabilidadLLuvia = 1 Then
            'Envia Lluvia
            Nebando = True
            Lloviendo = True
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(404, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno
            Call SendData(SendTarget.ToAll, 0, PrepareMessageFlashScreen(&HD254D6, 250)) 'Rayo
            Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
            Call SendData(SendTarget.ToAll, 0, PrepareMessageNevarToggle())
            Call AgregarAConsola("Servidor » Lloviendo.")
            TimerMeteorologico = TimerMeteorologico - 1
        Else
            Nieblando = False
            Lloviendo = False
            ServidorNublado = False
            Truenos.Enabled = False
            Call SendData(SendTarget.ToAll, 0, PrepareMessageNieblandoToggle(IntensidadDeNubes))
            Call AgregarAConsola("Servidor » Truenos y nubes desactivados.")
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
        Lloviendo = False
        Truenos.Enabled = False
        Nebando = False
        Call SendData(SendTarget.ToAll, 0, PrepareMessageNieblandoToggle(IntensidadDeNubes))
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        Call SendData(SendTarget.ToAll, 0, PrepareMessageNevarToggle())
        Call AgregarAConsola("Servidor >Lluvia desactivada.")
        Call ResetMeteo
        Exit Sub
    End If
    Exit Sub
TimerMeteorologia_Timer_Err:
    Call TraceError(Err.Number, Err.Description, "frmMain.TimerMeteorologia_Timer", Erl)
End Sub

Private Sub TimerRespawn_Timer()

    On Error GoTo ErrorHandler
    Dim NpcIndex As Long
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    'Update NPCs
    For NpcIndex = 1 To MaxRespawn
        'Debug.Print RespawnList(NpcIndex).name
        If RespawnList(NpcIndex).flags.NPCActive Then  'Nos aseguramos que este muerto
            If RespawnList(NpcIndex).Contadores.IntervaloRespawn > 0 Then
                RespawnList(NpcIndex).Contadores.IntervaloRespawn = RespawnList(NpcIndex).Contadores.IntervaloRespawn - 1
            Else
                RespawnList(NpcIndex).flags.NPCActive = False
                If RespawnList(NpcIndex).InformarRespawn = 1 Then 
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1788, RespawnList(NpcIndex).Numero, e_FontTypeNames.FONTTYPE_EXP)) ' Msg1788=¬1 ha vuelto a este mundo.
                    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(257, NO_3D_SOUND, NO_3D_SOUND)) 'Para evento de respwan
                End If
                Call ReSpawnNpc(RespawnList(NpcIndex))
            End If
        End If
    Next NpcIndex
    Call PerformTimeLimitCheck(PerformanceTimer, "TimerRespawn_Timer")
    Exit Sub

ErrorHandler:
    Call TraceError(Err.Number, Err.Description & vbNewLine & _
                                    "NPC: " & NpcList(NpcIndex).Name & _
                                    " en la posicion: " & NpcList(NpcIndex).Pos.Map & "-" & NpcList(NpcIndex).Pos.X & "-" & NpcList(NpcIndex).Pos.Y, "frmMain.TimerRespawn_Timer", Erl)
    Call MuereNpc(NpcIndex, 0)

End Sub

Private Sub tPiqueteC_Timer()

    On Error GoTo ErrHandler

    Static segundos As Integer
    Dim NuevaA      As Boolean
    Dim NuevoL      As Boolean
    Dim GI          As Integer
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    segundos = segundos + 6

    Dim i As Long
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            If MapData(UserList(i).Pos.Map, UserList(i).Pos.X, UserList(i).Pos.Y).trigger = e_Trigger.ANTIPIQUETE Then
                UserList(i).Counters.PiqueteC = UserList(i).Counters.PiqueteC + 1
                ' Le empiezo a avisar a partir de los 18 segundos, para no spamear
                If UserList(i).Counters.PiqueteC > 3 Then
                    Call WriteLocaleMsg(i, "70", e_FontTypeNames.FONTTYPE_INFO)
                End If
            
                If UserList(i).Counters.PiqueteC > 10 Then
                    UserList(i).Counters.PiqueteC = 0
                    'Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
                    ' En vez de encarcelarlo, lo sacamos del juego.
                    'Ojo! No sï¿½ si se puede abusar de esto para evitar los 10 segundos al salir
                    Call WriteDisconnect(i)
                    Call CloseSocket(i)
                End If
            Else
                If UserList(i).Counters.PiqueteC > 0 Then UserList(i).Counters.PiqueteC = 0
            End If
            If segundos >= 18 Then
                If segundos >= 18 Then UserList(i).Counters.Pasos = 0
            End If
        End If
    Next i
    Call PerformTimeLimitCheck(PerformanceTimer, "tPiqueteC_Timer", 100)
    If segundos >= 18 Then segundos = 0
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "frmMain.tPiqueteC_Timer", Erl)
End Sub

Private Sub Truenos_Timer()
    On Error GoTo Truenos_Timer_Err
    Dim Enviar    As Byte
    Dim TruenoWav As Integer
    Enviar = RandomNumber(1, 15)
    Dim Duracion As Long
    If Enviar < 8 Then
        TruenoWav = 399 + Enviar
        If TruenoWav = 404 Then TruenoWav = 406
        Duracion = RandomNumber(80, 250)
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(TruenoWav, NO_3D_SOUND, NO_3D_SOUND))
        Call SendData(SendTarget.ToAll, 0, PrepareMessageFlashScreen(&HEFEECB, Duracion))
    End If
    Exit Sub
Truenos_Timer_Err:
    Call TraceError(Err.Number, Err.Description, "frmMain.Truenos_Timer", Erl)
End Sub

Private Sub UptimeTimer_Timer()
    On Error GoTo UptimeTimer_Timer_Err
    SERVER_UPTIME = SERVER_UPTIME + 1
    Exit Sub
UptimeTimer_Timer_Err:
    Call TraceError(Err.Number, Err.Description, "frmMain.UptimeTimer_Timer", Erl)
End Sub

