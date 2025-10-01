VERSION 5.00
Begin VB.Form frmServidor 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Servidor"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   327
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   502
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Apagar - Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3840
      TabIndex        =   22
      Top             =   2880
      Width           =   3615
      Begin VB.CommandButton Command4 
         Caption         =   "Hacer un Backup del mundo"
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
         TabIndex        =   26
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Guardar todos los personajes"
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
         TabIndex        =   25
         Top             =   520
         Width           =   3255
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cargar BackUp del mundo"
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
         TabIndex        =   24
         Top             =   820
         Width           =   3255
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Boton Magico para apagar server"
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
         TabIndex        =   23
         Top             =   1120
         Width           =   3255
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Otros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   3840
      TabIndex        =   14
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton cmdReloadGuilds 
         Caption         =   "Recargar base de datos de clanes"
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
         TabIndex        =   28
         Top             =   2430
         Width           =   3255
      End
      Begin VB.CommandButton cmdDumpLogs 
         Caption         =   "Dump Logs"
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
         TabIndex        =   27
         Top             =   2170
         Width           =   3255
      End
      Begin VB.CommandButton Command6 
         Caption         =   "ReSpawn Guardias en pos. originales"
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
         TabIndex        =   21
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Reiniciar"
         Enabled         =   0   'False
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
         TabIndex        =   20
         Top             =   510
         Width           =   3255
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Configurar intervalos"
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
         TabIndex        =   19
         Top             =   800
         Width           =   3255
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Debug Npcs"
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
         TabIndex        =   18
         Top             =   1080
         Width           =   3255
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Unban All (PELIGRO!)"
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
         TabIndex        =   17
         Top             =   1360
         Width           =   3255
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Unban All IPs (PELIGRO!)"
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
         TabIndex        =   16
         Top             =   1650
         Width           =   3255
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Pausar el servidor"
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
         TabIndex        =   15
         Top             =   1920
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Trafico - Sockets"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   3735
      Begin VB.CommandButton Command20 
         Caption         =   "Reset sockets"
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
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   3255
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Trafico"
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
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   3255
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Administración"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   3255
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Debug UserList"
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
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Actualizar - Recargar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton Command14 
         Caption         =   "Update MOTD"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   3255
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Reload Lista Nombres Prohibidos"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   3255
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Reload Server.ini"
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
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   3255
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Actualizar npcs.dat"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Actualizar objetos.dat"
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
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Actualizar hechizos"
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
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   3255
      End
      Begin VB.CommandButton Command28 
         Caption         =   "Recargar configuracion.ini"
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
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   0
      Top             =   4560
      Width           =   945
   End
End
Attribute VB_Name = "frmServidor"
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

Private Sub cmdReloadGuilds_Click()
    Call modGuilds.LoadGuildsDB
End Sub

Private Sub Command1_Click()
    On Error GoTo Command1_Click_Err
    Call LoadOBJData
    Call LoadPesca
    Call LoadRecursosEspeciales
    Exit Sub
Command1_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command1_Click", Erl)
End Sub

Private Sub Command10_Click()
    On Error GoTo Command10_Click_Err
    frmTrafic.Show
    Exit Sub
Command10_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command10_Click", Erl)
End Sub

Private Sub Command12_Click()
    On Error GoTo Command12_Click_Err
    frmDebugNpc.Show
    Exit Sub
Command12_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command12_Click", Erl)
End Sub

Private Sub Command14_Click()
    On Error GoTo Command14_Click_Err
    Call LoadMotd
    Exit Sub
Command14_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command14_Click", Erl)
End Sub

Private Sub Command15_Click()
    On Error GoTo Command15_Click_Err
    Dim Fn       As String
    Dim cad$
    Dim n        As Integer, K As Integer
    Dim sENtrada As String
    sENtrada = InputBox("Escribe ""estoy DE acuerdo"" entre comillas y con distición de mayusculas minusculas para desbanear a todos los personajes", "UnBan", "hola")
    If sENtrada = "estoy DE acuerdo" Then
        Fn = App.Path & "\logs\GenteBanned.log"
        If FileExist(Fn, vbNormal) Then
            n = FreeFile
            Open Fn For Input Shared As #n
            Do While Not EOF(n)
                K = K + 1
                Input #n, cad$
                Call UnBan(cad$)
            Loop
            Close #n
            MsgBox "Se han habilitado " & K & " personajes."
            Kill Fn
        End If
    End If
    Exit Sub
Command15_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command15_Click", Erl)
End Sub

Private Sub Command16_Click()
    On Error GoTo Command16_Click_Err
    Call LoadSini
    Call LoadMD5
    Call LoadPrivateKey
    Exit Sub
Command16_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command16_Click", Erl)
End Sub

Private Sub Command17_Click()
    On Error GoTo Command17_Click_Err
    Call CargaNpcsDat
    Exit Sub
Command17_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command17_Click", Erl)
End Sub

Private Sub Command18_Click()
    On Error GoTo Command18_Click_Err
    Me.MousePointer = 11
    Call GuardarUsuarios
    Me.MousePointer = 0
    MsgBox "Grabado de personajes OK!"
    Exit Sub
Command18_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command18_Click", Erl)
End Sub

Private Sub Command2_Click()
    On Error GoTo Command2_Click_Err
    frmServidor.Visible = False
    Exit Sub
Command2_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command2_Click", Erl)
End Sub

Private Sub Command20_Click()
    On Error GoTo Command20_Click_Err
    Dim LoopC As Long
    If MsgBox("Esta seguro que desea reiniciar los sockets ? Se cerrarán todas las conexiones activas.", vbYesNo, "Reiniciar Sockets") = vbYes Then
        Call modNetwork.Disconnect
        For LoopC = 1 To MaxUsers
            Call CloseSocket(LoopC)
        Next
        Call modNetwork.Listen(MaxUsers, ListenIp, CStr(Puerto))
    End If
    Exit Sub
Command20_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command20_Click", Erl)
End Sub

'Barrin 29/9/03
Private Sub Command21_Click()
    On Error GoTo Command21_Click_Err
    If EnPausa = False Then
        EnPausa = True
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
        Command21.Caption = "Reanudar el servidor"
    Else
        EnPausa = False
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
        Command21.Caption = "Pausar el servidor"
    End If
    Exit Sub
Command21_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command21_Click", Erl)
End Sub

Private Sub Command22_Click()
    On Error GoTo Command22_Click_Err
    Me.Visible = False
    frmAdmin.Show
    Exit Sub
Command22_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command22_Click", Erl)
End Sub

Private Sub Command23_Click()
    On Error GoTo Command23_Click_Err
    If MsgBox("Esta seguro que desea hacer WorldSave, guardar pjs y cerrar ?", vbYesNo, "Apagar Magicamente") = vbYes Then
        Me.MousePointer = 11
        FrmStat.Show
        Call GuardarUsuarios
        Unload frmMain
    End If
    Exit Sub
Command23_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command23_Click", Erl)
End Sub

Private Sub Command27_Click()
    On Error GoTo Command27_Click_Err
    frmUserList.Show
    Exit Sub
Command27_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command27_Click", Erl)
End Sub

Private Sub Command28_Click()
    On Error GoTo Command28_Click_Err
    Call LoadMainConfigFile
    Exit Sub
Command28_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command28_Click", Erl)
End Sub

Private Sub Command3_Click()
    On Error GoTo Command3_Click_Err
    If MsgBox("¡¡Atencion!! Si reinicia el servidor puede provocar la perdida de datos de los usarios. ¿Desea reiniciar el servidor de todas maneras?", vbYesNo) = vbYes Then
        Me.Visible = False
        Call Restart
    End If
    Exit Sub
Command3_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command3_Click", Erl)
End Sub

Private Sub Command4_Click()
    On Error GoTo eh
    Me.MousePointer = 11
    FrmStat.Show
    Call DoBackUp
    Me.MousePointer = 0
    MsgBox "WORLDSAVE OK!!"
    Exit Sub
eh:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command4_Click", Erl)
End Sub

Private Sub Command5_Click()
    On Error GoTo Command5_Click_Err
    'Se asegura de que los sockets estan cerrados e ignora cualquier err
    If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."
    FrmStat.Show
    If FileExist(App.Path & "\logs\errores.log", vbNormal) Then Kill App.Path & "\logs\errores.log"
    If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\Connect.log"
    If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
    If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
    If FileExist(App.Path & "\logs\Resurrecciones.log", vbNormal) Then Kill App.Path & "\logs\Resurrecciones.log"
    If FileExist(App.Path & "\logs\Teleports.Log", vbNormal) Then Kill App.Path & "\logs\Teleports.Log"
    Dim LoopC As Integer
    For LoopC = 1 To MaxUsers
        Call CloseSocket(LoopC)
    Next
    Call modNetwork.Disconnect
    LastUser = 0
    NumUsers = 0
    Call FreeNPCs
    Call FreeCharIndexes
    Call LoadSini
    Call LoadMD5
    Call LoadPrivateKey
    Call LoadIntervalos
    Call CargarBackUp
    Call LoadOBJData
    Call LoadPesca
    Call LoadRecursosEspeciales
    Call modNetwork.Listen(MaxUsers, ListenIp, CStr(Puerto))
    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    Exit Sub
Command5_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command5_Click", Erl)
End Sub

Private Sub Command6_Click()
    On Error GoTo Command6_Click_Err
    Call ReSpawnOrigPosNpcs
    Exit Sub
Command6_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command6_Click", Erl)
End Sub

Private Sub Command7_Click()
    On Error GoTo Command7_Click_Err
    FrmInterv.Show
    Exit Sub
Command7_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command7_Click", Erl)
End Sub

Private Sub Command8_Click()
    On Error GoTo Command8_Click_Err
    Call CargarHechizos
    Exit Sub
Command8_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command8_Click", Erl)
End Sub

Private Sub Command9_Click()
    On Error GoTo Command9_Click_Err
    Call CargarForbidenWords
    Exit Sub
Command9_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Command9_Click", Erl)
End Sub

Private Sub Form_Deactivate()
    On Error GoTo Form_Deactivate_Err
    frmServidor.Visible = False
    Exit Sub
Form_Deactivate_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Form_Deactivate", Erl)
End Sub

Private Sub Form_Load()
    On Error GoTo Form_Load_Err
    Command20.Visible = True
    Exit Sub
Form_Load_Err:
    Call TraceError(Err.Number, Err.Description, "frmServidor.Form_Load", Erl)
End Sub
