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
      TabIndex        =   27
      Top             =   2400
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
      Height          =   2415
      Left            =   3840
      TabIndex        =   19
      Top             =   0
      Width           =   3615
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   1940
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
      TabIndex        =   10
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
         TabIndex        =   18
         Top             =   1800
         Width           =   3255
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Reset Listen"
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
         TabIndex        =   17
         Top             =   2040
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
         TabIndex        =   16
         Top             =   360
         Width           =   3255
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Stats de los slots"
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
         TabIndex        =   15
         Top             =   600
         Width           =   3255
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Debug listening socket"
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
         TabIndex        =   14
         Top             =   840
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
         TabIndex        =   13
         Top             =   1080
         Width           =   3255
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Estadisticas"
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
         Top             =   1320
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
         TabIndex        =   11
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   960
         Width           =   3255
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Reload MD5s"
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
         Top             =   1200
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
      Top             =   4200
      Width           =   945
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "pablito_3_15@hotmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3960
      TabIndex        =   35
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "2008"
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
      Left            =   5520
      TabIndex        =   34
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo Daniel Mercavides"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3960
      TabIndex        =   33
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "RevolucionAo 1.1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   3960
      TabIndex        =   32
      Top             =   4080
      Width           =   2295
   End
End
Attribute VB_Name = "frmServidor"
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

Private Sub Command1_Click()
        
        On Error GoTo Command1_Click_Err
        
100     Call LoadOBJData
102     Call LoadPesca
104     Call LoadRecursosEspeciales

        
        Exit Sub

Command1_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command1_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command10_Click()
        
        On Error GoTo Command10_Click_Err
        
100     frmTrafic.Show

        
        Exit Sub

Command10_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command10_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command11_Click()
        
        On Error GoTo Command11_Click_Err
        
100     frmConID.Show

        
        Exit Sub

Command11_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command11_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command12_Click()
        
        On Error GoTo Command12_Click_Err
        
100     frmDebugNpc.Show

        
        Exit Sub

Command12_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command12_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command13_Click()
        
        On Error GoTo Command13_Click_Err
        
100     frmDebugSocket.Visible = True

        
        Exit Sub

Command13_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command13_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command14_Click()
        
        On Error GoTo Command14_Click_Err
        
100     Call LoadMotd

        
        Exit Sub

Command14_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command14_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command15_Click()

    On Error Resume Next

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

End Sub

Private Sub Command16_Click()
        
        On Error GoTo Command16_Click_Err
        
100     Call LoadSini

        
        Exit Sub

Command16_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command16_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command17_Click()
        
        On Error GoTo Command17_Click_Err
        
100     Call CargaNpcsDat

        
        Exit Sub

Command17_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command17_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command18_Click()
        
        On Error GoTo Command18_Click_Err
        
100     Me.MousePointer = 11
102     Call GuardarUsuarios
104     Me.MousePointer = 0
106     MsgBox "Grabado de personajes OK!"

        
        Exit Sub

Command18_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command18_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command19_Click()
        
        On Error GoTo Command19_Click_Err
        

        Dim i        As Long, n As Long

        Dim sENtrada As String

100     sENtrada = InputBox("Escribe ""estoy DE acuerdo"" sin comillas y con distición de mayusculas minusculas para desbanear a todos los personajes", "UnBan", "hola")

102     If sENtrada = "estoy DE acuerdo" Then
    
104         n = BanIps.Count

106         For i = 1 To BanIps.Count
108             BanIps.Remove 1
110         Next i
    
112         MsgBox "Se han habilitado " & n & " ipes"

        End If

        
        Exit Sub

Command19_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command19_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command2_Click()
        
        On Error GoTo Command2_Click_Err
        
100     frmServidor.Visible = False

        
        Exit Sub

Command2_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command2_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command20_Click()
        
        On Error GoTo Command20_Click_Err
        
        #If UsarQueSocket = 1 Then

100         If MsgBox("Esta seguro que desea reiniciar los sockets ? Se cerrarán todas las conexiones activas.", vbYesNo, "Reiniciar Sockets") = vbYes Then
102             Call WSApiReiniciarSockets

            End If

        #ElseIf UsarQueSocket = 2 Then

            Dim LoopC As Integer

104         If MsgBox("Esta seguro que desea reiniciar los sockets ? Se cerrarán todas las conexiones activas.", vbYesNo, "Reiniciar Sockets") = vbYes Then

106             For LoopC = 1 To MaxUsers

108                 If UserList(LoopC).ConnID <> -1 And UserList(LoopC).ConnIDValida Then
110                     Call CloseSocket(LoopC)

                    End If

112             Next LoopC
    
114             Call frmMain.Serv.Detener
116             Call frmMain.Serv.Iniciar(Puerto)

            End If

        #End If

        
        Exit Sub

Command20_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command20_Click", Erl)
        Resume Next
        
End Sub

'Barrin 29/9/03
Private Sub Command21_Click()
        
        On Error GoTo Command21_Click_Err
        

100     If EnPausa = False Then
102         EnPausa = True
104         Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
106         Command21.Caption = "Reanudar el servidor"
        Else
108         EnPausa = False
110         Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
112         Command21.Caption = "Pausar el servidor"

        End If

        
        Exit Sub

Command21_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command21_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command22_Click()
        
        On Error GoTo Command22_Click_Err
        
100     Me.Visible = False
102     frmAdmin.Show

        
        Exit Sub

Command22_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command22_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command23_Click()
        
        On Error GoTo Command23_Click_Err
        

100     If MsgBox("Esta seguro que desea hacer WorldSave, guardar pjs y cerrar ?", vbYesNo, "Apagar Magicamente") = vbYes Then
102         Me.MousePointer = 11
    
104         FrmStat.Show
   
            'WorldSave
            '   Call DoBackUp

            'Guardar Pjs
106         Call GuardarUsuarios
    
            'Chauuu
108         Unload frmMain

        End If

        
        Exit Sub

Command23_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command23_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command24_Click()
        
        On Error GoTo Command24_Click_Err
        
100     frmEstadisticas.Show

        
        Exit Sub

Command24_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command24_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command25_Click()
        
        On Error GoTo Command25_Click_Err
        
100     Call MD5sCarga

        
        Exit Sub

Command25_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command25_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command26_Click()
        
        On Error GoTo Command26_Click_Err
        
        #If UsarQueSocket = 1 Then

            'Cierra el socket de escucha
100         If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
            'Inicia el socket de escucha
102         SockListen = ListenForConnect(Puerto, hWndMsg, "")
        #End If

        
        Exit Sub

Command26_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command26_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command27_Click()
        
        On Error GoTo Command27_Click_Err
        
100     frmUserList.Show

        
        Exit Sub

Command27_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command27_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command28_Click()
        
        On Error GoTo Command28_Click_Err
        
100     Call LoadConfiguraciones

        
        Exit Sub

Command28_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command28_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command3_Click()
        
        On Error GoTo Command3_Click_Err
        

100     If MsgBox("¡¡Atencion!! Si reinicia el servidor puede provocar la perdida de datos de los usarios. ¿Desea reiniciar el servidor de todas maneras?", vbYesNo) = vbYes Then
102         Me.Visible = False
104         Call Restart

        End If

        
        Exit Sub

Command3_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command3_Click", Erl)
        Resume Next
        
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
    Call LogError("Error en WORLDSAVE")

End Sub

Private Sub Command5_Click()

    'Se asegura de que los sockets estan cerrados e ignora cualquier err
    On Error Resume Next

    If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

    FrmStat.Show

    If FileExist(App.Path & "\logs\errores.log", vbNormal) Then Kill App.Path & "\logs\errores.log"
    If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\Connect.log"
    If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
    If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
    If FileExist(App.Path & "\logs\Resurrecciones.log", vbNormal) Then Kill App.Path & "\logs\Resurrecciones.log"
    If FileExist(App.Path & "\logs\Teleports.Log", vbNormal) Then Kill App.Path & "\logs\Teleports.Log"

    #If UsarQueSocket = 1 Then
        Call apiclosesocket(SockListen)
    #ElseIf UsarQueSocket = 0 Then
        frmMain.Socket1.Cleanup
        frmMain.Socket2(0).Cleanup
    #ElseIf UsarQueSocket = 2 Then
        frmMain.Serv.Detener
    #End If

    Dim LoopC As Integer

    For LoopC = 1 To MaxUsers
        Call CloseSocket(LoopC)
    Next

    LastUser = 0
    NumUsers = 0

    Call FreeNPCs
    Call FreeCharIndexes

    Call LoadSini
    Call LoadIntervalos
    Call CargarBackUp
    Call LoadOBJData
    Call LoadPesca
    Call LoadRecursosEspeciales

    #If UsarQueSocket = 1 Then
        SockListen = ListenForConnect(Puerto, hWndMsg, "")

    #ElseIf UsarQueSocket = 0 Then
        frmMain.Socket1.AddressFamily = AF_INET
        frmMain.Socket1.Protocol = IPPROTO_IP
        frmMain.Socket1.SocketType = SOCK_STREAM
        frmMain.Socket1.Binary = False
        frmMain.Socket1.Blocking = False
        frmMain.Socket1.BufferSize = 1024

        frmMain.Socket2(0).AddressFamily = AF_INET
        frmMain.Socket2(0).Protocol = IPPROTO_IP
        frmMain.Socket2(0).SocketType = SOCK_STREAM
        frmMain.Socket2(0).Blocking = False
        frmMain.Socket2(0).BufferSize = 2048

        'Escucha
        frmMain.Socket1.LocalPort = Puerto
        frmMain.Socket1.listen
    #End If

    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

End Sub

Private Sub Command6_Click()
        
        On Error GoTo Command6_Click_Err
        
100     Call ReSpawnOrigPosNpcs

        
        Exit Sub

Command6_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command6_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command7_Click()
        
        On Error GoTo Command7_Click_Err
        
100     FrmInterv.Show

        
        Exit Sub

Command7_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command7_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command8_Click()
        
        On Error GoTo Command8_Click_Err
        
100     Call CargarHechizos

        
        Exit Sub

Command8_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command8_Click", Erl)
        Resume Next
        
End Sub

Private Sub Command9_Click()
        
        On Error GoTo Command9_Click_Err
        
100     Call CargarForbidenWords

        
        Exit Sub

Command9_Click_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Command9_Click", Erl)
        Resume Next
        
End Sub

Private Sub Form_Deactivate()
        
        On Error GoTo Form_Deactivate_Err
        
100     frmServidor.Visible = False

        
        Exit Sub

Form_Deactivate_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Form_Deactivate", Erl)
        Resume Next
        
End Sub

Private Sub Form_Load()
        
        On Error GoTo Form_Load_Err
        
        #If UsarQueSocket = 1 Then
100         Command20.Visible = True
102         Command26.Visible = True
        #ElseIf UsarQueSocket = 0 Then
104         Command20.Visible = False
106         Command26.Visible = False
        #ElseIf UsarQueSocket = 2 Then
108         Command20.Visible = True
110         Command26.Visible = False
        #End If

        
        Exit Sub

Form_Load_Err:
        Call RegistrarError(Err.Number, Err.description, "frmServidor.Form_Load", Erl)
        Resume Next
        
End Sub
