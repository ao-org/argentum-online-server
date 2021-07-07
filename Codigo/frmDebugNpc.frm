VERSION 5.00
Begin VB.Form frmDebugNpc 
   Caption         =   "DebugNpcs"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   300
      Left            =   90
      TabIndex        =   5
      Top             =   2085
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ActualizarInfo"
      Height          =   300
      Left            =   90
      TabIndex        =   2
      Top             =   1755
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   "MaxNpcs:"
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   1380
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "LastNpcIndex:"
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   1065
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Npcs Libres:"
      Height          =   285
      Left            =   105
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Npcs Activos:"
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "frmDebugNpc"
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
        

        Dim i As Integer, K As Integer

100     For i = 1 To LastNPC

102         If NpcList(i).flags.NPCActive Then K = K + 1
104     Next i

106     Label1.Caption = "Npcs Activos:" & K
108     Label2.Caption = "Npcs Libres:" & MaxNPCs - K
110     Label3.Caption = "LastNpcIndex:" & LastNPC
112     Label4.Caption = "MAXNPCS:" & MaxNPCs

        
        Exit Sub

Command1_Click_Err:
114     Call TraceError(Err.Number, Err.Description, "frmDebugNpc.Command1_Click", Erl)
116
        
End Sub

Private Sub Command2_Click()
        
        On Error GoTo Command2_Click_Err
        
100     Unload Me

        
        Exit Sub

Command2_Click_Err:
102     Call TraceError(Err.Number, Err.Description, "frmDebugNpc.Command2_Click", Erl)
104
        
End Sub

