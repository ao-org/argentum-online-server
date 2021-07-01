VERSION 5.00
Begin VB.Form frmConID 
   Caption         =   "ConID"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Liberar todos los slots"
      Height          =   390
      Left            =   135
      TabIndex        =   3
      Top             =   3495
      Width           =   4290
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ver estado"
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Top             =   3030
      Width           =   4290
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   180
      TabIndex        =   1
      Top             =   150
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   3975
      Width           =   4290
   End
   Begin VB.Label Label1 
      Height          =   510
      Left            =   180
      TabIndex        =   4
      Top             =   2430
      Width           =   4230
   End
End
Attribute VB_Name = "frmConID"
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
        
100     Unload Me

        
        Exit Sub

Command1_Click_Err:
102     Call TraceError(Err.Number, Err.Description, "frmConID.Command1_Click", Erl)
104
        
End Sub

Private Sub Command2_Click()
        
        On Error GoTo Command2_Click_Err
        

100     List1.Clear

        Dim c As Integer

        Dim i As Integer

102     For i = 1 To MaxUsers
104         List1.AddItem "UserIndex " & i & " -- " & UserList(i).ConnID

106         If UserList(i).ConnID <> -1 Then c = c + 1
108     Next i

110     If c = MaxUsers Then
112         Label1.Caption = "No hay slots vacios!"
        Else
114         Label1.Caption = "Hay " & MaxUsers - c & " slots vacios!"

        End If

        
        Exit Sub

Command2_Click_Err:
116     Call TraceError(Err.Number, Err.Description, "frmConID.Command2_Click", Erl)
118
        
End Sub

Private Sub Command3_Click()
        
        On Error GoTo Command3_Click_Err
        

        Dim i As Integer

100     For i = 1 To MaxUsers

102         If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida And Not UserList(i).flags.UserLogged Then Call CloseSocket(i)
104     Next i

        
        Exit Sub

Command3_Click_Err:
106     Call TraceError(Err.Number, Err.Description, "frmConID.Command3_Click", Erl)
108
        
End Sub

