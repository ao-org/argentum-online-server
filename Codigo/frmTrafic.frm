VERSION 5.00
Begin VB.Form frmTrafic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trafico"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   585
      Left            =   90
      TabIndex        =   1
      Top             =   2250
      Width           =   960
   End
   Begin VB.ListBox lstTrafico 
      Height          =   2010
      Left            =   60
      TabIndex        =   0
      Top             =   135
      Width           =   4455
   End
End
Attribute VB_Name = "frmTrafic"
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

Private Sub Command1_Click()
    On Error Goto Command1_Click_Err

100     Me.Visible = False
   
    Exit Sub
Command1_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmTrafic.Command1_Click", Erl)
End Sub
