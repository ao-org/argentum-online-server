VERSION 5.00
Begin VB.Form frmAdmin 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Administración del servidor"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Personajes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   495
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   720
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Echar todos los PJS no privilegiados"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   2280
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "R"
         Height          =   315
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.ComboBox cboPjs 
         Height          =   315
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Echar"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   1800
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmAdmin.frm
'
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

Private Sub Command1_Click()
        
        On Error GoTo Command1_Click_Err
        

        Dim uUser As t_UserReference

100     uUser = NameIndex(cboPjs.Text)

102     If IsValidUserRef(uUser) Then
104         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor » " & UserList(uUser.ArrayIndex).name & " ha sido hechado. ", e_FontTypeNames.FONTTYPE_SERVER))
106         Call CloseSocket(uUser.ArrayIndex)

        End If

        
        Exit Sub

Command1_Click_Err:
108     Call TraceError(Err.Number, Err.Description, "frmAdmin.Command1_Click", Erl)
110
        
End Sub

Public Sub ActualizaListaPjs()
        
        On Error GoTo ActualizaListaPjs_Err
        

        Dim LoopC As Long

100     With cboPjs
102         .Clear
    
104         For LoopC = 1 To LastUser

106             If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnIDValida Then
108                 If UserList(LoopC).flags.Privilegios And e_PlayerType.user Then
110                     .AddItem UserList(LoopC).Name
112                     .ItemData(.NewIndex) = LoopC

                    End If

                End If

114         Next LoopC

        End With

        
        Exit Sub

ActualizaListaPjs_Err:
116     Call TraceError(Err.Number, Err.Description, "frmAdmin.ActualizaListaPjs", Erl)
118
        
End Sub

Private Sub Command3_Click()
        
        On Error GoTo Command3_Click_Err
        
100     Call EcharPjsNoPrivilegiados

        
        Exit Sub

Command3_Click_Err:
102     Call TraceError(Err.Number, Err.Description, "frmAdmin.Command3_Click", Erl)
104
        
End Sub
