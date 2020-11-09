VERSION 5.00
Begin VB.Form frmEstadisticas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stats"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Colas"
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   5175
      Begin VB.CommandButton Command2 
         Caption         =   "Adm"
         Height          =   315
         Left            =   2640
         TabIndex        =   13
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cboUsusColas 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "R"
         Height          =   375
         Left            =   4560
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblColas 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.Label lblStat 
         Height          =   495
         Index           =   3
         Left            =   2400
         TabIndex        =   8
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblStat 
         Height          =   495
         Index           =   2
         Left            =   2400
         TabIndex        =   7
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lblStat 
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   6
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblStat 
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "MAX Bytes Enviados x Seg:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "MAX Bytes Recibidos x Seg:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Bytes Enviados x Seg:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Bytes Recibidos x Seg:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmEstadisticas.frm
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

    Dim LoopC As Integer, n As Long, M As Long

    n = 0 'numero de pjs
    M = 0 'numero total de elementos en cola

    If cboUsusColas.ListCount > 0 Then cboUsusColas.Clear

    For LoopC = 1 To LastUser

        If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
            If UserList(LoopC).outgoingData.length > 0 Then
                n = n + 1
                M = M + UserList(LoopC).outgoingData.length
                cboUsusColas.AddItem UserList(LoopC).name

            End If

        End If

    Next LoopC

    lblColas.Caption = n & " PJs, " & M & " elementos en las colas."

    If cboUsusColas.ListCount > 0 Then cboUsusColas.ListIndex = 0
    
End Sub

Private Sub Command2_Click()
    frmAdmin.Show
    frmAdmin.cboPjs.Text = cboUsusColas.Text

End Sub

Private Sub Form_Activate()
    Call ActualizaStats

End Sub

Public Sub ActualizaStats()
    lblStat(0).Caption = TCPESStats.BytesRecibidosXSEG
    lblStat(1).Caption = TCPESStats.BytesEnviadosXSEG
    lblStat(2).Caption = TCPESStats.BytesRecibidosXSEGMax & vbCrLf & TCPESStats.BytesRecibidosXSEGCuando
    lblStat(3).Caption = TCPESStats.BytesEnviadosXSEGMax & vbCrLf & TCPESStats.BytesEnviadosXSEGCuando

End Sub

Private Sub Form_Click()
    Call ActualizaStats

End Sub

Private Sub Frame1_Click()
    Call ActualizaStats

End Sub

Private Sub lblStat_Click(Index As Integer)
    Call ActualizaStats

End Sub
