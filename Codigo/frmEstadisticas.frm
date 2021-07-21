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
         TabIndex        =   12
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cboUsusColas 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblColas 
         Height          =   255
         Left            =   120
         TabIndex        =   10
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

Private Sub Command2_Click()
        
        On Error GoTo Command2_Click_Err
        
100     frmAdmin.Show
102     frmAdmin.cboPjs.Text = cboUsusColas.Text

        
        Exit Sub

Command2_Click_Err:
104     Call RegistrarError(Err.Number, Err.Description, "frmEstadisticas.Command2_Click", Erl)
106
        
End Sub

Private Sub Form_Activate()
        
        On Error GoTo Form_Activate_Err
        
100     Call ActualizaStats

        
        Exit Sub

Form_Activate_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmEstadisticas.Form_Activate", Erl)
104
        
End Sub

Public Sub ActualizaStats()
        
        On Error GoTo ActualizaStats_Err
        
100     lblStat(0).Caption = TCPESStats.BytesRecibidosXSEG
102     lblStat(1).Caption = TCPESStats.BytesEnviadosXSEG
104     lblStat(2).Caption = TCPESStats.BytesRecibidosXSEGMax & vbCrLf & TCPESStats.BytesRecibidosXSEGCuando
106     lblStat(3).Caption = TCPESStats.BytesEnviadosXSEGMax & vbCrLf & TCPESStats.BytesEnviadosXSEGCuando

        
        Exit Sub

ActualizaStats_Err:
108     Call RegistrarError(Err.Number, Err.Description, "frmEstadisticas.ActualizaStats", Erl)
110
        
End Sub

Private Sub Form_Click()
        
        On Error GoTo Form_Click_Err
        
100     Call ActualizaStats

        
        Exit Sub

Form_Click_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmEstadisticas.Form_Click", Erl)
104
        
End Sub

Private Sub Frame1_Click()
        
        On Error GoTo Frame1_Click_Err
        
100     Call ActualizaStats

        
        Exit Sub

Frame1_Click_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmEstadisticas.Frame1_Click", Erl)
104
        
End Sub

Private Sub lblStat_Click(Index As Integer)
        
        On Error GoTo lblStat_Click_Err
        
100     Call ActualizaStats

        
        Exit Sub

lblStat_Click_Err:
102     Call RegistrarError(Err.Number, Err.Description, "frmEstadisticas.lblStat_Click", Erl)
104
        
End Sub
