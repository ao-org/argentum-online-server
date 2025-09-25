VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmCargando 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online Server"
   ClientHeight    =   2820
   ClientLeft      =   1410
   ClientTop       =   3000
   ClientWidth     =   6570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCargando.frx":0000
   ScaleHeight     =   237.208
   ScaleMode       =   0  'User
   ScaleWidth      =   438
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar cargar 
      Height          =   300
      Left            =   1125
      TabIndex        =   1
      Top             =   2160
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   529
      _Version        =   327682
      Appearance      =   0
      Min             =   1e-4
   End
   Begin VB.Label ToMapLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000015&
      BackStyle       =   0  'Transparent
      Caption         =   " Cargando Mapas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   2
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label lblDragForm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   1905
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6525
   End
End
Attribute VB_Name = "frmCargando"
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
Private Declare Function SetWindowPos _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal x As Long, _
                              ByVal y As Long, _
                              ByVal cx As Long, _
                              ByVal cy As Long, _
                              ByVal wFlags As Long) As Long
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
' For the label that allows yo to move the form
Private mlngX As Long
Private mlngY As Long

Private Sub Form_Load()
    ' Mostramos este form arriba de todo.
    Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE)
    Call Me.ZOrder(0)
End Sub

Private Sub lblDragForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        mlngX = x
        mlngY = y
    End If
End Sub

Private Sub lblDragForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngLeft As Long
    Dim lngTop  As Long
    If Button = vbLeftButton Then
        lngLeft = Me.Left + x - mlngX
        lngTop = Me.Top + y - mlngY
        Call Me.Move(lngLeft, lngTop)
    End If
End Sub
