VERSION 5.00
Begin VB.Form frmUserList 
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Echar todos los no Logged"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualiza"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmUserList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmUserList.frm
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
        

        Dim LoopC As Integer

100     Text2.Text = "MaxUsers: " & MaxUsers & vbCrLf
102     Text2.Text = Text2.Text & "LastUser: " & LastUser & vbCrLf
104     Text2.Text = Text2.Text & "NumUsers: " & NumUsers & vbCrLf
        'Text2.Text = Text2.Text & "" & vbCrLf

106     List1.Clear

108     For LoopC = 1 To MaxUsers
110         List1.AddItem Format$(LoopC, "000") & " " & IIf(UserList(LoopC).flags.UserLogged, UserList(LoopC).name, "")
112         List1.ItemData(List1.NewIndex) = LoopC
114     Next LoopC

        
        Exit Sub

Command1_Click_Err:
116     Call TraceError(Err.Number, Err.Description, "frmUserList.Command1_Click", Erl)
118
        
End Sub

Private Sub Command2_Click()
        
        On Error GoTo Command2_Click_Err
        

        Dim LoopC As Integer

100     For LoopC = 1 To MaxUsers

102         If UserList(LoopC).ConnIDValida And Not UserList(LoopC).flags.UserLogged Then
104             Call CloseSocket(LoopC)

            End If

106     Next LoopC

        
        Exit Sub

Command2_Click_Err:
108     Call TraceError(Err.Number, Err.Description, "frmUserList.Command2_Click", Erl)
110
        
End Sub

Private Sub List1_Click()
        
        On Error GoTo List1_Click_Err
        

        Dim UserIndex As Integer

100     If List1.ListIndex <> -1 Then
102         UserIndex = List1.ItemData(List1.ListIndex)

104         If UserIndex > 0 And UserIndex <= MaxUsers Then

106             With UserList(UserIndex)
108                 Text1.Text = "UserLogged: " & .flags.UserLogged & vbCrLf
110                 Text1.Text = Text1.Text & "IdleCount: " & .Counters.IdleCount & vbCrLf
114                 Text1.Text = Text1.Text & "ConnIDValida: " & .ConnIDValida & vbCrLf

                End With

            End If

        End If

        
        Exit Sub

List1_Click_Err:
116     Call TraceError(Err.Number, Err.Description, "frmUserList.List1_Click", Erl)
118
        
End Sub
