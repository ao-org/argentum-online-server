VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Chr Editor - Por Gonza_M"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6045
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   2880
      TabIndex        =   14
      Top             =   1200
      Width           =   2895
      Begin VB.CommandButton Command4 
         Caption         =   "Aplicar a Seleccionados"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   21
         Top             =   3000
         Width           =   2655
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Borrar Variable"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Width           =   2055
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Borrar Campo"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Campo               ej:  [FLAGS]"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Variable             ej:  Muerto="
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   2880
      TabIndex        =   6
      Top             =   1200
      Width           =   2895
      Begin VB.CommandButton Command3 
         Caption         =   "Aplicar a Seleccionados"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Valor                  ej:  1"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Variable             ej:  Muerto="
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Campo               ej:  [FLAGS]"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Borrar"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Seleccionar Ninguno"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Seleccionar Todos"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      ItemData        =   "Form1.frx":08CA
      Left            =   120
      List            =   "Form1.frx":08CC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Editor de Multiples Charfiles"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim k As Integer

Private Sub Command1_Click()
For k = 0 To List1.ListCount - 1
    List1.Selected(k) = True
Next k
End Sub

Private Sub Command2_Click()
For k = 0 To List1.ListCount - 1
    List1.Selected(k) = False
Next k
End Sub

Private Sub Command3_Click()
If List1.SelCount > 0 Then
If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then 'validacion
    For k = 0 To List1.ListCount - 1
        If List1.Selected(k) Then
            WriteVar App.Path & "\" & List1.List(k) & ".chr", Text1.Text, Text2.Text, Text3.Text
        End If
    Next
    MsgBox "Los datos fueron agregados correctamente"
Else
    MsgBox "error: falta ingresar datos"
End If
Else
MsgBox "No hay ningun Charfile seleccionado"
End If

End Sub

Private Sub Command4_Click()
If Option4.value = True Then
    If Text4.Text <> "" Then
        If List1.SelCount > 0 Then
            For k = 0 To List1.ListCount - 1
                WriteVar App.Path & "\" & List1.List(k) & ".chr", Text4.Text, vbNullString, vbNullString
            Next k
            MsgBox "Los datos fueron borrados correctamente"
        Else
            MsgBox "No hay ningun Charfile seleccionado"
            Exit Sub
        End If
    Else
        MsgBox "error: falta ingresar datos"
        Exit Sub
    End If
ElseIf Option5.value = True Then
    If Text4.Text <> "" And Text5.Text <> "" Then
        If List1.SelCount > 0 Then
            For k = 0 To List1.ListCount - 1
                WriteVar App.Path & "\" & List1.List(k) & ".chr", Text4.Text, Text5.Text, vbNullString
            Next k
            MsgBox "Los datos fueron borrados correctamente"
        Else
            MsgBox "No hay ningun Charfile seleccionado"
            Exit Sub
        End If
    Else
        MsgBox "error: falta ingresar datos"
    End If
End If
End Sub


Private Sub Form_Activate()
Option1.value = True
Option4.value = True
Frame2.Visible = False

End Sub

Private Sub Form_Load()
Dim fso As New FileSystemObject
Dim Carpeta As Folder
Dim fil As file
Set Carpeta = fso.GetFolder(App.Path)
k = 0
For Each fil In Carpeta.Files
    If Right(fil.Name, 4) = ".chr" Then 'solo enlista a los arhivos .chr
        List1.AddItem Left(fil.Name, Len(fil.Name) - 4), k
    End If
Next
If List1.ListCount = 0 Then
    MsgBox "No se encontró ningún Charfile. Acordate que este programa debe ubicarse en la carpeta 'Charfile' del server."
    End
End If
End Sub


Private Sub Option1_Click()
Frame1.Visible = True
Frame2.Visible = False

End Sub

Private Sub Option2_Click()
Frame1.Visible = False
Frame2.Visible = True

End Sub

Private Sub Option3_Click()
Frame1.Visible = False
Frame2.Visible = False

End Sub
