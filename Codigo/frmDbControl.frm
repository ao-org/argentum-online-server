VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDbControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Db Control"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8610
   LinkTopic       =   "DbControl"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRastrearBoveda 
      Caption         =   "Rastrear OBJ boveda"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   670
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdRastrearInventario 
      Caption         =   "Rastrear OBJ inventario"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   420
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin ComctlLib.ProgressBar pbarDb 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   7320
      Visible         =   0   'False
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Max             =   1000
   End
   Begin VB.CommandButton cmdActualizarObjetos 
      Caption         =   "Actualizar objetos DB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   500
      Width           =   2175
   End
   Begin VB.CommandButton cmdBoveda 
      Caption         =   "Ver Boveda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   500
      Width           =   1695
   End
   Begin VB.CommandButton cmdInventario 
      Caption         =   "Ver inventario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   500
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get users"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   -840
      TabIndex        =   3
      Top             =   420
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   11033
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdEjecutarQuery 
      Caption         =   "Ejecutar Query"
      Height          =   255
      Left            =   7200
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtQuery 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmDbControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdActualizarObjetos_Click()
  On Error Resume Next
        If MsgBox("La siguiente acción es demasiado costosa para el servidor, ¿Desea continuar?", vbYesNo) = vbYes Then
            pbarDb.Visible = True
            Dim Object As Integer
            Dim RS As Recordset
                
            Dim Leer   As clsIniManager
102         Set Leer = New clsIniManager
104         Call Leer.Initialize(DatPath & "Obj.dat")
            Command3.Enabled = False
            
            Command3.Caption = "Actualizando..."
            'obtiene el numero de obj
106         NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))
            
            Dim ObjKey As String
            Set RS = Query("delete from object")
            pbarDb.max = NumObjDatas
            'Llena la lista
118         For Object = 1 To NumObjDatas
122             ObjKey = "OBJ" & Object
                Query ("INSERT INTO object (number, name) VALUES (" & Object & ", '" & Leer.GetValue(ObjKey, "Name") & "')")
                Debug.Print Object
                pbarDb.Value = Object
644         Next Object
    
646         Set Leer = Nothing
        
            Command3.Enabled = True
            
            Command3.Caption = "Actualizar objetos DB"
            pbarDb.Visible = False
        
        End If
End Sub

Private Sub cmdBoveda_Click()
    If txtQuery.Text <> "" Then
        Call getData("select o.number, o.name, i.amount from bank_item i inner join object o on i.item_id = o.number  where user_id = (select id from user where name = '" & txtQuery.Text & "') and amount > 0")
    End If
End Sub

Private Sub cmdEjecutarQuery_Click()
   Call getData(txtQuery.Text)
End Sub

Private Sub cmdInventario_Click()
    If txtQuery.Text <> "" Then
        Call getData("select o.number, o.name, i.amount from inventory_item i inner join object o on i.item_id = o.number  where user_id = (select id from user where name = '" & txtQuery.Text & "') and amount > 0")
    End If
End Sub

Private Sub getData(ByVal queryStr As String)
     
    Dim RS As Recordset
    
    Set RS = Query(queryStr)
        
    If Not RS Is Nothing Then
        Set DataGrid1.DataSource = RS
    End If
    DataGrid1.DefColWidth = 0
End Sub

Private Sub cmdRastrearBoveda_Click()
    Call getData("select u.name, o.name, bi.amount from user u inner join bank_item bi on u.id = bi.user_id inner join object o on bi.item_id = o.number where o.name like '%" & txtQuery.Text & "%' and bi.amount > 0 order by bi.amount desc")
End Sub

Private Sub cmdRastrearInventario_Click()
    Call getData("select u.name, o.name, ii.amount from user u inner join inventory_item ii on u.id = ii.user_id inner join object o on ii.item_id = o.number where o.name like '%" & txtQuery.Text & "%' and ii.amount > 0 order by ii.amount desc")
End Sub

Private Sub Command2_Click()
    Call getData("select * from user")
End Sub
