VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEstadisticasIPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsEstadisticasIPC.cls
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

Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private sMensaje As Long, hVentana As Long, hVentanaMia As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Private Const GW_HWNDFIRST = 0

Private Const GW_HWNDNEXT = 2

'*************************************************
Public Enum EstaNotificaciones

    CANTIDAD_ONLINE = 1
    RECORD_USUARIOS = 2
    UPTIME_SERVER = 3
    CANTIDAD_MAPAS = 4
    EVENTO_NUEVO_CLAN = 5
    
    HANDLE_WND_SERVER = 100

End Enum

'*************************************************

Private Function BuscaVentana(Wnd As Long, str As String) As Long
        
        On Error GoTo BuscaVentana_Err
        

        Dim W As Long

        Dim L As Long

        Dim T As String

100     W = GetWindow(Wnd, GW_HWNDFIRST)

102     While W <> 0

104         L = GetWindowTextLength(W)

106         If L > 0 Then
108             T = Space(L + 1)
110             L = GetWindowText(W, T, L + 1)

112             If Left$(T, Len(str)) = str Then
114                 BuscaVentana = W
                    Exit Function

                End If

            End If

116         W = GetWindow(W, GW_HWNDNEXT)
        Wend

118     BuscaVentana = 0

        
        Exit Function

BuscaVentana_Err:
120     Call TraceError(Err.Number, Err.Description, "clsEstadisticasIPC.BuscaVentana", Erl)
122
        
End Function

Public Function Informar(ByVal QueCosa As EstaNotificaciones, ByVal Parametro As Long) As Long
        
        On Error GoTo Informar_Err
        
100     Call BuscaWndEstadisticas

102     If hVentana <> 0 Then
104         Informar = SendMessageLong(hVentana, sMensaje, QueCosa, Parametro)

        End If

        
        Exit Function

Informar_Err:
106     Call TraceError(Err.Number, Err.Description, "clsEstadisticasIPC.Informar", Erl)
108
        
End Function

Public Function EstadisticasAndando() As Boolean
        
        On Error GoTo EstadisticasAndando_Err
        

100     Call BuscaWndEstadisticas
        'Ret = SendNotifyMessage(hVentana, sMensaje, 0, 0)
102     EstadisticasAndando = (hVentana <> 0)

        
        Exit Function

EstadisticasAndando_Err:
104     Call TraceError(Err.Number, Err.Description, "clsEstadisticasIPC.EstadisticasAndando", Erl)
106
        
End Function

Public Sub Inicializa(ByVal hWnd As Long)
        
        On Error GoTo Inicializa_Err
        
100     hVentanaMia = hWnd
102     sMensaje = RegisterWindowMessage("EstadisticasAO")

        
        Exit Sub

Inicializa_Err:
104     Call TraceError(Err.Number, Err.Description, "clsEstadisticasIPC.Inicializa", Erl)
106
        
End Sub

Private Sub BuscaWndEstadisticas()
        
        On Error GoTo BuscaWndEstadisticas_Err
        
100     hVentana = BuscaVentana(hVentanaMia, "Servidor de estadisticas AO")

        
        Exit Sub

BuscaWndEstadisticas_Err:
102     Call TraceError(Err.Number, Err.Description, "clsEstadisticasIPC.BuscaWndEstadisticas", Erl)
104
        
End Sub
