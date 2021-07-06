VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRequestHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_xmlHttp As MSXML2.ServerXMLHTTP60
 
Public Sub Initialize(ByRef xmlHttpRequest As MSXML2.ServerXMLHTTP60)
   Set m_xmlHttp = xmlHttpRequest
End Sub
 
Sub OnReadyStateChange()
Attribute OnReadyStateChange.VB_UserMemId = 0
   
   If m_xmlHttp.readyState = 4 Then
    
        If m_xmlHttp.Status <> 200 Then
        
            Call TraceErrorAPI(m_xmlHttp.Status, m_xmlHttp.statusText, m_xmlHttp.ResponseText)
        
        End If
     
   End If
   
End Sub
