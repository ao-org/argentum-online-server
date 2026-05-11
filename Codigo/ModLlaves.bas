Attribute VB_Name = "ModLlaves"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
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
' Cantidad máxima de llaves
Public Const MAXKEYS As Byte = 10

Public Sub SacarLlaveDeLLavero(ByVal UserIndex As Integer, ByVal Llave As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim i As Integer
        For i = 1 To MAXKEYS
            If .Keys(i) = Llave Then
                .Keys(i) = 0
                Call WriteUpdateUserKey(UserIndex, i, 0)
                Exit Sub
            End If
        Next
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModLlaves.SacarLlaveDeLLavero", Erl)
End Sub

Public Sub EnviarLlaves(ByVal UserIndex As Integer)
    On Error GoTo EnviarLlaves_Err
    With UserList(UserIndex)
        Dim i As Integer
        For i = 1 To MAXKEYS
            If .Keys(i) <> 0 Then
                Call WriteUpdateUserKey(UserIndex, i, .Keys(i))
            End If
        Next
    End With
    Exit Sub
EnviarLlaves_Err:
    Call TraceError(Err.Number, Err.Description, "ModLlaves.EnviarLlaves", Erl)
End Sub

Public Sub UsarLlave(ByVal UserIndex As Integer, ByVal Slot As Integer)
    On Error GoTo UsarLlave_Err

    ' Índices y coordenadas copiadas localmente para evitar re-leer estructuras mutables
    Dim targetObjIndex As Integer
    Dim keyObjIndex As Integer
    Dim targetMap As Integer
    Dim targetX As Integer
    Dim targetY As Integer

    ' Índices derivados del tile actual
    Dim currentTileObjIndex As Integer
    Dim newDoorObjIndex As Integer

    ' Copias locales de los objetos (más seguro y más claro)
    Dim TargObj As t_ObjData
    Dim LlaveObj As t_ObjData

    ' Anti-spam / cooldown de uso
    If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub

    ' Validación defensiva del UserIndex
    If UserIndex < LBound(UserList) Or UserIndex > UBound(UserList) Then Exit Sub

    With UserList(UserIndex)

        ' Validar que el slot esté dentro de rango permitido
        If Slot < 1 Or Slot > MAXKEYS Then
            Call LogEdicionPaquete("El usuario " & .Name & " editó el slot del llavero | Valor: " & Slot & ".")
            Exit Sub
        End If

        ' Obtener índice del objeto llave
        keyObjIndex = .Keys(Slot)
        If keyObjIndex = 0 Then Exit Sub   ' Slot vacío

        ' Obtener índice del objeto target (puerta)
        targetObjIndex = .flags.TargetObj
        If targetObjIndex = 0 Then Exit Sub   ' No hay target

        ' Copiar coordenadas del target (evita inconsistencias si cambian durante ejecución)
        targetMap = .flags.TargetObjMap
        targetX = .flags.TargetObjX
        targetY = .flags.TargetObjY
    End With

    ' Validar índices de objetos antes de indexar ObjData
    If keyObjIndex < LBound(ObjData) Or keyObjIndex > UBound(ObjData) Then Exit Sub
    If targetObjIndex < LBound(ObjData) Or targetObjIndex > UBound(ObjData) Then Exit Sub

    ' Validar posición del mapa ANTES de acceder a MapData (crítico para evitar crash)
    If Not LegalPos(targetMap, targetX, targetY) Then Exit Sub

    ' Leer UNA sola vez el objeto del tile
    currentTileObjIndex = MapData(targetMap, targetX, targetY).ObjInfo.ObjIndex

    ' Validar índice del tile
    If currentTileObjIndex <= 0 Then Exit Sub
    If currentTileObjIndex < LBound(ObjData) Or currentTileObjIndex > UBound(ObjData) Then Exit Sub

    ' Copiar datos de los objetos (evita múltiples accesos a arrays globales)
    TargObj = ObjData(targetObjIndex)
    LlaveObj = ObjData(keyObjIndex)

    ' Validar que el target sea una puerta
    If TargObj.OBJType <> e_OBJType.otDoors Then Exit Sub

    ' Si no está cerrada, no hay nada que hacer
    If TargObj.Cerrada <> 1 Then
        Call WriteLocaleMsg(UserIndex, MSG_NO_CERRADA, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If

    ' Validar que la llave coincida
    If TargObj.clave <> LlaveObj.clave Then
        Call WriteLocaleMsg(UserIndex, MSG_NO_LLAVE_SIRVE, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If

    ' Determinar el nuevo estado de la puerta (índice destino)
    ' IMPORTANTE: siempre validar antes de usar como índice
    If TargObj.Llave > 0 Then
        ' Caso: puerta cerrada con llave ? abrir
        newDoorObjIndex = ObjData(currentTileObjIndex).IndexCerrada

        If newDoorObjIndex <= 0 Then Exit Sub
        If newDoorObjIndex < LBound(ObjData) Or newDoorObjIndex > UBound(ObjData) Then Exit Sub

        ' Aplicar cambio en el mapa
        MapData(targetMap, targetX, targetY).ObjInfo.ObjIndex = newDoorObjIndex

        ' Actualizar target del usuario para reflejar el nuevo estado
        UserList(UserIndex).flags.TargetObj = newDoorObjIndex

        Call WriteLocaleMsg(UserIndex, MSG_ABIERTO_PUERTA, e_FontTypeNames.FONTTYPE_INFO)

    Else
        ' Caso: puerta que pasa a estado "cerrada con llave"
        newDoorObjIndex = ObjData(currentTileObjIndex).IndexCerradaLlave

        If newDoorObjIndex <= 0 Then Exit Sub
        If newDoorObjIndex < LBound(ObjData) Or newDoorObjIndex > UBound(ObjData) Then Exit Sub

        MapData(targetMap, targetX, targetY).ObjInfo.ObjIndex = newDoorObjIndex
        UserList(UserIndex).flags.TargetObj = newDoorObjIndex

        Call WriteLocaleMsg(UserIndex, MSG_CERRADO_LLAVE_PUERTA, e_FontTypeNames.FONTTYPE_INFO)
    End If

    Exit Sub

UsarLlave_Err:
    ' Log centralizado de errores para debugging
    Call TraceError(Err.Number, Err.Description, "ModLlaves.UsarLlave", Erl)
End Sub
