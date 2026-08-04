# AI Context — Argentum Online (AO20)

**VB6 Official Documentation:** https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/visual-basic-6.0-documentation

## VB6 Build Commands

**Compiler:** `C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6.exe`

### Server

```cmd
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6.exe" /make ^
/d UsarQueSocket=1:ConUpTime=1:AntiExternos=0:Lac=1:DEBUGGING=0:PYMMO=1:UNLOCK_CPU=1:DIRECT_PLAY=0 ^
/out vb6build.log ^
Server.VBP
```

### Client

```cmd
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6.exe" /make ^
/d Compresion=1:DEBUGGING=1:PYMMO=1:ENABLE_ANTICHEAT=1:REMOTE_CLOSE=1:BATTLESERVER=1:DIRECT_PLAY=0:ENABLE_BASS=1 ^
/out vb6build.log ^
Client.VBP
```

---

## For AI Agents

---

## Project Context

You are an experienced Visual Basic 6 developer working on the legacy MMORPG Argentum Online.
The project uses a client-server architecture, where both the server and the client are written in VB6.
Communication between them occurs via a custom packet protocol.
All development must follow legacy-compatible practices and the coding standards below.
Prioritize clean, readable code with minimal risk of regression.
Avoid modern syntax not supported by VB6 (do NOT use VB.NET syntax).

### Repositories

| Repository | Type | Access | URL |
|------------|------|--------|-----|
| **Assets** | Resources/Config | Private | https://github.com/ao-org/argentum-online-assets/ |
| **Server** | VB6 Server Code | Public | https://github.com/ao-org/argentum-online-server/ |
| **Client** | VB6 Client Code | Public | https://github.com/ao-org/argentum-online-client |

- **Assets (Recursos)**: Graphics, sounds, maps, configuration files, shared resources
- **Server**: Game server logic, database, NPCs, combat systems
- **Client**: Game client, UI, rendering, input handling

---

## VB6 Coding Rules

### 1. Mandatory use of `Call` and parentheses

Always use `Call` when invoking any `Sub`. Always include parentheses, even if there are no arguments.

```vb
Call GuardarDatos()
Call EnviarMensaje("Hola", 2)
```

### 2. Functions always with parentheses

Even when ignoring the return value, functions must be called with parentheses.

```vb
Call ObtenerTiempoActual()
Dim puntos As Long
puntos = CalcularPuntos(usuarioId)
```

### 3. Naming conventions

| Element        | Convention              | Example                   |
| -------------- | ----------------------- | ------------------------- |
| Modules        | `mod` prefix            | `modNetwork`, `modLogin`  |
| Forms          | `frm` prefix            | `frmMain`, `frmLogin`     |
| Controls       | Hungarian notation      | `txtNombre`, `lblError`   |
| Variables      | `camelCase`             | `userIndex`, `goldAmount` |
| Constants      | `UPPER`                 | `GOLD_PRICE`              |
| Functions/Subs | `PascalCase`            | `Call ValidarUsuario()`   |
| Enums          | `e_` prefix             | `e_TipoPago`              |

### 4. Spacing and formatting

- Use 4-space indentation.
- One blank line between procedures.
- Align related variable declarations:

```vb
Dim userGold    As Long
Dim userSilver  As Long
Dim userName    As String
```

### 5. Avoid magic numbers

Declare all fixed values as constants in shared modules.

```vb
Public Const GOLD_PRICE As Long = 50000

If .Stats.GLD < GOLD_PRICE Then
    Call EscribirError("No tenés suficiente oro.")
End If
```

### 6. Standardized error handling

Always use `On Error GoTo Name_Err` with a label at the end of the `Sub`.

```vb
Private Sub ValidarSesion()
    On Error GoTo ValidarSesion_Err

    Call HacerAlgo()

    Exit Sub

ValidarSesion_Err:
    Call TraceError(Err.Number, Err.Description, "modLogin.ValidarSesion", Erl)
End Sub
```

### 7. Explicit identifiers

Use clear, specific names. Avoid generic names like `dato`, `res`, `temp`.

```vb
Dim creditAmount As Long
Dim connectionId As Integer
```

### 8. SQL queries

Always use parameterized queries (`?`). Always close all `Recordset` objects.

```vb
Dim RS As ADODB.Recordset
Set RS = Query("SELECT nivel FROM user WHERE id = ?;", userId)

If Not RS.EOF Then
    nivel = RS!nivel
End If

Call RS.Close
```

### 9. Localized messages

#### Server-side (`WriteLocaleMsg`)

The server sends a message ID; the client resolves it from a message index file.

```vb
Call WriteLocaleMsg(UserIndex, "1291", FONTTYPE_INFOBOLD, GOLD_PRICE)
```

`"1291"` maps to: *"You need at least ¬1 gold to sell your character"* — `GOLD_PRICE` replaces `¬1`.

Message files:
- [Spanish – SP_LocalMsg.dat](https://github.com/ao-org/Recursos/blob/master/init/SP_LocalMsg.dat)
- [English – EN_LocalMsg.dat](https://github.com/ao-org/Recursos/blob/master/init/EN_LocalMsg.dat)

To regenerate: run `python generar_localindex.py` or `generar_localindex.exe` from `/tools/`.

#### Client-side (`JsonLanguage`)

Client-side UI translation uses JSON files in the [`Languages`](https://github.com/ao-org/argentum-online-client/tree/master/Languages) folder:

- [Languages/1.json (Spanish)](https://github.com/ao-org/argentum-online-client/blob/master/Languages/1.json)
- [Languages/2.json (English)](https://github.com/ao-org/argentum-online-client/blob/master/Languages/2.json)

```vb
Call MsgBox(JsonLanguage.Item("MENSAJE_ERROR_CARGAR_OPCIONES"), vbCritical, JsonLanguage.Item("TITULO_ERROR_CARGAR"))
```

### 10. Clear control flow

Use `Exit Sub` for early exits. Avoid unnecessary nesting. Avoid GoTo flow control statements except for error handling, or unless absolutely necessary for performance issues.

```vb
If Not EstaAutenticado(UserIndex) Then
    Call Desconectar(UserIndex)
    Exit Sub
End If
```

### 11. Pure functions and no side effects

Prefer `Function` for validations or transformations. Avoid modifying global variables unnecessarily.

### 12. Files Encoding

All source code files must be saved with Windows-1252 encoding to ensure compatibility with VB6.

### 13. Function parameters

Avoid using explicit parameter declaration for syntactic sugar like
```vb
Public Function GetUserName(ByVal userId As Long = 0, Optional Something As Boolean = False) As String
End Function
Call GetUserName(userId, Something:= False)
```

Try to use simple ByVal parameters unless necessary. (for example when refactoring legacy code that uses ByRef).

### 14. Helper Functions

Matematicas.bas — Mathematical operations helper functions
