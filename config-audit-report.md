# Config File Audit Report — Argentum Online 20 Server

**Date:** April 19, 2026
**Method:** Cross-referenced every key in config files against VB6 source code (`Codigo/*.cls`, `Codigo/*.bas`) to identify dead/unused config entries.

---

## Summary

| File | Total Keys | Used | Dead/Unused | Notes |
|---|---|---|---|---|
| Configuracion.ini | ~75 keys | 73 | 2 | All sections are used |
| intervalos.ini | 45 keys | 40 | 5 | 5 keys never read by server |
| PacketRatePolicy.ini | 32 keys | 32 | 0 | All read dynamically in a loop |
| Partition.ini | — | — | — | Wrong filename (see below) |
| feature_toggle.ini | 43 keys | 43 | 0 | All read dynamically |
| Server.ini | ~35 keys | 35 | 0 | Clean |

---

## Dead Config — Safe to Remove

### intervalos.ini

| Line | Key | Reason |
|---|---|---|
| 18 | `IntervaloMover=200` | Never read by any VB6 source file. No variable references it. |
| 25 | `IntervaloPerdidaStaminaLluvia=9999` | **Commented out** in FileIO.bas (line 2247). Code reads: `'frmMain.tLluvia.Interval = val(...)` |
| 26 | `IntervaloTimerExec=50` | Never read by any VB6 source file. No variable references it. |
| 29 | `IntervaloAutoReiniciar=-1` | Never read by any VB6 source file. No variable references it. |
| 38 | `IntervaloMetamorfosis=800` | Never read by any VB6 source file. No variable references it. |
| 55 | `MargenDeIntervaloPorPing=30` | Never read by any VB6 source file. No variable references it. |

**Total: 6 dead keys in intervalos.ini**

---

## File Issues

### Partition.ini — Wrong Filename

The server code in `modPartition.bas` reads from `Partitions.ini` (with an 's'):
```vb
iniPath = App.path & "\Partitions.ini"
```

But the file in the repo is named `Partition.ini` (without 's'). **This file is never loaded by the server.** Either:
- Rename `Partition.ini` → `Partitions.ini`
- Or confirm this is intentional and the file is only used externally

---

## Config That Looks Suspicious But IS Used

These keys triggered warnings in the validator but are actually read by the server through patterns the auto-generator didn't catch:

| File | Key/Section | How It's Read |
|---|---|---|
| Configuracion.ini `[EVENTOS]` | Keys `0` through `23` | Read via `GetVar()` in `ModEventos.bas:180` |
| Configuracion.ini `OROPORNIVEL` | In `[CONFIGURACIONES]` | Read via `mSettings.Add` with `val()` wrapper (no type function) in `ServerConfig.cls:89` |
| Configuracion.ini `INSTANCEMAPS` | In `[INIT]` (not `[CONFIGURACIONES]`) | Read via `Lector.GetValue("INIT", "InstanceMaps")` in `FileIO.bas:1837` — **note: it's in `[INIT]` in the VB6 code but placed under `[CONFIGURACIONES]` in the config file. The server reads it from `[INIT]` in Server.ini, not from Configuracion.ini.** |
| Configuracion.ini `NPCMaxStepsPathFinding` | In `[AI]` | Read via `mSettings.Add` with `Min()` wrapper in `ServerConfig.cls:122` |
| Configuracion.ini `NPCMaxVisionRange` | In `[AI]` | Read via `mSettings.Add` with `Min()` wrapper in `ServerConfig.cls:124` |
| PacketRatePolicy.ini | All 16 sections | Read dynamically in a `For` loop in `FileIO.bas:LoadPacketRatePolicy` |

---

## Potential Data Issue

| File | Line | Key | Issue |
|---|---|---|---|
| Configuracion.ini | 38 | `InstanceMaps=100` | This key is under `[CONFIGURACIONES]` but the server reads `InstanceMaps` from `[INIT]` in **Server.ini** (FileIO.bas:1837). The value `50` in Server.ini is what the server actually uses. The `100` in Configuracion.ini is ignored. |
