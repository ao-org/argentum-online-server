# Copyright (C) 2026 Noland Studios LTD
# Licensed under the GNU Affero General Public License, version 3 or later.
"""Generate a native VB6 harness from the production binary reader and flag helpers.

The harness stubs gameplay services (NPC spawning, logging and graphics data),
but executes the actual section parsing, validation and map-property application.
No server process or database is started. Build/run commands are printed.
"""
from pathlib import Path
import re
import argparse

ROOT = Path(__file__).resolve().parents[1]

def procedure(text, name):
    return re.search(r'^(?:Public |Private )?(?:Sub|Function) '+name+r'\b.*?^End (?:Sub|Function)', text, re.M | re.S)[0]

def write(path, text):
    path.write_bytes(text.replace('\r\n', '\n').replace('\n', '\r\n').encode('cp1252'))

def generate(output):
    output.mkdir(parents=True, exist_ok=True)
    source = (ROOT/'Codigo/FileIO.bas').read_text(encoding='cp1252')
    types = source[source.index('Private Type t_Position'):source.index('Private FeatureToggles')]
    loader = procedure(source, 'CargarMapaFormatoCSM')
    # Metadata-to-game-policy processing after Close does not parse any bytes.
    loader = loader[:loader.index('    Close #fh')] + '    Close #fh\n    Exit Sub\n' + loader[loader.index('\nErrorHandler:'):]
    helpers = '\n\n'.join(procedure(source, name) for name in ['RequireCsmBytes','CsmSectionSize','ValidateCsmCoordinate','ReadCsmString','ReadCsmMetadata','LegacyMapZoneFlags','BuildLegacyRoofSeams','SaveMapCsm3','WriteCsmInt16','WriteCsmInt32','WriteCsmByte','WriteCsmPosition','WriteCsmString','WriteCsmMetadata','IsActiveRoofSeam'])
    tile_source = (ROOT/'Codigo/modTileProperties.bas').read_text(encoding='cp1252')
    helpers += '\n\n' + '\n\n'.join(procedure(tile_source, name) for name in ['HasTileFlag','HasZoneFlag','HasMapZoneFlag','SetMapZoneFlag','IsWaterTile','ConvertLegacyTrigger','TilePropertyKey','IsMapDataCoordinate','CopyPropertyDictionary','IsInPvPArena','BothInPvPArena','CrossesPvPArenaBoundary','IsPrisonMap','CastleAtTile','RegisterCastleEntrance','RemoveCastleEntrances','SetTileTriggerFlags','ReplayTileProperties'])
    helpers += '\n\n' + procedure((ROOT/'Codigo/Protocol_Writes.bas').read_text(encoding='cp1252'), 'WriteHooTileProperties')
    helpers += '\n\n' + procedure((ROOT/'Codigo/General.bas').read_text(encoding='cp1252'), 'RunScriptInFile')
    sql_source = (ROOT/'Codigo/modSqlScripts.bas').read_text(encoding='cp1252')
    helpers += '\n\n' + '\n\n'.join(procedure(sql_source, name) for name in ['SplitMigrationSql','AddSqlMigrationStatement'])
    declares = (ROOT/'Codigo/Declares.bas').read_text(encoding='cp1252')
    enums = '\n\n'.join(re.search('Public Enum '+name+r'\n.*?End Enum', declares, re.S)[0] for name in ['e_Trigger','e_ZoneFlags'])
    code = 'Attribute VB_Name = "MapReaderUnderTest"\nOption Explicit\nPrivate Const CSM_FIVE_LAYER_SIGNATURE As Long = &H324C3557\nPrivate Const CSM3_SIGNATURE As Long = &H334D5343\nPrivate Const MAX_RANDOM_TELEPORT_IN_MAP As Long = 20\n' + enums + '\n' + types + '\n' + loader + '\n' + helpers
    code += """
Public Function TestLegacyFlags(ByVal restrictions As String, ByVal oldBytes As Byte) As Long
    Dim metadata As t_MapDat
    metadata.restrict_mode = restrictions
    metadata.Seguro = oldBytes
    metadata.backup_mode = oldBytes
    metadata.lluvia = oldBytes
    metadata.Nieve = oldBytes
    metadata.niebla = oldBytes
    TestLegacyFlags = LegacyMapZoneFlags(1, metadata)
End Function
"""
    write(output/'reader.bas', code)
    harness = (ROOT/'tools/map_schema_harness.bas').read_text(encoding='cp1252')
    write(output/'harness.bas', harness)
    castle_source = (ROOT/'Codigo/ModCastle.bas').read_text(encoding='cp1252')
    castle_types = castle_source[castle_source.index('Private Type t_CastleCoordinates'):castle_source.index('Public Enum eCastleWhitelistOperation')]
    castle_constants = '\n'.join(line for line in castle_source.splitlines() if line.startswith('Private Const Castle') or line.startswith('Private Const CASTLE_MOCKUP_OBJ_INDEX'))
    castle_code = 'Attribute VB_Name = "CastleUnderTest"\nOption Explicit\n' + castle_types + '\nPublic CastleData() As t_CastleInfo\n' + castle_constants + '\n'
    castle_code += '\n\n'.join(procedure(castle_source, name).replace('Private Function', 'Public Function', 1) for name in ['GetCastleSlotById','IsCastleFootprintInMapBounds','IsEmperorCastleCreated','CheckCastleEntryWhiteList','CanPublishCastle'])
    write(output/'castle.bas', castle_code)
    write(output/'send_data.bas', '''Attribute VB_Name = "modSendData"
Option Explicit
Public Sub SendData(ByVal target As Long, ByVal recipient As Integer)
    Call CapturePacket(recipient)
End Sub
''')
    write(output/'tests.vbp', '''Type=Exe
Reference=*\\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\\Windows\\SysWOW64\\scrrun.dll#Microsoft Scripting Runtime
Reference=*\\G{88D89F7C-BB59-4F13-9850-93CBD874039D}#1.5#0#..\\..\\Aurora.Network.dll#Aurora.Network
Module=MapReaderUnderTest; reader.bas
Module=MapSchemaHarness; harness.bas
Module=CastleUnderTest; castle.bas
Module=modSendData; send_data.bas
Startup="Sub Main"
Name="MapSchemaTests"
ExeName32="map_schema_tests.exe"
CompilationType=0
''')
    project = (output/'tests.vbp').read_text(encoding='cp1252')
    ado_reference = next(line for line in (ROOT/'Server.VBP').read_text(encoding='cp1252').splitlines() if 'Microsoft ActiveX Data Objects' in line)
    write(output/'tests.vbp', project.replace('Type=Exe', 'Type=Exe\n' + ado_reference, 1))
    print(output/'tests.vbp')

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, default=ROOT/'build/map-schema-tests')
    generate(parser.parse_args().output)
