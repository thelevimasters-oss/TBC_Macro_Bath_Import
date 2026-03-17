Hydro: EchoSounding -> GNSS Vertical Offset (TBC Macro)
=====================================================

This macro converts Trimble Access Hydro JobXML (*.jxl) EchoSounding depths into
GNSS Vertical Offsets, saves a new JXL file, and (optionally) imports it into the
current Trimble Business Center project.

Why?
-----
Some TBC versions cannot import newer Hydro JobXML files that include <EchoSounding>
(and related records). By converting EchoSounding depth to the older/standard
<GNSSVerticalOffset> structure, the file becomes compatible and the depth is preserved
as a GNSS Vertical Offset.


Update
------
This package includes two fixes:

1) Startup fix for TBC MacroCommands3 / IronPython 3 where the macro could fail with:

   ImportError: No module named Xml

2) Import fix where some TBC builds do not expose a hardcoded ImportFilesServiceItem type name.
   The macro now uses reflection to discover the available import API at runtime.

Files in this package
---------------------
- HydroEchoToGnssVerticalOffset.py
- HydroEchoToGnssVerticalOffset.xaml
- (optional) HydroEchoToGnssVerticalOffset.bmp  (icon; not required)

Installation
------------
1) Close TBC.
2) Copy the folder containing these files into your MacroCommands folder.

   Typical locations:

   - TBC v5.90+ (IronPython 3):
     %ProgramData%\Trimble\Trimble Business Center\MacroCommands3\

   - Older TBC versions (IronPython 2.7):
     %ProgramData%\Trimble\Trimble Business Center\MacroCommands\

   Example:
   %ProgramData%\Trimble\Trimble Business Center\MacroCommands3\HydroEchoToGnssVerticalOffset\

3) Start TBC.
4) Open the command finder and search for:
   "Hydro: EchoSounding -> GNSS Vertical Offset"

Usage
-----
1) Run the macro.
2) Browse to the Hydro *.jxl file.
3) Choose an output *.jxl path.
4) Select options:
   - Remove <EchoSounding> blocks: recommended.
   - Remove <EchoSounderConfigurationRecord> blocks: recommended for older TBC.
   - Store EchoSounding values as point Notes: optional, but keeps Battery/Quality/Flags.
   - Convert depth from US survey feet to meters: only enable if you know your depths
     are in feet.
   - Import converted JXL: enable to import automatically.
5) Click Convert (or Convert & Import).

How conversion is computed
--------------------------
VerticalOffset = -Depth

Depth is assumed positive down. GNSS Vertical Offset in JobXML is stored in meters.

Troubleshooting
---------------
- If the macro says conversion succeeded but import failed:
  * Use File > Import in TBC to import the output JXL manually.
  * If it still fails, you likely need to remove additional unsupported records.

- If the imported point elevations look wrong by ~3.28x:
  * Your depth values might be in feet. Enable the option to convert feet -> meters.

Disclaimer
----------
This macro was generated without access to your specific TBC version, so it uses
best-effort reflection to call the TBC ImportService. If Trimble changes internal
API names, the import step may require a small tweak.

