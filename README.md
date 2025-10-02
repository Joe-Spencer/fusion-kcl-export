# Fusion 360 KCL Export Add-in

A Fusion 360 add-in for exporting designs to KCL (KittyCAD Language) format with both single file and batch processing capabilities. Built on top of the proven KCL export script logic.

## Features

- **Single File Export**: Export the currently active design to KCL format
- **Batch Processing**: Automatically export all .f3d files from your project folder
- **Clean Output**: Produces readable KCL code without verbose debug information
- **User-Friendly Interface**: Integrated dialog boxes with file/folder selection
- **Progress Tracking**: Real-time progress updates during batch processing
- **Modular Architecture**: Uses the proven standalone script for consistent export logic

## Installation

1. Copy the entire `fusion-kcl-export` folder to your Fusion 360 Add-ins directory:
   - **Windows**: `%APPDATA%\Autodesk\Autodesk Fusion 360\API\AddIns\`
   - **Mac**: `~/Library/Application Support/Autodesk/Autodesk Fusion 360/API/AddIns/`

2. Start Fusion 360

3. Go to **Tools** > **Add-Ins**

4. Click on the **Add-Ins** tab

5. Find "Fusion KCL Export" in the list and click **Run**

## Usage

### Single File Export

1. Open a design in Fusion 360
2. Click the **Export to KCL** button in the toolbar (under Scripts & Add-Ins panel)
3. In the dialog:
   - Click **Browse...** to select the output file location
   - Click **OK** to export

### Batch Processing

1. Open any design file in your Fusion 360 project
2. Click the **Batch Export to KCL** button in the toolbar
3. In the dialog:
   - Click **Browse Output...** to select the output folder for KCL files
   - Monitor progress in the progress text box
   - Click **OK** to start batch processing
4. The tool will automatically:
   - Find the project folder of your currently open design
   - Export ALL .f3d files in that folder (including the current one)
   - Save each as a .kcl file in your output folder

## KCL Output Format

The add-in exports designs to clean, readable KCL (KittyCAD Language) format with functional programming syntax:

```kcl
// Generated from Fusion 360
// Design: MyDesign.f3d

// Set units
@settings(defaultLengthUnit = mm)

// Component: MyDesign
// Found 1 sketches and 1 features

// === SKETCHES ===
// Sketch: Rectangle
rectangle = startSketchOn(XY)
  |> startProfile(at = [0.0, 0.0], %)
  |> line(endAbsolute = [50.0, 0.0], %)
  |> line(endAbsolute = [50.0, 25.0], %)
  |> line(endAbsolute = [0.0, 25.0], %)
  |> close(%)

// === FEATURES ===
// Extrude: Extrude1
extrude1 = rectangle |> extrude(length = 10.0)
```

The output is clean and concise, with verbose debug information disabled by default for production use.

## Supported Elements

### Parameters
- **Dimensional Parameters**: Top Level Parametric data is exported, however relationships are not currently retained

### Sketches
- **Lines**: Converted to KCL `line()` functions with absolute endpoints
- **Circles**: Exported as `circle()` with center and diameter
- **Arcs**: Converted to `arc()` with start/end angles and radius
- **Splines**: Approximated as connected line segments
- **Connectivity**: Smart curve ordering to maintain profile continuity

### Features
- **Extrudes**: Full support for distance, through-all, symmetric, and two-sided extents
- **Revolves**: Support for angle and full-sweep operations
- **Boolean Operations**: Join (union), Cut (subtract), and Intersect operations
- **Coordinate Systems**: Automatic conversion between Fusion 360 and KCL coordinate systems

### Advanced Capabilities
- **Plane Detection**: Automatic detection of XY, XZ, and YZ sketch planes
- **Unit Conversion**: Automatic conversion between Fusion 360 and KCL units (mm, in, cm, etc.)
- **Feature Tracking**: Maintains relationships between sketches and features for boolean operations
- **Smart Project Discovery**: Uses active design to automatically find and process project folders
- **Error Handling**: Comprehensive error handling with detailed logging

## Development

The add-in is structured using the standard Fusion 360 add-in template with integrated script logic:

- `fusion-kcl-export.py` - Main add-in entry point
- `commands/commandDialog/` - Single file export command
- `commands/batchProcess/` - Batch processing command  
- `fusion-kcl-export-script/` - Standalone script with core KCL export logic
- `config.py` - Configuration and global variables
- `lib/fusionAddInUtils/` - Utility functions for add-in development

### Core Export Logic

The heart of the add-in is the `KCLExporter` class in `fusion-kcl-export-script/fusion-kcl-export.py`, which contains the complete and proven export logic. The add-in commands import this script directly to ensure modularity and avoid code duplication. This ensures that the batch processing uses exactly the same export functionality as the standalone script.

## Troubleshooting

### Common Issues

1. **Batch export finds no files**: 
   - Ensure you have a design file open in Fusion 360
   - Make sure the design is saved to a Fusion 360 project folder (not just local)
   - Verify there are other .f3d files in the same project folder

2. **Export failed**: Check that:
   - Output folder has write permissions
   - Fusion 360 files are not corrupted
   - Sufficient disk space is available
   - You're logged into Fusion 360 with project access

3. **Add-in not appearing**: Verify that:
   - Files are in the correct Add-ins directory
   - Fusion 360 has been restarted after installation
   - Add-in is enabled in the Add-Ins dialog

### Debug Mode

For troubleshooting, you can enable debug mode by:
1. Opening `fusion-kcl-export-script/fusion-kcl-export.py`
2. Changing `KCLExporter(debug_planes=False)` to `KCLExporter(debug_planes=True)` in the add-in commands
3. This will output detailed conversion information and plane debugging data

## Author

Joseph Spencer

## Special Thanks

Special thanks to the [Zoo.dev](https://zoo.dev) team
