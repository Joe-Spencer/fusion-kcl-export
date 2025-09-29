# Fusion 360 KCL Export Add-in

A Fusion 360 add-in for exporting designs to KCL (KittyCAD Language) format with both single file and batch processing capabilities. Built on top of the proven KCL export script logic.

## Features

- **Single File Export**: Export the currently active design to KCL format
- **Batch Processing**: Export entire project folders containing multiple .f3d files
- **Configurable Export Options**: Choose to include/exclude sketches and features
- **User-Friendly Interface**: Integrated dialog boxes with file/folder selection
- **Progress Tracking**: Real-time progress updates during batch processing

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

1. Click the **Batch Export to KCL** button in the toolbar
2. In the dialog:
   - Click **Browse Folder...** to select the project folder containing .f3d files
   - Click **Browse Output...** to select the output folder for KCL files
   - Monitor progress in the progress text box
   - Click **OK** to start batch processing

## KCL Output Format

The add-in exports designs to proper KCL (KittyCAD Language) format with functional programming syntax:

```kcl
// Generated from Fusion 360
// Design: MyDesign.f3d

// Set units
@settings(defaultLengthUnit = mm)

// Component: MyDesign
// Found 1 sketches and 1 features

// === SKETCHES ===
// Processing sketch 1/1: Rectangle
// Sketch: Rectangle
rectangle = startSketchOn(XY)
  |> startProfile(at = [0.0, 0.0], %)
  |> line(endAbsolute = [50.0, 0.0], %)
  |> line(endAbsolute = [50.0, 25.0], %)
  |> line(endAbsolute = [0.0, 25.0], %)
  |> close(%)

// === FEATURES ===
// Processing feature 1/1: Extrude1
// Extrude: Extrude1
extrude1 = rectangle |> extrude(length = 10.0)
```

## Supported Elements

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
- **Unit Conversion**: Automatic conversion from Fusion 360's cm to KCL's mm
- **Feature Tracking**: Maintains relationships between sketches and features for boolean operations
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

1. **No .f3d files found**: Ensure the selected project folder contains Fusion 360 design files (.f3d extension)

2. **Export failed**: Check that:
   - Output folder has write permissions
   - Fusion 360 files are not corrupted
   - Sufficient disk space is available

3. **Add-in not appearing**: Verify that:
   - Files are in the correct Add-ins directory
   - Fusion 360 has been restarted after installation
   - Add-in is enabled in the Add-Ins dialog

### Debug Mode

Set `DEBUG = True` in `config.py` to enable detailed logging in the Text Command window.

## License

This add-in is provided as-is for educational and development purposes.

## Author

Joseph Spencer - Initial development and KCL export functionality
