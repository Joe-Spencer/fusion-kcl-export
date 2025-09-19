# Fusion 360 to KCL Export Script

This script exports Fusion 360 designs to KCL (KittyCAD Language) format, attempting to preserve as much of the feature tree, timeline, parameters, and constraints as possible.

## Overview

KCL (KittyCAD Language) is a scripting language developed by KittyCAD for defining geometry and interacting efficiently with their Geometry Engine. This script converts Fusion 360 designs into KCL format, enabling interoperability between Fusion 360 and the KittyCAD ecosystem.

## Features

### Currently Supported:
- ✅ **Sketches**: Lines, arcs, circles, and splines
- ✅ **Features**: Extrudes and revolves
- ✅ **Boolean Operations**: Join (union), Cut (subtract), and Intersect
- ✅ **Coordinate System**: Automatic conversion from cm to mm
- ✅ **File Export**: User-friendly file dialog for saving KCL files
- ✅ **Error Handling**: Comprehensive error reporting and logging

### Planned Features:
- 🔄 **Additional Features**: Fillets, chamfers, patterns
- 🔄 **Constraints**: Dimensional and geometric constraints
- 🔄 **Parameters**: User parameters and expressions
- 🔄 **Assembly Support**: Multi-component designs
- 🔄 **Material Properties**: Material assignments

## Installation

1. Copy the script files to your Fusion 360 Scripts folder:
   - Windows: `%APPDATA%\Autodesk\Autodesk Fusion 360\API\Scripts\`
   - macOS: `~/Library/Application Support/Autodesk/Autodesk Fusion 360/API/Scripts/`

2. The script folder should contain:
   ```
   fusion-kcl-export/
   ├── fusion-kcl-export.py
   ├── fusion-kcl-export.manifest
   ├── ScriptIcon.svg
   └── README.md
   ```

## Usage

### Running the Script

1. Open Fusion 360
2. Navigate to **Utilities** → **ADD-INS** → **Scripts and Add-Ins**
3. Under **Scripts**, find "fusion-kcl-export"
4. Click **Run**

### Export Process

1. **Design Selection**: The script automatically uses the currently active design
2. **File Dialog**: Choose where to save your KCL file
3. **Export**: The script processes your design and generates the KCL file

### Example Output

For a simple rectangular extrusion, the script generates KCL like this:

```kcl
// Generated from Fusion 360
// Design: MyDesign.f3d

// Component: MyDesign
// Sketch: Rectangle
const Rectangle = startSketchOn("XY")
  |> startProfileAt([0, 0], %)
  |> lineTo([50.0, 0], %)
  |> lineTo([50.0, 25.0], %)
  |> lineTo([0, 25.0], %)
  |> close(%)

// Extrude: Extrude1
const extrude_1 = Rectangle |> extrude(10.0, %)
```

### Boolean Operations Example

For designs with boolean operations (Combine features), the script generates:

```kcl
// Two extrudes
extrude1 = sketch1 |> extrude(length = 10.0)
extrude2 = sketch2 |> extrude(length = 15.0)

// Boolean operations
solid001 = union(extrude1, extrude2)        // Join operation
solid002 = subtract(extrude1, tools = extrude2)  // Cut operation  
solid003 = intersect(extrude1, extrude2)    // Intersect operation
```

## KCL Language Reference

### Basic Concepts

KCL uses a pipe operator (`|>`) to chain operations together, creating a functional programming approach to CAD modeling.

### Supported Operations

#### Sketching
- `startSketchOn("plane")` - Start a new sketch on a reference plane
- `startProfileAt([x, y], %)` - Begin a profile at coordinates
- `lineTo([x, y], %)` - Draw a line to coordinates
- `arc({center: [x, y], radius: r, angleStart: a1, angleEnd: a2}, %)` - Draw an arc
- `circle([x, y], radius)` - Create a circle
- `close(%)` - Close the current profile

#### 3D Operations
- `extrude(distance, %)` - Extrude a profile by distance
- `revolve(angle, %)` - Revolve a profile by angle

#### Boolean Operations
- `union(solid1, solid2, ...)` - Join multiple solids together
- `subtract(target, tools = tool1)` - Subtract tool solid(s) from target solid
- `intersect(solid1, solid2, ...)` - Keep only the intersecting volume of solids

### Coordinate System

- Fusion 360 uses centimeters internally
- KCL typically uses millimeters
- The script automatically converts units (cm → mm)

## Technical Details

### Architecture

The script is organized around the `KCLExporter` class with the following key methods:

- `export_design()` - Main entry point for design export
- `export_component()` - Processes individual components
- `export_sketch()` - Converts sketches to KCL
- `export_feature()` - Converts features (extrudes, revolves, etc.)

### Coordinate Conversion

```python
def convert_length(self, length_cm: float) -> float:
    """Convert length from Fusion 360 internal units (cm) to mm."""
    return round(length_cm * 10, 3)
```

### Name Sanitization

KCL variable names must follow specific rules. The script automatically:
- Replaces invalid characters with underscores
- Ensures names start with letters or underscores
- Handles duplicate names with unique IDs

## Limitations

### Current Limitations

1. **Sketch Complexity**: Complex sketches with many constraints may not export perfectly
2. **Feature Support**: Only basic extrudes and revolves are currently supported
3. **Assembly Models**: Multi-component assemblies are processed as single components
4. **Materials**: Material properties are not exported
5. **Splines**: Converted to line segments (approximation)

### Known Issues

- Construction geometry is not exported
- Sketch constraints are not preserved
- User parameters are not included
- Timeline dependencies may not be maintained

## Contributing

### Development Setup

The script is written in Python using the Fusion 360 API. Key dependencies:
- `adsk.core` - Fusion 360 core API
- `adsk.fusion` - Fusion 360 design API
- Standard Python libraries: `math`, `os`, `traceback`, `re`

### Adding New Features

To add support for new Fusion 360 features:

1. Add a new export method in the `KCLExporter` class
2. Update the `export_feature()` method to handle the new feature type
3. Add appropriate KCL syntax generation
4. Test with sample models

### Code Style

- Follow PEP 8 Python style guidelines
- Use type hints where possible
- Include comprehensive docstrings
- Add error handling for API calls

## Troubleshooting

### Common Issues

**"No active design found"**
- Ensure you have a design open in Fusion 360
- The design must be in Design workspace (not CAM or Simulation)

**"Export failed" errors**
- Check the Text Commands window for detailed error messages
- Verify the design doesn't contain unsupported features
- Try with a simpler test model first

**Invalid KCL output**
- The generated KCL may need manual adjustment
- Complex geometries might require simplification
- Check the KittyCAD documentation for syntax updates

### Debugging

Enable detailed logging by checking the Fusion 360 Text Commands window:
- **Utilities** → **Text Commands**
- Error messages and stack traces appear here

## Resources

### KittyCAD Documentation
- [KCL Language Reference](https://zoo.dev/docs/kcl-book/)
- [KCL Standard Library](https://zoo.dev/docs/kcl-std/)
- [KittyCAD GitHub](https://github.com/orgs/KittyCAD/repositories)

### Fusion 360 API
- [Fusion 360 API Documentation](https://help.autodesk.com/view/fusion360/ENU/?guid=GUID-A92A4B10-3781-4925-94C6-47DA85A4F65A)
- [API Samples](https://github.com/AutodeskFusion360/Fusion360Samples)

## License

This script is provided as-is for educational and development purposes. Please refer to Autodesk's terms of service for Fusion 360 API usage and KittyCAD's licensing terms for KCL usage.

## Version History

### v1.0.0 (Current)
- Initial release
- Basic sketch export (lines, arcs, circles, splines)
- Extrude and revolve feature support
- Unit conversion (cm to mm)
- File dialog interface
- Error handling and logging

---

**Note**: This is an experimental script. Always verify the generated KCL files before using them in production workflows. The KCL language and KittyCAD ecosystem are rapidly evolving, so syntax and features may change.
