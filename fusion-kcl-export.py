"""
Fusion 360 to KCL Export Script

This script exports Fusion 360 designs to KCL (KittyCAD Language) format.
It attempts to preserve as much of the feature tree, timeline, parameters 
and constraints as possible.
"""

import traceback
import os
import math
import adsk.core
import adsk.fusion

# Initialize the global variables for the Application and UserInterface objects.
app = adsk.core.Application.get()
ui = app.userInterface


class KCLExporter:
    """Main class for exporting Fusion 360 designs to KCL format."""
    
    def __init__(self):
        self.kcl_content = []
        self.units = "mm"  # Default units
        self.indent_level = 0
        
    def add_line(self, line: str):
        """Add a line to the KCL content with proper indentation."""
        indent = "  " * self.indent_level
        self.kcl_content.append(f"{indent}{line}")
    
    def add_comment(self, comment: str):
        """Add a comment to the KCL content."""
        self.add_line(f"// {comment}")
    
    def export_design(self, design: adsk.fusion.Design) -> str:
        """Export a Fusion 360 design to KCL format."""
        self.kcl_content = []
        
        # Add header comment and settings (like bone-plate example)
        self.add_comment("Generated from Fusion 360")
        self.add_comment(f"Design: {design.parentDocument.name}")
        self.add_line("")
        
        # Add KCL settings
        self.add_line("// Set units")
        self.add_line("@settings(defaultLengthUnit = mm)")
        self.add_line("")
        
        # Process the root component
        root_component = design.rootComponent
        self.export_component(root_component)
        
        return "\n".join(self.kcl_content)
    
    def export_component(self, component: adsk.fusion.Component):
        """Export a Fusion 360 component to KCL."""
        self.add_comment(f"Component: {component.name}")
        self.add_comment(f"Found {component.sketches.count} sketches and {component.features.count} features")
        self.add_line("")
        
        # Export sketches FIRST - features depend on them
        if component.sketches.count > 0:
            self.add_comment("=== SKETCHES ===")
            for i in range(component.sketches.count):
                sketch = component.sketches.item(i)
                self.add_comment(f"Processing sketch {i+1}/{component.sketches.count}: {sketch.name}")
                self.export_sketch(sketch)
        
        # Export features AFTER sketches
        if component.features.count > 0:
            self.add_comment("=== FEATURES ===")
            for i in range(component.features.count):
                feature = component.features.item(i)
                self.add_comment(f"Processing feature {i+1}/{component.features.count}: {feature.name} ({feature.objectType})")
                self.export_feature(feature)
    
    def export_sketch(self, sketch: adsk.fusion.Sketch):
        """Export a Fusion 360 sketch to KCL."""
        self.add_comment(f"Sketch: {sketch.name}")
        
        # Get the sketch plane
        plane_name = self.get_plane_name(sketch.referencePlane)
        sketch_var_name = self.get_safe_name(sketch.name)
        
        # Debug: Check if sketch has any curves
        total_curves = (sketch.sketchCurves.sketchLines.count + 
                       sketch.sketchCurves.sketchArcs.count + 
                       sketch.sketchCurves.sketchCircles.count +
                       sketch.sketchCurves.sketchFittedSplines.count)
        
        self.add_comment(f"Sketch has {total_curves} total curves (lines: {sketch.sketchCurves.sketchLines.count}, arcs: {sketch.sketchCurves.sketchArcs.count}, circles: {sketch.sketchCurves.sketchCircles.count})")
        
        if total_curves == 0:
            self.add_comment(f"Skipping {sketch.name} - no curves found")
            return
        
        # Start the sketch with proper KCL syntax (no const/let)
        first_point = self.find_sketch_start_point(sketch.sketchCurves)
        
        # Create the sketch and profile in one chain
        self.add_line(f'{sketch_var_name} = startSketchOn({plane_name})')
        self.indent_level += 1
        
        # Start with the first point - using correct KCL syntax
        self.add_line(f"|> startProfile(at = [{first_point[0]}, {first_point[1]}], %)")
        
        # Export sketch curves
        self.export_sketch_curve(sketch.sketchCurves)
        
        # Close the profile
        self.add_line("|> close(%)")
        
        self.indent_level -= 1
        self.add_line("")
    
    def export_sketch_curve(self, curves):
        """Export sketch curves to KCL."""
        # Handle lines
        for line in curves.sketchLines:
            self.export_line(line)
        
        # Handle arcs
        for arc in curves.sketchArcs:
            self.export_arc(arc)
        
        # Handle circles
        for circle in curves.sketchCircles:
            self.export_circle(circle)
        
        # Handle splines
        for spline in curves.sketchFittedSplines:
            self.export_spline(spline)
    
    def export_line(self, line: adsk.fusion.SketchLine):
        """Export a sketch line to KCL."""
        start = line.startSketchPoint.geometry
        end = line.endSketchPoint.geometry
        
        start_x, start_y = self.convert_point_2d(start)
        end_x, end_y = self.convert_point_2d(end)
        
        # Use KCL line function with proper labeled arguments (like the bone-plate example)
        self.add_line(f"  |> line(endAbsolute = [{end_x}, {end_y}], %)")
    
    def export_arc(self, arc: adsk.fusion.SketchArc):
        """Export a sketch arc to KCL."""
        center = arc.centerSketchPoint.geometry
        start = arc.startSketchPoint.geometry
        end = arc.endSketchPoint.geometry
        
        center_x, center_y = self.convert_point_2d(center)
        start_x, start_y = self.convert_point_2d(start)
        end_x, end_y = self.convert_point_2d(end)
        
        # Calculate radius and angle for KCL arc
        radius = self.convert_length(arc.radius)
        sweep_angle = math.degrees(arc.sweepAngle)
        
        # Use arc syntax from bone-plate example - need start and end angles
        start_angle = 0  # Default start angle
        end_angle = sweep_angle
        self.add_line(f"  |> arc(angleStart = {start_angle}, angleEnd = {end_angle}, radius = {radius}, %)")
    
    def export_circle(self, circle: adsk.fusion.SketchCircle):
        """Export a sketch circle to KCL."""
        center = circle.centerSketchPoint.geometry
        radius = circle.radius
        
        center_x, center_y = self.convert_point_2d(center)
        radius_mm = self.convert_length(radius)
        
        # For circles, use the correct KCL syntax (center and radius/diameter)
        diameter_mm = radius_mm * 2
        self.add_line(f"  |> circle(center = [{center_x}, {center_y}], diameter = {diameter_mm}, %)")
    
    def export_spline(self, spline: adsk.fusion.SketchFittedSpline):
        """Export a sketch spline to KCL (simplified as connected lines)."""
        # For now, approximate splines as a series of line segments
        # This is a simplification - KCL may have better spline support
        points = []
        for point in spline.fitPoints:
            x, y = self.convert_point_2d(point.geometry)
            points.append([x, y])
        
        # Create line segments between consecutive points
        for i in range(len(points) - 1):
            start = points[i]
            end = points[i + 1]
            self.add_line(f"  |> line(endAbsolute = [{end[0]}, {end[1]}], %)")
    
    def export_feature(self, feature):
        """Export a Fusion 360 feature to KCL."""
        if feature.objectType == adsk.fusion.ExtrudeFeature.classType():
            self.export_extrude(feature)
        elif feature.objectType == adsk.fusion.RevolveFeature.classType():
            self.export_revolve(feature)
        # Add more feature types as needed
    
    def export_extrude(self, extrude: adsk.fusion.ExtrudeFeature):
        """Export an extrude feature to KCL."""
        self.add_comment(f"Extrude: {extrude.name}")
        
        try:
            # Get the primary extent (extentOne)
            extent_one = extrude.extentOne
            distance = None
            
            # Handle different extent types
            if extent_one.objectType == adsk.fusion.DistanceExtentDefinition.classType():
                distance = self.convert_length(extent_one.distance.value)
            elif extent_one.objectType == adsk.fusion.ThroughAllExtentDefinition.classType():
                # For through-all, we'll use a default distance
                distance = 100.0  # Default 100mm
                self.add_comment("Note: Through-all extent converted to 100mm")
            elif extent_one.objectType == adsk.fusion.ToEntityExtentDefinition.classType():
                # For to-entity, we'll use a default distance
                distance = 50.0  # Default 50mm
                self.add_comment("Note: To-entity extent converted to 50mm")
            elif extent_one.objectType == adsk.fusion.SymmetricExtentDefinition.classType():
                # For symmetric extent, get the distance and use it
                distance = self.convert_length(extent_one.distance.value)
                self.add_comment("Note: Symmetric extent - using total distance")
            elif extent_one.objectType == adsk.fusion.TwoSidesExtentDefinition.classType():
                # For two-sided extent, use the first side distance
                distance = self.convert_length(extent_one.distanceOne.value)
                self.add_comment("Note: Two-sided extent - using first side distance only")
            else:
                # Log the actual extent type for debugging
                self.add_comment(f"Unsupported extent type: {extent_one.objectType}")
                distance = 10.0  # Default fallback distance
            
            if distance is not None:
                # Find the associated sketch/profile
                profile_obj = extrude.profile
                if profile_obj:
                    # Handle both single profile and ObjectCollection
                    if hasattr(profile_obj, 'objectType') and profile_obj.objectType == adsk.core.ObjectCollection.classType():
                        # Multiple profiles - use the first one
                        if profile_obj.count > 0:
                            first_profile = profile_obj.item(0)
                            if hasattr(first_profile, 'parentSketch') and first_profile.parentSketch:
                                sketch_name = self.get_safe_name(first_profile.parentSketch.name)
                                self.add_line(f"extrude{self.get_unique_id()} = {sketch_name} |> extrude(length = {distance})")
                            else:
                                self.add_line(f"extrude{self.get_unique_id()} = sketch |> extrude(length = {distance})")
                        else:
                            self.add_comment("Warning: Empty profile collection")
                    else:
                        # Single profile
                        if hasattr(profile_obj, 'parentSketch') and profile_obj.parentSketch:
                            sketch_name = self.get_safe_name(profile_obj.parentSketch.name)
                            self.add_line(f"extrude{self.get_unique_id()} = {sketch_name} |> extrude(length = {distance})")
                        else:
                            self.add_line(f"extrude{self.get_unique_id()} = sketch |> extrude(length = {distance})")
                else:
                    self.add_comment("Warning: No profile found for extrude")
            else:
                self.add_comment("Warning: Unsupported extent type")
                
        except Exception as e:
            self.add_comment(f"Error processing extrude: {str(e)}")
        
        self.add_line("")
    
    def export_revolve(self, revolve: adsk.fusion.RevolveFeature):
        """Export a revolve feature to KCL."""
        self.add_comment(f"Revolve: {revolve.name}")
        
        try:
            # Get the primary extent (extentDefinition)
            extent_def = revolve.extentDefinition
            angle = None
            
            # Handle different extent types
            if extent_def.objectType == adsk.fusion.AngleExtentDefinition.classType():
                angle = math.degrees(extent_def.angle.value)
            elif extent_def.objectType == adsk.fusion.FullSweepExtentDefinition.classType():
                angle = 360.0  # Full revolution
                self.add_comment("Note: Full sweep converted to 360 degrees")
            
            if angle is not None:
                # Find the associated sketch/profile
                profile_obj = revolve.profile
                if profile_obj:
                    # Handle both single profile and ObjectCollection
                    if hasattr(profile_obj, 'objectType') and profile_obj.objectType == adsk.core.ObjectCollection.classType():
                        # Multiple profiles - use the first one
                        if profile_obj.count > 0:
                            first_profile = profile_obj.item(0)
                            if hasattr(first_profile, 'parentSketch') and first_profile.parentSketch:
                                sketch_name = self.get_safe_name(first_profile.parentSketch.name)
                                self.add_line(f"revolve{self.get_unique_id()} = {sketch_name} |> revolve(axis = Y, angle = {angle})")
                            else:
                                self.add_line(f"revolve{self.get_unique_id()} = sketch |> revolve(axis = Y, angle = {angle})")
                        else:
                            self.add_comment("Warning: Empty profile collection")
                    else:
                        # Single profile
                        if hasattr(profile_obj, 'parentSketch') and profile_obj.parentSketch:
                            sketch_name = self.get_safe_name(profile_obj.parentSketch.name)
                            self.add_line(f"revolve{self.get_unique_id()} = {sketch_name} |> revolve(axis = Y, angle = {angle})")
                        else:
                            self.add_line(f"revolve{self.get_unique_id()} = sketch |> revolve(axis = Y, angle = {angle})")
                else:
                    self.add_comment("Warning: No profile found for revolve")
            else:
                self.add_comment("Warning: Unsupported revolve extent type")
                
        except Exception as e:
            self.add_comment(f"Error processing revolve: {str(e)}")
        
        self.add_line("")
    
    def get_plane_name(self, plane) -> str:
        """Get the KCL plane name for a Fusion 360 reference plane."""
        if plane.objectType == adsk.fusion.ConstructionPlane.classType():
            # Handle construction planes
            return "XY"  # Default fallback
        
        # Handle standard planes
        plane_name = str(plane)
        if "XY" in plane_name:
            return "XY"
        elif "XZ" in plane_name:
            return "XZ"
        elif "YZ" in plane_name:
            return "YZ"
        else:
            return "XY"  # Default fallback
    
    def convert_point_2d(self, point) -> tuple:
        """Convert a 2D point to KCL format."""
        x = self.convert_length(point.x)
        y = self.convert_length(point.y)
        return (x, y)
    
    def convert_length(self, length_cm: float) -> float:
        """Convert length from Fusion 360 internal units (cm) to mm."""
        return round(length_cm * 10, 3)  # Convert cm to mm and round to 3 decimal places
    
    def get_safe_name(self, name: str) -> str:
        """Convert a name to a safe KCL variable name in lowerCamelCase."""
        import re
        # Remove special characters and split on spaces/underscores
        words = re.findall(r'[a-zA-Z0-9]+', name)
        if not words:
            return "unnamed"
        
        # Convert to lowerCamelCase
        safe_name = words[0].lower()
        for word in words[1:]:
            safe_name += word.capitalize()
        
        # Ensure it starts with a letter
        if safe_name and safe_name[0].isdigit():
            safe_name = f"sketch{safe_name.capitalize()}"
        
        return safe_name or "unnamed"
    
    def get_unique_id(self) -> str:
        """Generate a unique ID for naming KCL entities."""
        if not hasattr(self, '_counter'):
            self._counter = 0
        self._counter += 1
        return str(self._counter)
    
    def find_sketch_start_point(self, curves) -> tuple:
        """Find a good starting point for the sketch profile."""
        # Try to find the first line's start point
        if curves.sketchLines.count > 0:
            first_line = curves.sketchLines.item(0)
            start_point = first_line.startSketchPoint.geometry
            return self.convert_point_2d(start_point)
        
        # Try to find the first arc's start point
        if curves.sketchArcs.count > 0:
            first_arc = curves.sketchArcs.item(0)
            start_point = first_arc.startSketchPoint.geometry
            return self.convert_point_2d(start_point)
        
        # Try to find a circle center (though circles don't have start points)
        if curves.sketchCircles.count > 0:
            first_circle = curves.sketchCircles.item(0)
            center_point = first_circle.centerSketchPoint.geometry
            return self.convert_point_2d(center_point)
        
        # Default fallback
        return (0.0, 0.0)


def run(_context: str):
    """This function is called by Fusion when the script is run."""

    try:
        # Get the active design
        design = app.activeProduct
        if not design:
            ui.messageBox('No active design found.')
            return
        
        if design.objectType != adsk.fusion.Design.classType():
            ui.messageBox('Active product is not a Fusion 360 design.')
            return
        
        # Create the exporter
        exporter = KCLExporter()
        
        # Export the design to KCL
        kcl_content = exporter.export_design(design)
        
        # Get the save location
        file_dialog = ui.createFileDialog()
        file_dialog.isMultiSelectEnabled = False
        file_dialog.title = "Save KCL File"
        file_dialog.filter = "KCL files (*.kcl)"
        file_dialog.filterIndex = 0
        
        # Set default filename based on design name
        design_name = design.parentDocument.name
        if design_name.endswith('.f3d'):
            design_name = design_name[:-4]  # Remove .f3d extension
        
        file_dialog.initialFilename = f"{design_name}.kcl"
        
        dialog_result = file_dialog.showSave()
        if dialog_result == adsk.core.DialogResults.DialogOK:
            filename = file_dialog.filename
            
            # Write the KCL file
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(kcl_content)
            
            ui.messageBox(f'Successfully exported to KCL:\n{filename}')
        else:
            ui.messageBox('Export cancelled.')
            
    except Exception as e:
        # Write the error message to the TEXT COMMANDS window
        error_msg = f'Failed to export KCL:\n{traceback.format_exc()}'
        app.log(error_msg)
        
        # Also log some debugging information
        try:
            design = app.activeProduct
            if design:
                app.log(f'Active design: {design.parentDocument.name}')
                root_component = design.rootComponent
                app.log(f'Root component: {root_component.name}')
                app.log(f'Number of sketches: {root_component.sketches.count}')
                app.log(f'Number of features: {root_component.features.count}')
        except:
            app.log('Could not gather debugging information')
        
        ui.messageBox(f'Export failed. Check the Text Commands window for details.\n\nError: {str(e)}')
