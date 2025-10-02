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
    
    def __init__(self, debug_planes=False):
        self.kcl_content = []
        self.indent_level = 0
        self.debug_planes = debug_planes  # Enable detailed plane debugging
        self.body_to_feature_map = {}  # Maps BRepBody entity token to the KCL feature name that created it
        self.feature_to_kcl_name = {}  # Maps Fusion feature entity token to KCL variable name
        self.units = "mm"  # Will be set during export_design
        
    def add_line(self, line: str):
        """Add a line to the KCL content with proper indentation."""
        indent = "  " * self.indent_level
        self.kcl_content.append(f"{indent}{line}")
    
    def add_comment(self, comment: str):
        """Add a comment to the KCL content."""
        self.add_line(f"// {comment}")
    
    def detect_document_units(self) -> str:
        """Detect the current document units using Fusion API."""
        try:
            # Get the active design
            design = app.activeProduct
            if not design or design.objectType != adsk.fusion.Design.classType():
                if self.debug_planes:
                    self.add_comment("No active design found, defaulting to mm")
                return "mm"  # Default fallback
            
            # Try fusionUnitsManager first (returns enum values)
            try:
                fusion_units_manager = design.fusionUnitsManager
                length_unit_enum = fusion_units_manager.distanceDisplayUnits
                if self.debug_planes:
                    self.add_comment(f"fusionUnitsManager.distanceDisplayUnits enum: {length_unit_enum}")
                
                # Convert enum values to KCL unit strings
                if length_unit_enum == adsk.fusion.DistanceUnits.InchDistanceUnits:
                    if self.debug_planes:
                        self.add_comment("Detected inches from enum")
                    return "in"
                elif length_unit_enum == adsk.fusion.DistanceUnits.MillimeterDistanceUnits:
                    if self.debug_planes:
                        self.add_comment("Detected millimeters from enum")
                    return "mm"
                elif length_unit_enum == adsk.fusion.DistanceUnits.CentimeterDistanceUnits:
                    if self.debug_planes:
                        if self.debug_planes:
                            self.add_comment("Detected centimeters from enum")
                    return "cm"
                elif length_unit_enum == adsk.fusion.DistanceUnits.MeterDistanceUnits:
                    if self.debug_planes:
                        if self.debug_planes:
                            self.add_comment("Detected meters from enum")
                    return "m"
                elif length_unit_enum == adsk.fusion.DistanceUnits.FootDistanceUnits:
                    if self.debug_planes:
                        if self.debug_planes:
                            self.add_comment("Detected feet from enum")
                    return "ft"
                else:
                    if self.debug_planes:
                        self.add_comment(f"Unsupported enum value {length_unit_enum}, defaulting to mm")
                    return "mm"
                    
            except Exception as e1:
                if self.debug_planes:
                    if self.debug_planes:
                        self.add_comment(f"fusionUnitsManager failed: {str(e1)}")
                # Fallback to regular unitsManager
                try:
                    units_manager = design.unitsManager
                    length_unit_string = units_manager.defaultLengthUnits
                    if self.debug_planes:
                        self.add_comment(f"unitsManager.defaultLengthUnits: '{length_unit_string}'")
                    
                    # Convert string values to KCL unit strings
                    if length_unit_string == "in" or length_unit_string == "inch":
                        if self.debug_planes:
                            self.add_comment("Detected inches from string")
                        return "in"
                    elif length_unit_string == "mm" or length_unit_string == "millimeter":
                        if self.debug_planes:
                            self.add_comment("Detected millimeters from string")
                        return "mm"
                    elif length_unit_string == "cm" or length_unit_string == "centimeter":
                        if self.debug_planes:
                            self.add_comment("Detected centimeters from string")
                        return "cm"
                    elif length_unit_string == "m" or length_unit_string == "meter":
                        if self.debug_planes:
                            self.add_comment("Detected meters from string")
                        return "m"
                    elif length_unit_string == "ft" or length_unit_string == "foot":
                        if self.debug_planes:
                            self.add_comment("Detected feet from string")
                        return "ft"
                    else:
                        if self.debug_planes:
                            self.add_comment(f"Unsupported string unit '{length_unit_string}', defaulting to mm")
                        return "mm"
                        
                except Exception as e2:
                    if self.debug_planes:
                        self.add_comment(f"unitsManager also failed: {str(e2)}")
                    return "mm"
                
        except Exception as e:
            # If detection fails, default to mm
            if self.debug_planes:
                self.add_comment(f"Unit detection failed: {str(e)}, defaulting to mm")
            return "mm"
    
    def export_design(self, design: adsk.fusion.Design) -> str:
        """Export a Fusion 360 design to KCL format."""
        self.kcl_content = []
        
        # Detect units first
        self.units = self.detect_document_units()
        
        # Add header comment and settings (like bone-plate example)
        self.add_comment("Generated from Fusion 360")
        self.add_comment(f"Design: {design.parentDocument.name}")
        self.add_line("")
        
        # Add KCL settings
        self.add_line("// Set units")
        self.add_line(f"@settings(defaultLengthUnit = {self.units})")
        self.add_line("")
        
        # Export parameters
        self.export_parameters(design)
        
        # Process the root component
        root_component = design.rootComponent
        self.export_component(root_component)
        
        return "\n".join(self.kcl_content)
    
    def export_parameters(self, design: adsk.fusion.Design):
        """Export design parameters to KCL format."""
        try:
            # Get all parameters in the design
            all_params = design.allParameters
            
            if all_params.count == 0:
                if self.debug_planes:
                    self.add_comment("No parameters found in design")
                return
            
            self.add_comment("=== PARAMETERS ===")
            
            # Separate user parameters from model parameters
            user_params = []
            model_params = []
            
            # Also get user parameters specifically
            user_param_collection = design.userParameters
            user_param_names = set()
            for i in range(user_param_collection.count):
                user_param = user_param_collection.item(i)
                user_param_names.add(user_param.name)
            
            for i in range(all_params.count):
                param = all_params.item(i)
                
                # Check if this is a user parameter by name
                if param.name in user_param_names:
                    user_params.append(param)
                else:
                    model_params.append(param)
            
            # Export user parameters first (these are the important ones)
            if user_params:
                self.add_comment("User Parameters:")
                for param in user_params:
                    self.export_parameter(param)
                self.add_line("")
            
            # Export model parameters if debug mode is enabled
            if model_params and self.debug_planes:
                self.add_comment("Model Parameters (auto-generated):")
                for param in model_params:
                    self.export_parameter(param)
                self.add_line("")
            
            if not user_params and not self.debug_planes:
                self.add_comment("No user parameters defined")
                self.add_line("")
                
        except Exception as e:
            if self.debug_planes:
                self.add_comment(f"Error exporting parameters: {str(e)}")
            self.add_line("")
    
    def export_parameter(self, param):
        """Export a single parameter to KCL format."""
        try:
            param_name = param.name
            param_value = param.value
            param_units = param.unit if hasattr(param, 'unit') and param.unit else ""
            param_comment = param.comment if hasattr(param, 'comment') and param.comment else ""
            param_expression = param.expression if hasattr(param, 'expression') and param.expression else str(param_value)
            
            # Clean up parameter name for KCL (replace invalid characters)
            kcl_param_name = self.get_safe_name(param_name)
            
            # Format the parameter value
            if param_units:
                # Convert units if needed
                if param_units in ['cm', 'mm', 'in', 'm', 'ft']:
                    # Length parameter - convert to display units
                    display_value = self.convert_internal_to_display_units(param_value)
                    param_line = f"{kcl_param_name} = {display_value}"
                else:
                    # Other units (angles, etc.) - use as is
                    param_line = f"{kcl_param_name} = {param_value}"
            else:
                # Dimensionless parameter
                param_line = f"{kcl_param_name} = {param_value}"
            
            # Add comment with original name and description if different
            if param_comment or param_name != kcl_param_name:
                comment_parts = []
                if param_name != kcl_param_name:
                    comment_parts.append(f"Original: {param_name}")
                if param_comment:
                    comment_parts.append(param_comment)
                if param_units:
                    comment_parts.append(f"Units: {param_units}")
                if param_expression != str(param_value):
                    comment_parts.append(f"Expression: {param_expression}")
                
                if comment_parts:
                    param_line += f"  // {' | '.join(comment_parts)}"
            
            self.add_line(param_line)
            
        except Exception as e:
            if self.debug_planes:
                self.add_comment(f"Error exporting parameter {param.name}: {str(e)}")
    
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
                if self.debug_planes:
                    self.add_comment(f"Processing sketch {i+1}/{component.sketches.count}: {sketch.name}")
                self.export_sketch(sketch)
        
        # Export features AFTER sketches
        if component.features.count > 0:
            self.add_comment("=== FEATURES ===")
            
            # Process all features using proper Fusion 360 API
            for i in range(component.features.count):
                feature = component.features.item(i)
                if self.debug_planes:
                    self.add_comment(f"Processing feature {i+1}/{component.features.count}: {feature.name} ({feature.objectType})")
                self.export_feature(feature)
    
    def export_sketch(self, sketch: adsk.fusion.Sketch):
        """Export a Fusion 360 sketch to KCL."""
        self.add_comment(f"Sketch: {sketch.name}")
        
        # Get the sketch plane
        plane_name = self.get_plane_name(sketch.referencePlane)
        sketch_var_name = self.get_safe_name(sketch.name)
        
        # Store current sketch plane for coordinate conversion
        self.current_sketch_plane = plane_name
        
        # Reset profile tracking for this sketch
        self.current_profile_position = None
        
        # Debug: Check if sketch has any curves
        total_curves = (sketch.sketchCurves.sketchLines.count + 
                       sketch.sketchCurves.sketchArcs.count + 
                       sketch.sketchCurves.sketchCircles.count +
                       sketch.sketchCurves.sketchFittedSplines.count)
        
        if self.debug_planes:
            self.add_comment(f"Sketch has {total_curves} total curves (lines: {sketch.sketchCurves.sketchLines.count}, arcs: {sketch.sketchCurves.sketchArcs.count}, circles: {sketch.sketchCurves.sketchCircles.count})")
        
        if total_curves == 0:
            self.add_comment(f"Skipping {sketch.name} - no curves found")
            return
        
        # Create the sketch and profile in one chain
        self.add_line(f'{sketch_var_name} = startSketchOn({plane_name})')
        self.indent_level += 1
        
        # Export sketch curves in the correct order (this will handle the starting point)
        has_circles = self.export_sketch_curve(sketch.sketchCurves)
        
        # Only close the profile if it's not already closed (circles are self-closing)
        if not has_circles:
            self.add_line("|> close(%)")
        
        self.indent_level -= 1
        self.add_line("")
    
    def export_sketch_curve(self, curves):
        """Export sketch curves to KCL in the correct order."""
        # Collect all curves into a single list with their types
        all_curves = []
        has_circles = False
        
        # Add lines
        for i in range(curves.sketchLines.count):
            line = curves.sketchLines.item(i)
            all_curves.append(('line', line))
        
        # Add arcs
        for i in range(curves.sketchArcs.count):
            arc = curves.sketchArcs.item(i)
            all_curves.append(('arc', arc))
        
        # Add circles (these are typically standalone, not part of profiles)
        for i in range(curves.sketchCircles.count):
            circle = curves.sketchCircles.item(i)
            all_curves.append(('circle', circle))
            has_circles = True
        
        # Add splines
        for i in range(curves.sketchFittedSplines.count):
            spline = curves.sketchFittedSplines.item(i)
            all_curves.append(('spline', spline))
        
        # Sort curves by their order in the sketch profile
        sorted_curves = self.sort_curves_by_connectivity(all_curves)
        
        if not sorted_curves:
            return has_circles
        
        # Get the starting point from the first curve
        first_curve_type, first_curve = sorted_curves[0]
        if first_curve_type == 'circle':
            # For circles, use center point
            if hasattr(first_curve, 'centerSketchPoint'):
                start_point_geom = first_curve.centerSketchPoint.geometry
            else:
                start_point_geom = None
        else:
            # For other curves, use start point
            if hasattr(first_curve, 'startSketchPoint'):
                start_point_geom = first_curve.startSketchPoint.geometry
            else:
                start_point_geom = None
        
        if start_point_geom:
            start_point = self.convert_point_2d(start_point_geom)
            self.add_line(f"|> startProfile(at = [{start_point[0]}, {start_point[1]}], %)")
            # Track the current position in the profile
            self.current_profile_position = start_point
        else:
            # Fallback
            self.add_line(f"|> startProfile(at = [0.0, 0.0], %)")
            self.current_profile_position = (0.0, 0.0)
        
        # Export curves in the correct order
        for i, (curve_type, curve) in enumerate(sorted_curves):
            if curve_type == 'line':
                self.export_line(curve)
            elif curve_type == 'arc':
                self.export_arc(curve)
            elif curve_type == 'circle':
                self.export_circle(curve)
            elif curve_type == 'spline':
                self.export_spline(curve)
        
        return has_circles
    
    def export_line(self, line: adsk.fusion.SketchLine):
        """Export a sketch line to KCL."""
        start = line.startSketchPoint.geometry
        end = line.endSketchPoint.geometry
        
        start_x, start_y = self.convert_point_2d(start)
        end_x, end_y = self.convert_point_2d(end)
        
        # Check for zero-length lines
        tolerance = 0.001  # 0.001 unit tolerance
        line_length = ((end_x - start_x) ** 2 + (end_y - start_y) ** 2) ** 0.5
        
        if line_length < tolerance:
            if self.debug_planes:
                self.add_comment(f"Skipping zero-length line: [{start_x}, {start_y}] -> [{end_x}, {end_y}]")
            return
        
        # Check if this endpoint is the same as our current position (duplicate endpoint)
        if hasattr(self, 'current_profile_position') and self.current_profile_position:
            current_x, current_y = self.current_profile_position
            if abs(end_x - current_x) < tolerance and abs(end_y - current_y) < tolerance:
                if self.debug_planes:
                    self.add_comment(f"Skipping duplicate endpoint: [{end_x}, {end_y}] (already at [{current_x}, {current_y}])")
                return
        
        # Use KCL line function with proper labeled arguments (like the bone-plate example)
        self.add_line(f"  |> line(endAbsolute = [{end_x}, {end_y}], %)")
        
        # Update current position
        self.current_profile_position = (end_x, end_y)
    
    def export_arc(self, arc: adsk.fusion.SketchArc):
        """Export a sketch arc to KCL."""
        center = arc.centerSketchPoint.geometry
        start = arc.startSketchPoint.geometry
        end = arc.endSketchPoint.geometry
        
        center_x, center_y = self.convert_point_2d(center)
        start_x, start_y = self.convert_point_2d(start)
        end_x, end_y = self.convert_point_2d(end)
        
        # Get arc geometry to access properties
        arc_geometry = arc.geometry
        
        # Calculate radius and angles for KCL arc
        radius = self.convert_internal_to_display_units(arc_geometry.radius)
        
        # Get start and end angles from Arc3D geometry (in radians)
        start_angle_rad = arc_geometry.startAngle
        end_angle_rad = arc_geometry.endAngle
        
        # Calculate sweep angle
        sweep_angle_rad = end_angle_rad - start_angle_rad
        
        # Ensure the sweep angle is positive (handle cases where arc crosses 0 degrees)
        if sweep_angle_rad < 0:
            sweep_angle_rad += 2 * math.pi
        
        # Convert to degrees for KCL
        start_angle_deg = math.degrees(start_angle_rad)
        end_angle_deg = math.degrees(end_angle_rad)
        
        # Ensure end angle is greater than start angle for KCL
        if end_angle_deg < start_angle_deg:
            end_angle_deg += 360
        
        # Use arc syntax from bone-plate example - need start and end angles
        self.add_line(f"  |> arc(angleStart = {start_angle_deg}, angleEnd = {end_angle_deg}, radius = {radius}, %)")
        
        # Update current position to arc end point
        self.current_profile_position = (end_x, end_y)
    
    def export_circle(self, circle: adsk.fusion.SketchCircle):
        """Export a sketch circle to KCL."""
        center = circle.centerSketchPoint.geometry
        radius = circle.radius
        
        center_x, center_y = self.convert_point_2d(center)
        radius_value = self.convert_internal_to_display_units(radius)
        
        # For circles, use the correct KCL syntax (center and radius/diameter)
        diameter_value = radius_value * 2
        self.add_line(f"  |> circle(center = [{center_x}, {center_y}], diameter = {diameter_value}, %)")
        
        # For circles, the current position remains at the center (circles are complete shapes)
        self.current_profile_position = (center_x, center_y)
    
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
            
            # Check for zero-length segments in splines too
            tolerance = 0.001
            segment_length = ((end[0] - start[0]) ** 2 + (end[1] - start[1]) ** 2) ** 0.5
            
            if segment_length >= tolerance:
                self.add_line(f"  |> line(endAbsolute = [{end[0]}, {end[1]}], %)")
                # Update current position
                self.current_profile_position = (end[0], end[1])
    
    def export_feature(self, feature):
        """Export a Fusion 360 feature to KCL."""
        if feature.objectType == adsk.fusion.ExtrudeFeature.classType():
            self.export_extrude(feature)
        elif feature.objectType == adsk.fusion.RevolveFeature.classType():
            self.export_revolve(feature)
        elif feature.objectType == adsk.fusion.CombineFeature.classType():
            self.export_combine(feature)
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
                raw_distance = extent_one.distance.value
                if self.debug_planes:
                    if self.debug_planes:
                        self.add_comment(f"Raw extrude distance (cm): {raw_distance}")
                distance = self.convert_internal_to_display_units(raw_distance)
            elif extent_one.objectType == adsk.fusion.ThroughAllExtentDefinition.classType():
                # For through-all, we'll use a default distance
                distance = 100.0  # Default 100 units
                self.add_comment("Note: Through-all extent converted to 100 units")
            elif extent_one.objectType == adsk.fusion.ToEntityExtentDefinition.classType():
                # For to-entity, we'll use a default distance
                distance = 50.0  # Default 50 units
                self.add_comment("Note: To-entity extent converted to 50 units")
            elif extent_one.objectType == adsk.fusion.SymmetricExtentDefinition.classType():
                # For symmetric extent, get the distance and use it
                distance = self.convert_internal_to_display_units(extent_one.distance.value)
                self.add_comment("Note: Symmetric extent - using total distance")
            elif extent_one.objectType == adsk.fusion.TwoSidesExtentDefinition.classType():
                # For two-sided extent, use the first side distance
                distance = self.convert_internal_to_display_units(extent_one.distanceOne.value)
                self.add_comment("Note: Two-sided extent - using first side distance only")
            else:
                # Log the actual extent type for debugging
                self.add_comment(f"Unsupported extent type: {extent_one.objectType}")
                distance = 10.0  # Default fallback distance
            
            if distance is not None:
                # Find the associated sketch/profile
                profile_obj = extrude.profile
                sketch_plane = None
                
                if profile_obj:
                    # Handle both single profile and ObjectCollection
                    if hasattr(profile_obj, 'objectType') and profile_obj.objectType == adsk.core.ObjectCollection.classType():
                        # Multiple profiles - use the first one
                        if profile_obj.count > 0:
                            first_profile = profile_obj.item(0)
                            if hasattr(first_profile, 'parentSketch') and first_profile.parentSketch:
                                sketch_name = self.get_safe_name(first_profile.parentSketch.name)
                                sketch_plane = self.get_plane_name(first_profile.parentSketch.referencePlane)
                                
                                # Adjust extrude distance for coordinate system differences
                                adjusted_distance = self.adjust_extrude_distance(distance, sketch_plane)
                                extrude_id = self.get_unique_id()
                                extrude_var_name = f"extrude{extrude_id}"
                                self.add_line(f"{extrude_var_name} = {sketch_name} |> extrude(length = {adjusted_distance})")
                                
                                # Track bodies created by this extrude
                                self.track_extrude_bodies(extrude, extrude_var_name)
                            else:
                                self.add_line(f"extrude{self.get_unique_id()} = sketch |> extrude(length = {distance})")
                        else:
                            self.add_comment("Warning: Empty profile collection")
                    else:
                        # Single profile
                        if hasattr(profile_obj, 'parentSketch') and profile_obj.parentSketch:
                            sketch_name = self.get_safe_name(profile_obj.parentSketch.name)
                            sketch_plane = self.get_plane_name(profile_obj.parentSketch.referencePlane)
                            
                            # Adjust extrude distance for coordinate system differences
                            adjusted_distance = self.adjust_extrude_distance(distance, sketch_plane)
                            extrude_id = self.get_unique_id()
                            extrude_var_name = f"extrude{extrude_id}"
                            self.add_line(f"{extrude_var_name} = {sketch_name} |> extrude(length = {adjusted_distance})")
                            
                            # Track bodies created by this extrude
                            self.track_extrude_bodies(extrude, extrude_var_name)
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
    
    def export_combine(self, combine: adsk.fusion.CombineFeature):
        """Export a combine feature to KCL using logical deduction since Fusion API body access fails."""
        self.add_comment(f"Combine: {combine.name}")
        
        try:
            # Get the operation type
            operation_name = "subtract"  # Default
            try:
                operation = combine.operation
                if operation == adsk.fusion.FeatureOperations.JoinFeatureOperation:
                    operation_name = "union"
                elif operation == adsk.fusion.FeatureOperations.CutFeatureOperation:
                    operation_name = "subtract"
                elif operation == adsk.fusion.FeatureOperations.IntersectFeatureOperation:
                    operation_name = "intersect"
                if self.debug_planes:
                    self.add_comment(f"Boolean operation: {operation_name}")
            except Exception as op_error:
                self.add_comment(f"Could not get operation type: {str(op_error)}")
            
            # Since Fusion API body access fails consistently, use logical deduction
            # Based on typical CAD workflow patterns
            if self.debug_planes:
                self.add_comment("Using logical deduction for combine operation (Fusion API body access unreliable)")
            
            # Count how many combines we've processed so far
            combine_count = 0
            for line in self.kcl_content:
                if 'solid' in line and ('subtract' in line or 'union' in line or 'intersect' in line):
                    combine_count += 1
            
            if self.debug_planes:
                self.add_comment(f"This is combine operation #{combine_count + 1}")
            
            # Get all extrude names in order
            extrude_names = []
            for token, kcl_name in self.feature_to_kcl_name.items():
                if kcl_name.startswith('extrude'):
                    extrude_names.append(kcl_name)
            extrude_names.sort(key=lambda x: int(x.replace('extrude', '')))
            
            if self.debug_planes:
                self.add_comment(f"Available extrudes: {extrude_names}")
            
            # Logical deduction based on typical CAD patterns:
            # - First extrude creates main body
            # - Subsequent extrudes create features to be subtracted
            # - Combines typically subtract newer features from older results
            
            target_kcl_name = None
            tool_kcl_name = None
            
            if combine_count == 0 and len(extrude_names) >= 2:
                # First combine: main body (extrude1) - first feature (extrude2)
                target_kcl_name = extrude_names[0]  # extrude1
                tool_kcl_name = extrude_names[1]    # extrude2
                self.add_comment(f"First combine: {target_kcl_name} - {tool_kcl_name}")
                
            else:
                # For all subsequent combines: most recent result - next available extrude
                # Find the most recent solid
                recent_solid = None
                for line in reversed(self.kcl_content):
                    if 'solid' in line and ('=' in line):
                        solid_name = line.split('=')[0].strip()
                        if solid_name.startswith('solid'):
                            recent_solid = solid_name
                            break
                
                # Find the next extrude to subtract (the one after the already used ones)
                # We need to find which extrudes have already been used in previous combines
                used_extrudes = set()
                for line in self.kcl_content:
                    if 'subtract' in line or 'union' in line or 'intersect' in line:
                        # Extract extrude names from the line
                        for extrude_name in extrude_names:
                            if extrude_name in line:
                                used_extrudes.add(extrude_name)
                
                if self.debug_planes:
                    self.add_comment(f"Used extrudes so far: {sorted(used_extrudes)}")
                
                # Find the first unused extrude (excluding the main body extrude1)
                unused_extrudes = [e for e in extrude_names if e not in used_extrudes and e != extrude_names[0]]
                
                if recent_solid and unused_extrudes:
                    target_kcl_name = recent_solid
                    tool_kcl_name = unused_extrudes[0]  # First unused extrude
                    self.add_comment(f"Combine #{combine_count + 1}: {target_kcl_name} - {tool_kcl_name}")
                else:
                    self.add_comment(f"Could not find unused extrudes. Recent solid: {recent_solid}, Unused: {unused_extrudes}")
            
            # Generate the boolean operation if we have deduced the parameters
            if target_kcl_name and tool_kcl_name:
                solid_id = str(self.get_unique_id()).zfill(3)
                solid_var_name = f"solid{solid_id}"
                
                if operation_name == "subtract":
                    self.add_line(f"{solid_var_name} = {operation_name}({target_kcl_name}, tools = {tool_kcl_name})")
                else:
                    self.add_line(f"{solid_var_name} = {operation_name}({target_kcl_name}, {tool_kcl_name})")
                
                if self.debug_planes:
                    self.add_comment(f"SUCCESS: Generated logical boolean operation")
                
            else:
                self.add_comment("Could not deduce combine parameters - SKIPPING")
                if not target_kcl_name:
                    self.add_comment("  Could not determine target")
                if not tool_kcl_name:
                    self.add_comment("  Could not determine tool")
                
        except Exception as e:
            self.add_comment(f"Error in combine processing: {str(e)}")
            self.add_comment("SKIPPING BOOLEAN OPERATION due to error")
        
        self.add_line("")
    
    def find_body_source_feature(self, body):
        """Find the source feature that created a body."""
        try:
            # Try to get the body's creation feature
            if hasattr(body, 'createdBy') and body.createdBy:
                feature = body.createdBy
                if feature.objectType == adsk.fusion.ExtrudeFeature.classType():
                    return f"extrude{self.get_feature_id(feature)}"
                elif feature.objectType == adsk.fusion.RevolveFeature.classType():
                    return f"revolve{self.get_feature_id(feature)}"
                else:
                    # For other feature types, use a generic name
                    return f"feature{self.get_feature_id(feature)}"
            else:
                return None
        except:
            return None
    
    def get_feature_id(self, feature) -> str:
        """Get a consistent ID for a feature."""
        # Use the feature's entity token as a unique identifier
        try:
            return str(hash(feature.entityToken) % 1000)
        except:
            return self.get_unique_id()
    
    def get_plane_name(self, plane) -> str:
        """Get the KCL plane name for a Fusion 360 reference plane."""
        try:
            # Add debugging information about the plane (if enabled)
            if self.debug_planes:
                self.add_comment(f"Plane debug - Object type: {plane.objectType}")
                self.add_comment(f"Plane debug - String representation: {str(plane)}")
            
            # Check if this is a BRepFace (planar face)
            if plane.objectType == adsk.fusion.BRepFace.classType():
                # Get the face's surface geometry
                surface = plane.geometry
                if surface.objectType == adsk.core.Plane.classType():
                    # Get the plane's normal vector
                    normal = surface.normal
                    if self.debug_planes:
                        self.add_comment(f"Face normal vector: ({normal.x:.3f}, {normal.y:.3f}, {normal.z:.3f})")
                    
                    # Determine plane based on normal vector
                    # XY plane has normal (0, 0, 1) or (0, 0, -1)
                    # XZ plane has normal (0, 1, 0) or (0, -1, 0)  
                    # YZ plane has normal (1, 0, 0) or (-1, 0, 0)
                    
                    tolerance = 0.1
                    if abs(normal.z) > (1.0 - tolerance):
                        if self.debug_planes:
                            self.add_comment("Detected XY plane (normal points in Z direction)")
                        return "XY"
                    elif abs(normal.y) > (1.0 - tolerance):
                        if self.debug_planes:
                            self.add_comment("Detected XZ plane (normal points in Y direction)")
                        return "XZ"
                    elif abs(normal.x) > (1.0 - tolerance):
                        if self.debug_planes:
                            self.add_comment("Detected YZ plane (normal points in X direction)")
                        return "YZ"
                    else:
                        if self.debug_planes:
                            self.add_comment(f"Custom plane orientation - defaulting to XY")
                        return "XY"
                else:
                    if self.debug_planes:
                        self.add_comment(f"Non-planar surface type: {surface.objectType}")
                    return "XY"
            
            # Check if this is a ConstructionPlane
            elif plane.objectType == adsk.fusion.ConstructionPlane.classType():
                # Get the construction plane's geometry
                plane_geometry = plane.geometry
                if plane_geometry.objectType == adsk.core.Plane.classType():
                    normal = plane_geometry.normal
                    if self.debug_planes:
                        self.add_comment(f"Construction plane normal: ({normal.x:.3f}, {normal.y:.3f}, {normal.z:.3f})")
                    
                    tolerance = 0.1
                    if abs(normal.z) > (1.0 - tolerance):
                        if self.debug_planes:
                            self.add_comment("Construction plane aligned with XY")
                        return "XY"
                    elif abs(normal.y) > (1.0 - tolerance):
                        if self.debug_planes:
                            self.add_comment("Construction plane aligned with XZ")
                        return "XZ"
                    elif abs(normal.x) > (1.0 - tolerance):
                        if self.debug_planes:
                            self.add_comment("Construction plane aligned with YZ")
                        return "YZ"
                    else:
                        if self.debug_planes:
                            self.add_comment("Custom construction plane orientation - defaulting to XY")
                        return "XY"
                else:
                    if self.debug_planes:
                        self.add_comment("Construction plane has non-standard geometry")
                    return "XY"
            
            # Fallback: try to parse the string representation for standard planes
            else:
                plane_str = str(plane).upper()
                if self.debug_planes:
                    self.add_comment(f"Using string parsing fallback for plane: {plane_str}")
                
                # Look for origin plane indicators
                if "XY" in plane_str or "ORIGIN XY" in plane_str:
                    if self.debug_planes:
                        self.add_comment("String parsing detected XY plane")
                    return "XY"
                elif "XZ" in plane_str or "ORIGIN XZ" in plane_str:
                    if self.debug_planes:
                        self.add_comment("String parsing detected XZ plane")
                    return "XZ"
                elif "YZ" in plane_str or "ORIGIN YZ" in plane_str:
                    if self.debug_planes:
                        self.add_comment("String parsing detected YZ plane")
                    return "YZ"
                else:
                    if self.debug_planes:
                        self.add_comment("String parsing failed - defaulting to XY")
                    return "XY"
                    
        except Exception as e:
            if self.debug_planes:
                self.add_comment(f"Error in plane detection: {str(e)}")
                self.add_comment("Using XY plane as fallback")
            return "XY"
    
    def convert_point_2d(self, point) -> tuple:
        """Convert a 2D point to KCL format, accounting for coordinate system differences."""
        if self.debug_planes:
            self.add_comment(f"Raw point values (cm): x={point.x}, y={point.y}")
        
        # Convert from internal centimeters to display units
        x = self.convert_internal_to_display_units(point.x)
        y = self.convert_internal_to_display_units(point.y)
        
        # Handle coordinate system differences between Fusion 360 and KCL
        if hasattr(self, 'current_sketch_plane'):
            if self.current_sketch_plane == "XZ":
                # For XZ plane, flip the Y coordinate to match KCL coordinate system
                original_y = y
                y = -y
                # Only log the first coordinate flip to avoid spam
                if self.debug_planes and not hasattr(self, '_xz_flip_logged'):
                    self.add_comment(f"XZ plane: Flipping Y coordinates (e.g., {original_y} -> {y})")
                    self._xz_flip_logged = True
        
        return (x, y)
    
    def convert_internal_to_display_units(self, value_cm: float) -> float:
        """Convert internal centimeter values to display units."""
        try:
            # Get the active design
            design = app.activeProduct
            if not design or design.objectType != adsk.fusion.Design.classType():
                return round(value_cm, 3)  # Fallback to cm
            
            # Get the units manager
            units_manager = design.unitsManager
            
            # Convert from centimeters to display units
            display_value = units_manager.convert(value_cm, 'cm', units_manager.defaultLengthUnits)
            
            if self.debug_planes:
                self.add_comment(f"Converted {value_cm} cm to {display_value} {units_manager.defaultLengthUnits}")
            
            return round(display_value, 3)
            
        except Exception as e:
            if self.debug_planes:
                self.add_comment(f"Unit conversion failed: {str(e)}, using raw value")
            return round(value_cm, 3)  # Fallback to raw value
    
    def adjust_extrude_distance(self, distance: float, sketch_plane: str) -> float:
        """Adjust extrude distance for coordinate system differences between Fusion 360 and KCL."""
        if sketch_plane == "XZ":
            # For XZ plane, flip the extrude direction to match KCL coordinate system
            adjusted_distance = -distance
            if self.debug_planes:
                self.add_comment(f"XZ plane: Flipped extrude direction from {distance} to {adjusted_distance}")
            return adjusted_distance
        else:
            # For XY and YZ planes, no adjustment needed
            return distance
    
    def track_extrude_bodies(self, extrude_feature: adsk.fusion.ExtrudeFeature, kcl_var_name: str):
        """Track the bodies created by an extrude feature for use in combine operations."""
        try:
            # Store the mapping from Fusion feature to KCL variable name using entity token
            feature_token = extrude_feature.entityToken
            self.feature_to_kcl_name[feature_token] = kcl_var_name
            
            
            if self.debug_planes:
                if self.debug_planes:
                    self.add_comment(f"Attempting to track bodies for {kcl_var_name}")
            
            # Check the extrude operation type
            operation_type = None
            operation_name = "unknown"
            try:
                operation_type = extrude_feature.operation
                
                # Map operation types to readable names
                if operation_type == adsk.fusion.FeatureOperations.JoinFeatureOperation:
                    operation_name = "Join"
                elif operation_type == adsk.fusion.FeatureOperations.CutFeatureOperation:
                    operation_name = "Cut"
                elif operation_type == adsk.fusion.FeatureOperations.IntersectFeatureOperation:
                    operation_name = "Intersect"
                elif operation_type == adsk.fusion.FeatureOperations.NewBodyFeatureOperation:
                    operation_name = "NewBody"
                elif operation_type == adsk.fusion.FeatureOperations.NewComponentFeatureOperation:
                    operation_name = "NewComponent"
                else:
                    operation_name = f"Unknown({operation_type})"
                
                if self.debug_planes:
                    if self.debug_planes:
                        self.add_comment(f"Extrude operation type: {operation_name}")
            except Exception as op_error:
                if self.debug_planes:
                    self.add_comment(f"Could not get operation type: {str(op_error)}")
            
            # Get all bodies created/modified by this extrude
            bodies = []
            
            # Check if bodies property exists and is accessible
            if hasattr(extrude_feature, 'bodies'):
                if self.debug_planes:
                    self.add_comment(f"Extrude has bodies property")
                try:
                    bodies_collection = extrude_feature.bodies
                    if bodies_collection:
                        body_count = bodies_collection.count
                        if self.debug_planes:
                            if self.debug_planes:
                                self.add_comment(f"Bodies collection has {body_count} bodies")
                        
                        for i in range(body_count):
                            try:
                                body = bodies_collection.item(i)
                                bodies.append(body)
                                if self.debug_planes:
                                    self.add_comment(f"Successfully accessed body {i}")
                            except Exception as body_error:
                                if self.debug_planes:
                                    self.add_comment(f"Error accessing body {i}: {str(body_error)}")
                    else:
                        if self.debug_planes:
                            if self.debug_planes:
                                self.add_comment("Bodies collection is None or empty")
                except Exception as bodies_error:
                    if self.debug_planes:
                        self.add_comment(f"Error accessing bodies collection: {str(bodies_error)}")
            else:
                if self.debug_planes:
                    self.add_comment("Extrude does not have bodies property")
            
            # Check linked features (for multi-component extrudes)
            if hasattr(extrude_feature, 'linkedFeatures'):
                try:
                    linked_features = extrude_feature.linkedFeatures
                    if linked_features and linked_features.count > 0:
                        if self.debug_planes:
                            self.add_comment(f"Found {linked_features.count} linked features")
                        
                        for i in range(linked_features.count):
                            linked_feature = linked_features.item(i)
                            if hasattr(linked_feature, 'bodies') and linked_feature.bodies:
                                for j in range(linked_feature.bodies.count):
                                    body = linked_feature.bodies.item(j)
                                    bodies.append(body)
                    else:
                        if self.debug_planes:
                            self.add_comment("No linked features found")
                except Exception as linked_error:
                    if self.debug_planes:
                        self.add_comment(f"Error accessing linked features: {str(linked_error)}")
            
            # Map each body to the KCL variable that created it using entity token
            for body in bodies:
                try:
                    body_token = body.entityToken
                    self.body_to_feature_map[body_token] = kcl_var_name
                    if self.debug_planes:
                        body_name = body.name if hasattr(body, 'name') else 'Unnamed body'
                        if self.debug_planes:
                            self.add_comment(f"Body tracking: {body_name} (token: {body_token}) created by {kcl_var_name}")
                except Exception as mapping_error:
                    if self.debug_planes:
                        self.add_comment(f"Error mapping body: {str(mapping_error)}")
            
            # Implement logical body tracking regardless of API access issues
            if len(bodies) == 0:
                if self.debug_planes:
                    if self.debug_planes:
                        self.add_comment(f"No bodies found via API - implementing logical tracking for {operation_name} operation")
                
                # Check component body count to understand the model state
                try:
                    component = extrude_feature.parentComponent
                    if component and hasattr(component, 'bRepBodies'):
                        component_bodies = component.bRepBodies
                        body_count = component_bodies.count
                        if self.debug_planes:
                            self.add_comment(f"Component has {body_count} total bodies")
                        
                        if body_count == 1:
                            # There's only one body in the component
                            single_body = component_bodies.item(0)
                            single_body_token = single_body.entityToken
                            self.body_to_feature_map[single_body_token] = kcl_var_name
                            bodies.append(single_body)
                            if self.debug_planes:
                                if self.debug_planes:
                                    self.add_comment(f"Logical tracking: {kcl_var_name} associated with single body")
                            
                        elif body_count > 1:
                            if self.debug_planes:
                                self.add_comment(f"Multiple bodies detected - this extrude created a new separate body")
                            # For multiple bodies, we need to find which one was created by this extrude
                            # Check each body to see if it was created by this extrude feature
                            found_new_body = False
                            for i in range(body_count):
                                body = component_bodies.item(i)
                                try:
                                    if hasattr(body, 'createdBy') and body.createdBy == extrude_feature:
                                        body_token = body.entityToken
                                        self.body_to_feature_map[body_token] = kcl_var_name
                                        bodies.append(body)
                                        found_new_body = True
                                        if self.debug_planes:
                                            self.add_comment(f"Found body created by {kcl_var_name} via createdBy check")
                                        break
                                except Exception as created_by_error:
                                    if self.debug_planes:
                                        self.add_comment(f"Error checking createdBy for body {i}: {str(created_by_error)}")
                            
                            # If we couldn't find via createdBy, fall back to assuming last body
                            if not found_new_body:
                                new_body = component_bodies.item(body_count - 1)
                                new_body_token = new_body.entityToken
                                self.body_to_feature_map[new_body_token] = kcl_var_name
                                bodies.append(new_body)
                                if self.debug_planes:
                                    self.add_comment(f"Fallback: assuming last body was created by {kcl_var_name}")
                        
                except Exception as comp_error:
                    if self.debug_planes:
                        self.add_comment(f"Error in logical body tracking: {str(comp_error)}")
            
            if self.debug_planes:
                if self.debug_planes:
                    self.add_comment(f"Successfully tracked {len(bodies)} bodies for {kcl_var_name}")
                if len(bodies) == 0:
                    self.add_comment("WARNING: No bodies were tracked for this extrude - this may cause issues with boolean operations")
                
        except Exception as e:
            if self.debug_planes:
                self.add_comment(f"Error tracking bodies for {kcl_var_name}: {str(e)}")
    
    
    def find_kcl_name_for_body(self, body) -> str:
        """Find the KCL variable name for a body by checking its creation feature."""
        try:
            # Check if we already have this body mapped using entity token
            body_token = body.entityToken
            if body_token in self.body_to_feature_map:
                kcl_name = self.body_to_feature_map[body_token]
                if self.debug_planes:
                    self.add_comment(f"Found cached mapping: {body_token} -> {kcl_name}")
                return kcl_name
            
            # Try to find the feature that created this body
            if hasattr(body, 'createdBy') and body.createdBy:
                creating_feature = body.createdBy
                creating_feature_token = creating_feature.entityToken
                if self.debug_planes:
                    self.add_comment(f"Body created by feature: {creating_feature.objectType}")
                    self.add_comment(f"Feature token: {creating_feature_token}")
                    
                if creating_feature_token in self.feature_to_kcl_name:
                    kcl_name = self.feature_to_kcl_name[creating_feature_token]
                    # Add this mapping for future use
                    self.body_to_feature_map[body_token] = kcl_name
                    if self.debug_planes:
                        self.add_comment(f"Found KCL name {kcl_name} for body via createdBy feature")
                    return kcl_name
                else:
                    if self.debug_planes:
                        self.add_comment(f"Creating feature token not found in feature_to_kcl_name mapping")
                        self.add_comment(f"Available feature tokens: {list(self.feature_to_kcl_name.keys())}")
            else:
                if self.debug_planes:
                    self.add_comment("Body has no createdBy feature")
            
            if self.debug_planes:
                self.add_comment("Could not find KCL name for body")
            return None
            
        except Exception as e:
            if self.debug_planes:
                self.add_comment(f"Error finding KCL name for body: {str(e)}")
            return None
    
    def track_combine_result(self, combine_feature: adsk.fusion.CombineFeature, kcl_var_name: str):
        """Track the result body created by a combine operation."""
        try:
            # Store the mapping from combine feature to KCL variable name using entity token
            combine_token = combine_feature.entityToken
            self.feature_to_kcl_name[combine_token] = kcl_var_name
            
            # Try to find the result body (this is tricky as combine operations modify existing bodies)
            # For now, we'll assume the target body becomes the result
            if hasattr(combine_feature, 'targetBody') and combine_feature.targetBody:
                target_body = combine_feature.targetBody
                target_body_token = target_body.entityToken
                # Update the mapping - the target body now represents the result of the combine
                self.body_to_feature_map[target_body_token] = kcl_var_name
                if self.debug_planes:
                    self.add_comment(f"Updated body mapping: target body (token: {target_body_token}) now maps to {kcl_var_name}")
            
        except Exception as e:
            if self.debug_planes:
                self.add_comment(f"Error tracking combine result: {str(e)}")
    
    
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
    
    def sort_curves_by_connectivity(self, all_curves):
        """Sort curves by their connectivity to form a continuous profile."""
        if not all_curves:
            return []
        
        if self.debug_planes:
            self.add_comment(f"Sorting {len(all_curves)} curves for connectivity")
        
        # Build a connectivity map
        connectivity_map = {}
        curve_endpoints = {}
        
        # First pass: collect all endpoints and build connectivity
        for i, (curve_type, curve) in enumerate(all_curves):
            start_point = self.get_curve_start_point(curve)
            end_point = self.get_curve_end_point(curve)
            
            if start_point and end_point:
                curve_endpoints[i] = {
                    'start': start_point,
                    'end': end_point,
                    'curve': (curve_type, curve),
                    'used': False
                }
                
                if self.debug_planes:
                    start_converted = self.convert_point_2d(start_point)
                    end_converted = self.convert_point_2d(end_point)
                    self.add_comment(f"Curve {i} ({curve_type}): {start_converted} -> {end_converted}")
        
        # Find the best starting curve (leftmost, then bottommost point)
        best_start_curve_idx = None
        best_start_point = None
        
        for i, curve_info in curve_endpoints.items():
            start_point = curve_info['start']
            start_converted = self.convert_point_2d(start_point)
            
            if best_start_point is None or (
                start_converted[0] < best_start_point[0] or 
                (abs(start_converted[0] - best_start_point[0]) < 0.001 and start_converted[1] < best_start_point[1])
            ):
                best_start_point = start_converted
                best_start_curve_idx = i
        
        if best_start_curve_idx is None:
            if self.debug_planes:
                self.add_comment("No valid starting curve found, using original order")
            return all_curves
        
        if self.debug_planes:
            self.add_comment(f"Starting with curve {best_start_curve_idx} at point {best_start_point}")
        
        # Trace the profile
        sorted_curves = []
        current_curve_idx = best_start_curve_idx
        current_end_point = curve_endpoints[current_curve_idx]['end']
        
        # Add the starting curve
        sorted_curves.append(curve_endpoints[current_curve_idx]['curve'])
        curve_endpoints[current_curve_idx]['used'] = True
        
        # Follow the chain
        while len(sorted_curves) < len(all_curves):
            found_next = False
            
            # Look for a curve that starts where we ended
            for i, curve_info in curve_endpoints.items():
                if curve_info['used']:
                    continue
                
                # Check if this curve starts where the current one ends
                if self.points_are_close(current_end_point, curve_info['start']):
                    sorted_curves.append(curve_info['curve'])
                    curve_info['used'] = True
                    current_end_point = curve_info['end']
                    current_curve_idx = i
                    found_next = True
                    if self.debug_planes:
                        end_converted = self.convert_point_2d(current_end_point)
                        self.add_comment(f"Connected to curve {i}, now at {end_converted}")
                    break
                
                # Check if this curve ends where the current one ends (reverse direction)
                elif self.points_are_close(current_end_point, curve_info['end']):
                    # We need to reverse this curve's direction conceptually
                    sorted_curves.append(curve_info['curve'])
                    curve_info['used'] = True
                    current_end_point = curve_info['start']
                    current_curve_idx = i
                    found_next = True
                    if self.debug_planes:
                        end_converted = self.convert_point_2d(current_end_point)
                        self.add_comment(f"Connected to curve {i} (reversed), now at {end_converted}")
                    break
            
            if not found_next:
                if self.debug_planes:
                    remaining = len(all_curves) - len(sorted_curves)
                    self.add_comment(f"Could not find next connected curve, {remaining} curves remaining")
                # Add any remaining curves
                for i, curve_info in curve_endpoints.items():
                    if not curve_info['used']:
                        sorted_curves.append(curve_info['curve'])
                        curve_info['used'] = True
                break
        
        if self.debug_planes:
            self.add_comment(f"Final curve order: {len(sorted_curves)} curves sorted")
        
        return sorted_curves
    
    def get_curve_start_point(self, curve):
        """Get the start point of a curve."""
        try:
            if hasattr(curve, 'startSketchPoint'):
                return curve.startSketchPoint.geometry
            elif hasattr(curve, 'centerSketchPoint'):
                # For circles, use center point
                return curve.centerSketchPoint.geometry
        except:
            pass
        return None
    
    def get_curve_end_point(self, curve):
        """Get the end point of a curve."""
        try:
            if hasattr(curve, 'endSketchPoint'):
                return curve.endSketchPoint.geometry
            elif hasattr(curve, 'centerSketchPoint'):
                # For circles, use center point (circles don't have end points)
                return curve.centerSketchPoint.geometry
        except:
            pass
        return None
    
    def points_are_close(self, point1, point2, tolerance=1e-6):
        """Check if two points are close enough to be considered the same."""
        if not point1 or not point2:
            return False
        try:
            dx = abs(point1.x - point2.x)
            dy = abs(point1.y - point2.y)
            return dx < tolerance and dy < tolerance
        except:
            return False

    def find_sketch_start_point(self, curves) -> tuple:
        """Find a good starting point for the sketch profile."""
        # Collect all curves to find the best starting point
        all_curves = []
        
        # Add lines
        for i in range(curves.sketchLines.count):
            line = curves.sketchLines.item(i)
            all_curves.append(('line', line))
        
        # Add arcs
        for i in range(curves.sketchArcs.count):
            arc = curves.sketchArcs.item(i)
            all_curves.append(('arc', arc))
        
        # Add circles (these are typically standalone, not part of profiles)
        for i in range(curves.sketchCircles.count):
            circle = curves.sketchCircles.item(i)
            all_curves.append(('circle', circle))
        
        # Add splines
        for i in range(curves.sketchFittedSplines.count):
            spline = curves.sketchFittedSplines.item(i)
            all_curves.append(('spline', spline))
        
        if not all_curves:
            return (0.0, 0.0)
        
        # Find the leftmost, then bottommost point among all curve start points
        best_point = None
        best_converted = None
        
        for curve_type, curve in all_curves:
            if curve_type == 'circle':
                # For circles, use center point
                if hasattr(curve, 'centerSketchPoint'):
                    point = curve.centerSketchPoint.geometry
                else:
                    continue
            else:
                # For other curves, use start point
                if hasattr(curve, 'startSketchPoint'):
                    point = curve.startSketchPoint.geometry
                else:
                    continue
            
            converted = self.convert_point_2d(point)
            
            if best_converted is None or (
                converted[0] < best_converted[0] or 
                (abs(converted[0] - best_converted[0]) < 0.001 and converted[1] < best_converted[1])
            ):
                best_point = point
                best_converted = converted
        
        if best_converted:
            return best_converted
        
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
        
        # Create the exporter (set debug_planes=True for detailed plane debugging)
        exporter = KCLExporter(debug_planes=True)
        
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
