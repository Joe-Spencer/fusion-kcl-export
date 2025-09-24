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
        self.units = "mm"  # Default units
        self.indent_level = 0
        self.debug_planes = debug_planes  # Enable detailed plane debugging
        
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
        
        # Store current sketch plane for coordinate conversion
        self.current_sketch_plane = plane_name
        
        # Reset profile tracking for this sketch
        self.current_profile_position = None
        
        # Debug: Check if sketch has any curves
        total_curves = (sketch.sketchCurves.sketchLines.count + 
                       sketch.sketchCurves.sketchArcs.count + 
                       sketch.sketchCurves.sketchCircles.count +
                       sketch.sketchCurves.sketchFittedSplines.count)
        
        self.add_comment(f"Sketch has {total_curves} total curves (lines: {sketch.sketchCurves.sketchLines.count}, arcs: {sketch.sketchCurves.sketchArcs.count}, circles: {sketch.sketchCurves.sketchCircles.count})")
        
        if total_curves == 0:
            self.add_comment(f"Skipping {sketch.name} - no curves found")
            return
        
        # Create the sketch and profile in one chain
        self.add_line(f'{sketch_var_name} = startSketchOn({plane_name})')
        self.indent_level += 1
        
        # Export sketch curves in the correct order (this will handle the starting point)
        self.export_sketch_curve(sketch.sketchCurves)
        
        # Close the profile
        self.add_line("|> close(%)")
        
        self.indent_level -= 1
        self.add_line("")
    
    def export_sketch_curve(self, curves):
        """Export sketch curves to KCL in the correct order."""
        # Collect all curves into a single list with their types
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
        
        # Sort curves by their order in the sketch profile
        sorted_curves = self.sort_curves_by_connectivity(all_curves)
        
        if not sorted_curves:
            return
        
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
    
    def export_line(self, line: adsk.fusion.SketchLine):
        """Export a sketch line to KCL."""
        start = line.startSketchPoint.geometry
        end = line.endSketchPoint.geometry
        
        start_x, start_y = self.convert_point_2d(start)
        end_x, end_y = self.convert_point_2d(end)
        
        # Check for zero-length lines
        tolerance = 0.001  # 0.001mm tolerance
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
        radius = self.convert_length(arc_geometry.radius)
        
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
        radius_mm = self.convert_length(radius)
        
        # For circles, use the correct KCL syntax (center and radius/diameter)
        diameter_mm = radius_mm * 2
        self.add_line(f"  |> circle(center = [{center_x}, {center_y}], diameter = {diameter_mm}, %)")
        
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
                                self.add_line(f"extrude{self.get_unique_id()} = {sketch_name} |> extrude(length = {adjusted_distance})")
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
                            self.add_line(f"extrude{self.get_unique_id()} = {sketch_name} |> extrude(length = {adjusted_distance})")
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
        """Export a combine feature to KCL (boolean operations)."""
        self.add_comment(f"Combine: {combine.name}")
        
        try:
            # Try to get the operation type safely
            operation = None
            operation_name = None
            
            try:
                operation = combine.operation
                self.add_comment(f"Raw operation value: {operation}")
            except Exception as op_error:
                self.add_comment(f"Could not get operation type: {str(op_error)}")
            
            # Map Fusion 360 operations to KCL operations
            if operation == adsk.fusion.FeatureOperations.JoinFeatureOperation:
                operation_name = "union"
                self.add_comment("Boolean operation: Join (Union)")
            elif operation == adsk.fusion.FeatureOperations.CutFeatureOperation:
                operation_name = "subtract"
                self.add_comment("Boolean operation: Cut (Subtract)")
            elif operation == adsk.fusion.FeatureOperations.IntersectFeatureOperation:
                operation_name = "intersect"
                self.add_comment("Boolean operation: Intersect")
            else:
                self.add_comment(f"Unsupported or unknown boolean operation: {operation}")
                # Try to infer from feature name or use a default
                feature_name_lower = combine.name.lower()
                if "join" in feature_name_lower or "union" in feature_name_lower:
                    operation_name = "union"
                    self.add_comment("Inferred operation: Union (from feature name)")
                elif "cut" in feature_name_lower or "subtract" in feature_name_lower:
                    operation_name = "subtract"
                    self.add_comment("Inferred operation: Subtract (from feature name)")
                elif "intersect" in feature_name_lower:
                    operation_name = "intersect"
                    self.add_comment("Inferred operation: Intersect (from feature name)")
                else:
                    operation_name = "union"  # Default fallback
                    self.add_comment("Using default operation: Union")
            
            # Try to get target and tool bodies safely
            target_bodies = None
            tool_bodies = None
            
            try:
                target_bodies = combine.targetBody
                self.add_comment(f"Target body found: {target_bodies is not None}")
            except Exception as target_error:
                self.add_comment(f"Could not get target body: {str(target_error)}")
            
            try:
                tool_bodies = combine.toolBodies
                self.add_comment(f"Tool bodies count: {tool_bodies.count if tool_bodies else 'None'}")
            except Exception as tool_error:
                self.add_comment(f"Could not get tool bodies: {str(tool_error)}")
            
            # If we can't get the bodies, create a generic boolean operation
            if not target_bodies or not tool_bodies:
                self.add_comment("Cannot access combine bodies - generating generic boolean operation")
                # Use the previous two features as a fallback
                prev_feature_1 = f"extrude{self.get_unique_id()}"
                prev_feature_2 = f"extrude{self.get_unique_id()}"
                
                solid_id = str(self.get_unique_id()).zfill(3)  # Format as 001, 002, etc.
                if operation_name == "subtract":
                    self.add_line(f"solid{solid_id} = {operation_name}(extrude1, tools = extrude2)")
                else:
                    self.add_line(f"solid{solid_id} = {operation_name}(extrude1, extrude2)")
                return
            
            # Find the source features for the bodies
            target_feature_name = self.find_body_source_feature(target_bodies)
            
            if tool_bodies and tool_bodies.count > 0:
                # Get tool body feature names
                tool_feature_names = []
                for i in range(tool_bodies.count):
                    tool_body = tool_bodies.item(i)
                    tool_name = self.find_body_source_feature(tool_body)
                    if tool_name:
                        tool_feature_names.append(tool_name)
                
                if target_feature_name and tool_feature_names:
                    # Generate KCL boolean operation
                    solid_id = str(self.get_unique_id()).zfill(3)  # Format as 001, 002, etc.
                    if operation_name == "subtract":
                        # subtract(target, tools = [tool1, tool2, ...])
                        if len(tool_feature_names) == 1:
                            tools_str = tool_feature_names[0]
                        else:
                            tools_str = f"[{', '.join(tool_feature_names)}]"
                        self.add_line(f"solid{solid_id} = {operation_name}({target_feature_name}, tools = {tools_str})")
                    else:
                        # union(feature1, feature2, ...) or intersect(feature1, feature2, ...)
                        all_features = [target_feature_name] + tool_feature_names
                        features_str = ', '.join(all_features)
                        self.add_line(f"solid{solid_id} = {operation_name}({features_str})")
                else:
                    self.add_comment("Warning: Could not identify source features - using generic names")
                    # Fallback with generic feature names
                    solid_id = str(self.get_unique_id()).zfill(3)  # Format as 001, 002, etc.
                    if operation_name == "subtract":
                        self.add_line(f"solid{solid_id} = {operation_name}(extrude1, tools = extrude2)")
                    else:
                        self.add_line(f"solid{solid_id} = {operation_name}(extrude1, extrude2)")
            else:
                self.add_comment("Warning: No tool bodies found for combine operation")
                
        except Exception as e:
            self.add_comment(f"Error processing combine: {str(e)}")
            # Provide a fallback boolean operation
            self.add_comment("Generating fallback boolean operation")
            solid_id = str(self.get_unique_id()).zfill(3)  # Format as 001, 002, etc.
            self.add_line(f"solid{solid_id} = union(extrude1, extrude2)")
        
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
        x = self.convert_length(point.x)
        y = self.convert_length(point.y)
        
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
    
    def convert_length(self, length_cm: float) -> float:
        """Convert length from Fusion 360 internal units (cm) to mm."""
        return round(length_cm * 10, 3)  # Convert cm to mm and round to 3 decimal places
    
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
