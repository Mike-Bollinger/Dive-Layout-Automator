"""
================================================================================
DIVE LAYOUT AUTOMATOR - ARCGIS PRO TOOLBOX
================================================================================

Title:          Dive Layout Automator
Author:         Mike Bollinger
Version:        1.0.0
Date Created:   June 2025
Date Updated:   June 11, 2025
Description:    Interactive ArcGIS Pro toolbox for automating the creation of 
                dive maps with manual extent control and scale bar management.

================================================================================
OVERVIEW
================================================================================

This toolbox automates the process of creating standardized maps for multiple 
dive sites while maintaining user control over map extent and scale bar units.
The tool provides an interactive workflow that combines automation with manual
fine-tuning capabilities.

Key Features:
- Automatically detects dive numbers from layer names
- Sets layer visibility for each dive sequentially
- Auto-zooms to dive extent as starting point
- Interactive GUI for manual extent adjustment
- Manual scale bar unit control (Meters, Kilometers, Feet, Miles)
- Customizable file naming with prefix/suffix options
- Progress tracking and error handling
- Scrollable completion log

================================================================================
WORKFLOW
================================================================================

1. User selects layout, layer groups, and output settings in toolbox
2. Tool identifies all dive numbers in selected groups
3. Interactive GUI opens with instructions
4. For each dive:
   a. Tool makes dive layers visible and hides others
   b. Tool auto-zooms to dive extent (15% buffer)
   c. User manually adjusts scale bar units if needed
   d. User fine-tunes map extent using ArcPro navigation
   e. User clicks "Export & Next" to save map and proceed
5. Process continues until all dives are completed

================================================================================
REQUIREMENTS
================================================================================

- ArcGIS Pro 2.8 or later
- Python 3.7+ with tkinter
- Layout with map frame and scale bar elements
- Layer groups containing dive-specific layers
- Dive layers named with recognizable patterns:
  * DIVE001, DIVE002, etc.
  * Dive1, Dive2, etc.
  * D001, D002, etc.
  * Or just numbers: 001, 002, etc.

================================================================================
PARAMETERS
================================================================================

Layout:                 Target layout for map export
Layer Groups:           Groups containing dive layers (multiselect)
Output Folder:          Destination for exported JPG files
File Name Prefix:       Optional text before dive name (e.g., "PC2403_")
File Name Suffix:       Optional text after dive name (e.g., "_Final")
Export DPI:             Resolution for exported images (default: 300)

================================================================================
OUTPUT
================================================================================

- High-quality JPG maps for each dive
- Custom naming: [prefix]DIVE001[suffix].jpg
- Consistent scale bar units across maps
- Detailed processing log in ArcPro Messages

================================================================================
TECHNICAL NOTES
================================================================================

- Uses ArcPy mapping module for layout control
- Implements CIM (Cartographic Information Model) for scale bar updates
- Tkinter GUI provides cross-platform interface
- Regex pattern matching for flexible dive name detection
- Auto-zoom targets largest map frame (avoids extent indicators)
- Modal GUI ensures user interaction before proceeding

================================================================================
VERSION HISTORY
================================================================================

v1.0.0 (June 11, 2025) - Initial release
- Interactive dive map generation
- Auto-zoom functionality
- Manual scale bar control
- Custom file naming
- Progress tracking GUI
- Comprehensive error handling

================================================================================
SUPPORT
================================================================================

For questions, issues, or feature requests:
- Check ArcPro Messages panel for detailed error information
- Ensure dive layer naming follows supported patterns
- Verify layout contains proper map frame and scale bar elements

================================================================================
"""

import arcpy
import os
import re
import tkinter as tk
from tkinter import ttk, messagebox

class Toolbox(object):
    def __init__(self):
        self.label = "Dive Layout Automator"
        self.alias = "DiveLayoutAutomator"
        self.tools = [DiveLayoutTool]

class DiveLayoutTool(object):
    def __init__(self):
        self.label = "Interactive Dive Layout Generator"
        self.description = "Generate maps for dive layers with interactive extent control"
        self.canRunInBackground = False
        
    def getParameterInfo(self):
        """Define tool parameters"""
        
        # Layout parameter
        layout_param = arcpy.Parameter(
            displayName="Layout",
            name="layout",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        layout_param.filter.type = "ValueList"
        
        # Layer groups parameter
        groups_param = arcpy.Parameter(
            displayName="Layer Groups (containing dive layers)",
            name="layer_groups", 
            datatype="GPString",
            parameterType="Required",
            direction="Input",
            multiValue=True
        )
        groups_param.filter.type = "ValueList"
        
        # Output folder parameter
        output_param = arcpy.Parameter(
            displayName="Output Folder",
            name="output_folder",
            datatype="DEFolder", 
            parameterType="Required",
            direction="Input"
        )
        
        # File prefix parameter
        prefix_param = arcpy.Parameter(
            displayName="File Name Prefix (optional)",
            name="file_prefix",
            datatype="GPString",
            parameterType="Optional",
            direction="Input"
        )
        prefix_param.value = ""
        
        # File suffix parameter
        suffix_param = arcpy.Parameter(
            displayName="File Name Suffix (optional)",
            name="file_suffix",
            datatype="GPString",
            parameterType="Optional",
            direction="Input"
        )
        suffix_param.value = ""
        
        # DPI parameter
        dpi_param = arcpy.Parameter(
            displayName="Export DPI",
            name="dpi",
            datatype="GPLong",
            parameterType="Optional",
            direction="Input"
        )
        dpi_param.value = 300
        
        return [layout_param, groups_param, output_param, prefix_param, suffix_param, dpi_param]
    
    def isLicensed(self):
        """Check if tool is licensed to execute"""
        return True
    
    def updateParameters(self, parameters):
        """Update parameter values and properties"""
        try:
            # Get current project
            aprx = arcpy.mp.ArcGISProject("CURRENT")
            
            # Update layout choices
            layouts = [layout.name for layout in aprx.listLayouts()]
            parameters[0].filter.list = layouts
            
            # Update layer group choices
            group_choices = []
            for map_obj in aprx.listMaps():
                for layer in map_obj.listLayers():
                    if hasattr(layer, 'isGroupLayer') and layer.isGroupLayer:
                        group_choices.append(f"{map_obj.name}:{layer.name}")
            
            parameters[1].filter.list = group_choices
            
        except Exception as e:
            arcpy.AddWarning(f"Could not update parameters: {str(e)}")
        
        return
    
    def updateMessages(self, parameters):
        """Update messages for parameters"""
        # Check for invalid filename characters in prefix/suffix
        invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
        
        if parameters[3].altered:  # prefix parameter
            prefix = parameters[3].valueAsText
            if prefix and any(char in prefix for char in invalid_chars):
                parameters[3].setErrorMessage("Prefix contains invalid filename characters: < > : \" / \\ | ? *")
        
        if parameters[4].altered:  # suffix parameter
            suffix = parameters[4].valueAsText
            if suffix and any(char in suffix for char in invalid_chars):
                parameters[4].setErrorMessage("Suffix contains invalid filename characters: < > : \" / \\ | ? *")
        
        return
    
    def execute(self, parameters, messages):
        """Execute the tool"""
        try:
            # Get parameters
            layout_name = parameters[0].valueAsText
            
            # Fix the parameter parsing for multiValue groups
            if parameters[1].valueAsText:
                # Handle both single and multiple selections
                raw_groups = parameters[1].valueAsText
                # Remove quotes and split properly
                raw_groups = raw_groups.replace("'", "").replace('"', '')
                selected_groups = [group.strip() for group in raw_groups.split(';') if group.strip()]
            else:
                selected_groups = []
                
            output_folder = parameters[2].valueAsText
            file_prefix = parameters[3].valueAsText if parameters[3].valueAsText else ""
            file_suffix = parameters[4].valueAsText if parameters[4].valueAsText else ""
            dpi = parameters[5].value if parameters[5].value else 300
            
            arcpy.AddMessage(f"Starting dive layout automation...")
            arcpy.AddMessage(f"Layout: {layout_name}")
            arcpy.AddMessage(f"Groups: {len(selected_groups)} selected")
            
            # Debug: Show the cleaned group names
            arcpy.AddMessage("Selected groups:")
            for i, group in enumerate(selected_groups):
                arcpy.AddMessage(f"  {i+1}. '{group}'")
            
            arcpy.AddMessage(f"Output: {output_folder}")
            
            # Show file naming pattern
            example_name = f"{file_prefix}DIVE001{file_suffix}.jpg" if file_prefix or file_suffix else "DIVE001.jpg"
            arcpy.AddMessage(f"File naming pattern: {example_name}")
            
            # Initialize the processor
            processor = DiveLayoutProcessor(layout_name, selected_groups, output_folder, dpi, file_prefix, file_suffix)
            
            # Run the interactive process
            processor.run_interactive_process()
            
        except Exception as e:
            arcpy.AddError(f"Tool execution failed: {str(e)}")
            import traceback
            arcpy.AddError(traceback.format_exc())

class DiveLayoutProcessor:
    """Main processor for dive layout automation"""
    
    def __init__(self, layout_name, selected_groups, output_folder, dpi, file_prefix="", file_suffix=""):
        self.layout_name = layout_name
        self.selected_groups = selected_groups
        self.output_folder = output_folder
        self.dpi = dpi
        self.file_prefix = file_prefix
        self.file_suffix = file_suffix
        self.aprx = arcpy.mp.ArcGISProject("CURRENT")
        self.dive_pattern = re.compile(r'DIVE(\d+)', re.IGNORECASE)
        
        # Create output folder
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
    
    def find_dive_numbers(self):
        """Find all dive numbers in selected layer groups"""
        dive_numbers = set()
        
        arcpy.AddMessage(f"Analyzing {len(self.selected_groups)} selected groups for dive layers...")
        
        # First, let's see what maps are available
        available_maps = [m.name for m in self.aprx.listMaps()]
        arcpy.AddMessage(f"Available maps in project: {available_maps}")
        
        for group_key in self.selected_groups:
            try:
                # Debug the splitting
                arcpy.AddMessage(f"Processing group_key: '{group_key}'")
                
                if ':' not in group_key:
                    arcpy.AddWarning(f"Invalid group format (expected 'Map:Group'): '{group_key}'")
                    continue
                    
                map_name, group_name = group_key.split(':', 1)
                arcpy.AddMessage(f"Parsed - Map: '{map_name}', Group: '{group_name}'")
                
                # Get map
                map_obj = None
                for m in self.aprx.listMaps():
                    arcpy.AddMessage(f"Checking map: '{m.name}' against '{map_name}'")
                    if m.name == map_name:
                        map_obj = m
                        break
                
                if not map_obj:
                    arcpy.AddWarning(f"Map '{map_name}' not found")
                    arcpy.AddWarning(f"Available maps: {[m.name for m in self.aprx.listMaps()]}")
                    continue
                
                arcpy.AddMessage(f"Found map: '{map_obj.name}'")
                
                # Get group layer
                group_layer = None
                available_groups = []
                for layer in map_obj.listLayers():
                    available_groups.append(layer.name)
                    if layer.name == group_name:
                        group_layer = layer
                        break
                
                if not group_layer:
                    arcpy.AddWarning(f"Group layer '{group_name}' not found in map '{map_name}'")
                    arcpy.AddWarning(f"Available groups in '{map_name}': {available_groups}")
                    continue
                
                arcpy.AddMessage(f"Found group layer: '{group_layer.name}'")
                
                # Find dive layers - let's be more flexible with the pattern
                layers_in_group = list(group_layer.listLayers())
                arcpy.AddMessage(f"Found {len(layers_in_group)} layers in group '{group_name}':")
                
                for layer in layers_in_group:
                    arcpy.AddMessage(f"  - {layer.name}")
                    
                    # Try multiple dive patterns
                    patterns_to_try = [
                        r'DIVE(\d+)',           # DIVE001, DIVE002, etc.
                        r'Dive(\d+)',           # Dive001, Dive002, etc.
                        r'dive(\d+)',           # dive001, dive002, etc.
                        r'D(\d+)',              # D001, D002, etc.
                        r'(\d+)',               # Just numbers: 001, 002, etc.
                    ]
                    
                    found_match = False
                    for pattern in patterns_to_try:
                        match = re.search(pattern, layer.name, re.IGNORECASE)
                        if match:
                            dive_num = int(match.group(1))
                            dive_numbers.add(dive_num)
                            arcpy.AddMessage(f"    ✓ Found dive number: {dive_num} (pattern: {pattern})")
                            found_match = True
                            break
                    
                    if not found_match:
                        arcpy.AddMessage(f"    - No dive pattern found in: {layer.name}")
            
            except Exception as e:
                arcpy.AddError(f"Error processing group '{group_key}': {str(e)}")
                import traceback
                arcpy.AddError(traceback.format_exc())
    
        if dive_numbers:
            sorted_dives = sorted(dive_numbers)
            arcpy.AddMessage(f"\nTotal dive numbers found: {sorted_dives}")
            return sorted_dives
        else:
            arcpy.AddWarning("No dive layers found. Please check:")
            arcpy.AddWarning("1. Layer names contain dive numbers (e.g., DIVE001, Dive1, D001)")
            arcpy.AddWarning("2. Selected groups contain the correct layers")
            arcpy.AddWarning("3. Layer names are spelled correctly")
            return []

    def set_dive_visibility(self, dive_number):
        """Set layer visibility for a specific dive"""
        dive_patterns = [
            f"DIVE{dive_number:03d}",  # DIVE001
            f"Dive{dive_number:03d}",  # Dive001
            f"dive{dive_number:03d}",  # dive001
            f"D{dive_number:03d}",     # D001
            f"DIVE{dive_number}",      # DIVE1
            f"Dive{dive_number}",      # Dive1
            f"dive{dive_number}",      # dive1
            f"D{dive_number}",         # D1
            f"{dive_number:03d}",      # 001
            f"{dive_number}",          # 1
        ]
        
        arcpy.AddMessage(f"Setting visibility for dive {dive_number}...")
        
        for group_key in self.selected_groups:
            map_name, group_name = group_key.split(':', 1)
            
            # Get map and group
            map_obj = None
            for m in self.aprx.listMaps():
                if m.name == map_name:
                    map_obj = m
                    break
            
            if not map_obj:
                continue
            
            group_layer = None
            for layer in map_obj.listLayers():
                if layer.name == group_name:
                    group_layer = layer
                    break
            
            if not group_layer:
                continue
            
            # Set visibility for all layers in group
            for layer in group_layer.listLayers():
                layer_name_upper = layer.name.upper()
                
                # Check if this layer matches any dive pattern
                is_dive_layer = False
                matches_current_dive = False
                
                # Check against all possible dive patterns
                for pattern in [r'DIVE(\d+)', r'D(\d+)', r'(\d+)']:
                    match = re.search(pattern, layer.name, re.IGNORECASE)
                    if match:
                        is_dive_layer = True
                        layer_dive_num = int(match.group(1))
                        if layer_dive_num == dive_number:
                            matches_current_dive = True
                        break
                
                if is_dive_layer:
                    if matches_current_dive:
                        layer.visible = True
                        arcpy.AddMessage(f"  ✓ Showing: {layer.name}")
                    else:
                        layer.visible = False
                        arcpy.AddMessage(f"  - Hiding: {layer.name}")
                else:
                    # Non-dive layer - leave as is
                    arcpy.AddMessage(f"  → Keeping: {layer.name} (not a dive layer)")
    
    def export_current_layout(self, dive_number):
        """Export the current layout state with custom naming"""
        dive_id = f"DIVE{dive_number:03d}"
        
        # Build the filename with prefix and suffix
        filename = f"{self.file_prefix}{dive_id}{self.file_suffix}.jpg"
        output_file = os.path.join(self.output_folder, filename)
        
        # Get layout
        layout = None
        for l in self.aprx.listLayouts():
            if l.name == self.layout_name:
                layout = l
                break
        
        if not layout:
            raise Exception(f"Layout '{self.layout_name}' not found")
        
        # Export
        layout.exportToJPEG(output_file, resolution=self.dpi, jpeg_quality=95)
        arcpy.AddMessage(f"  Exported: {filename}")
        
        return output_file
    
    def run_interactive_process(self):
        """Run the interactive process with GUI"""
        # Find all dives
        dive_numbers = self.find_dive_numbers()
        
        if not dive_numbers:
            arcpy.AddWarning("No dive layers found in selected groups")
            return
        
        arcpy.AddMessage(f"Found {len(dive_numbers)} dives: {dive_numbers}")
        
        # Create and run the interactive GUI
        gui = InteractiveExtentGUI(self, dive_numbers)
        self._current_gui = gui  # Store reference for auto-updates
        gui.run()

    def zoom_to_dive_layers(self, dive_number):
        """Zoom to the extent of the current dive's visible layers"""
        try:
            dive_id = f"DIVE{dive_number:03d}"
            arcpy.AddMessage(f"Auto-zooming to {dive_id} layers...")
            
            # Get the layout and map frame
            layout = None
            for l in self.aprx.listLayouts():
                if l.name == self.layout_name:
                    layout = l
                    break
            
            if not layout:
                arcpy.AddWarning(f"Layout '{self.layout_name}' not found for auto-zoom")
                return
            
            # Find the MAIN map frame (not extent indicators)
            map_frames = layout.listElements("MAPFRAME_ELEMENT")
            if not map_frames:
                arcpy.AddWarning("No map frame found in layout for auto-zoom")
                return
            
            # Find the largest map frame (main map, not extent indicators)
            main_map_frame = None
            largest_area = 0
            
            for mf in map_frames:
                # Calculate area of this map frame
                area = mf.elementWidth * mf.elementHeight
                arcpy.AddMessage(f"Map frame '{mf.name}' area: {area:.2f}")
                
                if area > largest_area:
                    largest_area = area
                    main_map_frame = mf
            
            if not main_map_frame:
                arcpy.AddWarning("Could not identify main map frame")
                return
            
            arcpy.AddMessage(f"Using main map frame: '{main_map_frame.name}' (area: {largest_area:.2f})")
            
            # Collect visible dive layers
            visible_layers = []
            for group_key in self.selected_groups:
                map_name, group_name = group_key.split(':', 1)
                
                # Get map and group
                map_obj = None
                for m in self.aprx.listMaps():
                    if m.name == map_name:
                        map_obj = m
                        break
                
                if not map_obj:
                    continue
                
                group_layer = None
                for layer in map_obj.listLayers():
                    if layer.name == group_name:
                        group_layer = layer
                        break
                
                if not group_layer:
                    continue
                
                # Get visible layers for this dive
                for layer in group_layer.listLayers():
                    if layer.visible:
                        # Check if it's a dive layer matching current dive
                        for pattern in [r'DIVE(\d+)', r'D(\d+)', r'(\d+)']:
                            match = re.search(pattern, layer.name, re.IGNORECASE)
                            if match:
                                layer_dive_num = int(match.group(1))
                                if layer_dive_num == dive_number:
                                    visible_layers.append(layer)
                                    arcpy.AddMessage(f"Found visible layer for {dive_id}: {layer.name}")
                                    break
    
            if not visible_layers:
                arcpy.AddWarning(f"No visible layers found for {dive_id} to zoom to")
                return
            
            arcpy.AddMessage(f"Zooming to {len(visible_layers)} visible layers for {dive_id}")
            
            # Calculate combined extent
            all_extents = []
            for layer in visible_layers:
                try:
                    desc = arcpy.Describe(layer)
                    extent = desc.extent
                    if extent.width > 0 and extent.height > 0:
                        all_extents.append(extent)
                        arcpy.AddMessage(f"Layer {layer.name} extent: ({extent.XMin:.6f}, {extent.YMin:.6f}) to ({extent.XMax:.6f}, {extent.YMax:.6f})")
                except Exception as e:
                    arcpy.AddWarning(f"Could not get extent for layer {layer.name}: {str(e)}")
            
            if not all_extents:
                arcpy.AddWarning(f"No valid extents found for {dive_id}")
                return
            
            # Calculate combined extent
            min_x = min(ext.XMin for ext in all_extents)
            min_y = min(ext.YMin for ext in all_extents)
            max_x = max(ext.XMax for ext in all_extents)
            max_y = max(ext.YMax for ext in all_extents)
            
            combined_extent = arcpy.Extent(min_x, min_y, max_x, max_y)
            
            # Add buffer (15% on each side)
            buffer_pct = 0.15
            buffer_x = combined_extent.width * buffer_pct
            buffer_y = combined_extent.height * buffer_pct
            
            buffered_extent = arcpy.Extent(
                combined_extent.XMin - buffer_x,
                combined_extent.YMin - buffer_y,
                combined_extent.XMax + buffer_x,
                combined_extent.YMax + buffer_y
            )
            
            # Set the extent in the MAIN map frame
            main_map_frame.camera.setExtent(buffered_extent)
            
            arcpy.AddMessage(f"Auto-zoomed MAIN map frame to {dive_id} with 15% buffer")
            arcpy.AddMessage(f"Extent: ({buffered_extent.XMin:.6f}, {buffered_extent.YMin:.6f}) to ({buffered_extent.XMax:.6f}, {buffered_extent.YMax:.6f})")
            
        except Exception as e:
            arcpy.AddWarning(f"Error auto-zooming to {dive_id}: {str(e)}")
            import traceback
            arcpy.AddWarning(traceback.format_exc())

    def set_scale_bar_units(self, units):
        """Set scale bar units using CIM (Cartographic Information Model)"""
        try:
            import json
            
            # Get the layout
            layout = None
            for l in self.aprx.listLayouts():
                if l.name == self.layout_name:
                    layout = l
                    break
            
            if not layout:
                arcpy.AddWarning(f"Layout '{self.layout_name}' not found for scale bar update")
                return
            
            # Map unit names to their WKID codes
            unit_codes = {
                "Meters": {"uwkid": 9001},
                "Kilometers": {"uwkid": 9036}, 
                "Feet": {"uwkid": 9002},
                "Miles": {"uwkid": 9093},
                "Nautical Miles": {"uwkid": 9030}
            }
            
            if units not in unit_codes:
                arcpy.AddWarning(f"Unsupported units: {units}. Supported: {list(unit_codes.keys())}")
                return
            
            # Get the layout CIM definition
            layout_cim = layout.getDefinition("V3")
            
            # Find scale bar elements
            scale_bars_found = 0
            
            for element in layout_cim.elements:
                # Check if this is a scale bar element
                if hasattr(element, 'unitLabel') and hasattr(element, 'units'):
                    arcpy.AddMessage(f"Found scale bar element: {element.name}")
                    
                    # Update the scale bar
                    element.unitLabel = units
                    element.units = unit_codes[units]
                    
                    # Optional: Set rounding for cleaner display
                    if hasattr(element, 'numberFormat') and hasattr(element.numberFormat, 'roundingValue'):
                        if units in ["Kilometers", "Miles"]:
                            element.numberFormat.roundingValue = 1  # Round to 1 decimal for larger units
                        else:
                            element.numberFormat.roundingValue = 0  # Round to whole numbers for smaller units
                    
                    scale_bars_found += 1
                    arcpy.AddMessage(f"Updated scale bar '{element.name}' to {units}")
            
            if scale_bars_found == 0:
                # If no scale bars found by properties, try searching by name
                arcpy.AddMessage("No scale bars found by properties, searching by name...")
                
                for element in layout_cim.elements:
                    if hasattr(element, 'name') and 'scale' in element.name.lower():
                        arcpy.AddMessage(f"Found potential scale bar by name: {element.name}")
                        arcpy.AddMessage(f"Element type: {type(element)}")
                        
                        # List all properties of this element
                        for attr in dir(element):
                            if not attr.startswith('_'):
                                try:
                                    value = getattr(element, attr)
                                    if not callable(value):
                                        arcpy.AddMessage(f"  {attr}: {value}")
                                except:
                                    pass
        
            if scale_bars_found > 0:
                # Apply the changes back to the layout
                layout.setDefinition(layout_cim)
                arcpy.AddMessage(f"Successfully updated {scale_bars_found} scale bar(s) to {units}")
            else:
                arcpy.AddWarning("No scale bar elements found to update")
                arcpy.AddMessage("Available elements in layout:")
                for element in layout_cim.elements:
                    if hasattr(element, 'name'):
                        arcpy.AddMessage(f"  - {element.name} (type: {type(element).__name__})")
                        
        except Exception as e:
            arcpy.AddWarning(f"Error setting scale bar units: {str(e)}")
            import traceback
            arcpy.AddWarning(traceback.format_exc())

    def set_dive_visibility(self, dive_number):
        """Set layer visibility for a specific dive, auto-zoom"""
        dive_patterns = [
            f"DIVE{dive_number:03d}",  # DIVE001
            f"Dive{dive_number:03d}",  # Dive001
            f"dive{dive_number:03d}",  # dive001
            f"D{dive_number:03d}",     # D001
            f"DIVE{dive_number}",      # DIVE1
            f"Dive{dive_number}",      # Dive1
            f"dive{dive_number}",      # dive1
            f"D{dive_number}",         # D1
            f"{dive_number:03d}",      # 001
            f"{dive_number}",          # 1
        ]
        
        arcpy.AddMessage(f"Setting visibility for dive {dive_number}...")
        
        for group_key in self.selected_groups:
            map_name, group_name = group_key.split(':', 1)
            
            # Get map and group
            map_obj = None
            for m in self.aprx.listMaps():
                if m.name == map_name:
                    map_obj = m
                    break
            
            if not map_obj:
                continue
            
            group_layer = None
            for layer in map_obj.listLayers():
                if layer.name == group_name:
                    group_layer = layer
                    break
            
            if not group_layer:
                continue
            
            # Set visibility for all layers in group
            for layer in group_layer.listLayers():
                layer_name_upper = layer.name.upper()
                
                # Check if this layer matches any dive pattern
                is_dive_layer = False
                matches_current_dive = False
                
                # Check against all possible dive patterns
                for pattern in [r'DIVE(\d+)', r'D(\d+)', r'(\d+)']:
                    match = re.search(pattern, layer.name, re.IGNORECASE)
                    if match:
                        is_dive_layer = True
                        layer_dive_num = int(match.group(1))
                        if layer_dive_num == dive_number:
                            matches_current_dive = True
                        break
                
                if is_dive_layer:
                    if matches_current_dive:
                        layer.visible = True
                        arcpy.AddMessage(f"  ✓ Showing: {layer.name}")
                    else:
                        layer.visible = False
                        arcpy.AddMessage(f"  - Hiding: {layer.name}")
                else:
                    # Non-dive layer - leave as is
                    arcpy.AddMessage(f"  → Keeping: {layer.name} (not a dive layer)")
    
        # Auto-zoom to the dive layers after setting visibility (with a small delay)
        try:
            import time
            time.sleep(0.5)  # Small delay to let visibility changes take effect
            self.zoom_to_dive_layers(dive_number)
            
        except Exception as e:
            arcpy.AddWarning(f"Error in auto-zoom: {str(e)}")

class InteractiveExtentGUI:
    """Simple GUI for interactive extent control"""
    
    def __init__(self, processor, dive_numbers):
        self.processor = processor
        self.dive_numbers = dive_numbers
        self.current_index = 0
        self.completed_dives = []
        
    def run(self):
        """Run the interactive GUI"""
        self.root = tk.Tk()
        self.root.title("Dive Layout Automator - Interactive Mode")
        self.root.geometry("550x700")  # Increased height for scale bar controls
        self.root.attributes('-topmost', True)  # Keep on top
        
        # Set custom NOAA logo as window icon
        try:
            # Look for NOAA_Logo in same folder as the .pyt file
            logo_path = os.path.join(os.path.dirname(__file__), "NOAA_Logo.ico")
            
            if os.path.exists(logo_path):
                self.root.iconbitmap(logo_path)
                arcpy.AddMessage(f"Using NOAA logo icon: {logo_path}")
            else:
                # Try PNG version if ICO not found
                png_logo_path = os.path.join(os.path.dirname(__file__), "NOAA_Logo.png")
                if os.path.exists(png_logo_path):
                    from PIL import Image, ImageTk
                    icon_image = Image.open(png_logo_path)
                    icon_image = icon_image.resize((32, 32), Image.Resampling.LANCZOS)
                    icon_photo = ImageTk.PhotoImage(icon_image)
                    self.root.iconphoto(True, icon_photo)
                    # Keep a reference to prevent garbage collection
                    self.root.icon_photo = icon_photo
                    arcpy.AddMessage(f"Using NOAA PNG logo: {png_logo_path}")
                else:
                    arcpy.AddMessage("NOAA_Logo not found, using default tkinter icon")
                    
        except Exception as e:
            arcpy.AddMessage(f"Could not load NOAA logo ({str(e)}), using default tkinter icon")
        
        # Make window modal
        self.root.grab_set()
        
        self.create_widgets()
        self.load_current_dive()
        
        # Center the window
        self.root.geometry("+%d+%d" % (
            (self.root.winfo_screenwidth() / 2) - 275,  # Adjusted for new width
            (self.root.winfo_screenheight() / 2) - 350   # Adjusted for new height
        ))
        
        self.root.mainloop()
    
    def create_widgets(self):
        """Create GUI widgets"""
        # Title frame with text logo
        title_frame = ttk.Frame(self.root)
        title_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Header with text logo
        header_frame = ttk.Frame(title_frame)
        header_frame.pack(fill=tk.X)
        
        # Text-based logo with coral reef emoji
        logo_label = ttk.Label(header_frame, text="🪸", font=("Arial", 32))  # Large coral reef emoji
        logo_label.pack(side=tk.LEFT, padx=(0, 15))
        
        # Title beside logo
        title_text_frame = ttk.Frame(header_frame)
        title_text_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Label(title_text_frame, text="Interactive Dive Layout Generator", 
                 font=("Arial", 14, "bold")).pack(anchor=tk.W)
        
        # Optional: Add subtitle
        ttk.Label(title_text_frame, text="Semi-Automated Dive Map Creation", 
                 font=("Arial", 9), foreground="gray").pack(anchor=tk.W)
        
        # Current dive info
        info_frame = ttk.LabelFrame(self.root, text="Current Dive", padding="10")
        info_frame.pack(fill=tk.X, padx=20, pady=10)
        
        self.dive_label = ttk.Label(info_frame, text="", font=("Arial", 12, "bold"))
        self.dive_label.pack()
        
        self.progress_label = ttk.Label(info_frame, text="")
        self.progress_label.pack()
        
        # Scale bar controls
        scale_frame = ttk.LabelFrame(self.root, text="Scale Bar Units (Manual Control)", padding="10")
        scale_frame.pack(fill=tk.X, padx=20, pady=10)
        
        scale_control_frame = ttk.Frame(scale_frame)
        scale_control_frame.pack(fill=tk.X)
        
        ttk.Label(scale_control_frame, text="Units:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.scale_var = tk.StringVar(value="Meters")
        scale_combo = ttk.Combobox(scale_control_frame, textvariable=self.scale_var, 
                                  values=["Meters", "Kilometers", "Feet", "Miles"], 
                                  state="readonly", width=12)
        scale_combo.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(scale_control_frame, text="Apply Units", 
                  command=self.apply_scale_units).pack(side=tk.LEFT)
        
        # Instructions
        instructions_frame = ttk.LabelFrame(self.root, text="Instructions", padding="10")
        instructions_frame.pack(fill=tk.X, padx=20, pady=10)
        
        instructions = """1. The layers for the current dive are now visible and auto-zoomed
2. Manually adjust scale bar units using the dropdown above if needed
3. Use ArcPro's navigation tools to fine-tune the map extent
4. Click 'Export & Next' when you're satisfied with the extent
5. Repeat for each dive"""
        
        ttk.Label(instructions_frame, text=instructions, justify=tk.LEFT).pack()
        
        # Status
        self.status_label = ttk.Label(self.root, text="Ready to start...", 
                                     foreground="blue")
        self.status_label.pack(pady=10)
        
        # Buttons
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=20, pady=20)
        
        # Add Re-zoom button
        self.rezoom_button = ttk.Button(button_frame, text="Re-zoom to Layers", 
                                       command=self.rezoom_to_layers)
        self.rezoom_button.pack(side=tk.LEFT, padx=5)
        
        self.export_button = ttk.Button(button_frame, text="Export & Next Dive", 
                                       command=self.export_and_next)
        self.export_button.pack(side=tk.RIGHT, padx=5)
        
        self.skip_button = ttk.Button(button_frame, text="Skip This Dive", 
                                     command=self.skip_dive)
        self.skip_button.pack(side=tk.RIGHT, padx=5)
        
        self.cancel_button = ttk.Button(button_frame, text="Cancel", 
                                       command=self.cancel)
        self.cancel_button.pack(side=tk.RIGHT, padx=5)
        
        # Results with scrollbar
        results_frame = ttk.LabelFrame(self.root, text="Completed", padding="10")
        results_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        
        # Create frame for listbox and scrollbar
        listbox_frame = ttk.Frame(results_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create listbox
        self.results_listbox = tk.Listbox(listbox_frame)
        self.results_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Create scrollbar
        scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=self.results_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Configure listbox to use scrollbar
        self.results_listbox.config(yscrollcommand=scrollbar.set)

    def load_current_dive(self):
        """Load the current dive"""
        if self.current_index >= len(self.dive_numbers):
            self.complete_process()
            return
        
        current_dive = self.dive_numbers[self.current_index]
        dive_id = f"DIVE{current_dive:03d}"
        
        # Build expected filename for display
        expected_filename = f"{self.processor.file_prefix}{dive_id}{self.processor.file_suffix}.jpg"
        
        # Update GUI
        self.dive_label.config(text=f"Current: {dive_id}")
        self.progress_label.config(text=f"Dive {self.current_index + 1} of {len(self.dive_numbers)}")
        
        progress_pct = int(((self.current_index + 1) / len(self.dive_numbers)) * 100)
        self.status_label.config(
            text=f"Set extent for {dive_id} in ArcPro, then click 'Export & Next' ({progress_pct}% complete)\nWill export as: {expected_filename}", 
            foreground="blue"
        )
        
        # Set layer visibility
        try:
            arcpy.AddMessage(f"\nLoading {dive_id}...")
            self.processor.set_dive_visibility(current_dive)
            arcpy.AddMessage(f"Layers updated for {dive_id}")
            
        except Exception as e:
            arcpy.AddError(f"Error loading {dive_id}: {str(e)}")
            messagebox.showerror("Error", f"Error loading {dive_id}: {str(e)}")
    
    def export_and_next(self):
        """Export current dive and move to next"""
        if self.current_index >= len(self.dive_numbers):
            return
        
        current_dive = self.dive_numbers[self.current_index]
        dive_id = f"DIVE{current_dive:03d}"
        
        try:
            # Export the layout
            self.status_label.config(text=f"Exporting {dive_id}...", foreground="orange")
            self.root.update()
            
            output_file = self.processor.export_current_layout(current_dive)
            
            # Add to completed list
            self.completed_dives.append(current_dive)
            self.results_listbox.insert(tk.END, f"{dive_id} ✓")
            
            # Auto-scroll to show the latest entry
            self.results_listbox.see(tk.END)
            
            arcpy.AddMessage(f"Successfully exported {dive_id}")
            
            # Move to next
            self.current_index += 1
            self.load_current_dive()
            
        except Exception as e:
            arcpy.AddError(f"Error exporting {dive_id}: {str(e)}")
            messagebox.showerror("Error", f"Error exporting {dive_id}: {str(e)}")
            self.status_label.config(text=f"Error exporting {dive_id}", foreground="red")
    
    def skip_dive(self):
        """Skip the current dive"""
        if self.current_index >= len(self.dive_numbers):
            return
        
        current_dive = self.dive_numbers[self.current_index]
        dive_id = f"DIVE{current_dive:03d}"
        
        self.results_listbox.insert(tk.END, f"{dive_id} (skipped)")
        
        # Auto-scroll to show the latest entry
        self.results_listbox.see(tk.END)
        
        arcpy.AddMessage(f"Skipped {dive_id}")
        
        self.current_index += 1
        self.load_current_dive()
    
    def complete_process(self):
        """Complete the process"""
        self.status_label.config(text="All dives completed!", foreground="green")
        self.export_button.config(state=tk.DISABLED)
        self.skip_button.config(state=tk.DISABLED)
        self.cancel_button.config(text="Close")
        
        arcpy.AddMessage(f"\nProcess completed! Exported {len(self.completed_dives)} dives")
        arcpy.AddMessage(f"Output folder: {self.processor.output_folder}")
        
        messagebox.showinfo("Complete", 
                           f"Process completed!\n\n"
                           f"Exported: {len(self.completed_dives)} dives\n"
                           f"Output: {self.processor.output_folder}")
    
    def cancel(self):
        """Cancel the process"""
        if self.completed_dives:
            if not messagebox.askyesno("Cancel", 
                                      f"You have completed {len(self.completed_dives)} dives.\n\n"
                                      "Do you want to cancel the process?"):
                return
        
        arcpy.AddMessage("Process cancelled by user")
        self.root.destroy()

    def apply_scale_units(self):
        """Apply the selected scale bar units"""
        try:
            units = self.scale_var.get()
            self.processor.set_scale_bar_units(units)
            self.status_label.config(text=f"Scale bar units set to {units}", foreground="green")
        except Exception as e:
            arcpy.AddError(f"Error setting scale bar units: {str(e)}")
            messagebox.showerror("Error", f"Error setting scale bar units: {str(e)}")

    def rezoom_to_layers(self):
        """Re-zoom to the current dive's layers"""
        if self.current_index < len(self.dive_numbers):
            current_dive = self.dive_numbers[self.current_index]
            try:
                self.processor.zoom_to_dive_layers(current_dive)
                self.status_label.config(text=f"Re-zoomed to DIVE{current_dive:03d} layers", foreground="blue")
            except Exception as e:
                arcpy.AddError(f"Error re-zooming: {str(e)}")
                messagebox.showerror("Error", f"Error re-zooming: {str(e)}")