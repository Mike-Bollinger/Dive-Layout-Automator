# Dive Layout Automator - ArcGIS Pro Toolbox

An interactive ArcGIS Pro toolbox for automating the creation of dive maps with manual extent control and scale bar management.

## Overview

This toolbox automates the process of creating standardized maps for multiple dive sites while maintaining user control over map extent and scale bar units. The tool provides an interactive workflow that combines automation with manual fine-tuning capabilities.

## Key Features

- **Automatic Dive Detection**: Detects dive numbers from layer names using flexible patterns
- **Sequential Layer Visibility**: Sets layer visibility for each dive sequentially
- **Auto-Zoom Functionality**: Automatically zooms to dive extent as starting point (15% buffer)
- **Interactive GUI**: User-friendly interface for manual extent adjustment
- **Manual Scale Bar Control**: Support for Meters, Kilometers, Feet, and Miles
- **Customizable File Naming**: Prefix/suffix options for output files
- **Progress Tracking**: Real-time progress updates and error handling
- **Scrollable Completion Log**: Track completed dives during processing

## Workflow

1. User selects layout, layer groups, and output settings in toolbox
2. Tool identifies all dive numbers in selected groups
3. Interactive GUI opens with step-by-step instructions
4. For each dive:
   - Tool makes dive layers visible and hides others
   - Tool auto-zooms to dive extent with 15% buffer
   - User manually adjusts scale bar units if needed
   - User fine-tunes map extent using ArcGIS Pro navigation tools
   - User clicks "Export & Next" to save map and proceed to next dive
5. Process continues until all dives are completed

## Requirements

- **ArcGIS Pro**: Version 2.8 or later
- **Python**: 3.7+ with tkinter support
- **Layout Components**: Layout must contain map frame and scale bar elements
- **Layer Organization**: Dive layers organized in groups with recognizable naming patterns

### Supported Dive Layer Naming Patterns

The tool supports flexible dive layer naming:
- `DIVE001`, `DIVE002`, etc.
- `Dive1`, `Dive2`, etc.
- `D001`, `D002`, etc.
- Just numbers: `001`, `002`, etc.

## Installation

1. Download the `DiveLayoutAutomator.pyt` file
2. Optional: Include `NOAA_Logo.ico` in the same directory for custom window icon
3. Add the toolbox to ArcGIS Pro:
   - Open ArcGIS Pro
   - In the Catalog pane, right-click "Toolboxes"
   - Select "Add Toolbox"
   - Browse to and select `DiveLayoutAutomator.pyt`

## Usage

### Tool Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| **Layout** | Required | Target layout for map export |
| **Layer Groups** | Required (Multi-select) | Groups containing dive layers |
| **Output Folder** | Required | Destination for exported JPG files |
| **File Name Prefix** | Optional | Text before dive name (e.g., "PC2403_") |
| **File Name Suffix** | Optional | Text after dive name (e.g., "_Final") |
| **Export DPI** | Optional | Resolution for exported images (default: 300) |

### Step-by-Step Instructions

1. **Setup**: Open your ArcGIS Pro project with dive layers organized in groups
2. **Run Tool**: Launch "Interactive Dive Layout Generator" from the toolbox
3. **Configure Parameters**: 
   - Select your target layout
   - Choose layer groups containing dive layers
   - Set output folder and naming preferences
4. **Interactive Process**:
   - GUI opens showing current dive and progress
   - Layers automatically become visible for current dive
   - Map auto-zooms to dive extent with buffer
   - Adjust scale bar units using dropdown (Meters, Kilometers, Feet, Miles)
   - Fine-tune map extent using ArcGIS Pro navigation tools
   - Click "Export & Next" when satisfied with extent
5. **Completion**: Process continues until all dives are exported

### GUI Controls

- **Scale Bar Units**: Dropdown to change units (Meters, Kilometers, Feet, Miles)
- **Apply Units**: Button to apply selected scale bar units
- **Re-zoom to Layers**: Re-centers map to current dive's extent
- **Export & Next Dive**: Exports current map and proceeds to next dive
- **Skip This Dive**: Skips current dive without exporting
- **Cancel**: Stops the process (confirms if dives already completed)

## Output

- **File Format**: High-quality JPG images
- **Naming Convention**: `[prefix]DIVE001[suffix].jpg`
- **Resolution**: Configurable DPI (default: 300)
- **Scale Bar**: Consistent units across all maps
- **Processing Log**: Detailed information in ArcGIS Pro Messages panel

## Technical Details

### Architecture
- **ArcPy Integration**: Uses ArcPy mapping module for layout control
- **CIM Implementation**: Cartographic Information Model for scale bar updates
- **Cross-Platform GUI**: Tkinter interface for user interaction
- **Pattern Matching**: Regex-based flexible dive name detection
- **Smart Auto-Zoom**: Targets largest map frame to avoid extent indicators

### Error Handling
- Comprehensive validation of inputs
- Detailed error messages in ArcGIS Pro Messages
- Graceful handling of missing layers or layouts
- User confirmation for process cancellation

## Troubleshooting

### Common Issues

**No dive layers found:**
- Check layer names contain dive numbers (DIVE001, Dive1, D001, etc.)
- Verify selected groups contain correct layers
- Ensure layer names are spelled correctly

**Layout not updating:**
- Verify layout contains proper map frame and scale bar elements
- Check that selected layout is the active layout
- Ensure map frame is properly linked to map containing dive layers

**Scale bar not updating:**
- Confirm layout contains scale bar elements
- Check scale bar is properly configured in layout
- Verify scale bar is linked to correct map frame

**Export failures:**
- Check output folder permissions
- Verify sufficient disk space
- Ensure no invalid characters in file prefix/suffix

## Version History

- **v1.0.0** (June 11, 2025) - Initial release
  - Interactive dive map generation
  - Auto-zoom functionality
  - Manual scale bar control
  - Custom file naming
  - Progress tracking GUI
  - Comprehensive error handling

## Support

For questions, issues, or feature requests:
- Check ArcGIS Pro Messages panel for detailed error information
- Ensure dive layer naming follows supported patterns
- Verify layout contains proper map frame and scale bar elements

## Author

Mike Bollinger  
Version 1.0.0  
Created: June 2025  
Updated: June 11, 2025

## Disclaimer

This repository is a scientific product and is not official communication of the National
Oceanic and Atmospheric Administration, or the United States Department of Commerce. All NOAA
GitHub project code is provided on an 'as is' basis and the user assumes responsibility for its
use. Any claims against the Department of Commerce or Department of Commerce bureaus stemming from
the use of this GitHub project will be governed by all applicable Federal law. Any reference to
specific commercial products, processes, or services by service mark, trademark, manufacturer, or
otherwise, does not constitute or imply their endorsement, recommendation or favoring by the
Department of Commerce. The Department of Commerce seal and logo, or the seal and logo of a DOC
bureau, shall not be used in any manner to imply endorsement of any commercial product or activity
by DOC or the United States Government.

## License

Software code created by U.S. Government employees is not subject to copyright in the United States (17 U.S.C. §105). The United States/Department of Commerce reserve all rights to seek and obtain copyright protection in countries other than the United States for Software authored in its entirety by the Department of Commerce. To this end, the Department of Commerce hereby grants to Recipient a royalty-free, nonexclusive license to use, copy, and create derivative works of the Software outside of the United States.
