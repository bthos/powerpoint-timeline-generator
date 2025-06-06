# Enhanced PowerPoint Timeline Generator

## Overview
This VBA script automatically creates professional, multi-lane timelines in PowerPoint that resemble OfficeTimeline.com layouts. It solves the common problem of overlapping events by intelligently placing them on separate lanes.

## 🚀 Quick Start

**New to the Timeline Generator?** Check out our [Quick Start Guide](QUICK_START.md) for a 5-minute walkthrough that gets you creating professional timelines immediately!

## Key Features

### 🛳️ **Multi-Lane Timeline Support with Swimlanes**
- **Swimlane Organization**: Events can be organized into horizontal swimlanes for better categorization
- **Automatic Overlap Detection**: Events that would overlap are automatically moved to separate lanes within each swimlane
- **Smart Lane Assignment**: Optimized algorithm minimizes the number of lanes needed per swimlane
- **Visual Connectors**: Dashed lines connect off-axis events to their respective swimlane axis

### 🎨 **Professional Styling**
- **Swimlane Headers**: Labeled sections for organizing different project tracks or teams
- **Lane Separators**: Optional horizontal grid lines for better organization within swimlanes
- **Color Coding**: Support for red, blue, green, orange, and custom colors
- **Milestone & Phase Support**: Different visual treatments for different event types
- **Smart Spacing**: Buffer zones prevent label overlap

### 📊 **Data Integration**
- **Excel Integration**: Reads data directly from Excel "TimelineData" sheet
- **Flexible Date Handling**: Supports both milestones (single date) and phases (date ranges)
- **Error Handling**: Comprehensive validation and user-friendly error messages

## Excel Data Structure

Create a sheet named "TimelineData" with the following columns:

| Column | Header | Description | Example |
|--------|--------|-------------|---------|
| A | Task Name | Event or phase name | "Project Kickoff" |
| B | Start Date | Start date (required) | 1/15/2025 |
| C | End Date | End date (optional for milestones) | 2/28/2025 |
| D | Type | "Milestone" or "Phase" | "Milestone" |
| E | Color | "red", "blue", "green", "orange" | "blue" |
| F | Swimlane | Swimlane category/track | "Planning" |

## Sample Data

```
Task Name          | Start Date | End Date   | Type      | Color  | Swimlane
Project Kickoff    | 1/15/2025  |            | Milestone | blue   | Planning
Requirements Phase | 1/16/2025  | 2/15/2025  | Phase     | green  | Planning
Design Review      | 2/15/2025  |            | Milestone | orange | Design
Development Phase  | 2/16/2025  | 4/30/2025  | Phase     | blue   | Development
Testing Phase      | 4/15/2025  | 5/15/2025  | Phase     | red    | Testing
Go Live           | 5/16/2025  |            | Milestone | green  | Deployment
```

## How to Use

1. **Prepare Data**: Create an Excel workbook with a "TimelineData" sheet containing your timeline data
2. **Open PowerPoint**: Ensure PowerPoint is running with a presentation open
3. **Run Macro**: Execute the `CreateTimelineFromData()` macro
4. **Review Timeline**: The script will create a new slide with your multi-lane timeline

## Customization Options

### Constants You Can Modify
```vba
Const LaneHeight As Integer = 50          ' Vertical spacing between lanes
Const SwimlaneHeight As Integer = 120     ' Vertical spacing between swimlanes
Const SwimlaneHeaderWidth As Integer = 150 ' Width for swimlane labels
Const CircleSize As Integer = 14          ' Size of milestone markers
Const BarHeight As Integer = 12           ' Height of phase bars
Const TimelineTop As Single = 80          ' Vertical position of timeline
```

### Adding Custom Colors
Extend the `GetColor()` function:
```vba
Case "purple": GetColor = RGB(128, 0, 128)
Case "yellow": GetColor = RGB(255, 255, 0)
```

## Troubleshooting

### Common Issues
- **"Excel is not open"**: Ensure Excel is running with your data file open
- **"Sheet 'TimelineData' not found"**: Verify the sheet name matches exactly
- **"No valid data found"**: Check that your data starts in row 2 (row 1 should contain headers)

### Performance Notes
- Works best with 50 or fewer timeline events
- Very dense timelines may require manual adjustment of `LaneHeight` constant
- Large date ranges may need font size adjustments

## Technical Details

### Swimlane Organization Algorithm
The script uses an intelligent swimlane system that:
1. **Groups Events**: Automatically groups events by their swimlane designation
2. **Creates Headers**: Adds labeled headers for each swimlane with professional styling
3. **Independent Lane Management**: Each swimlane manages its own lanes for overlapping events
4. **Optimal Spacing**: Calculates vertical spacing to accommodate multiple swimlanes and lanes

### Overlap Detection Algorithm
Within each swimlane, the script uses a sophisticated algorithm that:
1. Converts all dates to X-coordinates on the timeline
2. Adds buffer zones around milestones for label space
3. Compares each event with all previous events in the same swimlane
4. Assigns the lowest available lane that avoids conflicts

### Visual Enhancements
- **Connector Lines**: Subtle dashed lines link off-axis events to main timeline
- **Lane Separators**: Light gray horizontal lines separate lanes visually
- **Smart Labels**: Milestone labels include both name and date
- **Professional Colors**: Carefully chosen color palette for business presentations

## Version History
- **v2.0**: Added multi-lane support with automatic overlap detection
- **v1.0**: Basic single-line timeline generation

---
*Created for project managers, analysts, and consultants who need professional timeline visualizations in PowerPoint.*
