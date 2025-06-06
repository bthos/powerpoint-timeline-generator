# Quick Start Guide - Timeline Generator

## 🚀 Getting Started in 5 Minutes

### Step 1: Prepare Your Excel Data
1. Open Excel and create a new workbook
2. Rename Sheet1 to "TimelineData"
3. Add headers in row 1:
   - **A1**: Task Name
   - **B1**: Start Date
   - **C1**: End Date
   - **D1**: Type
   - **E1**: Color
   - **F1**: Swimlane

### Step 2: Add Your Timeline Events
Starting from row 2, add your events:
```
Project Kickoff    | 1/15/2025 |           | Milestone | blue    | Planning
Requirements       | 1/16/2025 | 2/15/2025 | Phase     | green   | Planning
Design Phase       | 2/16/2025 | 3/15/2025 | Phase     | blue    | Design
Development        | 3/16/2025 | 5/15/2025 | Phase     | orange  | Development
Testing           | 5/1/2025  | 6/1/2025  | Phase     | red     | Testing
Go Live           | 6/15/2025 |           | Milestone | green   | Deployment
```

### Step 3: Run the Timeline Generator
1. Open PowerPoint (any presentation)
2. Press `Alt + F11` to open the VBA Editor
3. Insert → Module
4. Copy and paste the entire timeline.bas code (or File → Import File)
5. Press `F5` or Run → Run Sub to execute `CreateTimelineFromData`
6. View → Immediate Window (or `Ctrl+G`) to see outout and error messages

### Step 4: Review Your Timeline
- A new slide will be created with your timeline
- Each swimlane appears as a separate horizontal track
- Overlapping events are automatically placed on different lanes
- Professional styling with colors, connectors, and labels

## 💡 Pro Tips

### Data Entry Tips:
- **Dates**: Use any standard Excel date format (1/15/2025, 15-Jan-2025, etc.)
- **Milestones**: Leave End Date empty for milestone events
- **Phases**: Always provide both Start Date and End Date
- **Colors**: Use predefined colors for consistency: red, blue, green, orange, purple, yellow, gray

### Swimlane Best Practices:
- **Logical Grouping**: Group related activities (Planning, Development, Testing)
- **Team-Based**: Use swimlanes for different teams or departments
- **Project Phases**: Separate major project phases into different swimlanes
- **Resource Types**: Different swimlanes for different types of resources

### Visual Optimization:
- **Keep Names Short**: Task names should be concise for better readability
- **Color Consistency**: Use the same color for related event types
- **Balanced Distribution**: Try to distribute events across swimlanes evenly

## 🔧 Customization Options

### Modify Timeline Appearance:
Edit these constants in the VBA code:
```vba
Const SwimlaneHeight As Integer = 120     ' Space between swimlanes
Const LaneHeight As Integer = 50          ' Space between lanes
Const TimelineTop As Single = 80          ' Distance from top
```

### Add Custom Colors:
Extend the `GetColor()` function:
```vba
Case "teal": GetColor = RGB(0, 128, 128)
Case "maroon": GetColor = RGB(128, 0, 0)
```

## 🆘 Troubleshooting

| Problem | Solution |
|---------|----------|
| "Excel is not open" | Make sure Excel is running with your data file open |
| "Sheet 'TimelineData' not found" | Verify sheet name is exactly "TimelineData" |
| "No valid data found" | Check that data starts in row 2 with headers in row 1 |
| Events overlapping | Use more swimlanes or check date formats |
| Timeline too cramped | Increase SwimlaneHeight or LaneHeight constants |

## 📈 Advanced Features

### Multiple Project Tracking:
Use different swimlanes for different projects or workstreams running in parallel.

### Resource Management:
Assign swimlanes to different teams, departments, or resource types.

### Milestone Dependencies:
Use color coding to show relationships between milestones across swimlanes.

### Progress Tracking:
Use different colors to indicate completion status (green=complete, yellow=in progress, red=delayed).

---
*Need help? Check the main README.md for detailed documentation or the EXCEL_STRUCTURE.md for data format examples.*
