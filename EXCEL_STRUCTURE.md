# Sample Timeline Data Structure

## Excel Sheet: "TimelineData"

### Headers (Row 1):
| A | B | C | D | E | F |
|---|---|---|---|---|---|
| Task Name | Start Date | End Date | Type | Color | Swimlane |

### Sample Data (Row 2 onwards):

| Task Name | Start Date | End Date | Type | Color | Swimlane |
|-----------|------------|----------|------|-------|----------|
| Project Kickoff | 2025-01-15 | | Milestone | blue | Planning |
| Requirements Gathering | 2025-01-16 | 2025-02-15 | Phase | green | Planning |
| Stakeholder Review | 2025-02-15 | | Milestone | orange | Planning |
| System Design | 2025-02-16 | 2025-03-15 | Phase | blue | Design |
| Architecture Review | 2025-03-15 | | Milestone | red | Design |
| UI/UX Design | 2025-02-20 | 2025-03-20 | Phase | purple | Design |
| Development Sprint 1 | 2025-03-16 | 2025-04-15 | Phase | blue | Development |
| Development Sprint 2 | 2025-04-16 | 2025-05-15 | Phase | blue | Development |
| Code Review | 2025-05-15 | | Milestone | orange | Development |
| Unit Testing | 2025-04-01 | 2025-05-20 | Phase | red | Testing |
| Integration Testing | 2025-05-16 | 2025-06-01 | Phase | red | Testing |
| User Acceptance Testing | 2025-05-25 | 2025-06-10 | Phase | yellow | Testing |
| Go-Live Preparation | 2025-06-01 | 2025-06-15 | Phase | green | Deployment |
| Production Deployment | 2025-06-15 | | Milestone | green | Deployment |
| Post-Launch Support | 2025-06-16 | 2025-07-15 | Phase | gray | Support |

## Key Points:

1. **Date Format**: Use standard Excel date format (YYYY-MM-DD or MM/DD/YYYY)
2. **End Date**: Leave empty for milestones, required for phases
3. **Type**: Must be exactly "Milestone" or "Phase" (case-insensitive)
4. **Color**: Supported colors: red, blue, green, orange, purple, yellow, gray
5. **Swimlane**: Any text label to group related events

## Tips:

- Events in the same swimlane will appear on the same horizontal track
- Overlapping events within a swimlane will automatically be placed on separate lanes
- Use meaningful swimlane names like "Planning", "Development", "Testing", etc.
- Keep task names concise for better label readability
- Ensure phase end dates are after start dates

## Advanced Features:

- **Custom Colors**: You can extend the color palette by modifying the `GetColor()` function
- **Multiple Projects**: Use different swimlanes for different project streams
- **Resource Allocation**: Use swimlanes to represent different teams or resources
