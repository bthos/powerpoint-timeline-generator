' Simple test script to verify swimlane splitting functionality
' This script simulates testing the swimlane splitting logic

' Test data that would likely cause swimlane splitting
Dim testData(15, 5) As Variant

' Headers (row 0)
testData(0, 0) = "Task Name"
testData(0, 1) = "Start Date"
testData(0, 2) = "End Date"
testData(0, 3) = "Type"
testData(0, 4) = "Color"
testData(0, 5) = "Swimlane"

' Large swimlane with many events - likely to be split
testData(1, 0) = "Task 1": testData(1, 1) = "2025-01-15": testData(1, 2) = "2025-01-30": testData(1, 3) = "Feature": testData(1, 4) = "blue": testData(1, 5) = "BigSwimlane"
testData(2, 0) = "Task 2": testData(2, 1) = "2025-02-01": testData(2, 2) = "2025-02-15": testData(2, 3) = "Feature": testData(2, 4) = "green": testData(2, 5) = "BigSwimlane"
testData(3, 0) = "Task 3": testData(3, 1) = "2025-02-16": testData(3, 2) = "2025-03-01": testData(3, 3) = "Feature": testData(3, 4) = "red": testData(3, 5) = "BigSwimlane"
testData(4, 0) = "Task 4": testData(4, 1) = "2025-03-02": testData(4, 2) = "2025-03-15": testData(4, 3) = "Feature": testData(4, 4) = "orange": testData(4, 5) = "BigSwimlane"
testData(5, 0) = "Task 5": testData(5, 1) = "2025-03-16": testData(5, 2) = "2025-03-30": testData(5, 3) = "Feature": testData(5, 4) = "blue": testData(5, 5) = "BigSwimlane"
testData(6, 0) = "Task 6": testData(6, 1) = "2025-04-01": testData(6, 2) = "2025-04-15": testData(6, 3) = "Feature": testData(6, 4) = "green": testData(6, 5) = "BigSwimlane"
testData(7, 0) = "Task 7": testData(7, 1) = "2025-04-16": testData(7, 2) = "2025-04-30": testData(7, 3) = "Feature": testData(7, 4) = "red": testData(7, 5) = "BigSwimlane"
testData(8, 0) = "Task 8": testData(8, 1) = "2025-05-01": testData(8, 2) = "2025-05-15": testData(8, 3) = "Feature": testData(8, 4) = "orange": testData(8, 5) = "BigSwimlane"

' Medium swimlane
testData(9, 0) = "Medium 1": testData(9, 1) = "2025-01-20": testData(9, 2) = "2025-02-05": testData(9, 3) = "Feature": testData(9, 4) = "blue": testData(9, 5) = "MediumSwimlane"
testData(10, 0) = "Medium 2": testData(10, 1) = "2025-02-06": testData(10, 2) = "2025-02-20": testData(10, 3) = "Feature": testData(10, 4) = "green": testData(10, 5) = "MediumSwimlane"
testData(11, 0) = "Medium 3": testData(11, 1) = "2025-02-21": testData(11, 2) = "2025-03-10": testData(11, 3) = "Feature": testData(11, 4) = "red": testData(11, 5) = "MediumSwimlane"
testData(12, 0) = "Medium 4": testData(12, 1) = "2025-03-11": testData(12, 2) = "2025-03-25": testData(12, 3) = "Feature": testData(12, 4) = "orange": testData(12, 5) = "MediumSwimlane"

' Small swimlane
testData(13, 0) = "Small 1": testData(13, 1) = "2025-01-25": testData(13, 2) = "2025-02-10": testData(13, 3) = "Feature": testData(13, 4) = "blue": testData(13, 5) = "SmallSwimlane"
testData(14, 0) = "Small 2": testData(14, 1) = "2025-02-11": testData(14, 2) = "2025-02-25": testData(14, 3) = "Feature": testData(14, 4) = "green": testData(14, 5) = "SmallSwimlane"

' Milestones
testData(15, 0) = "Milestone 1": testData(15, 1) = "2025-03-15": testData(15, 2) = "": testData(15, 3) = "Milestone": testData(15, 4) = "red": testData(15, 5) = "BigSwimlane"

WScript.Echo "Test data created with scenarios likely to trigger swimlane splitting:"
WScript.Echo "- BigSwimlane: 8 features + 1 milestone (likely to be split)"
WScript.Echo "- MediumSwimlane: 4 features (might be split depending on overlap)"
WScript.Echo "- SmallSwimlane: 2 features (should fit on one slide)"
WScript.Echo ""
WScript.Echo "To test:"
WScript.Echo "1. Create Excel file with this data in 'TimelineData' sheet"
WScript.Echo "2. Run CreateTimelineFromData() in PowerPoint"
WScript.Echo "3. Check if BigSwimlane gets split across slides"
WScript.Echo "4. Verify continuation markers '(cont.)' appear"