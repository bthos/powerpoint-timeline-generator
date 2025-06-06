' ===================================================================
' ENHANCED TIMELINE GENERATOR WITH HIERARCHICAL MULTI-LANE SUPPORT
' ===================================================================
'
' Professional PowerPoint timeline generator with automatic overlap detection
' and hierarchical event organization (Phase > Feature > Milestone).
'
' FEATURES:
' ? Three-level hierarchy: Phase (collections) > Feature (tasks) > Milestone (points)
' ? Multi-lane timeline: Automatic overlap detection and lane assignment
' ? Visual enhancements: Semi-transparent phases, connector lines, smart spacing
' ? Professional styling: Color-coded swimlanes, enhanced labels, today markers
' ? Smart color detection: Auto-assigns colors based on task names and types
' ? Timestamped debug messages for development tracking
'
' DATA STRUCTURE (Excel "TimelineData" sheet):
' Column A: Task Name
' Column B: Start Date
' Column C: End Date (required for Features and Phases)
' Column D: Type ("Milestone", "Feature", or "Phase")
' Column E: Color (optional - auto-detected if empty)
' Column F: Swimlane (Features and Milestone grouping based on swimlane name)
'
' VISUAL HIERARCHY:
' - Phase: Semi-transparent bars with dashed borders (top level)
' - Feature: Solid rounded bars (mid level)
' - Milestone: Diamond shapes with labels (point events)
'
' DEBUG MESSAGES:
' The generator outputs timestamped completion messages to the Immediate Window:
' - "dd-mmm-yyyy hh:mm:ss> Timeline generation completed successfully - Single slide created with N swimlanes"
' - "dd-mmm-yyyy hh:mm:ss> Timeline generation completed successfully - N slides created with N swimlanes distributed across slides"
'

' ===================================================================
' USER-DEFINED TYPES (MUST BE AT MODULE LEVEL)
' ===================================================================
Type TimelineConfig
    slideWidth As Single
    slideHeight As Single
    timelineTop As Single
    TimelineAxisTop As Single
    axisY As Single
    axisPadding As Integer
    circleSize As Integer
    elementHeight As Integer
    laneHeight As Integer
    swimlaneHeight As Integer
    SwimlaneHeaderWidth As Integer
    fontName As String
    fontSize As Integer
    ' === DYNAMIC LABEL WIDTH CONSTRAINTS ===
    featureNameLabelMinWidth As Single      ' Minimum width for feature name labels
    featureNameLabelMaxWidth As Single      ' Maximum width for feature name labels
    featureDurationLabelMinWidth As Single  ' Minimum width for feature duration labels
    featureDurationLabelMaxWidth As Single  ' Maximum width for feature duration labels
    featureDateRangeLabelMinWidth As Single ' Minimum width for feature date range labels
    featureDateRangeLabelMaxWidth As Single ' Maximum width for feature date range labels
    milestoneLabelMinWidth As Single        ' Minimum width for milestone labels
    milestoneLabelMaxWidth As Single        ' Maximum width for milestone labels
    phaseLabelMinWidth As Single            ' Minimum width for phase labels
    phaseLabelMaxWidth As Single            ' Maximum width for phase labels
    ' === LABEL POSITIONING CONSTANTS ===
    labelVerticalOffset As Integer          ' Vertical offset for centering labels to shapes (-8)
    labelHeight As Integer                  ' Standard height for event labels (16px)
    labelSpacingToShape As Integer          ' Standard spacing from shape edge to label (23px)
    labelInternalPadding As Single          ' Padding for labels inside bars (20px, 10px each side)
    ' === SPACING AND LAYOUT CONSTRAINTS ===
    swimlanePadding As Single               ' Padding between swimlanes (5px)
    swimlaneContentPadding As Single        ' Buffer from swimlane top to first element (10px)
    laneSpacingWithTopLabels As Single      ' Lane spacing when labels are on top (35px)
    laneSpacingWithInsideLabels As Single   ' Lane spacing when labels are inside (20px)
    elementBottomPadding As Single          ' Padding below bottom-most element (5px)    
    milestoneExtraSpace As Single           ' Extra space below milestones for date labels (15px)
    ' === MINIMUM DIMENSIONS AND CONSTRAINTS ===
    minimumBarWidth As Single               ' Minimum width for bars with invalid dates (10px)
    minimumShapeDimension As Single         ' Minimum dimension for shapes to prevent errors (1px)
    maximumSlidePosition As Single          ' Maximum slide position for validation (2000px)
    bottomMarginForSlides As Single         ' Bottom margin for multi-slide calculations (30px)
    ' === SLIDE LAYOUT CONFIGURATION ===
    slideLayoutName As String               ' User-configurable slide layout name (e.g., "Timeline Layout")
End Type

Type TimelineDateRange
    minDate As Date
    maxDate As Date
    scaleFactor As Double
End Type

Type SwimlaneOrganization
    swimlanes() As String
    swimlaneEvents() As Variant
    Count As Integer
End Type

' ===================================================================
' MAIN TIMELINE GENERATOR ENTRY POINT
' ===================================================================
Sub CreateTimelineFromData()
    ' Initialize configuration
    Dim config As TimelineConfig
    config = GetDefaultTimelineConfig()
    
    ' Load and validate data
    Dim timelineData() As Variant
    timelineData = LoadAndValidateData("TimelineData")
    If IsEmpty(timelineData) Then Exit Sub
    
    ' Calculate timeline bounds
    Dim dateRange As TimelineDateRange
    dateRange = CalculateTimelineDateRange(timelineData)
    
    ' Organize data by swimlanes
    Dim swimlaneOrg As SwimlaneOrganization
    swimlaneOrg = OrganizeTimelineData(timelineData)
    
    ' Check if multi-slide generation is needed
    Dim requiredSlides As Integer
    requiredSlides = CalculateRequiredSlides(swimlaneOrg, config)
    
    If requiredSlides = 1 Then
        ' Single slide - use existing logic
        Dim sld As Slide: sld = CreateTimelineSlide()
        Call RenderTimeline(sld, config, dateRange, swimlaneOrg, timelineData)
        Debug.Print Format(Now, "dd-mmm-yyyy hh:mm:ss") & "> Timeline generation completed successfully - Single slide created with " & swimlaneOrg.Count & " swimlanes"
    Else
        ' Multi-slide generation - debug message handled within CreateMultiSlideTimeline
        Call CreateMultiSlideTimeline(config, dateRange, swimlaneOrg, timelineData, requiredSlides)
    End If
    
    ' Title and subtitle removed to optimize space allocation
    ' More room now available for timeline content
End Sub

' ===================================================================
' CONFIGURATION FUNCTIONS
' ===================================================================
Function GetDefaultTimelineConfig() As TimelineConfig
    ' Professional timeline configuration with optimized spacing and white space utilization
    ' Enhanced for proper vertical spacing when feature bar labels are positioned on top
    With GetDefaultTimelineConfig
        ' === SLIDE LAYOUT CONFIGURATION ===
        .slideLayoutName = "Blank"          ' User-configurable slide layout name (empty = use ppLayoutBlank fallback)
        .fontName = "Calibri"               ' Professional font
        .fontSize = 9                       ' Standard font size for labels
        .slideWidth = 960                   ' 16:9 aspect ratio
        .slideHeight = 540                  '
        .timelineTop = 110                  ' Moved up to optimize space after phase area reduction
        .TimelineAxisTop = 50               ' Calendar header area (50-70px)
        .axisY = .timelineTop + 5           ' Swimlanes start at 115px with 5px buffer
        .axisPadding = 40                   ' Padding for timeline space
        .circleSize = 16                    ' Increased milestone size for better visibility
        .elementHeight = 16                 ' Slightly increased element height for better visibility
        .laneHeight = 48                    ' Increased lane spacing to accommodate top labels with proper gaps
        .swimlaneHeight = 85                ' Slightly increased swimlane spacing for more content
        .SwimlaneHeaderWidth = 100          ' Header width for swimlane labels
        
        ' === CONFIGURABLE DYNAMIC LABEL WIDTH CONSTRAINTS ===
        ' Feature name labels (inside bars or on top)
        .featureNameLabelMinWidth = 30   ' Minimum width for feature name labels
        .featureNameLabelMaxWidth = 300  ' Maximum width for feature name labels
        
        ' Feature duration labels (left side of bars: "N d")
        .featureDurationLabelMinWidth = 25   ' Minimum width for duration labels
        .featureDurationLabelMaxWidth = 50   ' Maximum width for duration labels
        
        ' Feature date range labels (right side of bars: "dd mmm - dd mmm")
        .featureDateRangeLabelMinWidth = 80  ' Minimum width for date range labels
        .featureDateRangeLabelMaxWidth = 150 ' Maximum width for date range labels
        
        ' Milestone labels (positioned intelligently left/right with DYNAMIC WIDTH based on text length)
        .milestoneLabelMinWidth = 30     ' Minimum width for milestone labels (dynamic sizing)
        .milestoneLabelMaxWidth = 300    ' Maximum width for milestone labels (dynamic sizing)
        
        ' Phase labels (two-line labels inside phase bars)
        .phaseLabelMinWidth = 80         ' Minimum width for phase labels
        .phaseLabelMaxWidth = 200        ' Maximum width for phase labels
        
        ' === LABEL POSITIONING CONSTANTS ===
        .labelVerticalOffset = -8        ' Vertical offset for centering labels to shapes
        .labelHeight = 16                ' Standard height for event labels
        .labelSpacingToShape = 23        ' Standard spacing from shape edge to label
        .labelInternalPadding = 1        ' Padding for labels inside bars (10px on each side)
        
        ' === SPACING AND LAYOUT CONSTRAINTS ===
        .swimlanePadding = 5             ' Padding between swimlanes
        .swimlaneContentPadding = 2     ' Buffer from swimlane top to first element
        .laneSpacingWithTopLabels = 35   ' Lane spacing when labels are on top
        .laneSpacingWithInsideLabels = 20 ' Lane spacing when labels are inside
        .elementBottomPadding = 2        ' Padding below bottom-most element
        .milestoneExtraSpace = 15        ' Extra space below milestones for date labels
        
        ' === MINIMUM DIMENSIONS AND CONSTRAINTS ===
        .minimumBarWidth = 10            ' Minimum width for bars with invalid dates
        .minimumShapeDimension = 1       ' Minimum dimension for shapes to prevent errors
        .maximumSlidePosition = 2000     ' Maximum slide position for validation
        .bottomMarginForSlides = 400      ' Bottom margin for multi-slide calculations
        
    End With
End Function

Function LoadAndValidateData(sheetName As String) As Variant
    ' Load data from Excel and validate structure
    Dim data() As Variant
    data = ReadDataFromExcel(sheetName)
    
    If IsEmpty(data) Then
        MsgBox "No valid data found in Excel sheet '" & sheetName & "'", vbExclamation
        Exit Function
    End If
    
    If Not ValidateTimelineData(data) Then
        Exit Function
    End If
    
    LoadAndValidateData = data
End Function

Function CreateTimelineSlide() As Slide
    ' Create new PowerPoint slide with user-configurable layout and fallback
    Dim config As TimelineConfig: config = GetDefaultTimelineConfig()
    
    ' Create slide with default layout first
    Set CreateTimelineSlide = ActivePresentation.Slides.Add( _
        ActivePresentation.Slides.Count + 1, ppLayoutBlank)
    
    ' Apply user-configured layout if specified
    If Trim(config.slideLayoutName) <> "" Then
        Call ApplyCustomSlideLayout(CreateTimelineSlide, config.slideLayoutName)
    Else
        ' Use default fallback layout
        On Error Resume Next
        CreateTimelineSlide.customLayout = ActivePresentation.SlideMaster.CustomLayouts(1)
        On Error GoTo 0
    End If
End Function

Function CalculateTimelineDateRange(data() As Variant) As TimelineDateRange
    ' Calculate minimum and maximum dates from timeline data
    If Not IsDate(data(0, 1)) Then
        MsgBox "First row Start Date is missing or invalid.", vbCritical
        Exit Function
    End If
    
    Dim minDate As Date, maxDate As Date
    minDate = Int(data(0, 1))
    maxDate = Int(data(0, 1))
    
    Dim i As Integer
    For i = 0 To UBound(data, 1)
        If IsDate(data(i, 1)) Then
            Dim startDate As Date: startDate = Int(data(i, 1))
            If startDate < minDate Then minDate = startDate
            If startDate > maxDate Then maxDate = startDate
            
            If IsDate(data(i, 2)) Then
                Dim endDate As Date: endDate = Int(data(i, 2))
                If endDate > maxDate Then maxDate = endDate
            End If
        End If
    Next i
    
    With CalculateTimelineDateRange
        .minDate = minDate
        .maxDate = maxDate
        .scaleFactor = 0 ' Will be calculated in RenderTimeline
    End With
End Function

Function OrganizeTimelineData(data() As Variant) As SwimlaneOrganization
    ' Organize timeline data into swimlanes
    Dim result As SwimlaneOrganization
    result.Count = OrganizeEventsBySwimlanes(data, result.swimlanes, result.swimlaneEvents)
    OrganizeTimelineData = result
End Function

' ===================================================================
' MAIN RENDERING ENGINE
' ===================================================================
Sub RenderTimeline(sld As Slide, config As TimelineConfig, ByRef dateRange As TimelineDateRange, _
                   swimlaneOrg As SwimlaneOrganization, data() As Variant)
    ' Main rendering pipeline for timeline visualization
    
    ' Calculate scale factor
    dateRange.scaleFactor = (config.slideWidth - config.SwimlaneHeaderWidth - config.axisPadding) / _
                           (dateRange.maxDate - dateRange.minDate)
    
    ' Store values in local variables to avoid ByRef issues with UDT members
    Dim minDate As Date: minDate = dateRange.minDate
    Dim maxDate As Date: maxDate = dateRange.maxDate
    Dim scaleFactor As Double: scaleFactor = dateRange.scaleFactor
    Dim headerWidth As Single: headerWidth = config.SwimlaneHeaderWidth
    Dim timelineTop As Single: timelineTop = config.TimelineAxisTop
    Dim fontName As String: fontName = config.fontName
    Dim slideHeight As Single: slideHeight = config.slideHeight
    
    ' Render swimlane structure
    Call RenderSwimlanes(sld, config, swimlaneOrg)
    
    ' Render top timeline axis with enhanced features
    Call DrawEnhancedTopTimelineAxis(sld, minDate, maxDate, _
        scaleFactor, headerWidth, timelineTop, fontName)
    
    ' Add vertical grid lines for better visual reference - REMOVED per user request
    ' Call AddVerticalGridLines(sld, minDate, maxDate, _
    '     scaleFactor, headerWidth, timelineTop, slideHeight)
    
    ' Render events in each swimlane
    Call RenderSwimlaneEvents(sld, config, dateRange, swimlaneOrg)
    
    ' Render phases in their dedicated area (separate from swimlanes)
    Call RenderPhasesInDedicatedArea(sld, config, dateRange, data)
End Sub

Sub RenderSwimlanes(sld As Slide, config As TimelineConfig, swimlaneOrg As SwimlaneOrganization)
    ' Render swimlane headers, backgrounds, and axes with dynamic heights
    
    ' Extract config values to avoid ByRef issues
    Dim axisY As Single: axisY = config.axisY
    Dim baseSwimlaneHeight As Integer: baseSwimlaneHeight = config.swimlaneHeight
    Dim fontName As String: fontName = config.fontName
    Dim headerWidth As Single: headerWidth = config.SwimlaneHeaderWidth
    Dim slideWidth As Single: slideWidth = config.slideWidth
    Dim axisPadding As Integer: axisPadding = config.axisPadding
    Dim laneHeight As Integer: laneHeight = config.laneHeight
    
    ' Calculate dynamic positions for each swimlane
    Dim currentY As Single: currentY = axisY
    Dim swimlanePaddingValue As Single: swimlanePaddingValue = config.swimlanePadding ' User-configurable padding between swimlanes
    
    Dim i As Integer
    For i = 0 To swimlaneOrg.Count - 1
        ' Calculate required lanes for this swimlane
        Dim requiredLanes As Integer: requiredLanes = 1 ' Default minimum
        If Not IsEmpty(swimlaneOrg.swimlaneEvents(i)) Then
            Dim tempEvents() As Variant: tempEvents = swimlaneOrg.swimlaneEvents(i)
            Dim tempEventLanes() As Integer
            ReDim tempEventLanes(0 To UBound(tempEvents))
            requiredLanes = CalculateSwimlaneRequiredLanes(tempEvents, tempEventLanes, config)
        End If
        
        ' Calculate actual height based on content (NEW SYSTEM)
        Dim dynamicSwimlaneHeight As Single
        If Not IsEmpty(swimlaneOrg.swimlaneEvents(i)) Then
            Dim tempEvents() As Variant: tempEvents = swimlaneOrg.swimlaneEvents(i)
            Dim tempEventLanes() As Integer
            ReDim tempEventLanes(0 To UBound(tempEvents))
            
            ' Extract values to avoid ByRef issues with UDT members
            Dim placeholderScaleFactor As Double: placeholderScaleFactor = 1
            Dim placeholderMinDate As Date: placeholderMinDate = Date
            
            ' Use the new content-based height calculation with proper parameters
            dynamicSwimlaneHeight = CalculateSwimlaneActualHeight(tempEvents, tempEventLanes, config, _
                placeholderScaleFactor, headerWidth, placeholderMinDate)
        Else
            dynamicSwimlaneHeight = 0 ' Empty swimlanes collapse to 0 height
        End If
        
        ' No minimum height constraints - swimlanes collapse to actual content size
        
        ' Enhanced swimlane header with matching height and vertical centering
        'Call AddEnhancedSwimlaneHeader(sld, 10, currentY - 2, _
        '    swimlaneOrg.swimlanes(i), fontName, 11, dynamicSwimlaneHeight)
        
        ' Dynamic background size based on actual content - EXTENDED BY 25PX LEFT AND RIGHT
        Call DrawSwimlaneBackground(sld, headerWidth - 25, currentY, _
            slideWidth - headerWidth - axisPadding + 50, _
            dynamicSwimlaneHeight, i)
        
        ' Move to next swimlane position with padding
        currentY = currentY + dynamicSwimlaneHeight + swimlanePaddingValue
    Next i
End Sub

Sub RenderSwimlaneEvents(sld As Slide, config As TimelineConfig, dateRange As TimelineDateRange, _
                        swimlaneOrg As SwimlaneOrganization)
    ' Render events within each swimlane with dynamic positioning based on actual heights
    
    ' Extract values to avoid ByRef issues with UDT members
    Dim scaleFactor As Double: scaleFactor = dateRange.scaleFactor
    Dim minDate As Date: minDate = dateRange.minDate
    Dim headerWidth As Single: headerWidth = config.SwimlaneHeaderWidth
    Dim fontName As String: fontName = config.fontName
    Dim circleSize As Integer: circleSize = config.circleSize
    Dim elementHeight As Integer: elementHeight = config.elementHeight
    Dim laneHeight As Integer: laneHeight = config.laneHeight
    
    ' Extract additional config values to avoid ByRef issues
    Dim axisY As Single: axisY = config.axisY
    Dim baseSwimlaneHeight As Integer: baseSwimlaneHeight = config.swimlaneHeight
    
    ' Calculate dynamic positions for each swimlane (same logic as RenderSwimlanes)
    Dim currentY As Single: currentY = axisY
    Dim swimlanePaddingValue As Single: swimlanePaddingValue = config.swimlanePadding ' User-configurable padding between swimlanes - same as RenderSwimlanes
    
    Dim i As Integer
    For i = 0 To swimlaneOrg.Count - 1
        Dim currentEvents() As Variant: currentEvents = swimlaneOrg.swimlaneEvents(i)
        
        If Not IsEmpty(currentEvents) Then
            ' Detect overlapping events and assign lanes
            Dim eventLanes() As Integer
            ReDim eventLanes(0 To UBound(currentEvents))
            Dim totalLanes As Integer
            totalLanes = AssignLanesToEvents(currentEvents, eventLanes, _
                scaleFactor, headerWidth, minDate)
            
            ' Draw lane separator lines if multiple lanes are needed
            If totalLanes > 1 Then
                Call DrawLaneSeparators(sld, config, currentY, totalLanes)
            End If
            
            ' Place events with enhanced styling using dynamic Y position
            Call PlaceEventsInSwimlane(sld, currentEvents, eventLanes, currentY, _
                scaleFactor, headerWidth, minDate, _
                fontName, circleSize, elementHeight, laneHeight)
        End If
        
        ' Calculate actual height based on content (NEW SYSTEM - matching RenderSwimlanes)
        Dim dynamicSwimlaneHeight As Single
        If Not IsEmpty(currentEvents) Then
            ' Use the new content-based height calculation with proper parameters
            dynamicSwimlaneHeight = CalculateSwimlaneActualHeight(currentEvents, eventLanes, config, _
                scaleFactor, headerWidth, minDate) ' Use actual parameters available in this function
        Else
            dynamicSwimlaneHeight = 0 ' Empty swimlanes collapse to 0 height
        End If
        
        ' No minimum height constraints - swimlanes collapse to actual content size
        
        ' Move to next swimlane position with padding - same as RenderSwimlanes
        currentY = currentY + dynamicSwimlaneHeight + swimlanePaddingValue
    Next i
End Sub

Sub DrawLaneSeparators(sld As Slide, config As TimelineConfig, swimlaneY As Single, totalLanes As Integer)
    ' Draw subtle separator lines between lanes within a swimlane - REMOVED per user request
    
    ' Lane separators disabled to clean up timeline appearance
    ' No lines will be drawn in the bar area
    Exit Sub
End Sub

Sub RenderPhasesInDedicatedArea(sld As Slide, config As TimelineConfig, dateRange As TimelineDateRange, data() As Variant)
    ' Render all phases in the optimized dedicated area between calendar header and swimlanes
    
    ' Extract values to avoid ByRef issues with UDT members
    Dim scaleFactor As Double: scaleFactor = dateRange.scaleFactor
    Dim minDate As Date: minDate = dateRange.minDate
    Dim headerWidth As Single: headerWidth = config.SwimlaneHeaderWidth
    Dim fontName As String: fontName = config.fontName
    Dim elementHeight As Integer: elementHeight = config.elementHeight
    
    ' Process all events to find and render phases
    Dim i As Integer
    For i = 0 To UBound(data)
        Dim eventType As String: eventType = UCase(CStr(data(i, 3)))
        
        ' Only process Phase events
        If eventType = "PHASE" Then
            Dim label As String: label = data(i, 0)
            Dim startDate As Date: startDate = Int(data(i, 1))
            Dim endDate As Date
            
            ' Phases must have end dates (validated earlier)
            If IsDate(data(i, 2)) Then
                endDate = Int(data(i, 2))
            Else
                ' Skip invalid phases
                GoTo NextPhase
            End If
            
            Dim colorName As String: colorName = data(i, 4)
            Dim xPos As Single: xPos = headerWidth + (startDate - minDate) * scaleFactor
            Dim phaseEndX As Single: phaseEndX = headerWidth + (endDate - minDate) * scaleFactor
            Dim phaseelementHeight As Single: phaseelementHeight = CSng(elementHeight + 8) ' Slightly larger for two-line labels
            
            ' Validate date order and calculate proper width for phases
            Dim phaseWidth As Single: phaseWidth = phaseEndX - xPos
            If phaseWidth <= 0 Then
                phaseWidth = 10 ' Minimum width for invalid dates
                phaseEndX = xPos + phaseWidth
            End If
            
            ' Get smart color with transparency for phases
            Dim phaseColor As Long: phaseColor = GetColorFromTaskName(label, colorName)
            
            ' === PHASES DIRECTLY BELOW CALENDAR HEADER: Ultra-minimal padding for maximum space utilization ===
            ' Calendar header ends at Y=45 (topY - 5, where topY=50)
            ' Apply ultra-small 1.5px padding between calendar header and phases
            Dim phaseAreaTop As Single: phaseAreaTop = 30   ' Calendar header bottom + 1.5px ultra-minimal padding
            Dim phaseAreaBottom As Single: phaseAreaBottom = 105  ' Before swimlanes (110px) - 5px padding
            Dim phaseYPos As Single: phaseYPos = phaseAreaTop + ((phaseAreaBottom - phaseAreaTop) / 2) ' Center in phase area
            
            ' Draw the phase bar with enhanced styling (semi-transparent overlay)
            Call DrawPhaseBar(sld, xPos, phaseYPos - phaseelementHeight / 2, phaseWidth, phaseelementHeight, phaseColor)
            
            ' === Two-line labels inside phase bars: Main label + Duration on separate lines ===
            Dim phaseCenterX As Single: phaseCenterX = xPos + (phaseWidth / 2)
            Dim phaseDuration As Long: phaseDuration = endDate - startDate
            Dim phaseDurationText As String: phaseDurationText = ""
            If phaseDuration > 0 Then
                phaseDurationText = phaseDuration & " days"
            End If
            
            ' Add two-line phase labels (main label + duration) vertically centered in block
            Call AddTwoLinePhaseLabels(sld, phaseCenterX, phaseYPos, label, phaseDurationText, fontName)
        End If
        
NextPhase:
    Next i
End Sub

' === Swimlane Organization Functions ===

Function OrganizeEventsBySwimlanes(milestoneData() As Variant, ByRef swimlanes() As String, ByRef swimlaneEvents() As Variant) As Integer
    ' Organize events by swimlane and return the number of unique swimlanes
    
    Dim i As Integer, j As Integer
    Dim uniqueSwimlanes As String, swimlaneName As String
    Dim swimlaneCount As Integer: swimlaneCount = 0
    
    ' Find unique swimlanes (only for Features and Milestones)
    For i = 0 To UBound(milestoneData)
        Dim eventType As String: eventType = UCase(CStr(milestoneData(i, 3)))
        
        ' Phase swimlane validation: Phases should not have swimlanes
        If eventType = "PHASE" Then
            swimlaneName = CStr(milestoneData(i, 5))
            If Trim(swimlaneName) <> "" And LCase(Trim(swimlaneName)) <> "default" Then
                Debug.Print "WARNING: Phase '" & CStr(milestoneData(i, 0)) & "' has swimlane '" & swimlaneName & "' - ignoring swimlane (Phases are displayed in dedicated area)"
            End If
            ' Skip phases when organizing swimlanes - they go to dedicated phase area
            GoTo NextEvent
        End If
        
        ' Only process Features and Milestones for swimlanes
        If eventType = "FEATURE" Or eventType = "MILESTONE" Then
            swimlaneName = CStr(milestoneData(i, 5)) ' Swimlane is in column F (index 5)
            If InStr(uniqueSwimlanes, swimlaneName & "|") = 0 Then
                uniqueSwimlanes = uniqueSwimlanes & swimlaneName & "|"
                swimlaneCount = swimlaneCount + 1
            End If
        End If
        
NextEvent:
    Next i
    
    ' Create swimlanes array
    ReDim swimlanes(0 To swimlaneCount - 1)
    ReDim swimlaneEvents(0 To swimlaneCount - 1)
    
    Dim parts() As String
    parts = Split(Left(uniqueSwimlanes, Len(uniqueSwimlanes) - 1), "|")
    For i = 0 To UBound(parts)
        swimlanes(i) = parts(i)
    Next i
    
    ' Group events by swimlane (only Features and Milestones)
    Dim validSwimlaneCount As Integer: validSwimlaneCount = 0
    ReDim validSwimlanes(0 To swimlaneCount - 1) As String
    ReDim validSwimlaneEvents(0 To swimlaneCount - 1) As Variant
    
    For i = 0 To swimlaneCount - 1
        Dim eventsInSwimlane() As Variant
        Dim eventCount As Integer: eventCount = 0
        
        ' Count ONLY Features and Milestones in this swimlane
        For j = 0 To UBound(milestoneData)
            Dim currentEventType As String: currentEventType = UCase(CStr(milestoneData(j, 3)))
            If CStr(milestoneData(j, 5)) = swimlanes(i) And (currentEventType = "FEATURE" Or currentEventType = "MILESTONE") Then
                eventCount = eventCount + 1
            End If
        Next j
        
        If eventCount > 0 Then
            ' Valid swimlane with Features/Milestones
            ReDim eventsInSwimlane(0 To eventCount - 1, 0 To 5)
            Dim eventIndex As Integer: eventIndex = 0
            
            ' Copy only Features and Milestones to swimlane array
            For j = 0 To UBound(milestoneData)
                Dim copyEventType As String: copyEventType = UCase(CStr(milestoneData(j, 3)))
                If CStr(milestoneData(j, 5)) = swimlanes(i) And (copyEventType = "FEATURE" Or copyEventType = "MILESTONE") Then
                    Dim k As Integer
                    For k = 0 To 5
                        eventsInSwimlane(eventIndex, k) = milestoneData(j, k)
                    Next k
                    eventIndex = eventIndex + 1
                End If
            Next j
            
            ' Add to valid swimlanes
            validSwimlanes(validSwimlaneCount) = swimlanes(i)
            validSwimlaneEvents(validSwimlaneCount) = eventsInSwimlane
            validSwimlaneCount = validSwimlaneCount + 1
        Else
            ' Empty swimlane - show warning
            Debug.Print "WARNING: Swimlane '" & swimlanes(i) & "' contains no Features or Milestones - skipping swimlane"
        End If
    Next i
    
    ' Update arrays to only include valid swimlanes
    If validSwimlaneCount > 0 Then
        ReDim swimlanes(0 To validSwimlaneCount - 1)
        ReDim swimlaneEvents(0 To validSwimlaneCount - 1)
        
        For i = 0 To validSwimlaneCount - 1
            swimlanes(i) = validSwimlanes(i)
            swimlaneEvents(i) = validSwimlaneEvents(i)
        Next i
    Else
        ' No valid swimlanes found
        ReDim swimlanes(0 To 0)
        ReDim swimlaneEvents(0 To 0)
        Debug.Print "WARNING: No valid swimlanes found with Features or Milestones"
    End If
    
    OrganizeEventsBySwimlanes = validSwimlaneCount
End Function

Sub PlaceEventsInSwimlane(sld As Slide, events() As Variant, eventLanes() As Integer, swimlaneY As Single, _
                         scaleFactor As Double, leftPadding As Single, minDate As Date, _
                         fontName As String, circleSize As Integer, elementHeight As Integer, laneHeight As Integer)
    ' Place all events within a specific swimlane with enhanced styling
    ' Ensures all events stay within their designated swimlane boundaries
    ' Uses dynamic lane spacing based on whether labels are positioned on top
    
    ' Get configuration values for consistent behavior
    Dim config As TimelineConfig: config = GetDefaultTimelineConfig()
    
    ' Find maximum lane number to size the array properly
    Dim maxLane As Integer: maxLane = 0
    Dim i As Integer
    For i = 0 To UBound(eventLanes)
        If eventLanes(i) > maxLane Then maxLane = eventLanes(i)
    Next i
    
    ' First pass: determine which lanes have labels on top (using proper array sizing)
    Dim lanesWithTopLabels() As Boolean
    ReDim lanesWithTopLabels(0 To maxLane)
    
    For i = 0 To UBound(events)
        If UCase(events(i, 3)) = "FEATURE" And IsDate(events(i, 2)) Then
            Dim startDateCheck As Date: startDateCheck = Int(events(i, 1))
            Dim endDateCheck As Date: endDateCheck = Int(events(i, 2))
            Dim featureEndXCheck As Single: featureEndXCheck = leftPadding + (endDateCheck - minDate) * scaleFactor
            Dim xPosCheck As Single: xPosCheck = leftPadding + (startDateCheck - minDate) * scaleFactor
            Dim barWidthCheck As Single: barWidthCheck = featureEndXCheck - xPosCheck
            
            ' Calculate required width based on label text length
            Dim labelText As String: labelText = CStr(events(i, 0)) ' Task name is in column A (index 0)
            Dim labelWidth As Single: labelWidth = CalculateDynamicLabelWidth(labelText, config.fontSize, 30, 300) ' Minimum 30, Maximum 300
            Dim requiredWidth As Single: requiredWidth = labelWidth + config.labelInternalPadding ' Add 20px padding (10px each side)
            
            ' Compare bar width with required space for label
            If barWidthCheck < requiredWidth Then
                ' Ensure we don't exceed array bounds
                If eventLanes(i) <= maxLane Then
                    lanesWithTopLabels(eventLanes(i)) = True
                End If
            End If
        ElseIf UCase(events(i, 3)) = "MILESTONE" Then
            ' === MILESTONE TOP LABEL DETECTION ===
            ' Check if milestone will have label on top (insufficient left space)
            Dim milestoneStartXCheck As Single: milestoneStartXCheck = leftPadding + (Int(events(i, 1)) - minDate) * scaleFactor
            Dim availableLeftSpaceCheck As Single: availableLeftSpaceCheck = milestoneStartXCheck - leftPadding
            Dim estimatedLabelWidthCheck As Single: estimatedLabelWidthCheck = (Len(CStr(events(i, 0))) * 6) + 20 ' Approximate
            Dim requiredLeftSpaceCheck As Single: requiredLeftSpaceCheck = estimatedLabelWidthCheck + 23 + 8 ' labelWidth + spacing + diamondHalfSize
            
            If availableLeftSpaceCheck < requiredLeftSpaceCheck Then
                ' Milestone will have label on top - mark lane for extra spacing
                If eventLanes(i) <= maxLane Then
                    lanesWithTopLabels(eventLanes(i)) = True
                End If
            End If
        End If
    Next i
    
    ' Second pass: place events with dynamic spacing
    For i = 0 To UBound(events)
        Dim label As String: label = events(i, 0)
        Dim startDateLoop As Date: startDateLoop = Int(events(i, 1))
        Dim endDateLoop As Date
        If IsDate(events(i, 2)) Then endDateLoop = Int(events(i, 2)) Else endDateLoop = Empty
        Dim typ As String: typ = events(i, 3)
        Dim colorName As String: colorName = events(i, 4)

        Dim xPosLoop As Single: xPosLoop = leftPadding + (startDateLoop - minDate) * scaleFactor
        
        ' OPTIMIZED LANE SPACING: Use config values instead of hardcoded margins
        Dim currentLane As Integer: currentLane = eventLanes(i)
        
        Dim yPos As Single: yPos = swimlaneY + config.swimlaneContentPadding ' Use configurable padding from swimlane top
        
        ' Calculate Y position with dynamic spacing for ALL lanes (including lane 0)
        Dim laneIndex As Integer
        For laneIndex = 0 To currentLane
            ' Ensure we don't exceed array bounds and apply consistent spacing
            If laneIndex <= maxLane And lanesWithTopLabels(laneIndex) Then
                yPos = yPos + 35 ' Top padding for lanes with top labels
            Else
                yPos = yPos + 20 ' Top padding for lanes with inside labels
            End If
        Next laneIndex
        
        ' Use calculated yPos for all events consistently

        If typ = "Milestone" Then
            ' Use consistent positioning for all lanes
            Dim milestoneYPos As Single: milestoneYPos = yPos
            
            ' Draw milestone with enhanced styling
            Call DrawCircle(sld, xPosLoop - circleSize / 2, milestoneYPos - circleSize / 2, circleSize, GetColor(colorName))
            
            ' Use intelligent label positioning same as feature bars
            Call AddIntelligentMilestoneLabel(sld, xPosLoop, xPosLoop, milestoneYPos, label, fontName, 9, _
                0, leftPadding, scaleFactor, minDate)
            
            ' Add date label vertically centered to the diamond (moved up by 4px for better positioning)
            Call AddDateLabel(sld, xPosLoop + 15, milestoneYPos - 8, Format(startDateLoop, "dd-mmm"), fontName, 8)
            
            ' Draw connector line from swimlane axis to milestone
            If eventLanes(i) > 0 Then
                Call DrawConnectorLine(sld, xPosLoop, swimlaneY, milestoneYPos - circleSize / 2)
            End If
            
        ElseIf typ = "Feature" And IsDate(endDateLoop) Then
            ' Use consistent positioning for all lanes
            Dim featureYPos As Single: featureYPos = yPos
            
            ' Draw feature bar (replaces previous phase functionality)
            Dim featureEndXLoop As Single: featureEndXLoop = leftPadding + (endDateLoop - minDate) * scaleFactor
            Dim elementHeightSingle As Single: elementHeightSingle = CSng(elementHeight)
            
            ' Validate date order and calculate proper width
            Dim barWidthLoop As Single: barWidthLoop = featureEndXLoop - xPosLoop
            If barWidthLoop <= 0 Then
                barWidthLoop = 10 ' Minimum width for invalid dates
                featureEndXLoop = xPosLoop + barWidthLoop
            End If
            
            ' Get smart color based on task name if color not specified
            Dim taskColor As Long: taskColor = GetColorFromTaskName(label, colorName)
            
            ' Draw the feature bar (clean bar without internal text)
            Call DrawBar(sld, xPosLoop, featureYPos - elementHeightSingle / 2, barWidthLoop, elementHeightSingle, taskColor)
            
            ' === NEW FEATURE BAR LABELING SYSTEM ===
            ' 1. Name label: inside bar if space allows, otherwise on top
            ' 2. Date range label: on the right side (dd mmm - dd mmm)
            ' 3. Duration label: on the left side (N d)
            Call AddEnhancedFeatureLabels(sld, xPosLoop, featureEndXLoop, featureYPos, label, startDateLoop, endDateLoop, fontName, barWidthLoop)
            
            ' Draw connector lines from swimlane axis to feature bar
            If eventLanes(i) > 0 Then
                Call DrawConnectorLine(sld, xPosLoop, swimlaneY, featureYPos)
                Call DrawConnectorLine(sld, featureEndXLoop, swimlaneY, featureYPos)
            End If
            
        ElseIf typ = "Phase" And IsDate(endDateLoop) Then
            ' Draw phase bar (collection of features - positioned between calendar and timeline)
            Dim phaseEndXLoop As Single: phaseEndXLoop = leftPadding + (endDateLoop - minDate) * scaleFactor
            Dim phaseelementHeightLoop As Single: phaseelementHeightLoop = CSng(elementHeight + 6) ' Slightly larger for phases
            
            ' Validate date order and calculate proper width for phases
            Dim phaseWidthLoop As Single: phaseWidthLoop = phaseEndXLoop - xPosLoop
            If phaseWidthLoop <= 0 Then
                phaseWidthLoop = 10 ' Minimum width for invalid dates
                phaseEndXLoop = xPosLoop + phaseWidthLoop
            End If
            
            ' Get smart color with transparency for phases
            Dim phaseColorLoop As Long: phaseColorLoop = GetColorFromTaskName(label, colorName)
            
            ' === DEDICATED PHASE AREA: Minimal whitespace for compact layout ===
            ' Phase area: 80px (calendar end) to 140px (swimlane start) = 60px dedicated space
            Dim phaseAreaTopLoop As Single: phaseAreaTopLoop = 80   ' After calendar header (50-70px) + 10px buffer
            Dim phaseAreaBottomLoop As Single: phaseAreaBottomLoop = 140 ' Before swimlanes (150px) - 10px buffer
            Dim phaseYPosLoop As Single: phaseYPosLoop = phaseAreaTopLoop + ((phaseAreaBottomLoop - phaseAreaTopLoop) / 2) ' Center in phase area
            
            ' Draw the phase bar with enhanced styling (semi-transparent overlay)
            Call DrawPhaseBar(sld, xPosLoop, phaseYPosLoop - phaseelementHeightLoop / 2, phaseWidthLoop, phaseelementHeightLoop, phaseColorLoop)
            
            ' === Position phase labels INSIDE phase bars ===
            Call AddPhaseBarLabel(sld, xPosLoop + (phaseEndXLoop - xPosLoop) / 2, phaseYPosLoop, label, fontName, 10, True, RGB(255, 255, 255))
            
            ' Add phase duration info below the bar (but still within phase area)
            Dim phaseDurationLoop As Long: phaseDurationLoop = endDateLoop - startDateLoop
            If phaseDurationLoop > 0 Then
                Dim phaseDurationTextLoop As String: phaseDurationTextLoop = "Phase: " & phaseDurationLoop & " days"
                Call AddDateLabel(sld, xPosLoop + (phaseEndXLoop - xPosLoop) / 2 - 30, phaseYPosLoop + 15, phaseDurationTextLoop, fontName, 8)
            End If
            
            ' Remove connector lines for phases since they're in their own dedicated area
            ' No connector lines needed for phases positioned in dedicated space
        End If
    Next i
End Sub

Sub AddEnhancedSwimlaneHeader(sld As Slide, x As Single, y As Single, txt As String, fontName As String, fontSize As Integer, swimlaneHeight As Single)
    ' Add swimlane label with hex #1F3763 background that matches swimlane height and aligns with background left edge
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=x, Top:=y, width:=65, height:=swimlaneHeight)
    
    With shp.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.name = fontName
        .TextRange.Font.size = fontSize
        .TextRange.Font.Bold = True
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255) ' White text on colored background
        .TextRange.ParagraphFormat.alignment = ppAlignRight ' Right-align to be close to timeline
        .VerticalAnchor = msoAnchorMiddle ' Vertically center the text
        .MarginLeft = 5
        .MarginRight = 5
        .MarginTop = 2
        .MarginBottom = 2
    End With
    
    ' Add hex #1F3763 background fill
    With shp.Fill
        .ForeColor.RGB = RGB(31, 55, 99) ' Hex #1F3763 converted to RGB
        .Solid
        .Visible = msoTrue
    End With
    shp.Line.Visible = msoFalse
End Sub

Sub DrawSwimlaneBackground(sld As Slide, x As Single, y As Single, width As Single, height As Single, swimlaneIndex As Integer)
    ' Draw swimlane section background with consistent hex #EDEDED color
    Dim bgColor As Long
    bgColor = RGB(237, 237, 237) ' Hex #EDEDED for all swimlanes
    
    ' Background section (consistent color, extends full width)
    Dim shp As Shape
    Set shp = sld.Shapes.AddShape(msoShapeRectangle, x, y, width, height)
    With shp
        .Fill.ForeColor.RGB = bgColor
        .Fill.Solid
        .Fill.Transparency = 0 ' Solid background as requested
        .Line.Visible = msoFalse
        .ZOrder msoSendToBack ' Send to back so other elements appear on top
    End With
    
    ' Add horizontal separator line at bottom of swimlane section - REMOVED per user request
    ' If swimlaneIndex > 0 Then ' Don't add line above first swimlane
    '     Dim separatorLine As Shape
    '     Set separatorLine = sld.Shapes.AddLine(x, y, x + width, y)
    '     With separatorLine.Line
    '         .ForeColor.RGB = RGB(220, 220, 220) ' Light gray separator
    '         .Weight = 1
    '         .DashStyle = msoLineSolid
    '     End With
    '     separatorLine.ZOrder msoSendBackward ' Behind events but above background
    ' End If
End Sub

' Note: AddBarLabel and AddDateLabel functions moved to TEXT LABEL UTILITIES section to avoid duplication

Sub DrawEnhancedTopTimelineAxis(sld As Slide, minDate As Date, maxDate As Date, scaleFactor As Double, leftPadding As Single, topY As Single, fontName As String)
    ' Draw enhanced top timeline axis with weekly segments, red timeline bar, and today marker
    
    Dim timelineLength As Single: timelineLength = (maxDate - minDate) * scaleFactor
    
    ' === Calculate Today marker position first ===
    Dim today As Date: today = Date
    Dim todayX As Single: todayX = -1 ' Default to -1 if today is not in range
    Dim redBarEndX As Single: redBarEndX = leftPadding + timelineLength ' Default to full length
    
    If today >= minDate And today <= maxDate Then
        todayX = leftPadding + (today - minDate) * scaleFactor
        redBarEndX = todayX ' Red bar ends at Today marker
    End If
    
    ' === Draw Red Timeline Bar (thinner, on top of week block, ending at Today marker) ===
    If redBarEndX > leftPadding Then ' Only draw if there's meaningful length
        Dim redBarWidth As Single: redBarWidth = redBarEndX - leftPadding
        Dim timelineBar As Shape
        Set timelineBar = sld.Shapes.AddShape(msoShapeRoundedRectangle, leftPadding, topY - 25, redBarWidth, 3)
        With timelineBar
            .Fill.ForeColor.RGB = RGB(220, 20, 60) ' Crimson red timeline bar
            .Fill.Solid
            .Line.Visible = msoFalse
            .Adjustments(1) = 0.15 ' Same rounded corners as calendar block
            .ZOrder msoBringToFront ' Bring to front so it appears on top of calendar block
        End With
    End If
    
    ' === Main Timeline Axis Removed for Cleaner Appearance ===
    ' Horizontal line under week days block removed per user request
    ' Clean separation between calendar header and timeline content
    
    ' === Add Enhanced Calendar Header Block with Weekly Date Segments ===
    ' Note: timelineLength already declared above, reusing the variable
    
    ' Draw rounded calendar header block with hex color #323E4F
    Dim calendarHeaderBlock As Shape
    Set calendarHeaderBlock = sld.Shapes.AddShape(msoShapeRoundedRectangle, leftPadding, topY - 25, timelineLength, 20)
    With calendarHeaderBlock
        .Fill.ForeColor.RGB = RGB(50, 62, 79) ' Hex #323E4F converted to RGB
        .Fill.Solid
        .Line.Visible = msoFalse
        .Adjustments(1) = 0.15 ' Rounded corners (15% radius)
        .ZOrder msoSendToBack ' Send to back so text appears on top
    End With
    
    Dim currentDate As Date: currentDate = minDate
    ' Start from the beginning of the week containing minDate
    currentDate = minDate - Weekday(currentDate, vbMonday) + 1
    
    ' Track the last week date to avoid adding vertical line after it
    Dim lastWeekDate As Date: lastWeekDate = minDate
    Do While lastWeekDate <= maxDate + 7
        If lastWeekDate >= minDate And lastWeekDate <= maxDate Then
            ' This is a valid week date, keep updating lastWeekDate
            lastWeekDate = lastWeekDate
        End If
        lastWeekDate = DateAdd("ww", 1, lastWeekDate)
    Loop
    lastWeekDate = DateAdd("ww", -1, lastWeekDate) ' Go back to the actual last week
    
    Do While currentDate <= maxDate + 7 ' Add buffer for complete weeks
        If currentDate >= minDate And currentDate <= maxDate Then
            Dim xPos As Single: xPos = leftPadding + (currentDate - minDate) * scaleFactor
            
            ' Vertical separator line between weeks (thin neutral grey, not touching edges)
            ' Do not add line after the last week date
            If currentDate > minDate And currentDate < lastWeekDate Then
                Dim separatorLine As Shape
                Set separatorLine = sld.Shapes.AddLine(xPos, topY - 22, xPos, topY - 8)
                With separatorLine.Line
                    .ForeColor.RGB = RGB(160, 160, 160) ' Neutral grey color
                    .Weight = 0.75 ' Thin line
                End With
            End If
            
            ' Weekly date label in dd-mmm format, left-aligned with normal font weight
            Dim weekLabel As Shape
            Set weekLabel = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
                Left:=xPos + 2, Top:=topY - 25, width:=60, height:=20)
            With weekLabel.TextFrame2
                .TextRange.Text = Format(currentDate, "dd-mmm")
                .TextRange.Font.name = fontName
                .TextRange.Font.size = 9
                .TextRange.Font.Bold = False ' Normal font weight instead of bold
                .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255) ' White text on colored background
                .TextRange.ParagraphFormat.alignment = ppAlignLeft ' Left-aligned as requested
                .VerticalAnchor = msoAnchorMiddle
                .MarginLeft = 3
                .MarginRight = 3
                .MarginTop = 2
                .MarginBottom = 2
            End With
            weekLabel.Fill.Visible = msoFalse
            weekLabel.Line.Visible = msoFalse
            weekLabel.ZOrder msoBringToFront ' Bring text to front
        End If
        
        currentDate = DateAdd("ww", 1, currentDate) ' Move to next week
    Loop
    
    ' === Month Separator Lines Removed for Cleaner Appearance ===
    ' Month separators disabled per user request to minimize visual clutter
    ' Weekly date labels provide sufficient time reference without additional lines
    
    ' === Add Enhanced "Today" Marker with Triangle Arrow ===
    If todayX > 0 Then ' Only draw if today is within timeline range
        ' Position triangle so bottom touches bottom of red timeline bar (topY - 22 is bottom of red bar)
        Call DrawTodayArrow(sld, todayX, topY - 22, fontName, RGB(220, 20, 60))
        
        ' Vertical red line removed per user request for cleaner appearance
    End If
End Sub

Sub DrawTodayArrow(sld As Slide, x As Single, y As Single, fontName As String, arrowColor As Long)
    ' Draw a simple triangle pointing down with "Today" label positioned above triangle
    
    ' Create simple triangle shape pointing down (bottom of triangle touches y position) - half height
    Dim triangle As Shape
    Set triangle = sld.Shapes.AddShape(msoShapeIsoscelesTriangle, x - 6, y - 7.5, 12, 7.5)
    With triangle
        .Fill.ForeColor.RGB = arrowColor ' Use same color as red timeline bar
        .Fill.Solid
        .Line.Visible = msoFalse ' No border for clean appearance
        .Rotation = 180 ' Point downward
    End With
    
    ' Add "Today" label ABOVE triangle (positioned above triangle top)
    Dim triangleTop As Single: triangleTop = y - 7.5  ' Top of triangle (half height)
    Dim todayLabel As Shape
    Set todayLabel = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=x - 20, Top:=triangleTop - 18, width:=40, height:=15)
    With todayLabel.TextFrame2
        .TextRange.Text = "Today"
        .TextRange.Font.name = fontName
        .TextRange.Font.size = 8 ' Smaller font size
        .TextRange.Font.Bold = False ' Normal font weight like week days
        .TextRange.Font.Fill.ForeColor.RGB = RGB(50, 62, 79) ' Same color as week days block background
        .TextRange.ParagraphFormat.alignment = ppAlignCenter
        .VerticalAnchor = msoAnchorMiddle
    End With
    todayLabel.Fill.Visible = msoFalse
    todayLabel.Line.Visible = msoFalse
End Sub

Sub DrawTopTimelineAxis(sld As Slide, minDate As Date, maxDate As Date, scaleFactor As Double, leftPadding As Single, topY As Single, fontName As String)
    ' Legacy function - replaced by DrawEnhancedTopTimelineAxis
    ' Keeping for backward compatibility
    Call DrawEnhancedTopTimelineAxis(sld, minDate, maxDate, scaleFactor, leftPadding, topY, fontName)
End Sub

' === Lane Assignment for Overlapping Events ===
Function AssignLanesToEvents(milestoneData() As Variant, ByRef eventLanes() As Integer, _
                            scaleFactor As Double, leftPadding As Single, minDate As Date) As Integer
    ' Enhanced lane assignment with smart conflict resolution
    ' Events ending later (extending further right) are moved to higher lanes
    
    Dim numEvents As Integer: numEvents = UBound(milestoneData) + 1
    Dim i As Integer, j As Integer
    Dim currentLanes As Integer: currentLanes = 0
    
    ' Initialize all events to lane 0 (main timeline)
    For i = 0 To UBound(eventLanes)
        eventLanes(i) = 0
    Next i
    
    ' Sort events by start date first for logical processing
    Call SortEventsByStartDate(milestoneData, eventLanes)
    
    ' Process each event to check for overlaps
    For i = 0 To numEvents - 1
        Dim assignedLane As Integer: assignedLane = 0
        Dim laneFound As Boolean: laneFound = False
        
        ' Find the lowest available lane for this event (unlimited lanes per swimlane)
        Do While Not laneFound
            laneFound = True
            
            ' Check if this lane is available (no overlaps)
            For j = 0 To i - 1
                If eventLanes(j) = assignedLane And EventsOverlap(milestoneData, i, j, scaleFactor, leftPadding, minDate) Then
                    ' Conflict detected - use CORRECTED logic
                    Dim currentEventEnd As Date, conflictEventEnd As Date
                    currentEventEnd = GetEventEndDate(milestoneData, i)
                    conflictEventEnd = GetEventEndDate(milestoneData, j)
                    
                    ' FIXED: Event ending LATER gets moved to higher lane (further down)
                    If currentEventEnd > conflictEventEnd Then
                        ' Current event ends later, so it moves to higher lane
                        assignedLane = assignedLane + 1
                        laneFound = False
                        Exit For
                    Else
                        ' Conflicting event ends later, move it to higher lane (no limit)
                        Call MoveEventToHigherLane(eventLanes, j, assignedLane + 1)
                        Exit For ' Re-check this lane since we moved the conflict
                    End If
                End If
            Next j
        Loop
        
        eventLanes(i) = assignedLane
        If assignedLane > currentLanes Then currentLanes = assignedLane
    Next i
    
    AssignLanesToEvents = currentLanes + 1 ' Return total number of lanes
End Function

Function EventsOverlap(milestoneData() As Variant, event1 As Integer, event2 As Integer, _
                      scaleFactor As Double, leftPadding As Single, minDate As Date) As Boolean
    ' Enhanced overlap detection considering bars, labels, and date labels
    ' This ensures proper spacing for all visual elements of timeline events
    
    Dim start1 As Date, end1 As Date, start2 As Date, end2 As Date
    Dim x1Start As Single, x1End As Single, x2Start As Single, x2End As Single
    Dim type1 As String, type2 As String
    Dim label1 As String, label2 As String
    
    ' Get event details
    start1 = Int(milestoneData(event1, 1))
    start2 = Int(milestoneData(event2, 1))
    type1 = CStr(milestoneData(event1, 3))
    type2 = CStr(milestoneData(event2, 3))
    label1 = CStr(milestoneData(event1, 0))
    label2 = CStr(milestoneData(event2, 0))
    
    ' Get end dates
    If IsDate(milestoneData(event1, 2)) Then
        end1 = Int(milestoneData(event1, 2))
    Else
        end1 = start1 ' Milestone has same start and end
    End If
    
    If IsDate(milestoneData(event2, 2)) Then
        end2 = Int(milestoneData(event2, 2))
    Else
        end2 = start2 ' Milestone has same start and end
    End If
    
    ' Convert dates to base X positions
    x1Start = leftPadding + (start1 - minDate) * scaleFactor
    x1End = leftPadding + (end1 - minDate) * scaleFactor
    x2Start = leftPadding + (start2 - minDate) * scaleFactor
    x2End = leftPadding + (end2 - minDate) * scaleFactor
    
    ' === ENHANCED OVERLAP DETECTION WITH LABEL SPACE ===
    ' Calculate extended boundaries including all visual elements
    
    ' Event 1 extended boundaries
    Dim event1ExtendedStart As Single, event1ExtendedEnd As Single
    Call CalculateEventExtendedBounds(x1Start, x1End, type1, label1, _
        event1ExtendedStart, event1ExtendedEnd)
    
    ' Event 2 extended boundaries
    Dim event2ExtendedStart As Single, event2ExtendedEnd As Single
    Call CalculateEventExtendedBounds(x2Start, x2End, type2, label2, _
        event2ExtendedStart, event2ExtendedEnd)
    
    ' Check for overlap using extended boundaries
    EventsOverlap = Not (event1ExtendedEnd < event2ExtendedStart Or event2ExtendedEnd < event1ExtendedStart)
End Function

' ===================================================================
' SHAPE DRAWING UTILITIES
' ===================================================================

' --- Basic Geometric Shapes ---
Sub DrawLine(sld As Slide, x1 As Single, y1 As Single, x2 As Single, y2 As Single, clr As Long)
    ' Draw a simple line with specified color
    Dim shp As Shape
    Set shp = sld.Shapes.AddLine(x1, y1, x2, y2)
    With shp.Line
        .ForeColor.RGB = clr
        .Weight = 2
    End With
End Sub

Sub DrawConnectorLine(sld As Slide, x As Single, y1 As Single, y2 As Single)
    ' Draw a subtle connector line with professional styling - REMOVED per user request
    
    ' Connector lines disabled to clean up timeline bar area
    ' No connector lines will be drawn
    Exit Sub
End Sub

' --- Event Shape Rendering ---
Sub DrawCircle(sld As Slide, x As Single, y As Single, size As Integer, clr As Long)
    ' Draw milestone with diamond shape for professional timelines
    Dim shp As Shape
    Set shp = sld.Shapes.AddShape(msoShapeDiamond, x, y, size, size)
    With shp.Fill
        .ForeColor.RGB = clr
        .Solid
    End With
    With shp.Line
        .ForeColor.RGB = RGB(255, 255, 255) ' White border for contrast
        .Weight = 2
    End With
End Sub

Sub DrawBar(sld As Slide, x As Single, y As Single, width As Single, height As Single, clr As Long)
    ' Draw feature bar with rounded rectangle styling and parameter validation
    
    ' Validate parameters to prevent runtime errors
    If width <= 0 Or height <= 0 Then
        Debug.Print "DrawBar: Invalid dimensions - width=" & width & ", height=" & height
        Exit Sub
    End If
    
    If x < 0 Or y < 0 Or x > 2000 Or y > 2000 Then
        Debug.Print "DrawBar: Invalid position - x=" & x & ", y=" & y
        Exit Sub
    End If
    
    ' Ensure minimum dimensions
    If width < 1 Then width = 1
    If height < 1 Then height = 1
    
    Dim shp As Shape
    Set shp = sld.Shapes.AddShape(msoShapeRoundedRectangle, x, y, width, height)
    With shp.Fill
        .ForeColor.RGB = clr
        .Solid
    End With
    With shp
        .Line.Visible = msoFalse
        .Adjustments(1) = 0.2 ' Subtle corner radius
    End With
End Sub

Sub DrawPhaseBar(sld As Slide, x As Single, y As Single, width As Single, height As Single, clr As Long)
    ' Draw phase bar with enhanced styling and semi-transparency with parameter validation
    
    ' Validate parameters to prevent runtime errors
    If width <= 0 Or height <= 0 Then
        Debug.Print "DrawPhaseBar: Invalid dimensions - width=" & width & ", height=" & height
        Exit Sub
    End If
    
    If x < 0 Or y < 0 Or x > 2000 Or y > 2000 Then
        Debug.Print "DrawPhaseBar: Invalid position - x=" & x & ", y=" & y
        Exit Sub
    End If
    
    ' Ensure minimum dimensions
    If width < 1 Then width = 1
    If height < 1 Then height = 1
    
    Dim shp As Shape
    Set shp = sld.Shapes.AddShape(msoShapeRoundedRectangle, x, y, width, height)
    With shp.Fill
        .ForeColor.RGB = clr
        .Solid
        .Transparency = 0.3 ' Semi-transparent to show underlying features
    End With
    With shp
        .Line.ForeColor.RGB = clr
        .Line.Weight = 2
        .Line.DashStyle = msoLineDash ' Dashed border to distinguish from features
        .Adjustments(1) = 0.25 ' Larger corner radius for phases
    End With
End Sub

Sub DrawArrowBar(sld As Slide, x As Single, y As Single, width As Single, height As Single, clr As Long)
    ' Draw arrow-shaped task bar for special event types
    Dim shp As Shape
    Set shp = sld.Shapes.AddShape(msoShapeRightArrow, x, y, width, height)
    With shp.Fill
        .ForeColor.RGB = clr
        .Solid
    End With
    With shp
        .Line.Visible = msoFalse
        .Adjustments(1) = 0.25 ' Arrow head width
        .Adjustments(2) = 0.5  ' Arrow head position
    End With
End Sub

' ===================================================================
' TEXT LABEL UTILITIES
' ===================================================================

' === SHARED LABEL POSITIONING FUNCTION ===
Sub AddEventLabel(sld As Slide, x As Single, y As Single, width As Single, height As Single, _
                  txt As String, fontName As String, fontSize As Integer, _
                  alignment As Long, textColor As Long, isBold As Boolean)
    ' Centralized function for creating event labels with consistent positioning
    ' Used by feature bars, milestones, and phases for standardized label creation
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=x, Top:=y, width:=width, height:=height)
    With shp.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.name = fontName
        .TextRange.Font.size = fontSize
        .TextRange.Font.Bold = isBold
        .TextRange.Font.Fill.ForeColor.RGB = textColor
        .TextRange.ParagraphFormat.alignment = alignment
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 3
        .MarginRight = 3
        .MarginTop = 1
        .MarginBottom = 1
    End With
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
End Sub

' --- Dynamic Label Width Helper Function ---
Function CalculateDynamicLabelWidth(labelText As String, fontSize As Integer, minWidth As Single, maxWidth As Single) As Single
    ' Calculate dynamic label width based on text length with configurable min/max bounds
    ' Approximate character width: varies by font size (6-8 pixels per character for Calibri)
    ' NOTE: minWidth and maxWidth are now REQUIRED parameters to ensure config values are always used
    
    Dim baseCharWidth As Single
    Select Case fontSize
        Case Is <= 8: baseCharWidth = 5.5
        Case 9: baseCharWidth = 6
        Case 10: baseCharWidth = 6.5
        Case 11: baseCharWidth = 7
        Case Is >= 12: baseCharWidth = 7.5
        Case Else: baseCharWidth = 6 ' Default
    End Select
    
    ' Calculate width with padding
    Dim calculatedWidth As Single
    calculatedWidth = (Len(labelText) * baseCharWidth) + 20 ' 10px padding on each side
    
    ' Apply min/max bounds from config
    If calculatedWidth < minWidth Then calculatedWidth = minWidth
    If calculatedWidth > maxWidth Then calculatedWidth = maxWidth
    
    CalculateDynamicLabelWidth = calculatedWidth
End Function

' --- Hierarchical Label System ---
Sub AddLabel(sld As Slide, x As Single, y As Single, txt As String, fontName As String, fontSize As Integer, center As Boolean)
    ' Generic label function for basic text placement
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=x - 50, Top:=y, width:=100, height:=40)
    With shp.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.name = fontName
        .TextRange.Font.size = fontSize
        .TextRange.ParagraphFormat.alignment = IIf(center, ppAlignCenter, ppAlignLeft)
    End With
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
End Sub

Sub AddPhaseLabel(sld As Slide, x As Single, y As Single, txt As String, fontName As String, fontSize As Integer, center As Boolean)
    ' Enhanced phase label for top-level hierarchy
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=x - 80, Top:=y, width:=160, height:=20)
    With shp.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.name = fontName
        .TextRange.Font.size = fontSize
        .TextRange.Font.Bold = True
        .TextRange.Font.Fill.ForeColor.RGB = RGB(30, 30, 30) ' Dark for hierarchy emphasis
        .TextRange.ParagraphFormat.alignment = IIf(center, ppAlignCenter, ppAlignLeft)
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 5
        .MarginRight = 5
        .MarginTop = 2
        .MarginBottom = 2
    End With
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
End Sub

Sub AddFeatureLabel(sld As Slide, x As Single, y As Single, txt As String, fontName As String, fontSize As Integer, center As Boolean)
    ' Enhanced feature label for mid-level hierarchy
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=x - 60, Top:=y, width:=120, height:=18)
    With shp.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.name = fontName
        .TextRange.Font.size = fontSize
        .TextRange.Font.Bold = True
        .TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50) ' Dark gray for visibility
        .TextRange.ParagraphFormat.alignment = IIf(center, ppAlignCenter, ppAlignLeft)
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 3
        .MarginRight = 3
        .MarginTop = 1
        .MarginBottom = 1
    End With
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
End Sub

' --- Specialized Label Functions ---
Sub AddBarLabel(sld As Slide, x As Single, y As Single, txt As String, fontName As String, fontSize As Integer, center As Boolean, textColor As Long)
    ' Label inside bars with custom text color
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=x - 50, Top:=y - 8, width:=100, height:=16)
    With shp.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.name = fontName
        .TextRange.Font.size = fontSize
        .TextRange.Font.Bold = True
        .TextRange.Font.Fill.ForeColor.RGB = textColor
        .TextRange.ParagraphFormat.alignment = IIf(center, ppAlignCenter, ppAlignLeft)
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 0
        .MarginBottom = 0
    End With
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
End Sub

Sub AddDateLabel(sld As Slide, x As Single, y As Single, txt As String, fontName As String, fontSize As Integer)
    ' Subtle date label with consistent styling
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=x, Top:=y, width:=120, height:=15)
    With shp.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.name = fontName
        .TextRange.Font.size = fontSize
        .TextRange.Font.Fill.ForeColor.RGB = RGB(100, 100, 100)
        .TextRange.ParagraphFormat.alignment = ppAlignLeft
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 2
        .MarginRight = 2
    End With
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
End Sub

Sub AddCenteredDateLabel(sld As Slide, centerX As Single, y As Single, txt As String, fontName As String, fontSize As Integer)
    ' Horizontally centered date label for phase duration text
    Dim labelWidth As Single: labelWidth = 80
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=centerX - (labelWidth / 2), Top:=y, width:=labelWidth, height:=15)
    With shp.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.name = fontName
        .TextRange.Font.size = fontSize
        .TextRange.Font.Fill.ForeColor.RGB = RGB(100, 100, 100)
        .TextRange.ParagraphFormat.alignment = ppAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 2
        .MarginRight = 2
    End With
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
End Sub

Sub AddPhaseBarLabel(sld As Slide, x As Single, y As Single, txt As String, fontName As String, fontSize As Integer, center As Boolean, textColor As Long)
    ' Enhanced phase label positioned INSIDE phase bars with high contrast
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=x - 80, Top:=y - 10, width:=160, height:=20)
    With shp.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.name = fontName
        .TextRange.Font.size = fontSize
        .TextRange.Font.Bold = True
        .TextRange.Font.Fill.ForeColor.RGB = textColor ' White text for contrast
        .TextRange.ParagraphFormat.alignment = IIf(center, ppAlignCenter, ppAlignLeft)
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 5
        .MarginRight = 5
        .MarginTop = 2
        .MarginBottom = 2
    End With
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    ' Bring to front so text appears on top of the phase bar
    shp.ZOrder msoBringToFront
End Sub

Sub AddIntelligentFeatureLabel(sld As Slide, barStartX As Single, barEndX As Single, barY As Single, _
                              txt As String, fontName As String, fontSize As Integer, _
                              barWidth As Single, leftPadding As Single, scaleFactor As Double, minDate As Date)
    ' Intelligent label positioning: inside bar if space allows, otherwise on appropriate side
    
    ' Get configuration values for consistent behavior
    Dim config As TimelineConfig: config = GetDefaultTimelineConfig()
    
    Const SlideMiddle As Single = config.slideWidth / 2                  ' Middle of slide (960/2)
    
    Dim labelX As Single
    Dim labelAlignment As Long
    
    ' Calculate actual label width based on text length
    Dim labelWidth As Single: labelWidth = CalculateDynamicLabelWidth(txt, fontSize, config.featureNameLabelMinWidth, config.featureNameLabelMaxWidth)
    
    ' Add some padding for comfortable fit (20px buffer: 10px on each side)
    Dim requiredWidth As Single: requiredWidth = labelWidth + config.labelInternalPadding
    
    ' Compare bar width with required width for label plus padding
    If barWidth >= requiredWidth Then
        ' === BAR IS WIDE ENOUGH: Place label INSIDE the bar ===
        labelX = barStartX + (barWidth / 2) - (labelWidth / 2)
        labelAlignment = ppAlignCenter
        
        ' Position label inside bar
        Dim shp As Shape
        Set shp = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
            Left:=labelX, Top:=barY - 8, width:=labelWidth, height:=16)
        With shp.TextFrame2
            .TextRange.Text = txt
            .TextRange.Font.name = fontName
            .TextRange.Font.size = fontSize
            .TextRange.Font.Bold = True
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255) ' White text inside bar
            .TextRange.ParagraphFormat.alignment = labelAlignment
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 2
            .MarginRight = 2
            .MarginTop = 0
            .MarginBottom = 0
        End With
        shp.Fill.Visible = msoFalse
        shp.Line.Visible = msoFalse
        shp.ZOrder msoBringToFront ' Ensure text is on top of bar
        
    Else
        ' === BAR IS TOO NARROW: Place label on appropriate SIDE ===
        Dim barCenterX As Single: barCenterX = barStartX + (barWidth / 2)
        
        If barCenterX <= SlideMiddle Then
            ' Bar is on LEFT side of slide -> Place label on RIGHT side of bar
            labelX = barEndX + 5
            labelAlignment = ppAlignLeft
        Else
            ' Bar is on RIGHT side of slide -> Place label on LEFT side of bar
            labelX = barStartX - labelWidth - 5
            labelAlignment = ppAlignRight
        End If
        
        ' Position label vertically centered to the bar
        Dim shp2 As Shape
        Set shp2 = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
            Left:=labelX, Top:=barY - 9, width:=labelWidth, height:=18)
        With shp2.TextFrame2
            .TextRange.Text = txt
            .TextRange.Font.name = fontName
            .TextRange.Font.size = fontSize
            .TextRange.Font.Bold = True
            .TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50) ' Dark gray for external labels
            .TextRange.ParagraphFormat.alignment = labelAlignment
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 3
            .MarginRight = 3
            .MarginTop = 1
            .MarginBottom = 1
        End With
        shp2.Fill.Visible = msoFalse
        shp2.Line.Visible = msoFalse
    End If
End Sub

Sub AddIntelligentMilestoneLabel(sld As Slide, milestoneStartX As Single, milestoneEndX As Single, milestoneY As Single, _
                                txt As String, fontName As String, fontSize As Integer, _
                                milestoneWidth As Single, leftPadding As Single, scaleFactor As Double, minDate As Date)
    ' Intelligent milestone label positioning with priority for LEFT side placement
    ' NEW RULES:
    ' 1. Always prefer LEFT side of milestone marker
    ' 2. Only place ON TOP if insufficient space on left side
    ' 3. Uses dynamic width calculation based on text length for optimal label sizing
    
    ' Get config values for milestone label constraints
    Dim config As TimelineConfig: config = GetDefaultTimelineConfig()
    
    ' DYNAMIC WIDTH: Calculate label width based on actual text length using config constraints
    Dim labelWidth As Single: labelWidth = CalculateDynamicLabelWidth(txt, fontSize, config.milestoneLabelMinWidth, config.milestoneLabelMaxWidth)
    
    ' Convert integer config values to Singles to avoid ByRef type mismatch
    Dim labelVerticalOffsetSingle As Single: labelVerticalOffsetSingle = CSng(config.labelVerticalOffset)
    Dim labelHeightSingle As Single: labelHeightSingle = CSng(config.labelHeight)
    Dim labelSpacingToShapeSingle As Single: labelSpacingToShapeSingle = CSng(config.labelSpacingToShape)
    
    ' Calculate space available on the left side of milestone
    Dim diamondHalfSize As Single: diamondHalfSize = 8 ' Half of 16px diamond
    Dim availableLeftSpace As Single: availableLeftSpace = milestoneStartX - leftPadding
    
    ' EXTREMELY CLOSE SPACING: Bring milestone markers extremely close to their labels
    Dim closeLabelSpacing As Single: closeLabelSpacing = 0.5 ' Reduced to 2px for extremely close positioning
    Dim requiredLeftSpace As Single: requiredLeftSpace = labelWidth + closeLabelSpacing + diamondHalfSize
    
    Dim labelX As Single, labelY As Single
    
    ' === POSITIONING LOGIC: LEFT PREFERRED, TOP AS FALLBACK ===
    If availableLeftSpace >= requiredLeftSpace Then
        ' === SUFFICIENT SPACE ON LEFT: Place label on LEFT side of milestone ===
        ' Apply 5px left shift for better visual spacing
        labelX = milestoneStartX - diamondHalfSize - closeLabelSpacing - labelWidth - 5
        labelY = milestoneY + labelVerticalOffsetSingle ' Vertically centered to milestone
        
        ' Use shared label function for LEFT positioning with RIGHT alignment
        Call AddEventLabel(sld, labelX, labelY, labelWidth, labelHeightSingle, _
                          txt, fontName, fontSize, ppAlignRight, RGB(50, 50, 50), True)
                          
    Else
        ' === INSUFFICIENT SPACE ON LEFT: Place label ON TOP of milestone ===
        labelX = milestoneStartX ' Left edge of label aligns with milestone center
        labelY = milestoneY - diamondHalfSize - closeLabelSpacing - labelHeightSingle ' Above milestone with closer spacing
        
        ' Use shared label function for TOP positioning with LEFT alignment
        Call AddEventLabel(sld, labelX, labelY, labelWidth, labelHeightSingle, _
                          txt, fontName, fontSize, ppAlignLeft, RGB(50, 50, 50), True)
    End If
End Sub

Sub AddTwoLinePhaseLabels(sld As Slide, centerX As Single, phaseY As Single, mainLabel As String, durationLabel As String, fontName As String)
    ' Add two-line phase labels with CONFIGURABLE DYNAMIC WIDTHS: main label + duration on separate lines, positioned slightly above center within phase bar
    
    ' Get config values for phase label constraints
    Dim config As TimelineConfig: config = GetDefaultTimelineConfig()
    
    ' === CALCULATE DYNAMIC WIDTHS FOR EACH LINE ===
    Dim mainLabelWidth As Single: mainLabelWidth = CalculateDynamicLabelWidth(mainLabel, 10, config.phaseLabelMinWidth, config.phaseLabelMaxWidth)
    Dim durationLabelWidth As Single: durationLabelWidth = CalculateDynamicLabelWidth(durationLabel, 8, config.phaseLabelMinWidth, config.phaseLabelMaxWidth)
    
    ' Use the larger width for consistent alignment
    Dim maxLabelWidth As Single: maxLabelWidth = mainLabelWidth
    If durationLabelWidth > maxLabelWidth Then maxLabelWidth = durationLabelWidth
    
    Dim totalBlockHeight As Single: totalBlockHeight = 24  ' Height for two lines with spacing
    Dim lineSpacing As Single: lineSpacing = 12           ' Spacing between lines
    
    ' Calculate starting Y position - moved up by 3px from center for better visual placement
    Dim blockStartY As Single: blockStartY = phaseY - (totalBlockHeight / 2) - 3
    
    ' === MAIN LABEL (Top line) WITH DYNAMIC WIDTH ===
    Dim mainLabelY As Single: mainLabelY = blockStartY
    Dim mainShape As Shape
    Set mainShape = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=centerX - (maxLabelWidth / 2), Top:=mainLabelY, width:=maxLabelWidth, height:=12)
    With mainShape.TextFrame2
        .TextRange.Text = mainLabel
        .TextRange.Font.name = fontName
        .TextRange.Font.size = 10
        .TextRange.Font.Bold = True
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255) ' White text for contrast
        .TextRange.ParagraphFormat.alignment = ppAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 0
        .MarginBottom = 0
    End With
    mainShape.Fill.Visible = msoFalse
    mainShape.Line.Visible = msoFalse
    mainShape.ZOrder msoBringToFront
    
    ' === DURATION LABEL (Bottom line) WITH DYNAMIC WIDTH - Only if duration text exists ===
    If Trim(durationLabel) <> "" Then
        Dim durationLabelY As Single: durationLabelY = blockStartY + lineSpacing
        Dim durationShape As Shape
        Set durationShape = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
            Left:=centerX - (maxLabelWidth / 2), Top:=durationLabelY, width:=maxLabelWidth, height:=12)
        With durationShape.TextFrame2
            .TextRange.Text = durationLabel
            .TextRange.Font.name = fontName
            .TextRange.Font.size = 8
            .TextRange.Font.Bold = False  ' Normal font weight as requested
            .TextRange.Font.Fill.ForeColor.RGB = RGB(220, 220, 220) ' Slightly smoother than main label (lighter gray)
            .TextRange.ParagraphFormat.alignment = ppAlignCenter
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 2
            .MarginRight = 2
            .MarginTop = 0
            .MarginBottom = 0
        End With
        durationShape.Fill.Visible = msoFalse
        durationShape.Line.Visible = msoFalse
        durationShape.ZOrder msoBringToFront
    End If
End Sub

Sub AddEnhancedFeatureLabels(sld As Slide, barStartX As Single, barEndX As Single, barY As Single, _
                            taskName As String, startDate As Date, endDate As Date, fontName As String, barWidth As Single)
    ' Enhanced feature bar labeling system with DYNAMIC LABEL WIDTHS:
    ' 1. Name label: inside bar if space allows, otherwise on top (close to bar) - DYNAMIC WIDTH
    ' 2. Date range label: on the right side (dd mmm - dd mmm format) - DYNAMIC WIDTH
    ' 3. Duration label: on the left side (N d format) - DYNAMIC WIDTH
    ' 4. All labels use dynamic sizing based on text content
    ' 5. Lane spacing provides gap between bars and top-positioned labels from other lanes
    
    ' Get config values for feature label constraints
    Dim config As TimelineConfig: config = GetDefaultTimelineConfig()
    
    ' Calculate duration
    Dim duration As Long: duration = endDate - startDate
    Dim durationText As String: durationText = duration & " d"
    
    ' Format date range
    Dim dateRangeText As String
    dateRangeText = Format(startDate, "dd mmm") & " - " & Format(endDate, "dd mmm")
    
    ' === CALCULATE DYNAMIC LABEL WIDTHS USING CONFIG VALUES ===
    Dim nameWidth As Single: nameWidth = CalculateDynamicLabelWidth(taskName, 9, config.featureNameLabelMinWidth, config.featureNameLabelMaxWidth)
    Dim durationWidth As Single: durationWidth = CalculateDynamicLabelWidth(durationText, 8, config.featureDurationLabelMinWidth, config.featureDurationLabelMaxWidth)
    Dim dateRangeWidth As Single: dateRangeWidth = CalculateDynamicLabelWidth(dateRangeText, 8, config.featureDateRangeLabelMinWidth, config.featureDateRangeLabelMaxWidth)
    
    ' === 1. NAME LABEL POSITIONING WITH DYNAMIC WIDTH ===
    ' Add padding for text to fit comfortably (20px buffer: 10px on each side)
    Dim requiredWidth As Single: requiredWidth = nameWidth + config.labelInternalPadding
    
    ' Compare bar width with required width for label plus padding
    If barWidth >= requiredWidth Then
        ' === NAME LABEL INSIDE BAR ===
        Dim nameInsideX As Single: nameInsideX = barStartX + (barWidth / 2) - (nameWidth / 2)
        Dim nameShape As Shape
        Set nameShape = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
            Left:=nameInsideX, Top:=barY - 8, width:=nameWidth, height:=16)
        With nameShape.TextFrame2
            .TextRange.Text = taskName
            .TextRange.Font.name = fontName
            .TextRange.Font.size = 9
            .TextRange.Font.Bold = True
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255) ' White text inside bar
            .TextRange.ParagraphFormat.alignment = ppAlignCenter
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 2
            .MarginRight = 2
            .MarginTop = 0
            .MarginBottom = 0
        End With
        nameShape.Fill.Visible = msoFalse
        nameShape.Line.Visible = msoFalse
        nameShape.ZOrder msoBringToFront ' Ensure text is on top of bar
    Else
        ' === NAME LABEL ON TOP OF BAR (close to bar, no additional gap) ===
        Dim nameTopX As Single: nameTopX = barStartX + (barWidth / 2) - (nameWidth / 2)
        Dim nameTopShape As Shape
        Set nameTopShape = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
            Left:=nameTopX, Top:=barY - 23, width:=nameWidth, height:=16)
        With nameTopShape.TextFrame2
            .TextRange.Text = taskName
            .TextRange.Font.name = fontName
            .TextRange.Font.size = 9
            .TextRange.Font.Bold = True
            .TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50) ' Dark gray for external labels
            .TextRange.ParagraphFormat.alignment = ppAlignCenter
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 2
            .MarginRight = 2
            .MarginTop = 0
            .MarginBottom = 0
        End With
        nameTopShape.Fill.Visible = msoFalse
        nameTopShape.Line.Visible = msoFalse
    End If
    
    ' === 2. DURATION LABEL ON LEFT SIDE WITH DYNAMIC WIDTH ===
    Dim durationX As Single: durationX = barStartX - durationWidth - 5 ' Positioned based on actual width
    Dim durationShape As Shape
    Set durationShape = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=durationX, Top:=barY - 8, width:=durationWidth, height:=16)
    With durationShape.TextFrame2
        .TextRange.Text = durationText
        .TextRange.Font.name = fontName
        .TextRange.Font.size = 8
        .TextRange.Font.Bold = False
        .TextRange.Font.Fill.ForeColor.RGB = RGB(100, 100, 100) ' Gray text
        .TextRange.ParagraphFormat.alignment = ppAlignRight ' Right-align to be close to bar
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 0
        .MarginBottom = 0
    End With
    durationShape.Fill.Visible = msoFalse
    durationShape.Line.Visible = msoFalse
    
    ' === 3. DATE RANGE LABEL ON RIGHT SIDE WITH DYNAMIC WIDTH ===
    Dim dateRangeX As Single: dateRangeX = barEndX + 5 ' Right side of bar
    Dim dateRangeShape As Shape
    Set dateRangeShape = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=dateRangeX, Top:=barY - 8, width:=dateRangeWidth, height:=16)
    With dateRangeShape.TextFrame2
        .TextRange.Text = dateRangeText
        .TextRange.Font.name = fontName
        .TextRange.Font.size = 8
        .TextRange.Font.Bold = False
        .TextRange.Font.Fill.ForeColor.RGB = RGB(100, 100, 100) ' Gray text
        .TextRange.ParagraphFormat.alignment = ppAlignLeft ' Left-align from bar edge
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 0
        .MarginBottom = 0
    End With
    dateRangeShape.Fill.Visible = msoFalse
    dateRangeShape.Line.Visible = msoFalse
End Sub

Function GetColorFromTaskName(taskName As String, colorName As String) As Long
    ' Smart color detection combining explicit color and task name analysis
    
    ' First, try to get color from explicit color column
    If Trim(colorName) <> "" And LCase(Trim(colorName)) <> "default" Then
        GetColorFromTaskName = GetColor(colorName)
        Exit Function
    End If
    
    ' If no explicit color, analyze task name for phase type
    Dim lowerTaskName As String: lowerTaskName = LCase(taskName)
    
    ' === Target Format Phase Detection ===
    If InStr(lowerTaskName, "discovery") > 0 Or InStr(lowerTaskName, "mapping") > 0 Or InStr(lowerTaskName, "analysis") > 0 Then
        GetColorFromTaskName = RGB(118, 203, 127) ' Green for discovery/mapping
    ElseIf InStr(lowerTaskName, "test") > 0 Or InStr(lowerTaskName, "qa") > 0 Or InStr(lowerTaskName, "validation") > 0 Then
        GetColorFromTaskName = RGB(160, 160, 160) ' Gray for testing
    ElseIf InStr(lowerTaskName, "build") > 0 Or InStr(lowerTaskName, "rollout") > 0 Or InStr(lowerTaskName, "deploy") > 0 Or InStr(lowerTaskName, "production") > 0 Then
        GetColorFromTaskName = RGB(68, 114, 196) ' Blue for build/rollout
    Else
        ' Default color based on position or swimlane context
        GetColorFromTaskName = RGB(68, 114, 196) ' Professional blue default
    End If
End Function

Function GetColor(name As String) As Long
    ' Enhanced color mapping with phase-specific target colors
    Select Case LCase(name)
        ' === Target Format Phase Colors ===
        Case "mapping", "green": GetColor = RGB(118, 203, 127)  ' Green for mapping phase
        Case "test", "testing", "gray", "grey": GetColor = RGB(160, 160, 160)  ' Gray for testing phase
        Case "rollout", "deployment", "blue": GetColor = RGB(68, 114, 196)     ' Blue for rollout phase
        
        ' === Standard Colors ===
        Case "red": GetColor = RGB(220, 20, 60)       ' Crimson red
        Case "orange": GetColor = RGB(255, 153, 0)    ' Vibrant orange
        Case "purple": GetColor = RGB(112, 48, 160)   ' Deep purple
        Case "yellow": GetColor = RGB(255, 192, 0)    ' Golden yellow
        Case "lightblue": GetColor = RGB(173, 216, 230)
        Case "lightgreen": GetColor = RGB(144, 238, 144)
        Case "pink": GetColor = RGB(255, 182, 193)
        Case "brown": GetColor = RGB(165, 42, 42)
        Case "navy": GetColor = RGB(25, 25, 112)
        Case "teal": GetColor = RGB(0, 128, 128)
        Case "darkgreen": GetColor = RGB(0, 100, 0)
        
        ' === Auto-detect phase type from task name ===
        Case Else:
            ' Smart color detection based on task name content
            If InStr(LCase(name), "map") > 0 Or InStr(LCase(name), "discover") > 0 Then
                GetColor = RGB(118, 203, 127) ' Green for mapping/discovery
            ElseIf InStr(LCase(name), "test") > 0 Or InStr(LCase(name), "qa") > 0 Then
                GetColor = RGB(160, 160, 160) ' Gray for testing
            ElseIf InStr(LCase(name), "rollout") > 0 Or InStr(LCase(name), "deploy") > 0 Or InStr(LCase(name), "build") > 0 Then
                GetColor = RGB(68, 114, 196) ' Blue for rollout/deployment
            Else
                GetColor = RGB(68, 114, 196) ' Default professional blue
            End If
    End Select
End Function

Function ReadDataFromExcel(sheetName As String) As Variant
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim lastRow As Long, i As Long, rawData() As Variant, result() As Variant

    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    
    If xlApp Is Nothing Then
        MsgBox "Excel is not open.", vbCritical
        Exit Function
    End If
    
    On Error GoTo 0

    Set xlBook = xlApp.ActiveWorkbook
    If xlBook Is Nothing Then
        MsgBox "No active Excel workbook.", vbCritical
        Exit Function
    End If
    
    On Error Resume Next
    Set xlSheet = xlBook.Sheets(sheetName)
    If xlSheet Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' not found.", vbCritical
        Exit Function
    End If
    On Error GoTo 0

    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).Row ' xlUp
    If lastRow < 2 Then Exit Function ' no data

    rawData = xlSheet.Range("A2:F" & lastRow).Value

    ReDim result(0 To UBound(rawData) - 1, 0 To 5)
    For i = 1 To UBound(rawData)
        result(i - 1, 0) = rawData(i, 1) ' Task Name
        result(i - 1, 1) = rawData(i, 2) ' Start Date
        result(i - 1, 2) = rawData(i, 3) ' End Date
        result(i - 1, 3) = rawData(i, 4) ' Type
        result(i - 1, 4) = rawData(i, 5) ' Color
        result(i - 1, 5) = IIf(IsEmpty(rawData(i, 6)), "Default", rawData(i, 6)) ' Swimlane
    Next i

    ReadDataFromExcel = result
End Function

Sub AddTimelineTitle(sld As Slide, title As String, dateRange As String, fontName As String)
    ' Add a professional title to the timeline
    Dim titleShape As Shape
    Set titleShape = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=170, Top:=10, width:=600, height:=30)
    With titleShape.TextFrame2
        .TextRange.Text = title
        .TextRange.Font.name = fontName
        .TextRange.Font.size = 18
        .TextRange.Font.Bold = True
        .TextRange.ParagraphFormat.alignment = ppAlignCenter
        .TextRange.Font.Fill.ForeColor.RGB = RGB(68, 114, 196)
        .VerticalAnchor = msoAnchorMiddle
    End With
    titleShape.Fill.Visible = msoFalse
    titleShape.Line.Visible = msoFalse
    
    ' Add subtitle with date range
    Dim subtitleShape As Shape
    Set subtitleShape = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=170, Top:=35, width:=600, height:=20)
    With subtitleShape.TextFrame2
        .TextRange.Text = dateRange
        .TextRange.Font.name = fontName
        .TextRange.Font.size = 12
        .TextRange.ParagraphFormat.alignment = ppAlignCenter
        .TextRange.Font.Fill.ForeColor.RGB = RGB(100, 100, 100)
        .VerticalAnchor = msoAnchorMiddle
    End With
    subtitleShape.Fill.Visible = msoFalse
    subtitleShape.Line.Visible = msoFalse
End Sub

Sub AddDateMarkers(sld As Slide, minDate As Date, maxDate As Date, scaleFactor As Double, _
                  leftPadding As Single, topY As Single, fontName As String)
    ' Add date markers along the timeline for reference
    Dim dateDiff As Long: dateDiff = maxDate - minDate
    Dim markerInterval As Long
    
    ' Determine appropriate interval based on timeline length
    If dateDiff <= 30 Then
        markerInterval = 7 ' Weekly markers for short timelines
    ElseIf dateDiff <= 180 Then
        markerInterval = 30 ' Monthly markers for medium timelines
    Else
        markerInterval = 90 ' Quarterly markers for long timelines
    End If
    
    Dim currentDate As Date: currentDate = minDate
    Do While currentDate <= maxDate
        Dim xPos As Single: xPos = leftPadding + (currentDate - minDate) * scaleFactor
        
        ' Draw marker line
        Dim markerShape As Shape
        Set markerShape = sld.Shapes.AddLine(xPos, topY - 10, xPos, topY + 10)
        With markerShape.Line
            .ForeColor.RGB = RGB(150, 150, 150)
            .Weight = 1
        End With
        
        ' Add date label
        Dim dateLabel As Shape
        Set dateLabel = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
            Left:=xPos - 25, Top:=topY - 35, width:=50, height:=20)
        With dateLabel.TextFrame2
            .TextRange.Text = Format(currentDate, "mmm-yy")
            .TextRange.Font.name = fontName
            .TextRange.Font.size = 8
            .TextRange.ParagraphFormat.alignment = ppAlignCenter
            .TextRange.Font.Fill.ForeColor.RGB = RGB(100, 100, 100)
        End With
        dateLabel.Fill.Visible = msoFalse
        dateLabel.Line.Visible = msoFalse
        
        currentDate = DateAdd("d", markerInterval, currentDate)
    Loop
End Sub

' === Enhanced Error Handling Function ===
Function ValidateTimelineData(milestoneData() As Variant) As Boolean
    ' Validate the timeline data for common issues
    Dim i As Integer
    Dim errorMessages As String
    
    For i = 0 To UBound(milestoneData)
        ' Check for missing task name
        If IsEmpty(milestoneData(i, 0)) Or Trim(CStr(milestoneData(i, 0))) = "" Then
            errorMessages = errorMessages & "Row " & (i + 2) & ": Missing task name" & vbCrLf
        End If
        
        ' Check for invalid start date
        If Not IsDate(milestoneData(i, 1)) Then
            errorMessages = errorMessages & "Row " & (i + 2) & ": Invalid start date" & vbCrLf
        End If
        
        ' Check for invalid type
        If Not (UCase(milestoneData(i, 3)) = "MILESTONE" Or UCase(milestoneData(i, 3)) = "FEATURE" Or UCase(milestoneData(i, 3)) = "PHASE") Then
            errorMessages = errorMessages & "Row " & (i + 2) & ": Type must be 'Milestone', 'Feature', or 'Phase'" & vbCrLf
        End If
        
        ' Check for features and phases without end dates
        If (UCase(milestoneData(i, 3)) = "FEATURE" Or UCase(milestoneData(i, 3)) = "PHASE") And Not IsDate(milestoneData(i, 2)) Then
            errorMessages = errorMessages & "Row " & (i + 2) & ": Feature and Phase events require an end date" & vbCrLf
        End If
    Next i
    
    If errorMessages <> "" Then
        MsgBox "Data validation errors found:" & vbCrLf & vbCrLf & errorMessages, vbCritical, "Timeline Data Validation"
        ValidateTimelineData = False
    Else
        ValidateTimelineData = True
    End If
End Function

Sub AddVerticalGridLines(sld As Slide, minDate As Date, maxDate As Date, scaleFactor As Double, leftPadding As Single, topY As Single, bottomY As Single)
    ' Add subtle vertical grid lines for better visual reference
    Dim currentDate As Date: currentDate = DateSerial(Year(minDate), Month(minDate), 1)
    
    Do While currentDate <= maxDate
        Dim xPos As Single: xPos = leftPadding + (currentDate - minDate) * scaleFactor
        
        ' Draw subtle vertical grid line
        Dim gridLine As Shape
        Set gridLine = sld.Shapes.AddLine(xPos, topY, xPos, bottomY)
        
        With gridLine.Line
            .ForeColor.RGB = RGB(240, 240, 240)
            .Weight = 0.5
            .DashStyle = msoLineDash
            .Transparency = 0.7
        End With
        gridLine.ZOrder msoSendToBack ' Send to back
        
        currentDate = DateAdd("m", 1, currentDate)
    Loop
End Sub

' ===================================================================
' ENHANCED LANE ASSIGNMENT HELPER FUNCTIONS
' ===================================================================

Sub SortEventsByStartDate(ByRef milestoneData() As Variant, ByRef eventLanes() As Integer)
    ' Sort events by start date for logical lane assignment processing
    Dim i As Integer, j As Integer
    Dim numEvents As Integer: numEvents = UBound(milestoneData) + 1
    
    ' Simple bubble sort by start date
    For i = 0 To numEvents - 2
        For j = i + 1 To numEvents - 1
            Dim startDate1 As Date, startDate2 As Date
            startDate1 = CDate(milestoneData(i, 1))
            startDate2 = CDate(milestoneData(j, 1))
            
            If startDate1 > startDate2 Then
                ' Swap events
                Call SwapEvents(milestoneData, i, j)
                ' Reset lane assignments since we changed order
                eventLanes(i) = 0
                eventLanes(j) = 0
            End If
        Next j
    Next i
End Sub

Sub SwapEvents(ByRef milestoneData() As Variant, index1 As Integer, index2 As Integer)
    ' Swap two events in the milestone data array
    Dim temp As Variant
    Dim col As Integer
    
    For col = 0 To 5
        temp = milestoneData(index1, col)
        milestoneData(index1, col) = milestoneData(index2, col)
        milestoneData(index2, col) = temp
    Next col
End Sub

Function GetEventEndDate(milestoneData() As Variant, eventIndex As Integer) As Date
    ' Get the end date for an event (start date for milestones)
    If IsDate(milestoneData(eventIndex, 2)) Then
        GetEventEndDate = CDate(milestoneData(eventIndex, 2))
    Else
        GetEventEndDate = CDate(milestoneData(eventIndex, 1)) ' Milestone uses start date
    End If
End Function

Sub MoveEventToHigherLane(ByRef eventLanes() As Integer, eventIndex As Integer, newLane As Integer)
    ' Move an event to a higher lane number
    eventLanes(eventIndex) = newLane
    
    ' Check if this creates new conflicts and recursively resolve them
    Dim i As Integer
    For i = 0 To UBound(eventLanes)
        If i <> eventIndex And eventLanes(i) = newLane Then
            ' Another event is already in this lane, move it up
            Call MoveEventToHigherLane(eventLanes, i, newLane + 1)
            Exit For
        End If
    Next i
End Sub

Sub CalculateEventExtendedBounds(baseStartX As Single, baseEndX As Single, eventType As String, _
                                eventLabel As String, ByRef extendedStartX As Single, ByRef extendedEndX As Single)
    ' Calculate the extended bounds of an event including all visual elements with ENHANCED LABELING SYSTEM:
    ' - Bar/milestone shape
    ' - Enhanced labels (name, duration, date range)
    ' - Consider all labels as single block for spacing calculations
    ' - Account for vertical space when name labels are positioned on top
    ' This ensures proper spacing between events considering all visual components
    
    ' Get configuration values for consistent behavior
    Dim config As TimelineConfig: config = GetDefaultTimelineConfig()
    
    Dim labelLength As Single
    Dim estimatedTextWidth As Single
    
    ' Estimate text width based on character count (approximate: 6 pixels per character)
    estimatedTextWidth = Len(eventLabel) * 6
    
    Select Case UCase(eventType)
        Case "MILESTONE"
            ' === MILESTONE BOUNDARIES WITH NEW POSITIONING RULES ===
            ' Milestones need space for:
            ' - Diamond shape (16px)
            ' - Label positioned on LEFT (preferred) or ON TOP (fallback)
            ' - Date label below milestone
            ' - Space calculations depend on positioning choice
            
            Dim milestoneRadius As Single: milestoneRadius = 8 ' Half of 16px diamond
            Dim labelBuffer As Single: labelBuffer = 120 ' Standard label width for left positioning
            Dim topLabelBuffer As Single: topLabelBuffer = 25 ' Reduced horizontal buffer when label is on top
            Dim dateBuffer As Single: dateBuffer = 15 ' Space below for date label
            
            ' Calculate if milestone would likely have label on left or top
            ' (This is approximate since we don't have leftPadding context here)
            Dim approximateLeftSpace As Single: approximateLeftSpace = baseStartX - 100 ' Estimate
            Dim estimatedLabelWidth As Single: estimatedLabelWidth = (Len(eventLabel) * 6) + 20 ' Approximate
            
            If approximateLeftSpace >= (estimatedLabelWidth + 23 + milestoneRadius) Then
                ' === LABEL LIKELY ON LEFT: Standard left-side buffer ===
                extendedStartX = baseStartX - milestoneRadius - labelBuffer
                extendedEndX = baseEndX + milestoneRadius + dateBuffer
            Else
                ' === LABEL LIKELY ON TOP: Reduced horizontal buffer, account for vertical space ===
                extendedStartX = baseStartX - milestoneRadius - topLabelBuffer
                extendedEndX = baseEndX + milestoneRadius + topLabelBuffer
                ' Note: Vertical space for top labels is handled by lane spacing, not horizontal bounds
            End If
            
        Case "FEATURE"
            ' === ENHANCED FEATURE BAR BOUNDARIES WITH VERTICAL SPACING ===
            ' Features need space for the ENHANCED LABELING SYSTEM:
            ' - Bar shape (baseStartX to baseEndX)
            ' - Name label (inside if bar > 80px, on top otherwise - REQUIRES VERTICAL SPACE)
            ' - Duration label on LEFT side (40px space needed)
            ' - Date range label on RIGHT side (100px space needed)
            ' - All labels considered as single block
            
            Dim barWidth As Single: barWidth = baseEndX - baseStartX
            ' Use TimelineConfig value for consistent behavior across all functions
            Const DurationLabelSpace As Single = 40  ' Space needed on left for "N d"
            Const DateRangeLabelSpace As Single = 100 ' Space needed on right for "dd mmm - dd mmm"
            
            ' === ALWAYS ACCOUNT FOR DURATION AND DATE RANGE LABELS ===
            ' Duration label is always on the left (40px)
            extendedStartX = baseStartX - DurationLabelSpace
            
            ' Date range label is always on the right (100px)
            extendedEndX = baseEndX + DateRangeLabelSpace
            
            ' === VERTICAL SPACE CONSIDERATION FOR NAME LABELS ON TOP ===
            ' When name label goes on top of bar, we need to account for vertical spacing
            ' This prevents overlap with bars in lanes above by increasing the effective "footprint"
            ' The lane assignment system will use this extended boundary to detect conflicts
            
            ' Calculate required width based on label text length
            Dim labelText As String: labelText = eventLabel ' Task name passed as parameter
            Dim labelWidth As Single: labelWidth = CalculateDynamicLabelWidth(labelText, config.fontSize, 30, 300) ' Min 30, Max 300
            Dim requiredWidth As Single: requiredWidth = labelWidth + config.labelInternalPadding ' Add 20px padding (10px each side)
            
            If barWidth < requiredWidth Then
                ' Name label goes on top - increase horizontal buffer to account for vertical space usage
                ' This creates more spacing between events when name labels are on top
                Const VerticalSpaceBuffer As Single = 25 ' Additional buffer for top-positioned labels
                extendedStartX = extendedStartX - VerticalSpaceBuffer
                extendedEndX = extendedEndX + VerticalSpaceBuffer
            End If
            
        Case "PHASE"
            ' === PHASE BAR BOUNDARIES ===
            ' Phases need space for:
            ' - Phase bar shape (baseStartX to baseEndX)
            ' - Label inside phase bar (no additional space)
            ' - Duration text inside bar (no additional space)
            
            ' Phase labels are inside bars, so no additional horizontal space
            extendedStartX = baseStartX
            extendedEndX = baseEndX
            
        Case Else
            ' === DEFAULT: Use basic buffer ===
            Dim defaultBuffer As Single: defaultBuffer = 20
            extendedStartX = baseStartX - defaultBuffer
            extendedEndX = baseEndX + defaultBuffer
    End Select
    
    ' Add minimum spacing buffer between any events
    Const MinimumEventSpacing As Single = 15
    extendedStartX = extendedStartX - MinimumEventSpacing
    extendedEndX = extendedEndX + MinimumEventSpacing
End Sub

' ===================================================================
' SWIMLANE DYNAMIC HEIGHT CALCULATION
' ===================================================================

Function CalculateSwimlaneActualHeight(events() As Variant, ByRef eventLanes() As Integer, config As TimelineConfig, _
                                      scaleFactor As Double, leftPadding As Single, minDate As Date) As Single
    ' Calculate the actual height needed for a swimlane based on the bottom edge of the last element + 10px
    ' This replaces the old lane-based calculation with content-based height calculation
    
    If IsEmpty(events) Then
        CalculateSwimlaneActualHeight = 0 ' Empty swimlanes collapse to 0 height
        Exit Function
    End If
    
    ' Perform lane assignment to determine actual positioning
    Dim totalLanes As Integer
    totalLanes = AssignLanesToEvents(events, eventLanes, scaleFactor, leftPadding, minDate)
    
    ' Find maximum lane number to size arrays properly
    Dim maxLane As Integer: maxLane = 0
    Dim i As Integer
    For i = 0 To UBound(eventLanes)
        If eventLanes(i) > maxLane Then maxLane = eventLanes(i)
    Next i
    
    ' Determine which lanes have labels on top (same logic as before)
    Dim lanesWithTopLabels() As Boolean
    ReDim lanesWithTopLabels(0 To maxLane)
    
    For i = 0 To UBound(events)
        If UCase(events(i, 3)) = "FEATURE" And IsDate(events(i, 2)) Then
            Dim startDateCheck As Date: startDateCheck = Int(events(i, 1))
            Dim endDateCheck As Date: endDateCheck = Int(events(i, 2))
            Dim barWidthCheck As Single: barWidthCheck = Abs(endDateCheck - startDateCheck) * scaleFactor
            
            ' Calculate required width based on label text length
            Dim labelText As String: labelText = CStr(events(i, 0)) ' Task name is in column A (index 0)
            Dim labelWidth As Single: labelWidth = CalculateDynamicLabelWidth(labelText, config.fontSize, 30, 300) ' Min 30, Max 300
            Dim requiredWidth As Single: requiredWidth = labelWidth + config.labelInternalPadding ' Add 20px padding (10px each side)
            
            ' Compare bar width with required space for label
            If barWidthCheck < requiredWidth Then
                If eventLanes(i) <= maxLane Then
                    lanesWithTopLabels(eventLanes(i)) = True
                End If
            End If
        ElseIf UCase(events(i, 3)) = "MILESTONE" Then
            ' Check if milestone will have label on top (insufficient left space)
            Dim milestoneStartXCheck As Single: milestoneStartXCheck = leftPadding + (Int(events(i, 1)) - minDate) * scaleFactor
            Dim availableLeftSpaceCheck As Single: availableLeftSpaceCheck = milestoneStartXCheck - leftPadding
            Dim estimatedLabelWidthCheck As Single: estimatedLabelWidthCheck = (Len(CStr(events(i, 0))) * 6) + 20
            Dim requiredLeftSpaceCheck As Single: requiredLeftSpaceCheck = estimatedLabelWidthCheck + 23 + 8
            
            If availableLeftSpaceCheck < requiredLeftSpaceCheck Then
                If eventLanes(i) <= maxLane Then
                    lanesWithTopLabels(eventLanes(i)) = True
                End If
            End If
        End If
    Next i
    
    ' Calculate the bottom-most position of all elements
    Dim maxBottomPosition As Single: maxBottomPosition = 0
    
    ' Calculate Y positions for each element and find the bottom-most position
    For i = 0 To UBound(events)
        Dim currentLane As Integer: currentLane = eventLanes(i)
        Dim elementY As Single: elementY = config.swimlaneContentPadding
        
        ' Calculate Y position with dynamic spacing (same logic as PlaceEventsInSwimlane)
        Dim laneIndex As Integer
        For laneIndex = 0 To currentLane
            If laneIndex <= maxLane And lanesWithTopLabels(laneIndex) Then
                elementY = elementY + config.laneSpacingWithTopLabels ' Top padding for lanes with top labels
            Else
                elementY = elementY + config.laneSpacingWithInsideLabels ' Top padding for lanes with inside labels
            End If
        Next laneIndex
        
        ' Calculate bottom position based on element type
        ' Default element height
        Dim elementBottomY As Single
        elementBottomY = elementY + CSng(config.elementHeight) / 2
        
        ' Track the maximum bottom position
        If elementBottomY > maxBottomPosition Then
            maxBottomPosition = elementBottomY
        End If
    Next i
       
    ' No minimum height constraints - swimlanes collapse completely to actual content size
    
    CalculateSwimlaneActualHeight = maxBottomPosition
End Function

Function CalculateSwimlaneRequiredLanes(events() As Variant, ByRef eventLanes() As Integer, config As TimelineConfig) As Integer
    ' Legacy function - now calculates equivalent lanes based on actual height for compatibility
    ' This maintains compatibility with existing code while using the new height-based calculation
    
    If IsEmpty(events) Then
        CalculateSwimlaneRequiredLanes = 1 ' Minimum lanes for empty swimlane
        Exit Function
    End If
    
    ' Extract config values
    Dim scaleFactor As Double: scaleFactor = 1 ' Placeholder - will be set by calling function
    Dim headerWidth As Single: headerWidth = config.SwimlaneHeaderWidth
    Dim minDate As Date: minDate = Date ' Placeholder - will be set by calling function
    
    ' Calculate actual height needed
    Dim actualHeight As Single
    actualHeight = CalculateSwimlaneActualHeight(events, eventLanes, config, scaleFactor, headerWidth, minDate)
    
    ' Convert to equivalent lanes for compatibility (minimum 1)
    Dim equivalentLanes As Integer
    equivalentLanes = Int(actualHeight / CSng(config.laneHeight)) + 1
    If equivalentLanes < 1 Then equivalentLanes = 1
    
    CalculateSwimlaneRequiredLanes = equivalentLanes
End Function

' ===================================================================
' MULTI-SLIDE SUPPORT FUNCTIONS
' ===================================================================

Function CalculateRequiredSlides(swimlaneOrg As SwimlaneOrganization, config As TimelineConfig) As Integer
    ' Calculate how many slides are needed based on actual swimlane content
    Dim totalRequiredHeight As Single
    Dim availableHeight As Single
    Dim currentY As Single
    
    ' Calculate available height for swimlane content (excluding calendar and phase areas)
    availableHeight = config.slideHeight - config.timelineTop - 30 ' 30px bottom margin
    
    ' Calculate total height needed for all swimlanes
    totalRequiredHeight = 0
    currentY = config.axisY
    Const SwimlonePadding As Single = 5
    
    Dim i As Integer
    For i = 0 To swimlaneOrg.Count - 1
        ' Calculate actual height based on content (NEW SYSTEM)
        Dim swimlaneHeight As Single
        If Not IsEmpty(swimlaneOrg.swimlaneEvents(i)) Then
            Dim tempEvents() As Variant: tempEvents = swimlaneOrg.swimlaneEvents(i)
            Dim tempEventLanes() As Integer
            ReDim tempEventLanes(0 To UBound(tempEvents))
            
            ' Extract values to avoid ByRef issues with UDT members
            Dim placeholderScaleFactor2 As Double: placeholderScaleFactor2 = 1
            Dim headerWidth2 As Single: headerWidth2 = config.SwimlaneHeaderWidth
            Dim placeholderMinDate2 As Date: placeholderMinDate2 = Date
            
            ' Use the new content-based height calculation
            swimlaneHeight = CalculateSwimlaneActualHeight(tempEvents, tempEventLanes, config, _
                placeholderScaleFactor2, headerWidth2, placeholderMinDate2)
        Else
            swimlaneHeight = 50 ' Minimum height for empty swimlane
        End If
        
        ' Ensure minimum height
        If swimlaneHeight < config.swimlaneHeight Then swimlaneHeight = config.swimlaneHeight
        
        totalRequiredHeight = totalRequiredHeight + swimlaneHeight + SwimlonePadding
    Next i
    
    ' Calculate number of slides needed
    If totalRequiredHeight <= availableHeight Then
        CalculateRequiredSlides = 1
    Else
        CalculateRequiredSlides = Int((totalRequiredHeight / availableHeight)) + 1
    End If
End Function

Sub CreateMultiSlideTimeline(config As TimelineConfig, dateRange As TimelineDateRange, _
                           swimlaneOrg As SwimlaneOrganization, timelineData() As Variant, requiredSlides As Integer)
    ' Create multiple slides with distributed swimlanes, duplicating calendar and phase sections on each slide
    
    Dim availableHeight As Single
    availableHeight = config.slideHeight - config.timelineTop - 30 ' Available height per slide
    
    ' Calculate swimlane heights for distribution
    Dim swimlaneHeights() As Single
    ReDim swimlaneHeights(0 To swimlaneOrg.Count - 1)
    Const SwimlonePadding As Single = 5
    
    Dim i As Integer
    For i = 0 To swimlaneOrg.Count - 1
        Dim requiredLanes As Integer: requiredLanes = 1
        If Not IsEmpty(swimlaneOrg.swimlaneEvents(i)) Then
            Dim tempEvents() As Variant: tempEvents = swimlaneOrg.swimlaneEvents(i)
            Dim tempEventLanes() As Integer
            ReDim tempEventLanes(0 To UBound(tempEvents))
            requiredLanes = CalculateSwimlaneRequiredLanes(tempEvents, tempEventLanes, config)
        End If
        
        swimlaneHeights(i) = (requiredLanes * config.laneHeight) + 40
        If swimlaneHeights(i) < config.swimlaneHeight Then swimlaneHeights(i) = config.swimlaneHeight
    Next i
    
    ' Distribute swimlanes across slides
    Dim currentSlide As Integer: currentSlide = 1
    Dim currentSlideHeight As Single: currentSlideHeight = 0
    Dim swimlaneStartIndex As Integer: swimlaneStartIndex = 0
    Dim actualSlidesCreated As Integer: actualSlidesCreated = 0
    
    For i = 0 To swimlaneOrg.Count - 1
        ' Check if current swimlane fits on current slide
        If currentSlideHeight + swimlaneHeights(i) + SwimlonePadding > availableHeight And i > swimlaneStartIndex Then
            ' Create slide for current batch of swimlanes
            Call CreateSingleSlideWithSwimlanes(config, dateRange, swimlaneOrg, timelineData, _
                swimlaneStartIndex, i - 1, currentSlide)
            actualSlidesCreated = actualSlidesCreated + 1
            
            ' Start new slide
            currentSlide = currentSlide + 1
            swimlaneStartIndex = i
            currentSlideHeight = swimlaneHeights(i) + SwimlonePadding
        Else
            ' Add to current slide
            currentSlideHeight = currentSlideHeight + swimlaneHeights(i) + SwimlonePadding
        End If
    Next i
    
    ' Create final slide with remaining swimlanes
    If swimlaneStartIndex <= swimlaneOrg.Count - 1 Then
        Call CreateSingleSlideWithSwimlanes(config, dateRange, swimlaneOrg, timelineData, _
            swimlaneStartIndex, swimlaneOrg.Count - 1, currentSlide)
        actualSlidesCreated = actualSlidesCreated + 1
    End If
    
    ' Debug message with actual slides created count
    Debug.Print Format(Now, "dd-mmm-yyyy hh:mm:ss") & "> Timeline generation completed successfully - " & actualSlidesCreated & " slides created with " & swimlaneOrg.Count & " swimlanes distributed across slides"
End Sub

Sub CreateSingleSlideWithSwimlanes(config As TimelineConfig, dateRange As TimelineDateRange, _
                                  swimlaneOrg As SwimlaneOrganization, timelineData() As Variant, _
                                  startSwimlaneIndex As Integer, endSwimlaneIndex As Integer, slideNumber As Integer)
    ' Create a single slide with specified range of swimlanes, including calendar and phase sections
    
    ' Create new slide
    Dim sld As Slide
    Set sld = CreateTimelineSlide()
    
    ' Calculate scale factor
    dateRange.scaleFactor = (config.slideWidth - config.SwimlaneHeaderWidth - config.axisPadding) / _
                           (dateRange.maxDate - dateRange.minDate)
    
    ' Store values in local variables
    Dim minDate As Date: minDate = dateRange.minDate
    Dim maxDate As Date: maxDate = dateRange.maxDate
    Dim scaleFactor As Double: scaleFactor = dateRange.scaleFactor
    Dim headerWidth As Single: headerWidth = config.SwimlaneHeaderWidth
    Dim timelineTop As Single: timelineTop = config.TimelineAxisTop
    Dim fontName As String: fontName = config.fontName
    
    ' === DUPLICATE CALENDAR SECTION ===
    Call DrawEnhancedTopTimelineAxis(sld, minDate, maxDate, scaleFactor, headerWidth, timelineTop, fontName)
    
    ' === DUPLICATE PHASES SECTION ===
    Call RenderPhasesInDedicatedArea(sld, config, dateRange, timelineData)
    
    ' === RENDER SUBSET OF SWIMLANES ===
    Call RenderSwimlaneSubset(sld, config, swimlaneOrg, startSwimlaneIndex, endSwimlaneIndex)
    Call RenderSwimlaneEventsSubset(sld, config, dateRange, swimlaneOrg, startSwimlaneIndex, endSwimlaneIndex)
    
    ' Add slide number indicator if multiple slides
    If slideNumber > 1 Then
        Call AddSlideNumberIndicator(sld, slideNumber, fontName)
    End If
End Sub

Sub RenderSwimlaneSubset(sld As Slide, config As TimelineConfig, swimlaneOrg As SwimlaneOrganization, _
                        startIndex As Integer, endIndex As Integer)
    ' Render swimlane headers and backgrounds for a subset of swimlanes
    
    Dim axisY As Single: axisY = config.axisY
    Dim baseSwimlaneHeight As Integer: baseSwimlaneHeight = config.swimlaneHeight
    Dim fontName As String: fontName = config.fontName
    Dim headerWidth As Single: headerWidth = config.SwimlaneHeaderWidth
    Dim slideWidth As Single: slideWidth = config.slideWidth
    Dim axisPadding As Integer: axisPadding = config.axisPadding
    Dim laneHeight As Integer: laneHeight = config.laneHeight
    
    Dim currentY As Single: currentY = axisY
    Const SwimlonePadding As Single = 5
    
    Dim i As Integer
    For i = startIndex To endIndex
        ' Calculate required lanes for this swimlane
        Dim requiredLanes As Integer: requiredLanes = 1
        If Not IsEmpty(swimlaneOrg.swimlaneEvents(i)) Then
            Dim tempEvents() As Variant: tempEvents = swimlaneOrg.swimlaneEvents(i)
            Dim tempEventLanes() As Integer
            ReDim tempEventLanes(0 To UBound(tempEvents))
            requiredLanes = CalculateSwimlaneRequiredLanes(tempEvents, tempEventLanes, config)
        End If
        
        ' Calculate dynamic height for this swimlane
        Dim dynamicSwimlaneHeight As Single
        dynamicSwimlaneHeight = (requiredLanes * laneHeight) + 40
        If dynamicSwimlaneHeight < baseSwimlaneHeight Then dynamicSwimlaneHeight = baseSwimlaneHeight
        
        ' Enhanced swimlane header with matching height and vertical centering
        Call AddEnhancedSwimlaneHeader(sld, 10, currentY - 1.5, _
            swimlaneOrg.swimlanes(i), fontName, 11, dynamicSwimlaneHeight)
        
        ' Dynamic background size based on actual content - EXTENDED BY 25PX LEFT AND RIGHT
        Call DrawSwimlaneBackground(sld, headerWidth - 25, currentY, _
            slideWidth - headerWidth - axisPadding + 50, dynamicSwimlaneHeight, i - startIndex)
        
        ' Move to next swimlane position with padding
        currentY = currentY + dynamicSwimlaneHeight + SwimlonePadding
    Next i
End Sub

Sub RenderSwimlaneEventsSubset(sld As Slide, config As TimelineConfig, dateRange As TimelineDateRange, _
                              swimlaneOrg As SwimlaneOrganization, startIndex As Integer, endIndex As Integer)
    ' Render events for a subset of swimlanes
    
    Dim scaleFactor As Double: scaleFactor = dateRange.scaleFactor
    Dim minDate As Date: minDate = dateRange.minDate
    Dim headerWidth As Single: headerWidth = config.SwimlaneHeaderWidth
    Dim fontName As String: fontName = config.fontName
    Dim circleSize As Integer: circleSize = config.circleSize
    Dim elementHeight As Integer: elementHeight = config.elementHeight
    Dim laneHeight As Integer: laneHeight = config.laneHeight
    Dim axisY As Single: axisY = config.axisY
    Dim baseSwimlaneHeight As Integer: baseSwimlaneHeight = config.swimlaneHeight
    
    Dim currentY As Single: currentY = axisY
    Const SwimlonePadding As Single = 5
    
    Dim i As Integer
    For i = startIndex To endIndex
        Dim currentEvents() As Variant: currentEvents = swimlaneOrg.swimlaneEvents(i)
        
        If Not IsEmpty(currentEvents) Then
            ' Detect overlapping events and assign lanes
            Dim eventLanes() As Integer
            ReDim eventLanes(0 To UBound(currentEvents))
            Dim totalLanes As Integer
            totalLanes = AssignLanesToEvents(currentEvents, eventLanes, scaleFactor, headerWidth, minDate)
            
            ' Place events with enhanced styling using dynamic Y position
            Call PlaceEventsInSwimlane(sld, currentEvents, eventLanes, currentY, _
                scaleFactor, headerWidth, minDate, fontName, circleSize, elementHeight, laneHeight)
        End If
        
        ' Calculate dynamic height for this swimlane to get next position
        Dim requiredLanes As Integer: requiredLanes = 1
        If Not IsEmpty(currentEvents) Then
            requiredLanes = CalculateSwimlaneRequiredLanes(currentEvents, eventLanes, config)
        End If
        
        Dim dynamicSwimlaneHeight As Single
        dynamicSwimlaneHeight = (requiredLanes * laneHeight) + 40
        If dynamicSwimlaneHeight < baseSwimlaneHeight Then dynamicSwimlaneHeight = baseSwimlaneHeight
        
        ' Move to next swimlane position with padding
        currentY = currentY + dynamicSwimlaneHeight + SwimlonePadding
    Next i
End Sub

Sub AddSlideNumberIndicator(sld As Slide, slideNumber As Integer, fontName As String)
    ' Add slide number indicator for multi-slide timelines
    Dim slideIndicator As Shape
    Set slideIndicator = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=10, Top:=10, width:=80, height:=20)
    With slideIndicator.TextFrame2
        .TextRange.Text = "Slide " & slideNumber
        .TextRange.Font.name = fontName
        .TextRange.Font.size = 10
        .TextRange.Font.Bold = True
        .TextRange.Font.Fill.ForeColor.RGB = RGB(100, 100, 100)
        .TextRange.ParagraphFormat.alignment = ppAlignLeft
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 3
        .MarginRight = 3
        .MarginTop = 2
        .MarginBottom = 2
    End With
    slideIndicator.Fill.Visible = msoFalse
    slideIndicator.Line.Visible = msoFalse
End Sub

' ===================================================================
' SLIDE LAYOUT CONFIGURATION HELPER
' ===================================================================
Sub ApplyCustomSlideLayout(sld As Slide, layoutName As String)
    ' Apply custom slide layout by name with error handling and fallback
    ' This allows users to specify their preferred slide layout in the configuration
    
    On Error GoTo LayoutError
    
    ' Search for the layout by name in the slide master
    Dim customLayout As customLayout
    Dim i As Integer
    Dim layoutFound As Boolean: layoutFound = False
    
    For i = 1 To ActivePresentation.SlideMaster.CustomLayouts.Count
        If LCase(Trim(ActivePresentation.SlideMaster.CustomLayouts(i).name)) = LCase(Trim(layoutName)) Then
            Set customLayout = ActivePresentation.SlideMaster.CustomLayouts(i)
            layoutFound = True
            Exit For
        End If
    Next i
    
    If layoutFound Then
        ' Apply the found layout
        sld.customLayout = customLayout
        Debug.Print sld.name & " - Applied custom layout: " & layoutName
    Else
        ' Layout not found - use fallback
        Debug.Print "WARNING: Layout '" & layoutName & "' not found. Using default layout."
        GoTo LayoutError
    End If
    
    Exit Sub
    
LayoutError:
    ' Fallback to default layout on any error
    On Error Resume Next
    sld.customLayout = ActivePresentation.SlideMaster.CustomLayouts(1)
    On Error GoTo 0
    Debug.Print sld.name & " - Applied fallback layout due to error with: " & layoutName
End Sub
