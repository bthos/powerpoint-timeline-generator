*** Settings ***
Documentation     Keywords for PowerPoint presentation validation
Resource          common.robot
Library           ../libraries/PowerPointValidator.py
Library           ../libraries/ExcelValidator.py
Library           Collections
Library           OperatingSystem

*** Variables ***
${PPTX_FILE_PATH}    %{PPTX_FILE_PATH=}
${BASELINE_PPTX}     ${CURDIR}/../data/baseline/expected_timeline.pptx

*** Keywords ***
Load PowerPoint Presentation
    [Documentation]    Load a PowerPoint presentation for validation
    [Arguments]    ${file_path}
    ${validator}=    Get Library Instance    PowerPointValidator
    ${result}=    Call Method    ${validator}    load_presentation    ${file_path}
    Should Be True    ${result}    Failed to load presentation: ${file_path}
    Log    Loaded presentation: ${file_path}    level=INFO

Close PowerPoint Presentation
    [Documentation]    Close the current PowerPoint presentation
    ${validator}=    Get Library Instance    PowerPointValidator
    Call Method    ${validator}    close_presentation

Get Slide Count
    [Documentation]    Get the number of slides in the presentation
    ${validator}=    Get Library Instance    PowerPointValidator
    ${count}=    Call Method    ${validator}    get_slide_count
    RETURN    ${count}

Get Slide Dimensions
    [Documentation]    Get slide dimensions in points
    ${validator}=    Get Library Instance    PowerPointValidator
    ${dimensions}=    Call Method    ${validator}    get_slide_dimensions
    RETURN    ${dimensions}

Get Shapes On Slide
    [Documentation]    Get all shapes on a slide
    [Arguments]    ${slide_index}=0
    ${validator}=    Get Library Instance    PowerPointValidator
    ${shapes}=    Call Method    ${validator}    get_shapes_on_slide    ${slide_index}
    RETURN    ${shapes}

Find Shapes By Text
    [Documentation]    Find shapes containing specific text
    [Arguments]    ${search_text}    ${slide_index}=0    ${partial_match}=${TRUE}
    ${validator}=    Get Library Instance    PowerPointValidator
    ${shapes}=    Call Method    ${validator}    find_shapes_by_text    ${search_text}    ${slide_index}    ${partial_match}
    RETURN    ${shapes}

Validate Timeline Structure
    [Documentation]    Validate the overall timeline structure on a slide
    [Arguments]    ${slide_index}=0
    ${validator}=    Get Library Instance    PowerPointValidator
    ${results}=    Call Method    ${validator}    validate_timeline_structure    ${slide_index}
    RETURN    ${results}

Validate Swimlanes Present
    [Documentation]    Validate that expected swimlanes are present (checks all slides by default)
    [Arguments]    ${expected_swimlanes}    ${slide_index}=${-1}
    ${validator}=    Get Library Instance    PowerPointValidator
    ${results}=    Call Method    ${validator}    validate_swimlanes    ${expected_swimlanes}    ${slide_index}
    IF    not ${results}[valid]
        Fail    Missing swimlanes: ${results}[missing]
    END
    RETURN    ${results}

Validate Events Present
    [Documentation]    Validate that expected events are present (checks all slides by default)
    [Arguments]    ${expected_events}    ${slide_index}=${-1}
    ${validator}=    Get Library Instance    PowerPointValidator
    ${results}=    Call Method    ${validator}    validate_events    ${expected_events}    ${slide_index}
    IF    not ${results}[valid]
        Fail    Missing events: ${results}[missing]
    END
    RETURN    ${results}

Get Presentation Summary
    [Documentation]    Get a summary of the presentation structure
    ${validator}=    Get Library Instance    PowerPointValidator
    ${summary}=    Call Method    ${validator}    get_presentation_summary
    RETURN    ${summary}

Compare With Baseline
    [Documentation]    Compare current presentation with a baseline
    [Arguments]    ${baseline_path}=${BASELINE_PPTX}    ${slide_index}=0
    ${validator}=    Get Library Instance    PowerPointValidator
    ${results}=    Call Method    ${validator}    compare_with_baseline    ${baseline_path}    ${slide_index}
    RETURN    ${results}

Verify Slide Created
    [Documentation]    Verify that at least one slide was created
    ${count}=    Get Slide Count
    Should Be True    ${count} > 0    No slides found in presentation

Verify Slide Dimensions
    [Documentation]    Verify slide has expected dimensions (16:9 aspect ratio)
    [Arguments]    ${expected_width}=960    ${expected_height}=540
    ${dimensions}=    Get Slide Dimensions
    ${width_ok}=    Evaluate    abs(${dimensions}[width] - ${expected_width}) < 10
    ${height_ok}=    Evaluate    abs(${dimensions}[height] - ${expected_height}) < 10
    Should Be True    ${width_ok}    Slide width ${dimensions}[width] does not match expected ${expected_width}
    Should Be True    ${height_ok}    Slide height ${dimensions}[height] does not match expected ${expected_height}

Verify Shape Count
    [Documentation]    Verify minimum number of shapes on slide
    [Arguments]    ${min_shapes}=5    ${slide_index}=0
    ${shapes}=    Get Shapes On Slide    ${slide_index}
    ${count}=    Get Length    ${shapes}
    Should Be True    ${count} >= ${min_shapes}    Expected at least ${min_shapes} shapes, found ${count}

Verify Text Present On Slide
    [Documentation]    Verify specific text is present on the slide
    [Arguments]    ${text}    ${slide_index}=0
    ${shapes}=    Find Shapes By Text    ${text}    ${slide_index}
    ${count}=    Get Length    ${shapes}
    Should Be True    ${count} > 0    Text "${text}" not found on slide ${slide_index}

Verify Timeline Axis Exists
    [Documentation]    Verify timeline axis exists at expected position
    [Arguments]    ${expected_y}=110    ${tolerance}=20    ${slide_index}=0
    ${validator}=    Get Library Instance    PowerPointValidator
    ${shapes}=    Call Method    ${validator}    find_shapes_by_position    ${expected_y}    ${tolerance}    ${slide_index}
    ${count}=    Get Length    ${shapes}
    Should Be True    ${count} > 0    No shapes found at timeline axis position (Y=${expected_y})

Load Excel And Get Expected Swimlanes
    [Documentation]    Load Excel data and extract expected swimlane names
    [Arguments]    ${excel_path}    ${sheet_name}=TimelineData
    ${validator}=    Get Library Instance    ExcelValidator
    ${summary}=    Call Method    ${validator}    get_data_summary    ${excel_path}    ${sheet_name}
    ${swimlanes}=    Get From Dictionary    ${summary}    swimlanes
    RETURN    ${swimlanes}

Load Excel And Get Expected Events
    [Documentation]    Load Excel data and get event list
    [Arguments]    ${excel_path}    ${sheet_name}=TimelineData
    ${validator}=    Get Library Instance    ExcelValidator
    ${data}=    Call Method    ${validator}    read_excel_data    ${excel_path}    ${sheet_name}
    RETURN    ${data}

Validate PowerPoint Against Excel Data
    [Documentation]    Full validation of PowerPoint against Excel source data
    [Arguments]    ${pptx_path}    ${excel_path}    ${sheet_name}=TimelineData
    
    # Load presentation
    Load PowerPoint Presentation    ${pptx_path}
    
    # Verify basic structure
    Verify Slide Created
    Verify Slide Dimensions
    
    # Get expected data from Excel
    ${expected_swimlanes}=    Load Excel And Get Expected Swimlanes    ${excel_path}    ${sheet_name}
    ${expected_events}=    Load Excel And Get Expected Events    ${excel_path}    ${sheet_name}
    
    # Validate swimlanes (if any exist in Excel)
    ${swimlane_count}=    Get Length    ${expected_swimlanes}
    IF    ${swimlane_count} > 0
        ${swimlane_results}=    Validate Swimlanes Present    ${expected_swimlanes}
        Log    Swimlane validation: ${swimlane_results}    level=INFO
    END
    
    # Validate events
    ${event_results}=    Validate Events Present    ${expected_events}
    Log    Event validation: ${event_results}    level=INFO
    
    # Close presentation
    Close PowerPoint Presentation
    
    RETURN    ${event_results}

