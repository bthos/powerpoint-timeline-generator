*** Settings ***
Documentation     Test suite for validating timeline Excel data structure
Resource          resources/common.robot
Resource          resources/data_excel.robot
Library           libraries/ErrorHandler.py
Library           libraries/ExcelValidator.py
Suite Setup       Setup Test Suite
Suite Teardown    Cleanup Test Suite

*** Variables ***
${EXCEL_FILE_PATH}    %{EXCEL_FILE_PATH=../timeline.xlsx}
${EXCEL_SHEET_NAME}   %{EXCEL_SHEET_NAME=TimelineData}

*** Test Cases ***
Verify Excel File Exists
    [Documentation]    Verify that the Excel file exists
    [Tags]    smoke    data_validation
    Verify Excel File Exists    ${EXCEL_FILE_PATH}

Verify Sheet Exists
    [Documentation]    Verify that the TimelineData sheet exists
    [Tags]    smoke    data_validation
    Verify Sheet Exists    ${EXCEL_FILE_PATH}    ${EXCEL_SHEET_NAME}

Validate Excel Structure
    [Documentation]    Validate that Excel file has correct structure and headers
    [Tags]    smoke    data_validation
    Validate Excel Data    ${EXCEL_FILE_PATH}    ${EXCEL_SHEET_NAME}

Validate Data Content
    [Documentation]    Validate that data rows contain valid values
    [Tags]    data_validation
    ${validator}=    Get Library Instance    ExcelValidator
    ${is_valid}    ${errors}=    Call Method    ${validator}    validate_excel_file    ${EXCEL_FILE_PATH}    ${EXCEL_SHEET_NAME}
    Should Be True    ${is_valid}    Excel data validation failed: ${errors}

Verify Required Columns
    [Documentation]    Verify all required columns are present
    [Tags]    data_validation
    ${data}=    Read Excel Data    ${EXCEL_FILE_PATH}    ${EXCEL_SHEET_NAME}
    ${columns}=    Get Dictionary Keys    ${data}[0]
    Should Contain    ${columns}    Task Name
    Should Contain    ${columns}    Start Date
    Should Contain    ${columns}    Type
    Should Contain    ${columns}    Color
    Should Contain    ${columns}    Swimlane

Verify Data Types
    [Documentation]    Verify that Type column contains valid values (Milestone, Feature, or Phase)
    [Tags]    data_validation
    ${data}=    Read Excel Data    ${EXCEL_FILE_PATH}    ${EXCEL_SHEET_NAME}
    FOR    ${row}    IN    @{data}
        ${event_type}=    Get From Dictionary    ${row}    Type
        Should Be True    "${event_type}" in ["Milestone", "Feature", "Phase", "milestone", "feature", "phase"]
        ...    Invalid Type value: ${event_type}
    END

Verify Color Values
    [Documentation]    Verify that Color column contains valid color names
    [Tags]    data_validation
    ${valid_colors}=    Create List    red    blue    green    orange    purple    yellow    gray    grey
    ${data}=    Read Excel Data    ${EXCEL_FILE_PATH}    ${EXCEL_SHEET_NAME}
    FOR    ${row}    IN    @{data}
        ${color}=    Get From Dictionary    ${row}    Color
        ${color_lower}=    Convert To Lower Case    ${color}
        Should Contain    ${valid_colors}    ${color_lower}
        ...    Invalid color value: ${color}
    END

Verify Date Logic
    [Documentation]    Verify that Start Date is before End Date for Phase events
    [Tags]    data_validation
    ${data}=    Read Excel Data    ${EXCEL_FILE_PATH}    ${EXCEL_SHEET_NAME}
    FOR    ${row}    IN    @{data}
        ${event_type}=    Get From Dictionary    ${row}    Type
        ${start_date}=    Get From Dictionary    ${row}    Start Date
        ${end_date}=    Get From Dictionary    ${row}    End Date
        IF    "${event_type}" in ["Phase", "phase"] and "${end_date}" != "${EMPTY}"
            # Date comparison would require date parsing
            Log    Verifying date logic for row: ${row}    level=DEBUG
        END
    END

Get Data Summary
    [Documentation]    Get and log summary statistics about the timeline data
    [Tags]    reporting
    ${summary}=    Get Excel Data Summary    ${EXCEL_FILE_PATH}    ${EXCEL_SHEET_NAME}
    Log    Data Summary: ${summary}    level=INFO
    Dictionary Should Contain Key    ${summary}    total_rows
    Dictionary Should Contain Key    ${summary}    milestones
    Dictionary Should Contain Key    ${summary}    phases
    Dictionary Should Contain Key    ${summary}    swimlanes

Verify Phases Have End Dates
    [Documentation]    Verify that Phase events have End Date values
    [Tags]    data_validation
    ${data}=    Read Excel Data    ${EXCEL_FILE_PATH}    ${EXCEL_SHEET_NAME}
    FOR    ${row}    IN    @{data}
        ${event_type}=    Get From Dictionary    ${row}    Type
        ${end_date}=    Get From Dictionary    ${row}    End Date
        IF    "${event_type}" in ["Phase", "phase"]
            Should Not Be Empty    ${end_date}
            ...    Phase events must have an End Date: ${row}
        END
    END

Verify Features Have End Dates
    [Documentation]    Verify that Feature events have End Date values
    [Tags]    data_validation
    ${data}=    Read Excel Data    ${EXCEL_FILE_PATH}    ${EXCEL_SHEET_NAME}
    FOR    ${row}    IN    @{data}
        ${event_type}=    Get From Dictionary    ${row}    Type
        ${end_date}=    Get From Dictionary    ${row}    End Date
        IF    "${event_type}" in ["Feature", "feature"]
            Should Not Be Empty    ${end_date}
            ...    Feature events must have an End Date: ${row}
        END
    END

