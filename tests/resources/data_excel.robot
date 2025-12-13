*** Settings ***
Documentation     Keywords for working with Excel data files
Resource          common.robot
Library           ../libraries/ExcelValidator.py
Library           Collections
Library           OperatingSystem

*** Variables ***
${EXCEL_FILE_PATH}    %{EXCEL_FILE_PATH=../timeline.xlsx}
${EXCEL_SHEET_NAME}   %{EXCEL_SHEET_NAME=TimelineData}

*** Keywords ***
Read Excel Data
    [Documentation]    Read data from Excel file using openpyxl
    [Arguments]    ${file_path}=${EXCEL_FILE_PATH}    ${sheet_name}=${EXCEL_SHEET_NAME}
    ${validator}=    Get Library Instance    ExcelValidator
    ${data}=    Call Method    ${validator}    read_excel_data    ${file_path}    ${sheet_name}
    RETURN    ${data}

Validate Excel Data
    [Documentation]    Validate required columns exist in Excel file
    [Arguments]    ${file_path}=${EXCEL_FILE_PATH}    ${sheet_name}=${EXCEL_SHEET_NAME}
    ${validator}=    Get Library Instance    ExcelValidator
    ${is_valid}    ${errors}=    Call Method    ${validator}    validate_excel_file    ${file_path}    ${sheet_name}
    IF    not ${is_valid}
        ${error_msg}=    Evaluate    '\\n'.join($errors)
        Fail    Excel validation failed:\n${error_msg}
    END
    Log    Excel file validation passed    level=INFO

Get Excel Data Summary
    [Documentation]    Get summary statistics about Excel data
    [Arguments]    ${file_path}=${EXCEL_FILE_PATH}    ${sheet_name}=${EXCEL_SHEET_NAME}
    ${validator}=    Get Library Instance    ExcelValidator
    ${summary}=    Call Method    ${validator}    get_data_summary    ${file_path}    ${sheet_name}
    RETURN    ${summary}

Verify Excel File Exists
    [Documentation]    Verify that Excel file exists
    [Arguments]    ${file_path}=${EXCEL_FILE_PATH}
    File Should Exist    ${file_path}    Excel file not found: ${file_path}
    Log    Excel file found: ${file_path}    level=INFO

Verify Sheet Exists
    [Documentation]    Verify that sheet exists in Excel file
    [Arguments]    ${file_path}=${EXCEL_FILE_PATH}    ${sheet_name}=${EXCEL_SHEET_NAME}
    ${validator}=    Get Library Instance    ExcelValidator
    ${sheets}=    Call Method    ${validator}    get_sheet_names    ${file_path}
    ${exists}=    Evaluate    '${sheet_name}' in $sheets
    IF    not ${exists}
        Fail    Sheet '${sheet_name}' not found in ${file_path}
    END
    Log    Sheet '${sheet_name}' found    level=INFO
