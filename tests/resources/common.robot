*** Settings ***
Documentation     Common keywords and settings for timeline generator tests
Library           ../libraries/ErrorHandler.py
Library           ../libraries/ExcelValidator.py
Library           Collections
Library           String
Library           OperatingSystem

*** Variables ***
${EXCEL_FILE_PATH}    %{EXCEL_FILE_PATH=../timeline.xlsx}
${EXCEL_SHEET_NAME}   %{EXCEL_SHEET_NAME=TimelineData}
${LOG_LEVEL}          %{LOG_LEVEL=INFO}
${REPORTS_DIR}        ../reports
${LOGS_DIR}           ../logs

*** Keywords ***
Setup Test Environment
    [Documentation]    Initialize test environment and load configuration
    Create Directory    ${REPORTS_DIR}
    Create Directory    ${LOGS_DIR}
    Load Environment Variables
    Log    Test environment initialized    level=${LOG_LEVEL}

Load Environment Variables
    [Documentation]    Load environment variables from .env file if it exists
    ${env_exists}=    Run Keyword And Return Status    File Should Exist    ../.env
    IF    ${env_exists}
        ${env_content}=    Get File    ../.env
        Log    Environment variables loaded from .env    level=DEBUG
    ELSE
        Log    No .env file found, using default values    level=DEBUG
    END

Take Screenshot On Failure
    [Documentation]    Take screenshot when test fails (placeholder for future PowerPoint automation)
    [Arguments]    ${test_name}
    Log    Screenshot would be taken for: ${test_name}    level=DEBUG

Log Audit Entry
    [Documentation]    Log audit entry with action, status, and details
    [Arguments]    ${action}    ${status}=info    ${details}=${EMPTY}
    ${handler}=    Get Library Instance    ErrorHandler
    Call Method    ${handler}    log_audit_entry    ${action}    ${status}    ${details}

Wait And Click Element
    [Documentation]    Wait for element and click with retry (placeholder for UI automation)
    [Arguments]    ${selector}    ${timeout}=5s
    Log    Would click element: ${selector}    level=DEBUG

Wait And Fill Text
    [Documentation]    Wait for element and fill text (placeholder for UI automation)
    [Arguments]    ${selector}    ${text}    ${timeout}=5s
    Log    Would fill text: ${text} into ${selector}    level=DEBUG

Test Teardown With Screenshot
    [Documentation]    Teardown that takes screenshot on failure
    Run Keyword If Test Failed    Take Screenshot On Failure    ${TEST_NAME}
    Save Audit Log

Save Audit Log
    [Documentation]    Save audit log to file
    ${handler}=    Get Library Instance    ErrorHandler
    ${log_path}=    Call Method    ${handler}    save_audit_log
    Log    Audit log saved to: ${log_path}    level=INFO

Setup Test Suite
    [Documentation]    Setup actions for test suite
    Setup Test Environment
    Log    Test suite setup completed    level=INFO

Cleanup Test Suite
    [Documentation]    Cleanup actions for test suite
    Save Audit Log
    Log    Test suite cleanup completed    level=INFO

