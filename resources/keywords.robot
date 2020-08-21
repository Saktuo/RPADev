*** Settings ***
# The libraries and variable files that are used for this automation

Library     RPA.Excel.Application
Library	    RPA.Excel.Files
Library     RPA.Desktop.Windows
Library     RPA.Browser
Library     String
Library     Collections
Library     Process
Variables   variables.py

# +
*** Keyword ***
Go to Wapice Family
    Open Browser    ${WAPICE_FAMILY_URL}    Chrome
    
    
Wait Until Success Filtering Team Leaders
    [Documentation]    Using an error management keyword 'Wait Until Keyword Succeeds' that runs a selected keyword and if it fails,
    ...                it waits 'GLOBAL_RETRY_INTERVAL' amount of time and runs the keyword again this is repeated 'GLOBAL_RETRY_AMOUNT' times.
    Wait Until Keyword Succeeds    ${GLOBAL_RETRY_AMOUNT}    ${GLOBAL_RETRY_INTERVAL}    Filter Team Leaders
    

Filter Team Leaders
    Wait Until Page Contains Element    xpath://*[@id="team-leader"]
    ${ELEMENT}=    Get WebElement    xpath://*[@id="team-leader"]
    Click Element    ${ELEMENT}
    Wait Until Page Contains Element    xpath://*[@id="counter"]/span
        
        
Wait Until Success Selectingt The First Team Leader
    ${COUNTER}=    Wait Until Keyword Succeeds    ${GLOBAL_RETRY_AMOUNT}    ${GLOBAL_RETRY_INTERVAL}    Select The First Team Leader
    Log    ${COUNTER}
    [Return]    ${COUNTER}
        
        
Select The First Team Leader
    ${RESULT}=    Get Text    xpath://*[@id="counter"]/span
    Log    ${RESULT}
    ${COUNTER}=    Fetch From Left    ${RESULT}    results
    Convert To Integer    ${COUNTER}
    Click Element    xpath://*[@id="employees"]/div[not(contains(@class, "isotope-hidden"))]
    [Return]    ${COUNTER}
    
    
Loop Team Leaders
    [Arguments]    ${COUNTER}
    ${emails}=    Create List
    Log    ${COUNTER}
    Append To List    ${emails}    Email
    ${table}=    Create Workbook    ${CURDIR}\\..\\WF_emails.xlsx
    FOR    ${i}    IN RANGE    1    ${COUNTER}+1
        Wait Until Success With Reading Email    ${emails}
    END
    Log List    ${emails}
    Append Rows To Worksheet    ${emails}
    Save Workbook
    Close Browser
    
    
Wait Until Success With Reading Email
    [Arguments]    ${emails}
    Wait Until Keyword Succeeds    ${GLOBAL_RETRY_AMOUNT}    ${GLOBAL_RETRY_INTERVAL}    Save Email Addresses    ${emails}
    [Return]    ${emails}
    
    
Save Email Addresses
    [Arguments]    ${emails}
    Wait Until Element Is Visible    css:.person_info .info .employee_email
    ${RECIPIENT}=    Get Text    css:.person_info .info .employee_email
    ${Length}=    Get Length    ${RECIPIENT}
    Should Be True    ${Length} > 1
    Log    ${Length}
    Append To List    ${emails}    ${RECIPIENT}
    Click Element    css:button.nextBtn.ui-button.ui-widget.ui-state-default.ui-corner-all.ui-button-text-only
    Wait Until Page Contains Element    xpath://*[@id="info_dialog"]/div[1]/div/img
    Log    ${RECIPIENT}
    [Return]    ${emails}
    

Run PowerShell Script
    #Open Using Run Dialog    powershell    Windows PowerShell
    #Send Keys    & {SPACE} ${POWERSHELL_SCRIPT} {ENTER}
    Run Process       Powershell.exe     ${CURDIR}\\send_email.ps1
    Sleep    3

    
    