*** Settings ***
Documentation   An automation that fetches Wapice team leaders' email from Wapice Family and sends a reminder to them with Power Shell script
Resource        keywords.robot
Library         RPA.Browser
Library         RPA.Desktop.Windows
Variables       variables.py
Library         Collections
Library         RPA.Excel.Files
Library         String

*** Variables ***
${GOOGLE_URL}    https://wf.wapice.com/
${GLOBAL_RETRY_AMOUNT}    3x
${GLOBAL_RETRY_INTERVAL}    1.0s
${COUNTER}

*** Keyword ***
Go to Wapice Family
    Open Available Browser    ${GOOGLE_URL}

*** Keyword ***
Filter Team Leaders
    Wait Until Page Contains Element    xpath://*[@id="team-leader"]
    ${ELEMENT}=    Get WebElement    xpath://*[@id="team-leader"]
    Capture Element Screenshot    xpath://*[@id="team-leader"]
    Click Element    ${ELEMENT}

*** Keyword ***
Wait Until Success Filtering Team Leaders
        Wait Until Keyword Succeeds    ${GLOBAL_RETRY_AMOUNT}    ${GLOBAL_RETRY_INTERVAL}    Filter Team Leaders

*** Keyword ***
Select The First Team Leader
    Wait Until Page Contains Element    xpath://*[@id="counter"]/span
    ${RESULT}=    Get Text    xpath://*[@id="counter"]/span
    Log    ${RESULT}
    ${COUNTER}=    Fetch From Left    ${RESULT}    results
    Convert To Integer    ${COUNTER}
    Capture Element Screenshot    xpath://*[@id="antonp"]/a/img[2]
    Click Element    xpath://*[@id="antonp"]/a/img[2]
    [Return]    ${COUNTER}

*** Keyword ***
Wait Until Success Selectingt The First Team Leader
    ${COUNTER}=    Wait Until Keyword Succeeds    ${GLOBAL_RETRY_AMOUNT}    ${GLOBAL_RETRY_INTERVAL}    Select The First Team Leader
    Log    ${COUNTER}
    [Return]    ${COUNTER}

*** Keyword ***
Loop Team Leaders
    [Arguments]    ${COUNTER}
    ${list}=    Create List
    Log    ${COUNTER}
    Append To List    ${list}    Email
    ${table}=    Create Workbook    C:\\Users\\saktuo\\Desktop\\WF_emails.xlsx
    FOR    ${i}    IN RANGE    1    ${COUNTER}+1
        Wait Until Success With Reading Email    ${list}
    END
    Log List    ${list}
    Append Rows To Worksheet    ${list}
    Save Workbook

*** Keyword ***
Wait Until Success With Reading Email
        [Arguments]    ${list}
        Wait Until Keyword Succeeds    ${GLOBAL_RETRY_AMOUNT}    ${GLOBAL_RETRY_INTERVAL}    Save Email Addresses    ${list}
        [Return]    ${list}

*** Keyword ***
Save Email Addresses
        [Arguments]    ${list}
        Wait Until Page Contains Element    xpath://*[@id="info_dialog"]/div[2]/div[9]/a
        ${RECIPIENT}=    Get Text    xpath://*[@id="info_dialog"]/div[2]/div[9]/a
        ${Length}=    Get Length    ${RECIPIENT}
        Should Be True    ${Length} > 1
        Log    ${Length}
        Wait Until Page Contains Element    xpath:/html/body/div[11]/div[3]/div/button[2]/span
        Click Element    xpath:/html/body/div[11]/div[3]/div/button[2]/span
        Wait Until Page Contains Element    xpath://*[@id="info_dialog"]/div[1]/div/img
        Append To List    ${list}    ${RECIPIENT}
        Log    ${RECIPIENT}
        [Return]    ${list}


*** Keywords ***
Close The Browser
    Close Browser

*** Keywords ***
Open PowerShell
    Open Using Run Dialog    powershell    Windows PowerShell
    Send Keys    & {SPACE} "C:\\Users\\saktuo\\Desktop\\emailtest"{ENTER}

*** Tasks ***
Get contact information from Wapice Family
    Go To Wapice Family
    Wait Until Success Filtering Team Leaders
    ${COUNTER}=    Wait Until Success Selectingt The First Team Leader
    Loop Team Leaders    ${COUNTER}
    Close The Browser
    Open PowerShell
