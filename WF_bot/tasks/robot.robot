*** Settings ***
Documentation   An automation that fetches Wapice team leaders' email from Wapice Family and sends a reminder to them with Power Shell script
Variables       variables.py
Resource        keywords.robot

# +
*** Tasks ***
Open Wapice Family and filter the page to see only team leaders
    Go To Wapice Family
    Wait Until Success Filtering Team Leaders
    
Loop throught the team leaders and store their emails into excel file
    ${COUNTER}=    Wait Until Success Selectingt The First Team Leader
    Loop Team Leaders    ${COUNTER}
    Close The Browser
    
Send emails to all team leaders with PowerShell script
    Open PowerShell
    [Teardown]    Close PowerShell
