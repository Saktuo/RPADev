*** Settings ***
# The libraries and variable files that are used for this automation
Variables       variables.py
Library         RPA.Browser
Library         RPA.Desktop.Windows
Library         Collections
Library         RPA.Excel.Files
Library         String

# +
*** Keyword ***
Go to Wapice Family
    # Goes to the URL defined in The Variables (The Wapice Family URL). The browser must be Google Chrome which is defined in the call.
    Open Available Browser    ${GOOGLE_URL}    googlechrome
    
    
Wait Until Success Filtering Team Leaders
    # Using an error management keyword 'Wait Until Keyword Succeeds' that runs a selected keyword and if it fails,
    #it waits 'GLOBAL_RETRY_INTERVAL' amount of time and runs the keyword again this is repeated 'GLOBAL_RETRY_AMOUNT' times.
    Wait Until Keyword Succeeds    ${GLOBAL_RETRY_AMOUNT}    ${GLOBAL_RETRY_INTERVAL}    Filter Team Leaders
    

Filter Team Leaders
    # The automation waits until the Wapice Family page is opened and then clicks the 'Team leader' filter and waits until the results are visible.
    Wait Until Page Contains Element    xpath://*[@id="team-leader"]
    ${ELEMENT}=    Get WebElement    xpath://*[@id="team-leader"]
    Click Element    ${ELEMENT}
    Wait Until Page Contains Element    xpath://*[@id="counter"]/span
        
        
Wait Until Success Selectingt The First Team Leader
    # Using an error management keyword 'Wait Until Keyword Succeeds' that runs a selected keyword and if it fails,
    # it waits 'GLOBAL_RETRY_INTERVAL' amount of time and runs the keyword again this is repeated 'GLOBAL_RETRY_AMOUNT' times.
    # They keyword 'Select The First Team Leader' return an argument '${COUNTER}' that has the amount of team leaders on Wapice Family page.
    # After this, the 'Wait Until Success Selectingt The First Team Leader' gives the value in log and returns the argument for other keywords to use
    ${COUNTER}=    Wait Until Keyword Succeeds    ${GLOBAL_RETRY_AMOUNT}    ${GLOBAL_RETRY_INTERVAL}    Select The First Team Leader
    Log    ${COUNTER}
    [Return]    ${COUNTER}
        
        
Select The First Team Leader
    # The automation reads the amount of team leaders shown in the results saves it to the variable '${RESULT}'.
    # Then it parses the stored variable and converts it to integer and saves it as '${COUNTER}' and returns it.
    # After this, the automation selects the first team leader from employee list using a xpath locator.
    ${RESULT}=    Get Text    xpath://*[@id="counter"]/span
    Log    ${RESULT}
    ${COUNTER}=    Fetch From Left    ${RESULT}    results
    Convert To Integer    ${COUNTER}
    Click Element    xpath://*[@id="antonp"]/a/img[2]
    [Return]    ${COUNTER}
    
    
Loop Team Leaders
    # This keyword recieves an argument '${COUNTER}' which is the amount of team leaders in Wapice Family page.
    # The keyword creates a list called '${emails}' and sets one the first item to be 'Email'
    # The keyword creates an excel file into a location '${EXCEL_FILE_LOCATION}'.
    # The keyword starts a FOR -loop that calls a keyword 'Wait Until Success With Reading Email' which recieves the '${emails}' as an argument.
    # After the loop, the keyword prints the items of the ${emails} into log, appends them into excel file and saves the excel file.
    [Arguments]    ${COUNTER}
    ${emails}=    Create List
    Log    ${COUNTER}
    Append To List    ${emails}    Email
    ${table}=    Create Workbook    ${EXCEL_FILE_LOCATION}
    FOR    ${i}    IN RANGE    1    ${COUNTER}+1
        Wait Until Success With Reading Email    ${emails}
    END
    Log List    ${emails}
    Append Rows To Worksheet    ${emails}
    Save Workbook
    
    
Wait Until Success With Reading Email
    # Using an error management keyword 'Wait Until Keyword Succeeds' that runs a selected keyword and if it fails,
    # it waits 'GLOBAL_RETRY_INTERVAL' amount of time and runs the keyword again this is repeated 'GLOBAL_RETRY_AMOUNT' times.
    # The keyword recieves an argument ${emails} and return is also to keyword that called this keyword.
    [Arguments]    ${emails}
    Wait Until Keyword Succeeds    ${GLOBAL_RETRY_AMOUNT}    ${GLOBAL_RETRY_INTERVAL}    Save Email Addresses    ${emails}
    [Return]    ${emails}
    
    
Save Email Addresses
    # This keyword recieves a list as an argument ${emails} that this keyword fills with emails fetched from Wapice Family.
    # The keyword gets the email address of a team leader and store it into a variable ${RECIPIENT} 
    # and checks that it read the right box by checking the length of the email.
    # If it's shorter than 1 char, automation throws an error and trries again as defined in 'Wait Until Keyword Succeeds' keyword.
    # When the email is ok, keyword adds it to the ${emails}.
    # After this, the automation clicks an element to get to the next employee and return the ${emails}.
    [Arguments]    ${emails}
    Wait Until Page Contains Element    xpath://*[@id="info_dialog"]/div[2]/div[9]/a
    ${RECIPIENT}=    Get Text    xpath://*[@id="info_dialog"]/div[2]/div[9]/a
    ${Length}=    Get Length    ${RECIPIENT}
    Should Be True    ${Length} > 1
    Log    ${Length}
    Append To List    ${emails}    ${RECIPIENT}
    Wait Until Page Contains Element    xpath:/html/body/div[11]/div[3]/div/button[2]/span
    Click Element    xpath:/html/body/div[11]/div[3]/div/button[2]/span
    Wait Until Page Contains Element    xpath://*[@id="info_dialog"]/div[1]/div/img
    Log    ${RECIPIENT}
    [Return]    ${emails}
    
    
Close The Browser
    # The keyword closes the browser where the Wapice Family is used.
    Close Browser
    
    
Open PowerShell
    # The automation opens PowerShell and sends a run command to run a PowerShell script stored in location '${POWERSHELL_SCRIPT}' 
    # that sends an email to all email addresses that are stored in the excel file in location '${EXCEL_FILE_LOCATION}'.
    Open Using Run Dialog    powershell    Windows PowerShell
    Send Keys    & {SPACE} ${POWERSHELL_SCRIPT} {ENTER}
    Sleep    3
    
    
Close PowerShell
    # Takes screencapture of the message sent and closes PowerShell application
    Quit Application
