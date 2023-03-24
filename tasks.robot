*** Settings ***
Documentation       Robot to solve the first challenge at rpachallenge.com,
...                 which consists of filling a form that randomly rearranges
...                 itself for ten times, with data taken from a provided
...                 Microsoft Excel file.

Library             RPA.Browser.Selenium
Library             RPA.Excel.Files
Library             RPA.HTTP


*** Tasks ***
Complete the challenge
    Start the challenge
    Fill the forms
    Collect the results


*** Keywords ***
Start the challenge
    Open Available Browser    http://rpachallenge.com/
    Sleep    5
    Download    http://rpachallenge.com/assets/downloadFiles/challenge.xlsx    overwrite=True
    Sleep    5
    Click Button    Start

Fill the forms
    ${people}=    Get the list of people from the Excel file
    FOR    ${person}    IN    @{people}
        Fill and submit the form    ${person}
        Sleep    3
    END

Get the list of people from the Excel file
    Open Workbook    challenge.xlsx
    ${table}=    Read Worksheet As Table    header=True
    Close Workbook
    RETURN    ${table}

Fill and submit the form
    [Arguments]    ${person}
    Input Text    css:input[ng-reflect-name="labelFirstName"]    ${person}[First Name]
    Sleep    1
    Input Text    css:input[ng-reflect-name="labelLastName"]    ${person}[Last Name]
    Sleep    1
    Input Text    css:input[ng-reflect-name="labelCompanyName"]    ${person}[Company Name]
    Sleep    1
    Input Text    css:input[ng-reflect-name="labelRole"]    ${person}[Role in Company]
    Sleep    1
    Input Text    css:input[ng-reflect-name="labelAddress"]    ${person}[Address]
    Sleep    1
    Input Text    css:input[ng-reflect-name="labelEmail"]    ${person}[Email]
    Sleep    1
    Input Text    css:input[ng-reflect-name="labelPhone"]    ${person}[Phone Number]
    Click Button    Submit

Set value by XPath
    [Arguments]    ${xpath}    ${value}
    ${result}=    Execute Javascript
    ...    document.evaluate('${xpath}',document.body,null,9,null).singleNodeValue.value='${value}';
    RETURN    ${result}

Collect the results
    Capture Element Screenshot    css:div.congratulations
    Close All Browsers
