*** Settings ***
Documentation       Template robot main suite.

Library             RPA.Browser.Selenium
Library             RPA.Excel.Files
Library             RPA.Robocorp.WorkItems
Library             String
Library             RPA.FTP
Library             RPA.Desktop
Library             XML
Library             OperatingSystem
Library             RPA.Notifier
Library             RPA.FileSystem
Library             RPA.RobotLogListener
Library             Collections
# Library    RPA.Excel.Application


*** Variables ***
${browser}                  Chrome
${url}                      https://www.laliga.com/laliga-easports/resultados
${btnAcceptAllCockies}      id=onetrust-accept-btn-handler
${btnStatistics}            //p[@class='styled__TextStyled-sc-1mby3k1-0 cYEkps' and text()='Estadísticas']
${dropDownJornada}          //div[@class='styled__DropdownContainer-sc-d9k1bl-0 gpdfZV']/ul[@class='styled__ItemsList-sc-d9k1bl-2 lofGQu']/li
# ${equipoCasa}    //div[@class='styled__ContainerName-sc-xzosab-2 bpVmwZ']//p[@class='styled__TextStyled-sc-1mby3k1-0 kFavhB']
# ${equipoFuera}    //div[@class='styled__ContainerName-sc-xzosab-2 fOhNvl']//p[@class='styled__TextStyled-sc-1mby3k1-0 kFavhB']
${totalPartidosJornada}     //td[@class='styled__TableCell-sc-43wy8s-5 styled__TableCellLink-sc-43wy8s-8 hhcbkq fVeghq']
${workbook_path}            C:\\cursos\\robot frameworks - selenium - python\\robot\\testData\\futbol.xlsx
${shell_name}               ${EMPTY}


*** Tasks ***
GetData
    Open laLiga
    ${items}=    getJornada
    selectJornada
    selectMatch
    # ${partido}=    getStatistics
    # log    ${partido}
    # Create Excel    ${partido}


*** Keywords ***
Open laLiga
    Open Browser    ${url}    ${browser}
    Maximize Browser Window
    Wait Until Element Is Visible    ${btnAcceptAllCockies}    timeout=30s
    Click Button    ${btnAcceptAllCockies}
    Title Should Be    Resultados de LALIGA EA SPORTS 2023/24 | LALIGA

getJornada
    ${items}=    Get Element Count    ${dropDownJornada}
    log    ${items}
    RETURN    ${items}

selectJornada
    Click Element
    ...    //div[@class='styled__DropdownContainer-sc-d9k1bl-0 gpdfZV']/ul[@class='styled__ItemsList-sc-d9k1bl-2 lofGQu']/li[1]
    Sleep    5s

getData_Statistics
    Wait Until Element Is Visible    ${btnStatistics}    timeout=30s
    Click Element    ${btnStatistics}
    ${listNames}=    Get Team_name
    ${listYellowCards}=    Get Yellow_Cards
    ${listRedCards}=    Get Red_Cards
    ${listCorners}=    Get Corners
    Create Excel    ${listNames}    ${listYellowCards}    ${listRedCards}    ${listCorners}
    Close Browser
    # RETURN    ${partido}

selectMatch
    ${matches}=    Get Element Count    ${totalPartidosJornada}
    Log    ${matches}
    FOR    ${i}    IN RANGE    ${1}
        Wait Until Page Contains Element    //tbody/tr[starts-with(@class, 'styled__TableRow')]/td/a
        ${td_elements}=    Get WebElements    //tbody/tr[starts-with(@class, 'styled__TableRow')]/td/a
        log    ${td_elements}
        Sleep    3s
        Close widgets
        Sleep    5s
        ${scroll_error}=    Run Keyword And Ignore Error    Scroll Element Into View    ${td_elements}[${i}]
        Wait Until Element Is Visible    ${td_elements}[${i}]    timeout=30s
        Sleep    3s
        Click Element    ${td_elements}[${i}]
        Log    Scrolling result: ${scroll_error}
        getData_Statistics
        Close Browser
        IF    ${i}< ${matches}-1
            Open laLiga
            SelectJornada
        END
    END

Create Excel
    [Arguments]    ${listNames}    ${listYellowCards}    ${listRedCards}    ${listCorners}
    ${file_exist}=    Does File Exist    ${workbook_path}
    Log    ${file_exist}
    IF    not ${file_exist}    Create new Excel
    Read excel file_exist    ${listNames}    ${listYellowCards}    ${listRedCards}    ${listCorners}

Create new sheet in Excel
    [Arguments]    ${team}
    Open Workbook    ${workbook_path}
    Create Worksheet    name=${team}
    Save Workbook    ${workbook_path}
    Close Workbook

Create new Excel
    Create Workbook    ${workbook_path}
    Save Workbook
    Close Workbook

Read excel file_exist
    [Arguments]    ${listNames}    ${listYellowCards}    ${listRedCards}    ${listCorners}
    Open Workbook    ${workbook_path}
    ${sheets}=    List Worksheets
    &{Dictionary1}=    Create Dictionary    teams=${listNames[0]}
    ...    yellowCards=${listYellowCards[0]}
    ...    redCards=${listRedCards[0]}
    ...    corners=${listCorners[0]}
    &{Dictionary2}=    Create Dictionary    teams=${listNames[1]}
    ...    yellowCards=${listYellowCards[1]}
    ...    redCards=${listRedCards[1]}
    ...    corners=${listCorners[1]}
    Log    ${Dictionary2}
    # FOR    ${teamName}    IN    @{listNames}
    #    Log    ${teamName}
    IF    ' ${listNames[0]}' not in ${sheets}
        Create new sheet in Excel    ${listNames[0]}
    ELSE
        log    sheet is present in the escel
    END
    IF    '${listNames[1]}' not in ${sheets}
        Create new sheet in Excel    ${listNames[1]}
    ELSE
        log    sheet is present in the escel
    END
    Add rows to excel for team1    &{Dictionary1}
    # Add rows to excel for team1    &{Dictionary2}

Add rows to excel for team1
    [Arguments]    &{Dictionary}
    Log    ${Dictionary}
    Open Workbook    ${workbook_path}
    ${teamname}=    Get From Dictionary    ${Dictionary}    teams
    Log    ${teamname}
    ${worksheet_name}=    Read Worksheet    name=${teamName}
    Append Rows To Worksheet    ${Dictionary}    header=${TRUE}
    Save Workbook
    Close Workbook

Close widgets
    ${propaganda1}=    Run Keyword And Return Status
    ...    Element Should Be Visible
    ...    (//div[@class='rctfl-close rctfl-widget-close rctfl-widget-close-image'])[2]
    ...    timeout=10s
    ${propaganda2}=    Run Keyword And Return Status
    ...    Element Should Be Visible
    ...    (//div[@class='rctfl-close rctfl-widget-close rctfl-widget-close-image'])[1]
    ...    timeout=10s
    IF    ${propaganda1}
        Click Element    (//div[@class='rctfl-close rctfl-widget-close rctfl-widget-close-image'])[2]
    END
    IF    ${propaganda2}
        Click Element    (//div[@class='rctfl-close rctfl-widget-close rctfl-widget-close-image'])[1]
    END

Get Team_name
    ${team1}=    Get Text
    ...    //div[@class='styled__ContainerName-sc-xzosab-2 bpVmwZ']//p[@class='styled__TextStyled-sc-1mby3k1-0 kFavhB']
    ${team2}=    Get Text
    ...    //div[@class='styled__ContainerName-sc-xzosab-2 fOhNvl']//p[@class='styled__TextStyled-sc-1mby3k1-0 kFavhB']
    ${listNames}=    Create List    ${team1}    ${team2}
    Log    ${listNames}
    RETURN    ${listNames}

Get Yellow_Cards
    ${yellowCard1}=    Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 iPYgyC'])[9]
    ${yellowCard2}=    Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 iPYgyC'])[10]
    ${listYellowCards}=    Create List    ${yellowCard1}    ${yellowCard2}
    Log    ${listYellowCards}
    RETURN    ${listYellowCards}

Get Red_Cards
    ${redCard1}=    Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 iPYgyC'])[11]
    ${redCard2}=    Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 iPYgyC'])[12]
    ${listRedCards}=    Create List    ${redCard1}    ${redCard2}
    Log    ${listRedCards}
    RETURN    ${listRedCards}

Get Corners
    ${corner1}=    Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 iPYgyC'])[15]
    ${corner2}=    Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 iPYgyC'])[16]
    ${listCorners}=    Create List    ${corner1}    ${corner2}
    Log    ${listCorners}
    RETURN    ${listCorners}
