*** Settings ***
Documentation       Template robot main suite.

Library             RPA.Browser.Selenium
Library             RPA.Excel.Files
# Library    RPA.Robocorp.WorkItems
# Library    String
# Library    RPA.FTP
# Library    RPA.Desktop
# Library    XML
# Library    OperatingSystem
# Library    RPA.Notifier
Library             RPA.FileSystem
# Library    RPA.RobotLogListener
Library             Collections
Library             RPA.Windows
Library             RPA.RobotLogListener
# Library    RPA.Windows
# Library    RPA.Smartsheet


*** Variables ***
${browser}                  Chrome
${url}                      https://www.laliga.com/laliga-easports/resultados
${btnAcceptAllCockies}      id=onetrust-accept-btn-handler
${btnStatistics}            //p[@class='styled__TextStyled-sc-1mby3k1-0 cYEkps' and text()='Estadísticas']
${dropDownJornada}          //div[@class='styled__DropdownContainer-sc-d9k1bl-0 gpdfZV']/ul[@class='styled__ItemsList-sc-d9k1bl-2 lofGQu']/li
${totalPartidosJornada}     //td[@class='styled__TableCell-sc-43wy8s-5 styled__TableCellLink-sc-43wy8s-8 hhcbkq fVeghq']
${workbook_path}            C:\\git\\laLiga_robotFramework\\testData\\futbol.xlsx
# ${i}    1


*** Tasks ***
GetData
    ${jornadaRegistrada}=    Validate sheet
    WHILE    '${jornadaRegistrada}' != 39
        Log    ${jornadaRegistrada}
        Open laLiga
        IF    ${jornadaRegistrada} == 1    selectJornada    ${jornadaRegistrada}
        IF    ${jornadaRegistrada} == 2    selectJornada    ${jornadaRegistrada}
        IF    ${jornadaRegistrada} == 3    selectJornada    ${jornadaRegistrada}
        IF    ${jornadaRegistrada} == 4    selectJornada    ${jornadaRegistrada}
        IF    ${jornadaRegistrada} == 5    selectJornada    ${jornadaRegistrada}
        IF    ${jornadaRegistrada} == 6    selectJornada    ${jornadaRegistrada}
        IF    ${jornadaRegistrada} == 7    selectJornada    ${jornadaRegistrada}
        IF    ${jornadaRegistrada} == 8    selectJornada    ${jornadaRegistrada}
        IF    ${jornadaRegistrada} == 9    selectJornada    ${jornadaRegistrada}
        IF    ${jornadaRegistrada} == 10
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 11
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 12
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 13
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 14
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 15
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 16
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 17
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 18
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 19
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 20
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 21
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 22
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 23
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 24
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 25
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 26
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 27
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 28
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 29
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 30
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 31
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 32
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 33
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 34
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada}== 35    selectJornada    ${jornadaRegistrada}
        IF    ${jornadaRegistrada} == 36
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 37
            selectJornada    ${jornadaRegistrada}
        END
        IF    ${jornadaRegistrada} == 38
            selectJornada    ${jornadaRegistrada}
        END
        selectMatch    ${jornadaRegistrada}
        ${jornadaRegistrada}=    Validate sheet
        # ${jornadaRegistrada}=    Set Variable    39
    END

    # Open laLiga
    # ${items}=    getJornada
    # selectJornada
    # selectMatch


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
    [Arguments]    ${jornada}
    Click Element    //a[text()='Jornada ${jornada}']
    # ...    //div[@class='styled__DropdownContainer-sc-d9k1bl-0 gpdfZV']/ul[@class='styled__ItemsList-sc-d9k1bl-2 lofGQu']/li[1]
    Sleep    10s

selectMatch
    [Arguments]    ${jornada}
    ${matches}=    Get Element Count    ${totalPartidosJornada}
    Log    ${matches}
    FOR    ${i}    IN RANGE    ${matches}
        Wait Until Page Contains Element    //tbody/tr[starts-with(@class, 'styled__TableRow')]/td/a
        ${td_elements}=    Get WebElements    //tbody/tr[starts-with(@class, 'styled__TableRow')]/td/a
        log    ${td_elements}
        Close widgets
        Sleep    1s
        ${scroll_error}=    Run Keyword And Ignore Error    Scroll Element Into View    ${td_elements}[${i}]
        Wait Until Element Is Visible    ${td_elements}[${i}]    timeout=30s
        Click Element    ${td_elements}[${i}]
        Log    Scrolling result: ${scroll_error}
        Sleep    3s
        ${element-found}=    Run Keyword And Return Status
        ...    Element Should Be Visible
        ...    ${btnStatistics}
        ...    timeout=30s
        IF    ${element-found}    getData_Statistics
        # getData_Statistics
        Close Browser
        IF    ${i}< ${matches}-1
            Open laLiga
            SelectJornada    ${jornada}
        END
    END

getData_Statistics
    # ${element-found}=    Run Keyword And Return Status    Element Should Be Visible    ${btnStatistics}    timeout=30s

    # IF    ${element-found}
    # ...    Click Element    ${btnStatistics}
    ${equipo1}=    Create ListEquipoCasa
    ${equipo2}=    Create ListEquipoFuera
    Create Excel    ${equipo1}    ${equipo2}
    # Wait Until Element Is Visible    ${btnStatistics}    timeout=30s
    # Click Element    ${btnStatistics}

Create Excel
    [Arguments]    ${equipo1}    ${equipo2}
    ${file_exist}=    Does File Exist    ${workbook_path}
    IF    not ${file_exist}    Create new Excel
    Read excel file_exist    ${equipo1}    ${equipo2}

Read excel file_exist
    [Arguments]    ${equipo1}    ${equipo2}
    Open Workbook    ${workbook_path}
    ${sheets}=    List Worksheets
    ${teamName}=    Get Equipo_casa
    log    ${teamName}
    IF    not '${teamName}' in ${sheets}
        Create new sheet in Excel    ${teamName}
    END
    Add rows to excel for team    ${equipo1}    ${teamName}

    ${teamName}=    Get Equipo_fuera

    IF    not '${teamName}' in ${sheets}
        Create new sheet in Excel    ${teamName}
    END
    Add rows to excel for team    ${equipo2}    ${teamName}

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

Add rows to excel for team
    [Arguments]    ${equipo}    ${teamname}
    Log    ${equipo}
    Open Workbook    ${workbook_path}
    Log    ${teamname}
    @{sheets}=    List Worksheets
    Read Worksheet    ${teamname}
    FOR    ${sheet}    IN    @{sheets}
        Log    ${sheet}
        IF    '${sheet}' == '${teamname}'
            Append Rows To Worksheet    ${equipo}    name=${teamname}    header=${TRUE}
            Save Workbook
            Close Workbook
        END
    END

Close widgets
    ${widget_elements}=    Get WebElements    //div[@class='rctfl-close rctfl-widget-close rctfl-widget-close-image']
    Log    ${widget_elements}
    FOR    ${element}    IN    @{widget_elements}
        ${is_present}=    Run Keyword And Return Status    Page Should Contain Element    ${element}
        Log    ${is_present}
        IF    ${is_present}
            ${is_interactable}=    Run Keyword And Return Status    Element Should Be Visible    ${element}
            IF    ${is_interactable}
                Click Element    ${element}
            ELSE
                Log    Element is not interactable: ${element}
            END
        ELSE
            Log    Element not found: ${element}
        END
        # IF    ${is_visible}    Click Element    ${element}
    END
    # ${propaganda1}=    Run Keyword And Return Status
    # ...    Element Should Be Visible
    # ...    (//div[@class='rctfl-close rctfl-widget-close rctfl-widget-close-image'])[2]
    # ...    timeout=20s
    # ${propaganda2}=    Run Keyword And Return Status
    # ...    Element Should Be Visible
    # ...    (//div[@class='rctfl-close rctfl-widget-close rctfl-widget-close-image'])[1]
    # ...    timeout=20s

    # IF    ${propaganda1}
    #    Click Element    (//div[@class='rctfl-close rctfl-widget-close rctfl-widget-close-image'])[2]
    # END
    # IF    ${propaganda2}
    #    Click Element    (//div[@class='rctfl-close rctfl-widget-close rctfl-widget-close-image'])[1]
    # END

Get Equipo_casa
    ${equipo_casa}=    RPA.Browser.Selenium.Get Text
    ...    //div[@class='styled__ContainerName-sc-xzosab-2 bpVmwZ']//p[@class='styled__TextStyled-sc-1mby3k1-0 kFavhB']
    RETURN    ${equipo_casa}

Get Equipo_fuera
    ${equipo_fuera}=    RPA.Browser.Selenium.Get Text
    ...    //div[@class='styled__ContainerName-sc-xzosab-2 fOhNvl']//p[@class='styled__TextStyled-sc-1mby3k1-0 kFavhB']
    # ${listNames}=    Create List    ${team1}    ${team2}
    RETURN    ${equipo_fuera}

Get Yellow_Cards_Casa
    ${yellowCard_casa}=    RPA.Browser.Selenium.Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 iPYgyC'])[9]
    RETURN    ${yellowCard_casa}

Get Yellow_Cards_Fuera
    ${yellowCard_fuera}=    RPA.Browser.Selenium.Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 iPYgyC'])[10]
    RETURN    ${yellowCard_fuera}

Get Red_Cards_Casa
    ${redCard_casa}=    RPA.Browser.Selenium.Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 iPYgyC'])[11]
    RETURN    ${redCard_casa}

Get Red_Cards_Fuera
    ${redCard_fuera}=    RPA.Browser.Selenium.Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 iPYgyC'])[12]
    RETURN    ${redCard_fuera}

Get Corners_Casa
    ${corner_casa}=    RPA.Browser.Selenium.Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 iPYgyC'])[15]
    RETURN    ${corner_casa}

Get Corners_Fuera
    ${corner_fuera}=    RPA.Browser.Selenium.Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 iPYgyC'])[16]
    RETURN    ${corner_fuera}

Get Matches
    ${team1}=    RPA.Browser.Selenium.Get Text
    ...    //div[@class='styled__ContainerName-sc-xzosab-2 bpVmwZ']//p[@class='styled__TextStyled-sc-1mby3k1-0 kFavhB']
    ${team2}=    RPA.Browser.Selenium.Get Text
    ...    //div[@class='styled__ContainerName-sc-xzosab-2 fOhNvl']//p[@class='styled__TextStyled-sc-1mby3k1-0 kFavhB']
    ${teamMatch}=    Catenate    SEPARATOR=-    ${team1} ${team2}
    Log    ${teamMatch}
    RETURN    ${teamMatch}

Create ListEquipoCasa
    ${partido}=    Get Matches
    ${tarjetasA_casa}=    Get Yellow_Cards_Casa
    ${tarjetasR_casa}=    Get Red_Cards_Casa
    ${corners_casa}=    Get Corners_Casa
    &{equipo1}=    Create Dictionary
    ...    partido=${partido}
    ...    tarjetasA=${tarjetasA_casa}
    ...    tarjetasR=${tarjetasR_casa}
    ...    corners=${corners_casa}
    RETURN    ${equipo1}

Create ListEquipoFuera
    ${partido}=    Get Matches
    ${tarjetasA_fuera}=    Get Yellow_Cards_Fuera
    ${tarjetasR_fuera}=    Get Red_Cards_Fuera
    ${corners_fuera}=    Get Corners_Fuera
    &{equipo2}=    Create Dictionary
    ...    partido=${partido}
    ...    tarjetasA=${tarjetasA_fuera}
    ...    tarjetasR=${tarjetasR_fuera}
    ...    corners=${corners_fuera}
    RETURN    ${equipo2}

Validate sheet
    ${file_exist}=    Does File Exist    ${workbook_path}
    IF    ${file_exist}
        Open Workbook    ${workbook_path}
        @{sheets}=    List Worksheets
        ${sheet_name}=    Get From List    ${sheets}    1
        ${sheet}=    Read Worksheet    ${sheet_name}
        ${count}=    Get Length    ${sheet}
        Log    ${count}
    ELSE
        ${count}=    Set Variable    1
    END
    RETURN    ${count}

CheckMatchNotPlayed
    ${played}=    Run Keyword And Return Status    Element Should Be Visible    //p[text()=' VS ']
    RETURN    ${played}
