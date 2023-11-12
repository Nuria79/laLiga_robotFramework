*** Settings ***
Documentation       Template robot main suite.

Library             RPA.Browser.Selenium
Library             RPA.Excel.Files
Library             RPA.FileSystem
Library             Collections
Library             RPA.Windows
Library             RPA.RobotLogListener
Library             RPA.Desktop


*** Variables ***
${browser}                  Chrome
${url}                      https://www.laliga.com/laliga-easports/resultados
${btnAcceptAllCockies}      id=onetrust-accept-btn-handler
${btnStatistics}            //p[@class='styled__TextStyled-sc-1mby3k1-0 cYEkps' and text()='Estadísticas']
${dropDownJornada}          //div[@class='styled__DropdownContainer-sc-d9k1bl-0 gpdfZV']/ul[@class='styled__ItemsList-sc-d9k1bl-2 lofGQu']/li
${totalPartidosJornada}     //td[@class='styled__TableCell-sc-43wy8s-5 styled__TableCellLink-sc-43wy8s-8 hhcbkq fVeghq']
${workbook_path}            C:\\git\\laLiga_robotFramework\\testData\\futbol.xlsx


*** Tasks ***
GetData
    ${jornadaRegistrada}=    Validate sheet
    WHILE    ${jornadaRegistrada} != 39
        Log    ${jornadaRegistrada}
        Open laLiga
        ${jornada_sel_default}=    Run Keyword And Return Status
        ...    Element Should Be Visible
        ...    //p[text()='Jornada ${jornadaRegistrada}']
        IF    not ${jornada_sel_default}    # jornada seleccionda por defecto en la lista
            selectJornada    ${jornadaRegistrada}
        END
        # comprobamos si la jornada se ha jugado viendo si los partidos estan vs
        ${jornada_jugada}=    getJornadaJugada
        IF    ${jornada_jugada} != 10    # jornada no jugada aun
            selectMatch    ${jornadaRegistrada}
            ${jornadaRegistrada}=    Validate sheet
        END
        Close Browser
        ${jornadaRegistrada}=    Evaluate    ${jornadaRegistrada} + 1
    END


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
    Sleep    4s

selectMatch
    [Arguments]    ${jornada}
    ${matches}=    Get Element Count    ${totalPartidosJornada}
    Log    ${matches}
    FOR    ${i}    IN RANGE    ${matches}
        Wait Until Page Contains Element    //tbody/tr[starts-with(@class, 'styled__TableRow')]/td/a
        ${td_elements}=    Get WebElements    //tbody/tr[starts-with(@class, 'styled__TableRow')]/td/a
        log    ${td_elements}
        ${scroll_error}=    Run Keyword And Ignore Error    Scroll Element Into View    ${td_elements}[${i}]
        Wait Until Element Is Visible    ${td_elements}[${i}]    timeout=30s

        ${enabled}=    Run Keyword And Return Status    Click Element    ${td_elements}[${i}]
        IF    not ${enabled}
            Close widgets
            Click Element    ${td_elements}[${i}]
        END
        Log    Scrolling result: ${scroll_error}
        Sleep    3s
        ${scroll_error}=    Run Keyword And Ignore Error    Scroll Element Into View    ${btnStatistics}
        ${element-found}=    Run Keyword And Return Status
        ...    Element Should Be Visible
        ...    ${btnStatistics}
        ...    timeout=30s
        IF    ${element-found}    getData_Statistics
        Close Browser
        # END

        IF    ${i}< ${matches}-1
            Open laLiga
            SelectJornada    ${jornada}
        END
    END

getData_Statistics
    Click Element    ${btnStatistics}
    ${equipo1}=    Create ListEquipoCasa
    ${equipo2}=    Create ListEquipoFuera
    Create Excel    ${equipo1}    ${equipo2}

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
    END

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
    ${teamMatch}=    Catenate    ${team1}    -    ${team2}
    Log    ${teamMatch}
    RETURN    ${teamMatch}

Create ListEquipoCasa
    ${partido}=    Get Matches
    ${resultado}=    getResultado
    ${tarjetasA_casa}=    Get Yellow_Cards_Casa
    ${tarjetasR_casa}=    Get Red_Cards_Casa
    ${total_tarjetas}=    getTotalTarjetas
    ${corners_casa}=    Get Corners_Casa
    ${total_corners}=    getTotalCorners
    &{equipo1}=    Create Dictionary
    ...    Partido=${partido}
    ...    Resultado=${resultado}
    ...    TarjetasA=${tarjetasA_casa}
    ...    TarjetasR=${tarjetasR_casa}
    ...    TotalTarjetas=${total_tarjetas}
    ...    Corners=${corners_casa}
    ...    TotalCorners=${total_corners}
    RETURN    ${equipo1}

Create ListEquipoFuera
    ${partido}=    Get Matches
    ${resultado}=    getResultado
    ${tarjetasA_fuera}=    Get Yellow_Cards_Fuera
    ${tarjetasR_fuera}=    Get Red_Cards_Fuera
    ${total_tarjetas}=    getTotalTarjetas
    ${corners_fuera}=    Get Corners_Fuera
    ${total_corners}=    getTotalCorners
    &{equipo2}=    Create Dictionary
    ...    Partido=${partido}
    ...    Resultado=${resultado}
    ...    TarjetasA=${tarjetasA_fuera}
    ...    TarjetasR=${tarjetasR_fuera}
    ...    TotalTarjetas=${total_tarjetas}
    ...    Corners=${corners_fuera}
    ...    TotalCorners=${total_corners}
    RETURN    ${equipo2}

Validate sheet
    ${file_exist}=    Does File Exist    ${workbook_path}
    IF    ${file_exist}
        Open Workbook    ${workbook_path}
        @{sheets}=    List Worksheets
        ${sheet_name}=    Get From List    ${sheets}    1
        ${sheet}=    Read Worksheet    ${sheet_name}
        ${count}=    Get Length    ${sheet}
        Log    ${count}-1
    ELSE
        ${count}=    Set Variable    1
    END
    RETURN    ${count}

getJornadaJugada
    ${jornada_aplazada}=    Get WebElements    //p[text()=' VS ']
    ${count}=    Get Element Count    ${jornada_aplazada}
    RETURN    ${count}

getResultado
    ${res_casa}=    RPA.Browser.Selenium.Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 kZeLZM'])[1]
    ${res_fuera}=    RPA.Browser.Selenium.Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 kZeLZM'])[2]
    ${resultado}=    Catenate    ${res_casa}    -    ${res_fuera}
    RETURN    ${resultado}

getTotalTarjetas
    ${yellowCard_fuera}=    RPA.Browser.Selenium.Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 iPYgyC'])[10]
    ${yellowCard_casa}=    RPA.Browser.Selenium.Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 iPYgyC'])[9]
    ${redCard_casa}=    RPA.Browser.Selenium.Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 iPYgyC'])[11]
    ${redCard_fuera}=    RPA.Browser.Selenium.Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 iPYgyC'])[12]
    ${total_tarjetas}=    Evaluate    ${yellowCard_casa} + ${yellowCard_fuera} + ${redCard_casa} + ${redCard_fuera}
    RETURN    ${total_tarjetas}

getTotalCorners
    ${corner_casa}=    RPA.Browser.Selenium.Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 iPYgyC'])[15]
    ${corner_fuera}=    RPA.Browser.Selenium.Get Text    (//p[@class='styled__TextStyled-sc-1mby3k1-0 iPYgyC'])[16]
    ${total_corners}=    Evaluate    ${corner_casa} + ${corner_fuera}
    RETURN    ${total_corners}
