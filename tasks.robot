*** Settings ***
Documentation     Insert the sales data for the week and export it as a PDF.
...               Collects the input Excel file from the user.
Library           RPA.Browser.Selenium    auto_close=${FALSE}
Library           RPA.Dialogs
Library           RPA.Excel.Files
Library           RPA.HTTP
Library           RPA.PDF

*** Tasks ***
Insert the sales data for the week and export it as a PDF
    ${excel_file_path}=    Collect Excel file from the user
    Open the intranet website
    Log in
    Fill the form using the data from the Excel file    ${excel_file_path}
    Collect the results
    Export the table as a PDF
    [Teardown]    Log out and close the browser

*** Keywords ***
Collect Excel file from the user
    Add heading    Upload Excel File
    Add file input
    ...    label=Upload the Excel file with sales data
    ...    name=fileupload
    ...    file_type=Excel files (*.xls;*.xlsx)
    ...    destination=${OUTPUT_DIR}
    ${response}=    Run dialog
    [Return]    ${response.fileupload}[0]

Open the intranet website
    Open Available Browser    https://robotsparebinindustries.com/

Log in
    Input Text    username    maria
    Input Password    password    thoushallnotpass
    Submit Form
    Wait Until Page Contains Element    id:sales-form

Fill and submit the form for one person
    [Arguments]    ${sales_rep}
    Input Text    firstname    ${sales_rep}[First Name]
    Input Text    lastname    ${sales_rep}[Last Name]
    Input Text    salesresult    ${sales_rep}[Sales]
    Select From List By Value    salestarget    ${sales_rep}[Sales Target]
    Click Button    Submit

Fill the form using the data from the Excel file
    [Arguments]    ${excel_file_path}
    Open Workbook    ${excel_file_path}
    ${sales_reps}=    Read Worksheet As Table    header=True
    Close Workbook
    FOR    ${sales_rep}    IN    @{sales_reps}
        Fill and submit the form for one person    ${sales_rep}
    END

Collect the results
    Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}sales_summary.png

Export the table as a PDF
    Wait Until Element Is Visible    id:sales-results
    ${sales_results_html}=    Get Element Attribute    id:sales-results    outerHTML
    Html To Pdf    ${sales_results_html}    ${OUTPUT_DIR}${/}sales_results.pdf

Log out and close the browser
    Click Button    Log out
    Close Browser
