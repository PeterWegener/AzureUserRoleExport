
# Create a user with the according rights before and exchange witht the correct tenant
Connect-AzAccount -Tenant '00000000-0000-0000-0000-000000000000' 
$tenantId = (Get-AzContext).Tenant.Id

$allResources = @()
$subscriptions=Get-AzureRMSubscription

ForEach ($vsub in $subscriptions){
    Select-AzureRmSubscription $vsub.SubscriptionID

    Write-Host
    Write-Host “Working on “ $vsub
    Write-Host

    $allResources += $allResources |Select-Object $vsub.SubscriptionID,$vsub.Name

    Get-AzureRmVM | Select Name,REsourceGroupName,Location

    # Fetch list of all directory roles with template ID
    Get-AzureADMSRoleDefinition

    # Fetch a specific directory role by ID
    $role = Get-AzureADMSRoleDefinition -Id $vsub

    # Fetch membership for a role
    Get-AzureADMSRoleAssignment -Filter "roleDefinitionId eq '$($role.Id)'" | Export-Csv az_user_role.csv


    ############### Convert to excel #####################


    ### Set input and output path
    $inputCSV = "C:\ps\az_user_role.csv"
    $outputXLSX = "C:\ps\az_user_role.xlsx"

    ### Create a new Excel Workbook with one empty sheet
    $excel = New-Object -ComObject excel.application 
    $workbook = $excel.Workbooks.Add(1)
    $worksheet = $workbook.worksheets.Item(1)

    ### Build the QueryTables.Add command
    ### QueryTables does the same as when clicking "Data » From Text" in Excel
    $TxtConnector = ("TEXT;" + $inputCSV)
    $Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
    $query = $worksheet.QueryTables.item($Connector.name)

    ### Set the delimiter (, or ;) according to your regional settings
    $query.TextFileOtherDelimiter = $Excel.Application.International(5)

    ### Set the format to delimited and text for every column
    ### A trick to create an array of 2s is used with the preceding comma
    $query.TextFileParseType  = 1
    $query.TextFileColumnDataTypes = ,2 * $worksheet.Cells.Columns.Count
    $query.AdjustColumnWidth = 1

    ### Execute & delete the import query
    $query.Refresh()
    $query.Delete()

    ### Save & close the Workbook as XLSX. Change the output extension for Excel 2003
    $Workbook.SaveAs($outputXLSX,51)
    $excel.Quit()

}

