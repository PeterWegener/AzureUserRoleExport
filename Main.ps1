
# Create a user with the according rights before and exchange witht the correct tenant
Connect-AzAccount -Tenant '00000000-0000-0000-0000-000000000000' 
# Variable for defining the location, which is used to filter 
# example "southeastasia"
$location = "southeastasia"
#Generate csv File to later convertation to Excel file. Fill in your path!
# example ".\output.xlsx"
$outfile = "filepath_for_temporary_csv_file"
$outputXLSX = "filepath_for_Excel_report"
#safe headers for csv file
Add-Content -Path $outfile  -Value '"Subscription_Name","Subscription_ID","Role_Name","Username","Email"'
$allResources = @()
$subscriptions=Get-AzSubscription

ForEach ($vsub in $subscriptions){
    #Select curent subscription to get ressource details
    Select-AZSubscription $vsub.SubscriptionID

    Write-Host
    Write-Host “Working on “ $vsub
    Write-Host
    # Fetch all Ressources in Subscription
    $allResources = Get-AzResource
    ForEach ($resource in $allResources){
        # Check if Ressourcelocation is the same as $location variable
        if($resource.Location -eq $location){
            # Get Role-Assignements of subscription
             $Roleassignment = Get-AzRoleAssignment
             # Import csv file for appending a new row
             $csvimport = Import-Csv $outfile
             # Create new Custom Object to append to CSV
             $newrow = [PSCustomObject] @{
                "Subscription_Name" = $vsub.Name;
                "Subscription_ID" = $vsub.SubscriptionId;
                "Role_Name" = $Roleassignment.RoleDefinitionName;
                "Username" = $Roleassignment.DisplayName;
                "Email" = $Roleassignment.SignInName;
            }
            $newrow | Export-CSV $outfile -Append -NoTypeInformation
        }
    }
}
        
    ############### Convert to excel #####################


    ### Set input and output path
    $inputCSV = $outfile

    ### Create a new Excel Workbook with one empty sheet
    $excel = New-Object -ComObject excel.application 
    $workbook = $excel.Workbooks.Add(1)
    $worksheet = $workbook.worksheets.Item(1)

    ### Build the QueryTables.Add command
    ### QueryTables does the same as when clicking "Data » From Text" in Excel
    $TxtConnector = ("TEXT;" + $outfile)
    $Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
    $query = $worksheet.QueryTables.item($Connector.name)

    ### Set the delimiter to ","
    $query.TextFileOtherDelimiter = ","

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
    # remove csv File
    Remove-Item $outfile


