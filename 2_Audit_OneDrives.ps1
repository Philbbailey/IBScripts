#Get the Input File of sites with SiteID column
$InputFile = Read-Host -Prompt “Enter the input file name and path”
$validPath = Test-Path $InputFile -IsValid
If ($ValidPath -eq $True) {Write-Host "You selected "$inputfile}
Else 
{Write-Host "The file you selected was not found"}
$Allsites = Import-CSV $InputFile

# Set the base URL for the SharePoint site
$base_url = "https://EDIT_ME-admin.sharepoint.com"

# Set the tenant ID, client ID, and client secret
$tenant_id = ""
$client_id = ""
$client_secret = ""

# Prompt the user for their username and password
$username = Read-Host -Prompt "Enter your username"
$password = Read-Host -Prompt "Enter your password" -AsSecureString


# Get an access token using the resource owner password credentials flow
$body = @{
    "grant_type" = "password"
    "client_id" = $client_id
    "client_secret" = $client_secret
    "username" = $username
    "password" = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
    "resource" = $base_url
}
$response = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenant_id/oauth2/token" -Body $body
$access_token = $response.access_token

# Set the headers with the access token
$headers = @{
    "Authorization" = "Bearer $access_token"
    "Accept" = "application/json;odata=verbose"
}

$start_row = 0

# Create a new Excel file
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

# Write the header row
$worksheet.Cells.Item(1, 1) = "IBMode"
$worksheet.Cells.Item(1, 2) = "URL"
$worksheet.Cells.Item(1, 3) = "Owner"
$worksheet.Cells.Item(1, 4) = "IBSegments"

$row = 2

$progressCounter = 0
$totalSites = $Allsites.Count
    # Iterate over each site and get more information
    foreach ($site in $allsites) {
        $site_id = $site.SiteID
        $response = Invoke-RestMethod -Uri ($base_url + "/_api/SPO.Tenant/sites('" + $site_id + "')") -Headers $headers
        $site_info = $response.d
        $IBMode     = if ($site_info.IBMode) { $site_info.IBMode } else { "" }
        $Owner      = if ($site_info.Owner) { $site_info.Owner } else { "" }
        $IBSegments = if ($site_info.IBSegments.results) { [string]::Join(",", $site_info.IBSegments.results) } else { "" }
        $URL        = if ($site_info.Url) { $site_info.Url } else { "" }
        $progressCounter++
        # Write the site information to the Excel file as we go through
        $worksheet.Cells.Item($row, 1) = $IBMode
        $worksheet.Cells.Item($row, 2) = $URL
        $worksheet.Cells.Item($row, 3) = $Owner
        $worksheet.Cells.Item($row, 4) = $IBSegments
        $row++
    }
    Write-Progress -Activity "Processing sites" -Status "$progressCounter of $totalSites completed" -PercentComplete (($progressCounter / $totalSites) * 100)
    # Increment the start row for the next batch of results
    $start_row += $row_limit
# Save and close the Excel file
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputfile = "C:\Temp\OD_Site_Audit_"+$timestamp+".xlsx"
$workbook.SaveAs($outputfile)
$workbook.Close()
$excel.Quit()
