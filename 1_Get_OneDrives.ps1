#PS script to get all OneDrive site IDs
$clientId = ""
$clientSecret = ""
$tenantId = ""
$resource = "https://graph.microsoft.com"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$Output = "C:\Temp\OD_SiteList_"+$timestamp+".csv"
#Prepare the authentication package
$body = @{
    client_id = $clientId
    client_secret = $clientSecret
    resource = $resource
    grant_type = "client_credentials"
}
#Get the Access Token
$access_token = (Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/token" -Body $body).access_token


# Connect to Microsoft Graph
$headers = @{
    "Authorization" = "Bearer $access_token"
}

$filter = "webUrl contains 'personal'"
 #Get all OneDrive for Business sites
$onedriveSites = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites?$filter&$select=id,webUrl" -Headers $headers 
$Sites = @()
$Sites+= $onedriveSites.value
#loop through pages
$Pages = $onedriveSites.'@odata.nextLink'

while($null -ne $Pages){
	Write-Warning "Checking next page results"
	$NewSites = Invoke-RestMethod -Uri $Pages -Headers $headers	
	if($Pages){
		$Pages=$NewSites.'@odata.nextLink'
	}
	$Sites+=$NewSites.value
}
# Create an array to store the Site URL and Site ID
$sitesArray = @()
# Iterate through each OneDrive site
foreach ($site in $Sites) {
	if($site.webUrl -like '*personal*'){ 
        #SiteID is returned with 3 values, the base URL, the parent site collection ID and the Site ID, the Site ID value is the middle value returned
        $SiteID = $Site.ID -split "," | Select-Object -Index 1
         # Add the Site URL and Site ID to the array using a custom object
         $sitesArray += New-Object PSObject -Property @{
            SiteURL = $site.WebUrl;
            SiteID  = $SiteID;
         }
		}
	}
$sitesarray | export-csv -path $output -NoTypeInformation   
Write-Host "All OneDrives and Site ID's were written to "$output 

