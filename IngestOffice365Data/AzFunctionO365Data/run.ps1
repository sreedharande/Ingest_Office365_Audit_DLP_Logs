# Input bindings are passed in via param block.
param($Timer)

# Get the current universal time in the default string format
$currentUTCtime = (Get-Date).ToUniversalTime()

# The 'IsPastDue' porperty is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}

# Write an information log with the current time.
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"

# Main
if ($env:MSI_SECRET -and (Get-Module -ListAvailable Az.Accounts)){
    Connect-AzAccount -Identity
}

#region Environment Variables

$Office365ContentTypes = $env:contentTypes
$Office365RecordTypes = $env:recordTypes
$AADAppClientId = $env:clientId
$AADAppClientSecret = $env:clientSecret 
$AADAppClientDomain = $env:domain
$AADAppPublisher = $env:publisher
$AzureTenantId = $env:tenantGuid
$AzureAADLoginUri = $env:AzureAADLoginUri
$OfficeLoginUri = $env:OfficeLoginUri
$DceEndPointUri = $env:DceEndPointUri
$DcrImmutableId = $env:DcrImmutableId
$StreamName = $env:streamName

#Azure Function State management between Executions
$AzureWebJobsStorage =$env:AzureWebJobsStorage  #Storage Account to use for table to maintain state for log queries between executions

#endregion


Function Send-DataToDCE {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)] $JsonPayload,        
		[parameter(Mandatory = $true)] $DceURI,
        [parameter(Mandatory = $true)] $DcrId,
        [parameter(Mandatory = $true)] $CustomTableName
    )

    $AccessToken = Get-AzMonitorBearerToken $AADAppClientId $AADAppClientSecret $AzureTenantId

    # Initialize Headers and URI for POST request to the Data Collection Endpoint (DCE)
    $headers = @{"Authorization" = "Bearer $AccessToken"; "Content-Type" = "application/json"}
    $uri = "$DceURI/dataCollectionRules/$DcrId/streams/$CustomTableName`?api-version=2023-01-01"    
   
    $payload_size = ([System.Text.Encoding]::UTF8.GetBytes($JsonPayload).Length)
    If ($payload_size -le 1mb) {
        Write-Host "Sending log events with size $dataset_size" 
        Try {
            # Sending data to Data Collection Endpoint (DCE) -> Data Collection Rule (DCR) -> Azure Monitor table
            $IngestionStatus = Invoke-RestMethod -Uri $uri -Method "POST" -Body $JsonPayload -Headers $headers -verbose
            Write-Host "Status : $IngestionStatus"
        }
        catch {
            Write-Host "Error occured in Send-DataToDCE :$($_)"
        }        
    }
    else {
        # Maximum size of API call: 1MB for both compressed and uncompressed data
        Write-Host "Log size is greater than APILimitBytes"
    }    
}

function Get-Office365AuthToken {
	[cmdletbinding()]
	Param(
		[Parameter(Mandatory = $true, Position = 0)]
		[string]$ClientID,
		[parameter(Mandatory = $true, Position = 1)]
		[string]$ClientSecret,
		[Parameter(Mandatory = $true, Position = 2)]
		[string]$TenantDomain,
		[Parameter(Mandatory = $true, Position = 3)]
		[string]$TenantGUID
	)
	try {
		$body = @{grant_type="client_credentials";resource=$OfficeLoginUri;client_id=$ClientID;client_secret=$ClientSecret}
		$oauth = Invoke-RestMethod -Method Post -Uri $AzureAADLoginUri/$TenantDomain/oauth2/token?api-version=1.0 -Body $body
		$headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
		return $headerParams
	}
	catch {
		Write-Host "Error occured in Get-Office365AuthToken :$($_)"
        exit
	}
	 
}

Function Get-AzMonitorBearerToken {
	[cmdletbinding()]
	Param(
		[Parameter(Mandatory = $true, Position = 0)]
		[string]$ClientID,
		[parameter(Mandatory = $true, Position = 1)]
		[string]$ClientSecret,		
		[Parameter(Mandatory = $true, Position = 2)]
		[string]$TenantGUID
	)
    Try {
        Add-Type -AssemblyName System.Web        
        $scope = [System.Web.HttpUtility]::UrlEncode("https://monitor.azure.com//.default")
        $body = "client_id=$ClientID&scope=$scope&client_secret=$ClientSecret&grant_type=client_credentials";
        $headers = @{"Content-Type" = "application/x-www-form-urlencoded"};
        $uri = "$AzureAADLoginUri/$TenantGUID/oauth2/v2.0/token"
        $bearerToken = (Invoke-RestMethod -Uri $uri -Method "POST" -Body $body -Headers $headers).access_token
        return $bearerToken
    }
    catch {
        Write-Host "Error occured in Get-AzMonitorBearerToken :$($_)"
        exit
    }
}

function Convert-ObjectToHashTable {
	[CmdletBinding()]
	param
	(
		[parameter(Mandatory=$true,ValueFromPipeline=$true)]
		[pscustomobject] $Object
	)
	$HashTable = @{}
	$ObjectMembers = Get-Member -InputObject $Object -MemberType *Property
	foreach ($Member in $ObjectMembers) 
	{
		$HashTable.$($Member.Name) = $Object.$($Member.Name)
	}
	return $HashTable
}																																				   

function Get-O365Data{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$startTime,
        [parameter(Mandatory = $true, Position = 1)]
        [string]$endTime,
        [Parameter(Mandatory = $true, Position = 2)]
        [psobject]$headerParams,
        [parameter(Mandatory = $true, Position = 3)]
        [string]$tenantGuid
    )
    #List Available Content
    $contentTypes = $Office365ContentTypes.split(",")
    #Loop for each content Type like Audit.General;
	
	#API front end for GCC-High is “manage.office365.us” instead of the commercial “manage.office.com”. 
	if ($OfficeLoginUri.split('.')[2] -eq "us") {
		$OfficeLoginUri = "https://manage.office365.us"
	}
	
	#Loop for each content Type like Audit.General; DLP.ALL
    foreach($contentType in $contentTypes){
		$contentType = $contentType.Trim()
        $listAvailableContentUri = "$OfficeLoginUri/api/v1.0/$tenantGuid/activity/feed/subscriptions/content?contentType=$contentType&PublisherIdentifier=$AADAppPublisher&startTime=$startTime&endTime=$endTime"        
		Write-Output $listAvailableContentUri
		
		do {
            #List Available Content
            $contentResult = Invoke-RestMethod -Method GET -Headers $headerParams -Uri $listAvailableContentUri
            Write-Output $contentResult.Count
            #Loop for each Content
            foreach($obj in $contentResult){
                #Retrieve Content
                $data = Invoke-RestMethod -Method GET -Headers $headerParams -Uri ($obj.contentUri)                
                #Loop through each Record in the Content
                foreach($event in $data){
                    #Filtering for Recrord types
                    #Get all Record Types
                    if($Office365RecordTypes -eq "0"){
                        #We dont need Cloud App Security Alerts due to MCAS connector
                        if(($event.Source) -ne "Cloud App Security") {
                            $ht = Convert-ObjectToHashTable $event     
                            $ht = $ht | ConvertTo-Json -Depth 5                            
                            $arrayOfObjects = "["+$ht+"]"                                                      
                            Send-DataToDCE -JsonPayload $arrayOfObjects -DceURI $DceEndPointUri -DcrId $DcrImmutableId -CustomTableName $StreamName
                        }
                    }
                    else {
                        #Get only certain record types
                        $types = ($Office365RecordTypes).split(",")
                        if(($event.RecordType) -in $types){
                            #We dont need Cloud App Security Alerts due to MCAS connector
                            if(($event.Source) -ne "Cloud App Security") {
                                $ht = Convert-ObjectToHashTable $event     
                                $ht = $ht | ConvertTo-Json -Depth 5                            
                                $arrayOfObjects = "["+$ht+"]"                                                      
                            	Send-DataToDCE -JsonPayload $arrayOfObjects -DceURI $DceEndPointUri -DcrId $DcrImmutableId -CustomTableName $StreamName
                            }
                        }
                        
                    }
                }
            }
            
            #Handles Pagination
            $nextPageResult = Invoke-WebRequest -Method GET -Headers $headerParams -Uri $listAvailableContentUri
            If($null -ne ($($nextPageResult.Headers.NextPageUrl))){
                $nextPage = $true
                $listAvailableContentUri = $nextPageResult.Headers.NextPageUrl
            }
            Else {
				$nextPage = $false
			}
        } until ($nextPage -eq $false)
    }

	#add last run time to ensure no missed packages
	$endTime = $currentUTCtime | Get-Date -Format yyyy-MM-ddThh:mm:ss
	Add-AzTableRow -table $o365TimeStampTbl -PartitionKey "Office365" -RowKey "lastExecutionEndTime" -property @{"lastExecutionEndTimeValue"=$endTime} -UpdateExisting
}

#region Driver Program
# Retrieve Timestamp from last records received 
# Check if Table has already been created and if not create it to maintain state between executions of Function
$storageAccountContext = New-AzStorageContext -ConnectionString $AzureWebJobsStorage
$storageAccountTableName = "O365ApiExecutions"
$StorageTable = Get-AzStorageTable -Name $storageAccountTableName -Context $storageAccountContext -ErrorAction Ignore
if($null -eq $StorageTable.Name){  
    $startTime = $currentUTCtime.AddHours(-24) | Get-Date -Format yyyy-MM-ddThh:mm:ss
	New-AzStorageTable -Name $storageAccountTableName -Context $storageAccountContext
    $o365TimeStampTbl = (Get-AzStorageTable -Name $storageAccountTableName -Context $storageAccountContext.Context).cloudTable    
    Add-AzTableRow -table $o365TimeStampTbl -PartitionKey "Office365" -RowKey "lastExecutionEndTime" -property @{"lastExecutionEndTimeValue"=$startTime} -UpdateExisting
}
Else {
    $o365TimeStampTbl = (Get-AzStorageTable -Name $storageAccountTableName -Context $storageAccountContext.Context).cloudTable
}
# retrieve the last execution values
$lastExecutionEndTime = Get-azTableRow -table $o365TimeStampTbl -partitionKey "Office365" -RowKey "lastExecutionEndTime" -ErrorAction Ignore

$startTime = $($lastExecutionEndTime.lastExecutionEndTimeValue)
$endTime = $currentUTCtime | Get-Date -Format yyyy-MM-ddThh:mm:ss

$O365Params = Get-Office365AuthToken $AADAppClientId $AADAppClientSecret $AADAppClientDomain $AzureTenantId
Get-O365Data $startTime $endTime $O365Params $AzureTenantId

#Updating
$endTime = $currentUTCtime | Get-Date -Format yyyy-MM-ddThh:mm:ss
Add-AzTableRow -table $o365TimeStampTbl -PartitionKey "Office365" -RowKey "lastExecutionEndTime" -property @{"lastExecutionEndTimeValue"=$endTime} -UpdateExisting
#endregion

# Write an information log with the current time.
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"