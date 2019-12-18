
<#

.COPYRIGHT
Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
See LICENSE in the project root for license information.

#>

####################################################

install-module AzureAD
Try{$aadConnectionDetail = Get-AzureADTenantDetail}
Catch
{
    connect-AzureAD
}

$csvPath = Read-Host -Prompt "Enter csv file path"
$csv = Import-Csv -path $csvPath
$outputLogPath = Read-Host -Prompt "Enter output log file path"

$inputCount = 0
$updatedCount = 0
$notFoundCount = 0
$failedToUpdateCount=0

if ([System.IO.File]::Exists($outputLogPath))
{
    $overwrite = Read-Host -Prompt "Log file already exists. Do you want to overwrite? yes|no"
    if ($overwrite -like "yes")
    {
        Remove-Item $outputLogPath
    }
    else
    {
        exit
    }
}

Add-Content $outputLogPath "UserName,OfficeLocation";

foreach ($inputUser in $csv)
{
    if ([string]::IsNullOrEmpty($inputUser.UserName))
    {
        continue
    }

    $inputCount = $inputCount + 1
    
    Try{$fetchedUser = Get-AzureADUser -ObjectId $inputUser.UserName}
    Catch
    {
        $notFoundCount = $notFoundCount + 1
        continue
    }

    $existingLocation = $fetchedUser.PhysicalDeliveryOfficeName
    
    Try{Set-AzureADUser -ObjectId $inputUser.UserName -PhysicalDeliveryOfficeName $inputUser.OfficeLocation}
    Catch
    {
        $failedToUpdateCount = $failedToUpdateCount + 1
    }
    
    $updatedCount = $updatedCount + 1
    
    $log = "{0},{1}" -f $inputUser.UserName,$existingLocation
    $log | add-content -path $outputLogPath
}

Write-Host "Input Users count: " $inputCount
Write-Host "Updated Users count: " $updatedCount
Write-Host "Failed to update Users count: " $failedToUpdateCount
Write-Host "Not found Users count: " $notFoundCount
