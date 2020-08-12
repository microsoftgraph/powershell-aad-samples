
<#


.COPYRIGHT

Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.

See LICENSE in the project root for license information.


#>


function Get-AuthToken 
{
    <#

    .SYNOPSIS

    This function is used to authenticate with the MS Graph.

    #>


    param
    (
        [string]
        [Parameter(Mandatory=$true)]
        $TenantId
    )

    Write-Host "Checking for AzureAD module..."

    $AadModule = Get-Module -Name "AzureAD" -ListAvailable

    if ($null -eq $AadModule)
    {
        Write-Host "`r`n"

        Write-Host "AzureAD Powershell module is not installed." -f Red

        Write-Host "`r`n"

        Write-Host "Install by running 'Install-Module AzureAD' from an elevated PowerShell prompt." -f Yellow

        exit
    }

    #
    # If the AzureAD PowerShell module count is greater than 1, then find the latest version.
    #
    if($AadModule.count -gt 1)
    {
        $Latest_Version = ($AadModule | select version | Sort-Object)[-1]

        $AadModule = $AadModule | ? { $_.version -eq $Latest_Version.version }

        #
        # Checking if there are multiple modules of the latest version.
        #
        if($AadModule.count -gt 1)
        {
            $AadModule = $AadModule | select -Unique
        }
    }
    
    #
    # Getting path to Active Directory assemblies.
    #
    $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"

    $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"

    #
    # Load Active Directory assemblies.
    #
    [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null

    [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null


    #
    # Well-Known client ID for PowerShell.
    #
    $clientId = "1b730954-1685-4b74-9bfd-dac224a7b894"

    $redirectUri = "urn:ietf:wg:oauth:2.0:oob"

    $resourceAppIdURI = "https://graph.microsoft.com"

    $authority = "https://login.microsoftonline.com/$TenantId"

    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority

    $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Always"

    $authResult = $authContext.AcquireTokenAsync($resourceAppIdURI, $clientId, $redirectUri, $platformParameters).Result

    #
    # If the access token is valid, then create the authorization header.
    #
    if($null -ne $authResult.AccessToken)
    {
        $authHeader = @{
            'Content-Type'='application/json'

            'Authorization'="Bearer " + $authResult.AccessToken

            'ExpiresOn'=$authResult.ExpiresOn
        }

        return $authHeader
    }
    else
    {
        Write-Host "`r`n"
        
        Write-Host "Authorization access token is null." -f Red
       
        exit
    }
}

function ReassignSensitivityLabelToO365Groups
{

    <#

    .SYNOPSIS

    This function is used to reassign the given label to associated O365 groups in order to resolve conflicts due to label policy updates.

    #>


    param
    (
        [string]
        [Parameter(Mandatory=$true)]
        $LabelId,

        [hashtable]
        [Parameter(Mandatory=$true)]
        $AuthHeader,

        [string]
        [Parameter(Mandatory=$true)]
        $LogFile
    )

    $retryInterval = 5

    $hostname = "https://graph.microsoft.com/beta"
    $queryPageSize = 20
    $contentType = "application/json"

    $groupsQueryUriBase = $hostname + "/groups?"
    $groupsFilterByAssignedLabelQuery = "`$top=$($queryPageSize)&`$select=id,displayName&`$filter=assignedLabels/any(x:x/labelId+eq+'$($LabelId)')"

    while ($true)
    {
        $groupsFilterQueryUri = $groupsQueryUriBase + $groupsFilterByAssignedLabelQuery

        #
        # If a graph call fails, then the same operation will be retried for 3 times.
        #
        $retries = 3
        $retryCount = 0

        do
        {
            try
            {
                #
                # GET request to filter groups by assigned label.
                #
                $groupsFilterQueryResult = Invoke-RestMethod -Uri $groupsFilterQueryUri -Headers $AuthHeader -Method Get
                break
            }
            catch
            {
                if ($retries-- -gt 0)
                {
                    $retryCount++

                    Write-Warning "Filter query to get groups by assigned label has failed and will be retried within $($retryInterval) seconds. Retry attempt: $($retryCount)"
                    Write-Host "`r`n"

                    Start-Sleep -Seconds $retryInterval
                }
                else
                {
                    Write-Host $_.Exception.Message -f Red

                    if ($null -ne $_.Exception.Response)
                    {
                        $responseStream = $_.Exception.Response.GetResponseStream()
                        $streamReader = New-Object System.IO.StreamReader($responseStream)
                        $responseBody = $streamReader.ReadToEnd()

                        if (![string]::IsNullOrEmpty($responseBody))
                        {
                            $responseObject = $responseBody | ConvertFrom-Json

                            if ($null -ne $responseObject -and $null -ne $responseObject.'error')
                            {
                                Write-Host "Error Code: " $responseObject.'error'.code -f Red
                                Write-Host "Error Message: " $responseObject.'error'.message -f Red
                            }
                        }
                    }

                    Write-Host "`r`n"
                    throw
                }
            }
        }
        while ($true)


        $groups = $groupsFilterQueryResult.value

        foreach ($group in $groups)
        {
            $groupUri = $hostname + "/groups/" + $group.id

            $requestBody = @{
                assignedLabels = @(
                    @{
                        labelId = $($LabelId)
                     }
                )
            }

            $requestBody = $requestBody | ConvertTo-Json


            $retries = 3
            $retryCount = 0

            do{
                try
                {
                    #
                    # PATCH request to reassign the given label to one of the associated groups.
                    #
                    Invoke-RestMethod -Uri $groupUri -Headers $AuthHeader -Method Patch -Body $requestBody -ContentType $contentType

                    Write-Host "The label has been reassigned to the following group object sucessfully:" -f Green
                    Write-Host "Group Object Id: " $group.id -f Green
                    Write-Host "Group Display Name: " $group.displayName -f Green
                    Write-Host "`r`n"

                    break
                }
                catch
                {
                    if ($retries-- -gt 0)
                    {
                        $retryCount++

                        Write-Warning "Reassignment of the label has failed and will be retried for the following group object within $($retryInterval) seconds. Retry attempt: $($retryCount)"
                        Write-Host "Group Object Id: " $group.id -f Yellow
                        Write-Host "Group Display Name: " $group.displayName -f Yellow
                        Write-Host "`r`n"

                        Start-Sleep -Seconds $retryInterval
                    }
                    else
                    {
                        Write-Host "Reassignment of the label has failed on all attempts for the following group object:" -f Red
                        Write-Host "Group Object Id: " $group.id -f Red 
                        Write-Host "Group Display Name: " $group.displayName -f Red
                        Write-Host "`r`n"
                        Write-Host $_.Exception.Message -f Red

                        #
                        # Write the group information and the exception message into a file.
                        #
                        Add-Content $LogFile -Value "Reassignment of the label has failed on all attempts for the following group object:"
                        Add-Content $LogFile -Value "Group Object Id: $($group.id)"
                        Add-Content $LogFile -Value "Group Display Name: $($group.displayName)"
                        Add-Content $LogFile -Value "`n"
                        Add-Content $LogFile -Value $_.Exception.Message

                        if ($null -ne $_.Exception.Response)
                        {
                            $responseStream = $_.Exception.Response.GetResponseStream()
                            $streamReader = New-Object System.IO.StreamReader($responseStream)
                            $responseBody = $streamReader.ReadToEnd()

                            if ($null -ne $responseBody)
                            {
                                $responseObject = $responseBody | ConvertFrom-Json

                                if ($null -ne $responseObject -and $null -ne $responseObject.'error')
                                {
                                    $errorCode = "Error Code: " + $responseObject.'error'.code
                                    $errorMessage = "Error Message: " + $responseObject.'error'.message

                                    Write-Host $errorCode -f Red
                                    Write-Host $errorMessage -f Red

                                    Add-Content $LogFile -Value $errorCode
                                    Add-Content $LogFile -Value $errorMessage
                                }
                            }
                        }

                        Add-Content $LogFile -Value "-------------------------------------------------------------------------------------"
                        Add-Content $LogFile -Value "`n`n"

                        Write-Host "`r`n"
                        break
                    }
                }
            }while ($true)
        }

        if ($null -eq $groupsFilterQueryResult.'@odata.nextLink')
        {
            break
        }

        $nextPageIndex = $groupsFilterQueryResult.'@odata.nextLink'.IndexOf('$')

        $groupsFilterByAssignedLabelQuery = $groupsFilterQueryResult.'@odata.nextLink'.Substring($nextPageIndex)
    }
}

function Main
{
    
    <#

    .SYNOPSIS

    This function is used to:
        
        1- Get an access token for MS Graph.
        
        2- Reassign the given label to associated O365 groups.

    .PARAMETER TenantId

    Specifies the identifier of the target tenant.

    .PARAMETER LabelId

    Specifies the identifier of a label in the Security and Compliance Center.

    .PARAMETER LogFile

    Specifies the path to a file that would be used to log errors during reassignment of labels to groups.

    .EXAMPLE

    PS> Main -TenantId "00000000-0000-0000-0000-000000000000" -LabelId "00000000-0000-0000-0000-000000000000" -LogFile "C:\errors.txt"

    #>
    

    param
    (
        [string]
        [Parameter(Mandatory=$true)]
        $TenantId,

        [string]
        [Parameter(Mandatory=$true)]
        $LabelId,

        [string]
        [Parameter(Mandatory=$true)]
        $LogFile
    )

    $authHeader = Get-AuthToken -TenantId $TenantId

    ReassignSensitivityLabelToO365Groups -LabelId $LabelId -AuthHeader $authHeader -LogFile $LogFile
}

Main
