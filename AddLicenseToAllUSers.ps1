#This script has very limited error checking and didn't change all licenses correctly so lots of manual changes need to be made
#Some Apps are not the same name in E1 and E3 so this script only deals with one license type at a time.

#gets all users with E3 licenses assigned - change to STANDARDPACK for E1
$users = Get-MsolUser -All | Where-Object {($_.Licenses).AccountSkuId -match "ENTERPRISEPACK"}

foreach ($user in $users)
{
    Write-Host "Checking " $user.UserPrincipalName -foregroundcolor "Cyan"
    $CurrentSku = $user.Licenses.Accountskuid
    #If more than one SKU, Have to check them all!  
    if ($currentSku.count -gt 1)  
    {  
        Write-Host $user.UserPrincipalName "Has Multiple SKU Assigned. Checking all of them" -foregroundcolor "White"
        for($i = 0; $i -lt $currentSku.count; $i++)
        {       
            #Loop trough Each SKU to see if one of their services Is TEAMS1
            if($user.Licenses[$i].ServiceStatus.ServicePlan.ServiceName -match "TEAMS1" )
                {
                    Write-host $user.Licenses[$i].AccountSkuid "has POWERAPPS. " -foregroundcolor "Yellow"
                    #Check it is disabled
                    if (($user.Licenses[$i].ServiceStatus | Where-Object { $_.ServicePlan.ServiceName.equals("TEAMS1") }).ProvisioningStatus -eq "Disabled" )
                    {   
                        Write-Host "Confirmed that has TEAMS1 disabled in their license."

                        # Build a hashtable with the service plan names and statuses for this person's license
                        Write-Host "Building hash table of service plans and current statuses..."
                        $plans = @{}
                        $user.Licenses[$i].ServiceStatus | ForEach-Object { $plans.add($_.serviceplan.servicename, $_.provisioningstatus) }
                        $plans.set_item("TEAMS1","Success")

                        $disabledPlans = @()  # this is a list (array of strings) variable
                        #create a list of disabled plans based on our change
                        $plans.Keys | ForEach-Object { if ($plans.get_item($_) -eq "Disabled") { $disabledPlans += $_ } }
                        
                        #If our change means no plans disabled then there doesn't need to be a disabled plans command
                        if ($disabledPlans) {
                            Write-host "Creating new msolLicenseOptions object for $userPrincipalName"
                            $NewSkU = New-MsolLicenseOptions -AccountSkuId $user.Licenses[$i].AccountSkuid -DisabledPlans $disabledPlans -verbose
                        } else {
                            Write-host "NOTE: All service plans will now be enabled for this user."
                            $NewSkU = New-MsolLicenseOptions -AccountSkuId $user.Licenses[$i].AccountSkuid
                        }

                        #set the new license options
                        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -LicenseOptions $NewSkU
                        Write-Host "TEAMS1 Enabled for " $user.UserPrincipalName " On SKU " $user.Licenses[$i].AccountSkuid -foregroundcolor "Green"
                    }

                }
                else
                {
                    Write-host $user.Licenses[$i].AccountSkuid " doesn't have TEAMS1, Skip" -foregroundcolor "Magenta"
                }
            }
    }  
    else  
    {  
        if (($user.Licenses.ServiceStatus | Where-Object { $_.ServicePlan.ServiceName.equals("TEAMS1") }).ProvisioningStatus -eq "Disabled" )
        {
            Write-Host "Confirmed that has TEAMS1 disabled in their license."

            # Build a hashtable with the service plan names and statuses for this person's license
            Write-Host "Building hash table of service plans and current statuses..."
            $plans = @{}
            $user.Licenses.ServiceStatus | ForEach-Object { $plans.add($_.serviceplan.servicename, $_.provisioningstatus) }
            $plans.set_item("TEAMS1","Success")

            $disabledPlans = @()  # this is a list (array of strings) variable
            $plans.Keys | ForEach-Object { if ($plans.get_item($_) -eq "Disabled") { $disabledPlans += $_ } }

            if ($disabledPlans) {
                Write-host "Creating new msolLicenseOptions object for $userPrincipalName"
                $NewSkU = New-MsolLicenseOptions -AccountSkuId $user.Licenses.AccountSkuid -DisabledPlans $disabledPlans -verbose
            } else {
                Write-host "NOTE: All service plans will now be enabled for this user."
                $NewSkU = New-MsolLicenseOptions -AccountSkuId $user.Licenses.AccountSkuid
            }

            Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -LicenseOptions $NewSkU
            Write-Host "TEAMS1 Enabled for " $user.UserPrincipalName " On SKU " $user.Licenses.AccountSkuid -foregroundcolor "Green"
        }
    }  
}
