<#  

    Auto Licensing

    Requirements:  Saved Service Credentials

    Author: Rachel Moorehead

 
 	Usage:  Load script into session.  May need to dot-invoke: . ./uga_licensing.ps1
 			License-Users -ResetAll
 			License-Users -Unlicensed
 			
 	Be sure to change the tenant name in the license variables: yourdomain
#>



<#

    Functions (Subroutines)

#>

Function Check-Connection{

    $testConnection = Get-MsolCompanyInformation -ErrorAction SilentlyContinue
    
    If(!($testConnection)){

        # Create a Connection to Microsoft 
        # To create the stored password,
        #$securestring = convertto-securestring "password" -AsPlainText -Force;
        #$credential = New-Object System.Management.Automation.PsCredential "username",$securestring;
        #$credential.Password | ConvertFrom-SecureString | Set-Content C:\PowerShell\license_account_cred.txt;

        $password = Get-Content C:\PowerShell\license_account_cred.txt | ConvertTo-SecureString

        $cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist "username",$password

        Import-Module MSOnline

        Connect-MsolService -Credential $cred
    }

}


Function Assign-Licenses{

    $type = $args[0]
    $users = $args[1]
	$log = $args[2]
	
	$license = "yourdomain:STANDARDWOFFPACK_STUDENT"
    $disabledPlans = "SHAREPOINTWAC_EDU","SHAREPOINTSTANDARD_EDU","MCOSTANDARD"
    $EO1Plan = "yourdomain:EXCHANGESTANDARD_STUDENT"

    $customOptions = New-MsolLicenseOptions -AccountSkuId $license -DisabledPlans $disabledPlans 
	#Write-Output "Created License Options" >> $log
	
    $rerun_users = @()
    
    $users | foreach { 
                            try{ 
                                    Check-Connection
									$u = $_.UserPrincipalName
									#Write-Output "Checked connection: $u" >> $log
                                    Set-MsolUser -UserPrincipalName $u -UsageLocation US -PasswordNeverExpires $true
									#Write-Output "Set UsageLocation: $u" >> $log
                                    if($type -eq "Reset"){
                                    
                                        $currentLicenses = (Get-MsolUser -UserPrincipalName $u).Licenses
										#Write-Output "Current licenses for $u : $currentLicenses" >> $log
                                        Set-MsolUserLicense -UserPrincipalName $u -AddLicenses $license -LicenseOptions $ugaOptions -RemoveLicenses $currentLicenses.AccountSkuId -ea 1
										Write-Output "Set License: $u" >> $log
                                    }
                                    elseif($type -eq "Unlicensed"){

                                        Set-MsolUserLicense -UserPrincipalName $u -AddLicenses $license -LicenseOptions $customOptions
										Write-Output "Set License: $u" >> $log
                                    }
                                    else{

                                        "Fatal Error with Script: $u" >> $log
                                    }
                            }
                            catch{
                                    $rerun_users += @($u);
									Write-Output "Error with Setting License: $u" >> $log
                            }
                   }

    # Report Reruns  
    if($rerun_users){
        $count = $rerun_users.Length  
        Write-Output "$count to rerun. Rerun:" >> $log
        $rerun_users | Format-Table >> $log
    }

}

<#

    Main Function

#>

Function License-Users {
    param(
        [parameter(ParameterSetName="Reset",HelpMessage="Use to reset all users back to Default SKU assignments (EO Only).")][switch]$ResetAll,
        [parameter(ParameterSetName="Unlicensed",HelpMessage="Use to assign licenses to Unlicensed Users only.")][switch]$Unlicensed)

        <#

            Global Variables

        #>
        $date = Get-Date -UFormat "%m%d%Y%H%M%S"
        $log = "licensing_$date.txt"

        Check-Connection

        if($ResetAll){

            #Load Users
            $allUsers = Get-MsolUser -All
            $num = $allUsers.Length
            Write-Output "Retrieved number of users: $num" >> $log
            #$allUsers | ft >> $log

            #Process Users
            Assign-Licenses "Reset" $allUsers $log

            Write-Host "Completed reset process.  Please review output.  Press Enter to exit."
            Read-Host
            
        
        }

        elseif($Unlicensed){

            # Load Users from O365
            $unlicensedUsers = Get-MsolUser -All -UnlicensedUsersOnly
            $num = $unlicensedUsers.Length
            Write-Output "Retrieved number of unlicensed users from O365: $num" >> $log
            #$unlicensedUsers | ft >> $log

            # Process Users
            Assign-Licenses "Unlicensed" $unlicensedUsers $log

        }

        else{
        
            Write-Output "Sorry, incorrect input.  Please review the usage instructions for this function."
        }

}
