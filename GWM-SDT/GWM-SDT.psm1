Function Connect-Office365 {
    <#
    .SYNOPSIS
        Function connects to all Office 365 services, or allows connection to individual services.

    .DESCRIPTION
        Get-SendOnBehalfPermissions shows what users have Send On Behalf rights on a specific mailbox.
    
    .PARAMETER DemoParam1
        The parameter DemoParam1 is used to define the value of blah and also blah.

    .PARAMETER DemoParam2
        The parameter DemoParam2 is used to define the value of blah and also blah.
    
    .EXAMPLE
        This example connects to Azure AD
        Connect-Office365 -Service AzureAD
        
        Connect to the Exchange Online PowerShell Module
        Connect-Office365 -Service ExO

        Connect to a service while also supplying your administration user principal name.
        Connect-Office365 -Service PIM -UPN adm.jesse.newell@dbschenker.com
    
        
    .NOTES
        Author: Jesse Newell
        Last Edit: 2019-03-05
        Version 1.0
    
    #>
    [CmdletBinding()]
    [CmdletBinding(DefaultParameterSetName)]  # Add cmdlet features.
    Param (
        # Define parameters below, each separated by a comma

        [Parameter(Mandatory = $True, Position = 1)]
        [ValidateSet('PIM', 'AzureAD', 'ExO', 'MSOnline', 'SharePoint')]
        [string[]]$Service,    
        [Parameter(DontShow, Mandatory = $false, Position = 2, ParameterSetName = 'UserPrincipalName')]
        [String]$UPN
        # Add additional parameters here.


    )
    
    $getModuleSplat = @{
        ListAvailable = $True
        Verbose       = $False
    }
    
            
    ForEach ($item in $PSBoundParameters.Service) {

        Write-Verbose "Attempting to connect to $Item"  
        Switch ($item) {

            AzureAD {
                if ($null -eq (Get-Module @getmodulesplat -name "AzureAD")) {

                    Write-Error "Azure AD module is not present!"
                    Confirm-AzureStorSimpleLegacyVolumeContainerStatus
                    Continue

                }
                else {
                    $Connect = Connect-AzureAD
                    if ($null -eq $Connect) {
                        If (($host.ui.RawUI.windowtitle) -notlike "*Connected To:*") {
                            $host.ui.RawUI.WindowTitle += " - Connected To: AzureAD"

                        }
                        Else {
                            $host.UI.RawUI.WindowTitle += " - AzureAD"

                        }


                    }
                            
                }
                        
                Continue
            }
                
            ExO {
                       
                $PSExoPowershellModuleRoot = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select-Object -Last 1).DirectoryName

                If ($null -eq $PSExoPowershellModuleRoot) {
                    Write-Error "The Exchange Online MFA Module was not found! https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps"
                    continue
                }
                Else {
                    Write-Verbose "Importing Exchange MFA Module"
                    $ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellModule.dll";
                    $ModulePath = [System.IO.Path]::Combine($PSExoPowershellModuleRoot, $ExoPowershellModule);
                    Import-Module -verbose $ModulePath; 
                    $Office365PSSession = New-ExoPssession -userprincipalname $UPN -ConnectionUri "https://outlook.office365.com/powershell-liveid/" 
                    Import-Module (Import-PSSession $Office365PSSession -AllowClobber) -Global
                    Write-Verbose "Connecting to Exchange Online"
                    Remove-Variable $UPN
                }
                Continue

            }

            PIM {
                if ($null -eq (Get-Module @getmodulesplat -name "Microsoft.Azure.ActiveDirectory.PIM.PSModule")) {

                    Write-Error "PIM Service is not enabled!"
                    Continue
    
                }
                else {

                    $Connect = Connect-PimService
                    if ($null -eq $Connect) {
                        If (($host.ui.RawUI.windowtitle) -notlike "*Connected To:*") {
                            $host.ui.RawUI.WindowTitle += " - Connected To: PIM Service"
    
                        }
                        Else {
                            $host.UI.RawUI.WindowTitle += " - PIM Service"
                        }
                    }
                }
                Continue
            }


            MSOnline {
                if ($null -eq (Get-Module @getmodulesplat -name "MSOnline")) {

                    Write-Error "MSOnline Service is not enabled!"
                    Continue
    
                }
                else {

                    $Connect = Connect-MsolService
                    if ($null -eq $Connect) {
                        If (($host.ui.RawUI.windowtitle) -notlike "*Connected To:*") {
                            $host.ui.RawUI.WindowTitle += " - Connected To: MSOL Service"
    
                        }
                        Else {
                            $host.UI.RawUI.WindowTitle += " - MSOL Service"
    
                        }

                    }

                }
                Continue

            }

            SharePoint {
                if ($null -eq (Get-Module @getmodulesplat -name "Microsoft.Online.SharePoint.PowerShell")) {

                    Write-Error "SharePoint Service is not enabled!"
                    Continue
    
                }
                else {

                    $Connect = Connect-SPOService
                    if ($null -eq $Connect) {
                        If (($host.ui.RawUI.windowtitle) -notlike "*Connected To:*") {
                            $host.ui.RawUI.WindowTitle += " - Connected To: SharePoint Service"
    
                        }
                        Else {
                            $host.UI.RawUI.WindowTitle += " - SharePoint Service"
    
                        }

                    }

                }
                Continue

            }
        }

    }

        
} # End Function




Function Get-SendOnBehalfPermissions {
    <#
        .SYNOPSIS
            This advanced function obtains the Send on Behalf permissions enabled for a mailbox
        
        .DESCRIPTION
            Get-SendOnBehalfPermissions shows what users have Send On Behalf rights on a specific mailbox.
        
        .PARAMETER DemoParam1
            The parameter DemoParam1 is used to define the value of blah and also blah.
        
        .PARAMETER DemoParam2
            The parameter DemoParam2 is used to define the value of blah and also blah.
        
        .EXAMPLE
            The example below does blah
            PS C:\> <Example>
            
        .NOTES
            Author: Jesse Newell
            Last Edit: 2019-02-05
            Version 1.0 - initial release of blah
            Version 1.1 - update for blah
        
        #>
    [CmdletBinding()]  # Add cmdlet features.
    Param (
        # Define parameters below, each separated by a comma
        
        [Parameter(Mandatory = $True)]
        [int]$DemoParam1,
        
        [Parameter(Mandatory = $False)]
        [ValidateSet('Alpha', 'Beta', 'Gamma')]
        [string]$DemoParam2,
         
        # you don’t have to use the full [Parameter()] decorator on every parameter
        [string]$DemoParam3,
                
        [string]$DemoParam4, $DemoParam5
                
        # Add additional parameters here.
    )
        
    Begin {
        # Start of the BEGIN block.
        Write-Verbose -Message "Entering the BEGIN block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
        
        # Add additional code here.
                
    } # End Begin block
        
    Process {
        # Start of PROCESS block.
        Write-Verbose -Message "Entering the PROCESS block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
        
        
        # Add additional code here.
        
                
    } # End of PROCESS block.
        
    End {
        # Start of END block.
        Write-Verbose -Message "Entering the END block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
        
        # Add additional code here.
        
    } # End of the END Block.
} # End Function



Function Move-GitModule {
    <#
            .SYNOPSIS
                Advanced function to move the downloaded git module to the necessary path. Reducing the 

            .DESCRIPTION
                Get-SendOnBehalfPermissions shows what users have Send On Behalf rights on a specific mailbox.
            
            .PARAMETER DemoParam1
                The parameter DemoParam1 is used to define the value of blah and also blah.
            
            .PARAMETER DemoParam2
                The parameter DemoParam2 is used to define the value of blah and also blah.
            
            .EXAMPLE
                The example below does blah
                PS C:\> <Example>
                
            .NOTES
                Author: Jesse Newell
                Last Edit: 2019-02-05
                Version 1.0 - initial release of blah
                Version 1.1 - update for blah
            
            #>
    [CmdletBinding()]  # Add cmdlet features.
    Param (
        # Define parameters below, each separated by a comma
            
        [Parameter(Mandatory = $True)]
        [int]$DemoParam1,
            
        [Parameter(Mandatory = $False)]
        [ValidateSet('Alpha', 'Beta', 'Gamma')]
        [string]$DemoParam2,
             
        # you don’t have to use the full [Parameter()] decorator on every parameter
        [string]$DemoParam3,
                    
        [string]$DemoParam4, $DemoParam5
                    
        # Add additional parameters here.
    )
            
    Begin {
        # Start of the BEGIN block.
        Write-Verbose -Message "Entering the BEGIN block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
            
        # Add additional code here.
                    
    } # End Begin block
            
    Process {
        # Start of PROCESS block.
        Write-Verbose -Message "Entering the PROCESS block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
            
            
        # Add additional code here.
            
                    
    } # End of PROCESS block.
            
    End {
        # Start of END block.
        Write-Verbose -Message "Entering the END block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
            
        # Add additional code here.
            
    } # End of the END Block.
} # End Function


Function Enable-PIMElevation {
    <#
        .SYNOPSIS
            This advanced function elevates the users administration roles.
        
        .DESCRIPTION
            Enable the PIM elevation for cloud enviornment.
        
        .PARAMETER DemoParam1
            The parameter DemoParam1 is used to define the value of blah and also blah.
        
        .PARAMETER DemoParam2
            The parameter DemoParam2 is used to define the value of blah and also blah.
        
        .EXAMPLE
            The example below does blah
            PS C:\> <Example>
            
        .NOTES
            Author: Jesse Newell
            Last Edit: 2019-03-05
            Version 1.0 - initial release of Enable-PIMElevation
        
        #>
    [CmdletBinding()]  # Add cmdlet features.
    Param (
        # Define parameters below, each separated by a comma
        
        [Parameter(Mandatory = $false, Valuefrompipeline = $true, Position = 1)]
        [bool]$AutoFill,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 2)]
        [int[]]$Duration,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 3)]
        [string[]]$TicketNumber,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 4)]
        [string[]]$Reason
            

    )
    
    
        <#Check if PIM Service module is installed#>
        $getModuleSplat = @{
            ListAvailable = $True
            Verbose       = $False
        }

        if ($null -eq (Get-Module @getmodulesplat -name "Microsoft.Azure.ActiveDirectory.PIM.PSModule")) {

            Write-Error "PIM Service module does not exist!" -ErrorAction Stop
    
        }

            if ($AutoFill -eq $true) {
                Connect-PimService
                $roles = Get-PrivilegedRoleAssignment | where-object {$_.isElevated -eq $false | Select-Object rolename, roleID
                Write-Verbose "Connected to PIM"
                Write-Verbose "Adding general reason and general ticket number"
                foreach ($role in $roles) {

                    
                    if (Get-PrivilegedRoleAssignment | Where-Object {$_.roleid -eq $role.roleid -and $_.isElevated -eq $true}){

                        Write-Output $role.name " is already elevated...skipping"
                        Continue
                    }

                    else {
                        Write-Output "Enabling " $role.name
                        Enable-PrivilegedRoleAssignment -roleID $role.roleid -Duration 9 -ticketnumber "SNOW Tickets" -Reason "SNOW Tickets"
                        Write-Output $role.name " is now enabled"
                    }

                }

            }
            else {
                Write-Verbose "Include duration, ticketnumber, and reason for elevation"

                foreach ($role in $roles) {
                    Enable-PrivilegedRoleAssignment -roleID $role -Duration $Duration -ticketnumber $TicketNumber -Reason $Reason
                }
                
            }


        }

    # End of the END Block.
} # End Function


