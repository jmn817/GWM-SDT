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

Function Complete-MigrationUser {

    <#
        .SYNOPSIS
           
        
        .DESCRIPTION
           
        
        .PARAMETER DemoParam1

        
        .EXAMPLE

            
        .NOTES

    #>


    Begin {

        $getModuleSplat = @{
            ListAvailable = $True
            Verbose       = $False
        }
        #Exchange Online Module Installed?
        try {

            $PSExoPowershellModuleRoot = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select-Object -Last 1).DirectoryName

        }

        Catch {

            If ($null -eq $PSExoPowershellModuleRoot) {
                Write-Error "The Exchange Online MFA Module was not found! https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps"

            }

        }

        #Checking if Exchange Online is connected
        Write-Host "Checking if Exchange Online is connected"
        $checkEXOPSSession = Get-PSSession

        if ($checkEXOPSSession.state -eq "Opened" -and $checkEXOPSSession.configurationname -eq "Microsoft.Exchange") {
    
            Write-Host "Exchange Online Powershell already running. Contining Get-SendonBehalfPermission command." -ForegroundColor Green
    
        }

        #If ExOPssesion is running, ignore module import. If not, import module.

        else {

            Write-Host "Not connected to Exchange Online Powershell...please sign in to complete authentication and import module" -ForegroundColor Yellow

            $PSExoPowershellModuleRoot = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select-Object -Last 1).DirectoryName
            $ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellModule.dll";
            $ModulePath = [System.IO.Path]::Combine($PSExoPowershellModuleRoot, $ExoPowershellModule);
            Import-Module -verbose $ModulePath;
            $Office365PSSession = New-ExoPSSession -ConnectionUri "https://outlook.office365.com/powershell-liveid/"
            Import-Module (Import-PSSession $Office365PSSession -AllowClobber) -Global
            Write-Host "Module imported successfully."

        }

    }

    Process {



    }


    End {


    }
}


Function Get-SendOnBehalfPermissions {
    <#
        .SYNOPSIS
            This advanced function obtains the Send on Behalf permissions enabled for a mailbox
        
        .DESCRIPTION
            Get-SendOnBehalfPermissions shows what users have Send On Behalf rights on a specific mailbox.
        
        .PARAMETER DemoParam1
            SharedMailbox parameter

            Get-SendOnBehalfPermissions -SharedMailbox Sharedmailbox@dbschenker.com
        

        
        .EXAMPLE
            
            Get-SendOnBehalfPermissions -SharedMailbox us.sm.amer.GWM-Migrations@dbschenker.com

            Checking if Exchange Online is connected
            Exchange Online Powershell already running. Continuing PST Import...
            user@dbschenker.com has the GrantSendOnBehalf permission set.
            user@dbschenker.com has the GrantSendOnBehalf permission set.
            user@dbschenker.com has the GrantSendOnBehalf permission set.
            
        .NOTES
            Author: Jesse Newell
            Last Edit: 2019-02-05
            Version 1.0 - initial release of blah
            Version 1.1 - update for blah
        
        #>
    [CmdletBinding()]  # Add cmdlet features.
    Param (
        # Define parameters below, each separated by a comma
        
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]
        [String[]]$SharedMailbox
        
        # Add additional parameters here.
    )
        
    Begin {
        $getModuleSplat = @{
            ListAvailable = $True
            Verbose       = $False
        }
        #Exchange Online Module Installed?
        try {

            $PSExoPowershellModuleRoot = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select-Object -Last 1).DirectoryName

        }

        Catch {

            If ($null -eq $PSExoPowershellModuleRoot) {
                Write-Error "The Exchange Online MFA Module was not found! https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps"

            }

        }

        #Checking if Exchange Online is connected
        Write-Host "Checking if Exchange Online is connected"
        $checkEXOPSSession = Get-PSSession

        if ($checkEXOPSSession.state -eq "Opened" -and $checkEXOPSSession.configurationname -eq "Microsoft.Exchange") {
    
            Write-Host "Exchange Online Powershell already running. Contining Get-SendonBehalfPermission command." -ForegroundColor Green
    
        }

        #If ExOPssesion is running, ignore module import. If not, import module.

        else {

            Write-Host "Not connected to Exchange Online Powershell...please sign in to complete authentication and import module" -ForegroundColor Yellow

            $PSExoPowershellModuleRoot = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select-Object -Last 1).DirectoryName
            $ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellModule.dll";
            $ModulePath = [System.IO.Path]::Combine($PSExoPowershellModuleRoot, $ExoPowershellModule);
            Import-Module -verbose $ModulePath;
            $Office365PSSession = New-ExoPSSession -ConnectionUri "https://outlook.office365.com/powershell-liveid/"
            Import-Module (Import-PSSession $Office365PSSession -AllowClobber) -Global
            Write-Host "Module imported successfully."

        }
                
    }# End Begin block
        
    Process {

        foreach ($mailbox in $SharedMailbox) {

            if ($mailbox) {
                $users = Get-Mailbox -Identity $mailbox | Select-Object GrantSendOnBehalfTo -ErrorAction SilentlyContinue | ForEach-Object { $_.GrantSendonBehalfTo }

                foreach ($user in $users) {
        
                    $username = get-mailbox -Identity $user | Select-Object -ExpandProperty primarysmtpaddress
        
                    if ($user) {
                
                        Write-Host "$username " -nonewline; Write-Host "has the GrantSendOnBehalf permission set." -f Green; -nonewline
                        Write-Host " on $mailbox"
                    }
                
                }


            }


        }
   
    } # End of PROCESS block.
        
    End {
        # Start of END block.
        Write-Verbose -Message "Entering the END block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."

        Remove-Variable Mailbox
        Remove-Variable SharedMailbox
        
        # Add additional code here.
        
    } # End of the END Block.
} # End Function


Function Set-SendOnBehalfPermissions {
    <#
        .SYNOPSIS
            This advanced function sets the Send on Behalf Permissions for Shared mailboxes
        
        .DESCRIPTION
            
        
        .PARAMETER DemoParam1
         
        
        .EXAMPLE
            
        
            
        .NOTES
            Author: Jesse Newell
            Last Edit: 2019-04-22
            Version 1.0 - initial release of blah
            Version 1.1 - update for blah
        
        #>
    [CmdletBinding()]  # Add cmdlet features.
    Param (
        # Define parameters below, each separated by a comma
            
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]
        [String[]]$SharedMailbox,
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]
        [String[]]$Mailbox
            
        # Add additional parameters here.
    )

    Begin {

        $getModuleSplat = @{
            ListAvailable = $True
            Verbose       = $False
        }
        #Exchange Online Module Installed?
        try {
    
            $PSExoPowershellModuleRoot = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select-Object -Last 1).DirectoryName
    
        }
    
        Catch {
    
            If ($null -eq $PSExoPowershellModuleRoot) {
                Write-Error "The Exchange Online MFA Module was not found! https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps"
    
            }
    
        }
    
        #Checking if Exchange Online is connected
        Write-Host "Checking if Exchange Online is connected"
        $checkEXOPSSession = Get-PSSession
    
        if ($checkEXOPSSession.state -eq "Opened" -and $checkEXOPSSession.configurationname -eq "Microsoft.Exchange") {
        
            Write-Host "Exchange Online Powershell already running. Contining Get-SendonBehalfPermission command." -ForegroundColor Green
        
        }
    
        #If ExOPssesion is running, ignore module import. If not, import module.
    
        else {
    
            Write-Host "Not connected to Exchange Online Powershell...please sign in to complete authentication and import module" -ForegroundColor Yellow
    
            $PSExoPowershellModuleRoot = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select-Object -Last 1).DirectoryName
            $ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellModule.dll";
            $ModulePath = [System.IO.Path]::Combine($PSExoPowershellModuleRoot, $ExoPowershellModule);
            Import-Module -verbose $ModulePath;
            $Office365PSSession = New-ExoPSSession -ConnectionUri "https://outlook.office365.com/powershell-liveid/"
            Import-Module (Import-PSSession $Office365PSSession -AllowClobber) -Global
            Write-Host "Module imported successfully."
    
        }

    }

    Process {


        #Check if the user has access rights to the shared mailbox already
        Foreach ($user in $mailbox) {

            $userPermissionCheck = Get-mailbox -identity $sharedmailbox | select-object GrantSendonBehalfTo -ErrorAction SilentlyContinue | ForEach-Object { $_.GrantSendOnBehalfTo }

            if ($user -contains $userPermissionCheck) {

                Write-Error "$user " -nonewline; Write-Error "already has Grant Send on Behalf permissions" -f Red; -nonewline
                Write-Error " on $sharedMailbox"
                Continue

 
            }
            if ($user -notcontains $userPermissionCheck) {

                #Assigning Permissions
                Write-Host "Assigning Grant Send on Behalf Permission to $user" -NoNewline
                Write-Host "on $sharedmailbox."

                #Adding users to collection array
                $user += $usersToEnable 


            }

        }

        Get-Mailbox -Identity $mailbox | set-mailbox -GrantSendOnBehalfTo $usersToEnable
        Write-Host "Permissions assigned to the following users:"
            
        foreach ($successUser in $usersToEnable) {

            $username = get-mailbox -Identity $successuser | Select-Object -ExpandProperty primarysmtpaddress
            Write-Host "$username " -nonewline; Write-Host "has the GrantSendOnBehalf permission set." -f Green; -nonewline
            Write-Host " on $sharedMailbox"

        }

    }

    End {

        Remove-Variable Mailbox
        Remove-Variable SharedMailbox
        Remove-Variable Username
        Remove-Variable 

    }


}

Function Enable-PIMElevation {
    <#
        .SYNOPSIS
            This advanced function elevates the users administration roles using the PIM Service. 
        
        .DESCRIPTION
            Enable the PIM elevation for administrtion roles. One function to elevate all user roles.
        
        .PARAMETER DemoParam1
            The parameter DemoParam1 is used to define the value of blah and also blah.
        
        .PARAMETER DemoParam2
            The parameter DemoParam2 is used to define the value of blah and also blah.
        
        .EXAMPLE
            
            To enable PIM for all administration roles with already prefilled ticket number, reason, and time use the AutoFill Parameter:

            Enable-PIMElevation -AutoFill:$true

            To enable PIM for all administration roles with specific ticket number, reason, and time:

            Enable-PIMElevation -Duration 9 -TicketNumber INC0093333 -Reason 'Device Administration for intall'
            
            
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
        [string[]]$Reason,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 5)]
        [string[]]$SpecificRole
            

    )

    Begin {
        <#Check if PIM Service module is installed#>
        $getModuleSplat = @{
            ListAvailable = $True
            Verbose       = $False
        }

        if ($null -eq (Get-Module @getmodulesplat -name "Microsoft.Azure.ActiveDirectory.PIM.PSModule")) {

            Write-Error "PIM Service module does not exist!" -ErrorAction Stop
    
        }
        Write-Output "Connecting to PIM Service"
        Connect-PimService
        Write-Output "PIM Service connected... obtaining administration roles."
        $roles = Get-PrivilegedRoleAssignment | Where-Object { $_.isElevated -eq $false } | Select-Object rolename, roleID
        $EmptyFields = @()

    }#End of the BEGIN Block
    
    
        
    Process {
            
        if ($AutoFill -eq $true) {
            Write-Verbose "Adding general reason and general ticket number"
            foreach ($role in $roles) {

                    
                if (Get-PrivilegedRoleAssignment | Where-Object { $_.roleid -eq $role.roleid -and $_.isElevated -eq $true }) {

                    Write-Output $role.name " is already elevated...skipping"
                    Continue
                }

                else {
                    Write-Output "Enabling " $role.name
                    Enable-PrivilegedRoleAssignment -roleID $role.roleid -Duration 9 -ticketnumber "SNOW Tickets" -Reason "SNOW Tickets"
                    Write-Output $role.name " is now enabled"
                }

            }

            Write-Output "All administration roles are now elevated"
            Continue 
        }
            
        else {
            Write-Verbose "Include duration, ticketnumber, and reason for elevation"

            foreach ($role in $roles) {
                Enable-PrivilegedRoleAssignment -roleID $role -Duration $Duration -ticketnumber $TicketNumber -Reason $Reason
            }
                
        }


        if ($null -eq $SpecificRole) {

            Write-Host "Please select an administration role you are currently assigned to"
            Write-Host "Currently you have the following administration roles:"
            foreach ($role in $roles) {

                Write-Output $role.name

            }
            Continue
        }

        if ($SpecificRole) {

            if ($null -eq $Duration -or $TicketNumber -or $Reason) {
                        
                if ($null -eq $Duration) {

                    $EmptyFields += "Duration"

                }

                if ($null -eq $TicketNumber) {

                    $EmptyFields += "TicketNumber"
                }

                if ($null -eq $Reason) {

                    $EmptyFields += "Reason"
                }

                foreach ($emptyfield in $emptyfields) {

                    Write-Output $emptyfield " is empty! Please fill in the parameter"

                }
                Continue

            }


           

                    


        }
        
    }#End of the PROCESS Block


    End {




    }# End of the END Block.
} # End Function




Function Get-MissingLocationPackage {
    <#
        .SYNOPSIS
        
           Get users that do not have a location package.

        .DESCRIPTION
           If a user is staged and has no location package assigned then this script will export them to a CSV file for further action.

           CSV will be located under the directory you ran the script from.If you run the script in memory without saving it first it will save the CSV in your OneDrive documents folder.

           Called Report_US_No_LocationPackage.csv


        .PARAMETER UsageLocation

        UsageLocation parameter is to specify which country to query users from. Use the country's short code to specify which country to run the script on.

        
        .EXAMPLE

        To search for users in the US, type in US. For Peru PE, Germany use DE, etc.

        Get-UsersLocationPackage -UsageLocation US

        This will return users that are staged and do not have a location package.
            
            
        .NOTES
            Author: Jesse Newell
            Last Edit: 2019-05-08
            Version 1.0
        
        #>


    [CmdletBinding()]  # Add cmdlet features.
    Param (
        # Define parameters below, each separated by a comma
            
        [Parameter(Mandatory = $true, Valuefrompipeline = $true, Position = 1)]
        [string]$UsageLocation
        
    )

    Begin {

        <#Verify connection to Azure AD#>
        try { 
            $var = Get-AzureADTenantDetail
        } 

        catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException] { 
            Write-Host "You're not connected to Azure AD." -ForegroundColor Yellow; Connect-AzureAD
        }

           
        $details = @()

        <#Setting file name#>
        $filename = "Report_$usagelocation`_Users_NoLocationPackage.csv"

        try {
        <#Getting directory information#>
        $filebase = Join-Path $PSScriptRoot $filename -ErrorAction SilentlyContinue
        }
        catch {
         
            <#If no directory is set then it will default to your OneDrive documents folder#>
            Write-Host "No path set, defaulting to the OneDrive documents folder"
            $userDocuments = -Global "$env:USERPROFILE\OneDrive - Schenker AG\Documents"
            $filebase = Join-Path $userDocuments $filename

        }

        
    }

    Process {

        <#Searching for users that are assigned to the usage location. Filtering user accounts only#>
        $users = Get-AzureADUser -Filter "usagelocation eq '$usagelocation'" -all $true | Where-Object { $_.displayname -notlike "*Shared*" } | 
        Where-Object { $_.displayname -notlike "*Technical*" } | Where-Object { $_.displayname -notlike "*(ADM)*" } | Where-Object { $_.displayname -notlike "*(SVC)*" } | 
        Where-Object { $_.displayName -notlike "*Room*" } | Where-Object { $_.DisplayName -notlike "*DEM*" } | Where-Object { $_.DisplayName -notlike "*.rm.*" } | 
        Select-Object UserPrincipalName, objectID

        Foreach ($user in $users) {

            <#Obtaining CFG Core Object ID to determine if the user is staged#>
            $cfgCore = (Get-AzureADUserMembership -ObjectId $user.objectID | Where-Object { $_.objectID -eq "78bbc179-984a-4c74-a4d0-ae68f852adf2" })

            if ($cfgCore) {
                <#Verifying user does not already have a location package#>
                $location = (Get-AzureADUserMembership -ObjectId $user.objectID | Where-Object { $_.DisplayName -like "*CFG - Core - $usagelocation*" }).displayname
                
                <#If so, skipping. User is good#>
                if ($location) {
                    
                    Write-Host $user.userprincipalname ' has ' $location -ForegroundColor Green


                }

                else {
                
                    Write-Host $user.userprincipalname ' does not have a location package assigned! Inputting user to CSV for tracking.' -BackgroundColor Red
                    
                    <#Inputting username to custom object#>
                    $details += New-Object PSObject -Property @{
                            
                        Username = $user.userprincipalname
                        
                    }
   
                }

            }

            else {
                <#If the user is not staged then report skips them. Determined if they are staged by the CFG Core group#>
                Write-Host $user.userprincipalname " is not staged, skipping..."
                Continue
            }
                
        }#end foreach
            
                
            
        <#Exporting results to CSV file, with a header of Username#>
        $details | export-csv -Path $filebase

        Write-Host "."
        Write-Host "."
        Write-Host "."
        Write-Host "Done! CSV file saved to $filebase" -ForegroundColor Yellow

            

    }#end Process


    End {
        <#Removing variables#>
        Remove-Variable UsageLocation
        Remove-Variable details
        Remove-Variable location


    }
}


Function Set-LocationPackage {


    <#
        .SYNOPSIS
        
           Set the user's location package. 

        .DESCRIPTION
          

        .PARAMETER UsageLocation


        
        .EXAMPLE

            
            
        .NOTES
            Author: Jesse Newell
            Last Edit: 2019-05-03
            Version 1.0
        
        #>

    [CmdletBinding()]  # Add cmdlet features.
    Param (
        # Define parameters below, each separated by a comma
            
        [Parameter(Mandatory = $true, Valuefrompipeline = $true, Position = 1)]
        [string[]]$UserPrincipalName,
        [parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 2)]
        [string[]]$LocationCode,
        [switch]$USUser,
        [Parameter(Mandatory = $false)]
        [array]$DomainList = @(
            "am.bax.global",
            "schenkerusa.com",
            "schenker.ca"
        )
        
    )

    Begin {
        try { 
            $var = Get-AzureADTenantDetail
        } 
        catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException] { 
            Write-Host "You're not connected to Azure AD." -ForegroundColor Yellow; Connect-AzureAD
        }
    }

    Process {

        foreach ($user in $UserPrincipalName) {
            $usersInfo = Get-AzureADUser -Filter "userprincipalname eq '$user'" | Select-Object objectID, usagelocation
            Write-Host "Verifying user does not have a location package assigned"

            $cfgCore = (Get-AzureADUserMembership -ObjectId $usersinfo.objectID | Where-Object { $_.objectID -eq "78bbc179-984a-4c74-a4d0-ae68f852adf2" })
            if ($cfgcore) {   
                $location = (Get-AzureADUserMembership -ObjectId $usersinfo.objectID | Where-Object { $_.DisplayName -like "*CFG - Core - $LocationCode*" }).displayname
                                
                if ($location) {
                    Write-Host "$user has a location package already" -foreground Red
                    Break
                }

                else {
                    Write-Host "$user does not have a location package, continuing..." -ForegroundColor Yellow
                    
                }

            }

            else {

                Write-Host "$user is not staged. Skipping..."
                Break
            }

            foreach ($domain in $DomainList) {
                $ADUser = Get-ADUser -Server $Domain -filter { emailaddress -eq $user } -ErrorAction SilentlyContinue | Select-Object SamAccountName
                if ($ADUser) {
                    Write-Host "$User found on $domain"
                    if ($domain -eq "am.bax.global") {
                        $domainPrefix = "AM"
                        Write-Host "$user is on the $domainPrefix, assigning location package based on the domain."
                        Break
                    }
                    if ($domain -eq "schenkerusa.com") {                        
                        $domainPrefix = "SUSA"
                        Write-Host "$user is on the $domainPrefix domain, assigning location package based on the domain."
                        Break
                    }
                    if ($domain -eq "schenker.ca") {                        
                        $domainPrefix = "CA"
                        Write-Host "$user is on the $domainPrefix, assigning location package based on the domain."
                        Break
                    }      
               
                }
                 
            }

        
              
            $groups = Get-AzureADGroup -all:$true -searchstring "CFG - Core - $LocationCode" | select-object displayname, objectid
               
            if ($groups) {
                            
                $specificGroup = $groups | Where-Object { $_.displayname -match $domainPrefix }

                Write-Host "Location package found called " $specificGroup.displayname
                Write-Host "With objectid of "$specificGroup.objectid

                Add-AzureADGroupMember -ObjectId $specificGroup.objectid -RefObjectId $usersInfo.objectid -ErrorAction SilentlyContinue
                Write-Host "Success! $user assigned to " $specificgroup.displayname -ForegroundColor Green
  
            }

            else {
                
                Write-Error "Location package containing $LocationCode cannot be found!"
            
            }

        }#End Foreach

    }#End Process

    End {
        
        
    }

}