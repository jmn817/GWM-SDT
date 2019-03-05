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
        The example below does blah
        PS C:\> <Example>
        
    .NOTES
        Author: Jesse Newell
        Last Edit: 2019-02-19
        Version 1.0 - initial release of blah
        Version 1.1 - update for blah
    
    #>
        [CmdletBinding()]
        [CmdletBinding(DefaultParameterSetName)]  # Add cmdlet features.
        Param (
            # Define parameters below, each separated by a comma

            [Parameter(Mandatory = $True, Position = 1)]
            [ValidateSet('PIM', 'AzureAD', 'ExO', 'MSOnline', 'SharePoint')]
            [string[]]$Service,    
            [Parameter(Mandatory =$false, Position = 2, ParameterSetName = 'Credential')]
            [String]$Credential
            # Add additional parameters here.
        )
    
        $getModuleSplat = @{
            ListAvailable = $True
            Verbose	      = $False
        }
    
            
            ForEach ($item in $PSBoundParameters.Service){

                Write-Verbose "Attempting to connect to $Item"  
                Switch ($item){

                    AzureAD {
                        if ($null -eq(Get-Module @getmodulesplat -name "AzureAD")){

                            Write-Error "Azure AD module is not present!"
                            Confirm-AzureStorSimpleLegacyVolumeContainerStatus

                        }
                        else {
                            $Connect = Connect-AzureAD
                            if($Connect -ne $null){
                                If(($host.ui.RawUI.windowtitle) -notlike "*Connected To:*"){
                                    $host.ui.RawUI.WindowTitle += " - Connected To: AzureAD"

                                }
                                Else{
                                    $host.UI.RawUI.WindowTitle += " - AzureAD"

                                }


                            }
                            
                        }
                        
                        Continue
                    }
                
                    ExO {

                    
                
                        $PSExoPowershellModuleRoot = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select-Object -Last 1).DirectoryName
					
					    If ($null -eq $PSExoPowershellModuleRoot)
					    {
						    Write-Error "The Exchange Online MFA Module was not found! https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps"
						    continue
					        }
                        Else
					    {
                          
                                
                                Write-Verbose "Importing Exchange MFA Module"
                                $ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellModule.dll";
                                $ModulePath = [System.IO.Path]::Combine($PSExoPowershellModuleRoot, $ExoPowershellModule);

                                Import-Module -verbose $ModulePath;

                                $Office365PSSession = New-ExoPSSession -userprincipalname $office365UserPrincipalName -ConnectionUri "https://outlook.office365.com/powershell-liveid/" 
                            
                                Write-Verbose "Connecting to Exchange Online"
                                Import-PSSession $Office365PSSession
                                Connect-ExoPssession
                            
                            
					
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
        
                [Parameter(Mandatory=$True)]
                [int]$DemoParam1,
        
                [Parameter(Mandatory=$False)]
                [ValidateSet('Alpha','Beta','Gamma')]
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



        Function Set-MigrationBatchJob {
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
            
                    [Parameter(Mandatory=$True)]
                    [int]$DemoParam1,
            
                    [Parameter(Mandatory=$False)]
                    [ValidateSet('Alpha','Beta','Gamma')]
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