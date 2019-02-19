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
        Last Edit: 2019-02-06
        Version 1.0 - initial release of blah
        Version 1.1 - update for blah
    
    #>
        [CmdletBinding()]  # Add cmdlet features.
        Param (
            # Define parameters below, each separated by a comma
    
            [Parameter(Mandatory=$False, Position = 1, ValueFromPipeline = $False)]
            [Boolean]$MFA,
            [Parameter(Mandatory=$False, Position = 2, ValueFromPipeline = $False)]
            [Switch]$PIM,
            [Parameter(Mandatory=$False, Position = 3, ValueFromPipeline = $False)]
            [Switch]$AzureAD,
            [Parameter(Mandatory=$False, Position = 4, ValueFromPipeline = $False)]
            [Switch]$ExO,
            [Parameter(Mandatory=$False, Position = 5, ValueFromPipeline = $False)]
            [Switch]$MSOL,
            [Parameter(Mandatory=$False, Position = 6, ValueFromPipeline = $False)]
            [Switch]$Sharepoint,
            [Parameter(Mandatory=$False, Position = 7, ValueFromPipeline = $False)]
            [Switch]$All
            # Add additional parameters here.
        )
    
       End Begin block
    
        Process {
            # Start of PROCESS block.
            Write-Verbose -Message "Entering the PROCESS block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
    
    
            # Add additional code here.
    
            
        } # End of PROCESS block.
    
       

        
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