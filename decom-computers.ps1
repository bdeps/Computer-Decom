function Move-ADComputerObject{
<#
    .SYNOPSIS
    Moves the specified computer to a new OU

    .DESCRIPTION
    Moves a single computer or a list of computers. Takes GUIDS

    .Parameter $Identity
    Takes computerGUIDS. Either a list or a single guid
    
    .Parameter $targetOUPath
    Path to where the computer objects should be moved.

    .Parameter $computerDomain
    Takes the domain the computers are in. All computers in the list needs to be from the same domain

    .EXAMPLE
     Move-ADComputerObject -Identity "ExampleGUID1,ExampleGUID2" -targetOuPath "OU=Example,DC=Contoso,DC=Com" -computeDomain "contoso.com"
    #>

    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$targetOUPath,

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$computerDomain,

        [parameter(Mandatory = $false, ParameterSetName = "Hidden")]
        [hashtable]$Config,

        [parameter(Mandatory = $false, ParameterSetName = "Hidden")]
        [hashtable]$Parameters
    )

    $HashParameters = @{
        Identity = $Identity
        Server = $computerDomain
        TargetPath = $targetOUPath
    }

    $Computers = $Identity.Split(",")
    Foreach($Computer in $Computers)
    {
        try{
            Write-Host "Moving computerObject $Identity to OU: $targetOUPath in domain $computerDomain" -ForegroundColor Green
            Get-ADComputer -Identity $Identity -Server $computerDomain | Move-ADObject -TargetPath $TargetOUPath
        }
        catch{
             Write-Host "Moving computerObject $Identity to OU: $targetOUPath in domain $computerDomain failed" -ForegroundColor Red
             Write-EventLog -LogName 'ComputerDecom' -Source 'Move-ADComputerObject' -Message "Unable to move $Identity to target path $TargetOUPath - Parameters: $HashParameters" -EventId '2200'
        }
    }
}

function Disable-Computers {
<#
    .SYNOPSIS
    Disables the provided computers

    .DESCRIPTION
    Disables a single computer or a list of computers. Takes GUIDS

    .Parameter $Identity
    Takes computerGUIDS. Either a list or a single guid

    .Parameter $computerDomain
    Takes the domain the computers are in. All computers in the list needs to be from the same domain

    .EXAMPLE
     Disable-Computers -Identity "exampleGUID1,exampleGUID2" -$computerDomain contoso.com
    #>

    [cmdletbinding()]
    param(
        [parameter(
            Mandatory = $true,
            HelpMessage = "computerGUID")]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [parameter(
            Mandatory = $true,
            HelpMessage="Domain which holds the user account")]
        [ValidateNotNullOrEmpty()]
        [string]$computerDomain,

        [parameter(Mandatory = $false, ParameterSetName = "Hidden")]
        [hashtable]$Config,

        [parameter(Mandatory = $false, ParameterSetName = "Hidden")]
        [hashtable]$Parameters
    )

    $HashParameters = @{
        Identity = $Identity
        Server = $computerDomain
        TargetPath = $targetOUPath
    }

    $Computers = $Identity.Split(",")
   

    Foreach($Computer in $Computers)
    {
        try{
            Write-Host "Disabling computerObject $Identity in domain $computerDomain" -ForegroundColor Green
            Disable-ADAccount -Identity (Get-ADComputer $Computer -Server $computerDomain) -Server $computerDomain
        }
        catch{
             Write-Host "Disabling computerObject $Identity in domain $computerDomain" -ForegroundColor Red
             Write-EventLog -LogName 'ComputerDecom' -Source 'Disable-Computers' -Message "Unable to disable $Identity - Parameters: $HashParameters" -EventId '2200'
        }
    }
}

function Update-ComputerDescription {
<#
    .SYNOPSIS
    Updates the description of a computer object

    .DESCRIPTION
    Updates the Description field of a single computer or a list of computers. Takes GUIDS

    .Parameter $Identity
    Takes computerGUIDS. Either a list or a single guid

    .Parameter $computerDomain
    Takes the domain the computers are in. All computers in the list needs to be from the same domain

    .Parameter $newDescription
    Takes the a string that contains an addition or replacement to the current description

    .Parameter $appendToCurrentDescription
    TRUE if you want to append $newDescription to the end of the existing one. False if you want to completely replace the current description

    .Parameter $orderID
    OrderID from the ZervicePoint order. Automatically appended at the end of the description

    .Parameter $orderDate
    OrderDate from the ZervicePoint order. Automatically appended at the end of the description

    .EXAMPLE
     Update-ComputerDescription -Identity "ExampleGUID1,ExampleGUID2" -computerDomain "contoso.com" -newDescription "Test Description" -appendToCurrentDescription $true -orderID "123" -$orderDate ""
    #>
    [cmdletbinding()]
    param(
        [parameter(
            Mandatory = $true,
            HelpMessage = "computerGUID")]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [parameter(
            Mandatory = $true,
            HelpMessage="Domain which holds the user account")]
        [ValidateNotNullOrEmpty()]
        [string]$computerDomain,

        [parameter(
            Mandatory = $true,
            HelpMessage="New Description")]
        [ValidateNotNullOrEmpty()]
        [string]$newDescription,

        [parameter(
            Mandatory = $true,
            HelpMessage="True if you want to keep the current Description and append an addition")]
        [ValidateNotNullOrEmpty()]
        [boolean]$appendToCurrentDescription,

        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$orderId,

        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [dateTime]$orderDate,

        [parameter(Mandatory = $false, ParameterSetName = "Hidden")]
        [hashtable]$Config,

        [parameter(Mandatory = $false, ParameterSetName = "Hidden")]
        [hashtable]$Parameters
    )

    $HashParameters = @{
        Identity = $Identity
        Server = $computerDomain
        TargetPath = $targetOUPath
        Description = $newDescription
        appendToCurrent = $appendToCurrentDescription
        orderID = $orderId
        orderDate = $orderDate
    }

    $Computers = $Identity.Split(",")
    if($orderDate){
        $strOrderDate = $orderDate.ToString("yyyy-MM-dd")
        $newDescription = "$newDescription $strOrderDate"
    }
    if($orderID){
        $newDescription = "$newDescription $orderID"
    }

    if($appendToCurrentDescription -eq $true){
        Foreach($Computer in $Computers){
            try{
                ## Append $newDescription to the Description Field
                Write-Host -InformationAction "Appending $newDescription to $Identity in domain $computerDomain" -ForegroundColor Green
                $conputerDN = Get-ADComputer $Computer -Server $computerDomain -Properties Description
                Set-ADComputer -Identity (Get-ADComputer $computer -Server $computerDomain) -Description "$($computerDN.Description) $newDescription"
            }
            catch{
                Write-Host "Appending $newDescription to $Identity in domain $computerDomain failed" -ForegroundColor Red
                Write-EventLog -LogName 'ComputerDecom' -Source 'Update-ComputerDescription' -Message "Unable to append $Identity description to $newDescription - Parameters: $HashParameters" -EventId '2200'
            }
        }
    }
    else{
        Foreach($Computer in $Computers){
            try{
                ## Change the Description Field to $newDescription
                Write-Host "Updating $newDescription to $Identity in domain $computerDomain" -ForegroundColor Green
                Set-ADComputer -Identity (Get-ADComputer $Computer -Server $computerDomain) -Server $computerDomain -Description $newDescription
            }
            catch{
                Write-Host "Updating $newDescription to $Identity in domain $computerDomain failed" -ForegroundColor Red
                Write-EventLog -LogName 'ComputerDecom' -Source 'Update-ComputerDescription' -Message "Unable to update $Identity description to $newDescription - Parameters: $HashParameters" -EventId '2200'
            }
        }
    }

}

function Decom-Computer{
<#
    .SYNOPSIS
    decoms a computer object

    .DESCRIPTION
    Disables, moves and updates the description of a computer object

    .Parameter $Identity
    Takes computerGUIDS. Either a list or a single guid

    .Parameter $computerDomain
    Takes the domain the computers are in. All computers in the list needs to be from the same domain

    .Parameter $newDescription
    Takes the a string that contains an addition or replacement to the current description

    .Parameter $targetOUPath
    Takes an OU path where the computer objects should be moved to "OU=DisabledComputers, DC=Contoso, DC=com"

    .Parameter $appendToCurrentDescription
    TRUE if you want to append $newDescription to the end of the existing one. False if you want to completely replace the current description

    .Parameter $orderID
    OrderID if running from an automatic service order. Automatically appended at the end of the description

    .Parameter $orderDate
    OrderDate if running from an automatic service order order. Automatically appended at the end of the description

    .EXAMPLE
    Decom-Computers -Identity "ExampleGUID1,ExampleGUID2" -computerDomain "contoso.com" -newDescription "Test Description" -appendToCurrentDescription $true -orderID "123" -$orderDate ""
    #>
    [cmdletbinding()]
    param(
        [parameter(
            Mandatory = $true,
            HelpMessage = "computerGUID")]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [parameter(
            Mandatory = $true,
            HelpMessage="Domain which holds the user account")]
        [ValidateNotNullOrEmpty()]
        [string]$computerDomain,

        [parameter(
            Mandatory = $true,
            HelpMessage="New Description")]
        [ValidateNotNullOrEmpty()]
        [string]$newDescription,

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$targetOUPath,

        [parameter(
            Mandatory = $true,
            HelpMessage="True if you want to keep the current Description and append an addition")]
        [ValidateNotNullOrEmpty()]
        [boolean]$appendToCurrentDescription,

        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$orderId,

        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [dateTime]$orderDate,

        [parameter(Mandatory = $false, ParameterSetName = "Hidden")]
        [hashtable]$Config,

        [parameter(Mandatory = $false, ParameterSetName = "Hidden")]
        [hashtable]$Parameters
    )

    ## Creates a hastable of parameters. Only adds non mandatory ones if provided.
    $HashParameters = @{
        Identity = $Identity
        Server = $computerDomain
        Description = $newDescription
        appendToCurrent = $appendToCurrentDescription
    }
    if($orderId){$HashParameters.add('orderID',$orderId)}
    if($orderDate){$HashParameters.Add('orderDate',$orderDate)}

    ## Disable computer objects
    Disable-Computers -Identity $Identity -computerDomain $computerDomain

    ## Wait for changes to replicate out to DC's
    Start-Sleep -Seconds 10

    ## Update computer description
    Update-ComputerDescription $HashParameters

    ## Wait for changes to replicate out to DC's
    Start-Sleep -Seconds 10

    ## Moves the computer objects
    Move-ADComputerObject -Identity $Identity -computerDomain $computerDomain -targetOUPath $targetOUPath
}